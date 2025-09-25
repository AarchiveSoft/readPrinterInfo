import os
import time
import ctypes
import smtplib
import win32print
from email.mime.text import MIMEText
import schedule

# -----------------------
# CONFIG
# -----------------------
PRINTER_NAME = "DP-DS620"   # adjust to your installed printer name
DLL_PATH = os.path.abspath(os.path.join("data", "CspStat.dll"))

# Email config (uses SMTP relay; add login if your relay requires it)
SMTP_SERVER = "smtp-relay.gmail.com"
SMTP_PORT = 587  # try 25 if 587 doesn’t work
EMAIL_FROM = "noreply@graphicart.ch"
EMAIL_TO = "a.hafner@graphicart.ch"

LOW_THRESHOLD_PERCENT = 10.0

# Grace window to avoid DLL calls right after Windows reports idle
# (customer observed ~10 s of printing after queue is empty)
GRACE_AFTER_IDLE_SECONDS = 15

# -----------------------
# Spooler-side checks
# -----------------------
JOB_BUSY_FLAGS = (
    getattr(win32print, "JOB_STATUS_PRINTING", 0x10)
    | getattr(win32print, "JOB_STATUS_SPOOLING", 0x08)
    | getattr(win32print, "JOB_STATUS_PAUSED", 0x01)
    | getattr(win32print, "JOB_STATUS_BLOCKED_DEVQ", 0x200)
    | getattr(win32print, "JOB_STATUS_RESTART", 0x800)
)

PRINTER_BUSY_FLAGS = (
    getattr(win32print, "PRINTER_STATUS_PRINTING", 0x400)
    | getattr(win32print, "PRINTER_STATUS_PROCESSING", 0x4000)
    | getattr(win32print, "PRINTER_STATUS_BUSY", 0x200)
    | getattr(win32print, "PRINTER_STATUS_WARMING_UP", 0x10000)
    | getattr(win32print, "PRINTER_STATUS_WAITING", 0x2000)
    | getattr(win32print, "PRINTER_STATUS_INITIALIZING", 0x8000)
)

_last_activity_ts = 0.0  # updated whenever we detect spooler activity


def _spooler_activity_snapshot(printer_name: str) -> dict:
    """
    Return a snapshot: {'busy': bool, 'cjobs': int, 'status_flags': int}.
    'busy' covers printer flags or job flags indicating active work.
    """
    h = win32print.OpenPrinter(printer_name)
    try:
        info = win32print.GetPrinter(h, 2)  # PRINTER_INFO_2
        status_flags = info.get("Status", 0) or 0
        cjobs = info.get("cJobs", 0) or 0

        # Printer-side busy flags
        if status_flags & PRINTER_BUSY_FLAGS:
            return {"busy": True, "cjobs": cjobs, "status_flags": status_flags}

        # Job-side busy flags
        if cjobs > 0:
            jobs = win32print.EnumJobs(h, 0, cjobs, 1)  # JOB_INFO_1
            for j in jobs:
                if (j.get("Status", 0) or 0) & JOB_BUSY_FLAGS:
                    return {"busy": True, "cjobs": cjobs, "status_flags": status_flags}

        return {"busy": False, "cjobs": cjobs, "status_flags": status_flags}
    finally:
        win32print.ClosePrinter(h)


def is_effectively_idle(printer_name: str) -> bool:
    """
    True only if the spooler is idle AND the grace window since the last activity
    has elapsed. This avoids calling the DLL while the engine is still finishing a print.
    """
    global _last_activity_ts
    snap = _spooler_activity_snapshot(printer_name)
    now = time.time()

    if snap["busy"] or snap["cjobs"] > 0:
        _last_activity_ts = now
        return False

    if (now - _last_activity_ts) < GRACE_AFTER_IDLE_SECONDS:
        return False

    return True

# -----------------------
# DLL binding
# -----------------------
def bind_dnp_functions(dll_path: str):
    dll = ctypes.WinDLL(dll_path)
    def bind(name, argtypes=(), restype=ctypes.c_int):
        try:
            fn = getattr(dll, name)
            fn.argtypes = list(argtypes)
            fn.restype = restype
            return fn
        except AttributeError:
            return None

    return {
        "GetPrinterPortNum": bind("GetPrinterPortNum"),
        "GetMediaCounter": bind("GetMediaCounter", [ctypes.c_int]),
        "GetInitialMediaCount": bind("GetInitialMediaCount", [ctypes.c_int]),
    }

def find_vendor_port(api, probe_max=16):
    if api["GetPrinterPortNum"] is not None:
        try:
            p = api["GetPrinterPortNum"]()
            if isinstance(p, int) and p >= 0:
                return p
        except OSError:
            pass
    for p in range(probe_max):
        try:
            r = api["GetMediaCounter"](p)
            if isinstance(r, int) and r >= 0:
                return p
        except OSError:
            continue
    return None

# -----------------------
# Email sending
# -----------------------
def send_email(subject: str, body: str):
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.ehlo()
        try:
            server.starttls()  # only needed if relay requires TLS
            server.ehlo()
        except smtplib.SMTPNotSupportedError:
            # server doesn’t support TLS, that’s fine
            pass

        # If your relay requires login, uncomment:
        # server.login("USERNAME", "PASSWORD")

        server.sendmail(EMAIL_FROM, [EMAIL_TO], msg.as_string())

# -----------------------
# Main check logic
# -----------------------
def check_remaining():
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] Checking printer status...")

    if not is_effectively_idle(PRINTER_NAME):
        print(f"Printer not past grace window; will try later (grace={GRACE_AFTER_IDLE_SECONDS}s).")
        return

    if not os.path.exists(DLL_PATH):
        print(f"CspStat.dll not found at {DLL_PATH}")
        return

    api = bind_dnp_functions(DLL_PATH)
    if api["GetMediaCounter"] is None:
        print("DLL missing required export: GetMediaCounter.")
        return

    port = find_vendor_port(api)
    if port is None:
        print("No DS620 found.")
        return

    try:
        remaining = api["GetMediaCounter"](port)
    except OSError as e:
        print(f"Error calling GetMediaCounter: {e}")
        return

    total = -1
    if api["GetInitialMediaCount"] is not None:
        try:
            total = api["GetInitialMediaCount"](port)
        except OSError as e:
            print(f"Error calling GetInitialMediaCount: {e}")

    print(f"Remaining prints: {remaining}/{total if total >= 0 else '?'}")

    if total and total > 0:
        pct = (remaining / total) * 100.0
        print(f"Remaining percent: {pct:.1f}%")
        if pct < LOW_THRESHOLD_PERCENT:
            subject = f"DNP DS620: Papierstand niedrig ({pct:.1f}%)"
            body = (
                f"Drucker {PRINTER_NAME}: Papier fast leer – noch {remaining} von {total} Ausdrucken übrig. "
                f"Bitte bald Papier nachlegen."
            )
            try:
                send_email(subject, body)
                print("Low-paper email sent.")
            except Exception as e:
                print(f"Error sending email: {e}")

# -----------------------
# Scheduler loop
# -----------------------
if __name__ == "__main__":
    schedule.every(5).minutes.do(check_remaining)

    print("Started monitoring loop. Ctrl+C to stop.")
    # Run an immediate check at startup (optional); comment out if not desired
    check_remaining()

    while True:
        schedule.run_pending()
        time.sleep(1)
