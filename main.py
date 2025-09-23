import os
import time
import ctypes
import win32print
import smtplib
from email.mime.text import MIMEText
import schedule

# -----------------------
# CONFIG
# -----------------------
PRINTER_NAME = "DP-DS620"   # adjust to your installed printer name
DLL_PATH = os.path.abspath(os.path.join("data", "CspStat.dll"))

SMTP_SERVER = "smtp-relay.gmail.com"
SMTP_PORT = 587  # try 25 if 587 doesn’t work
EMAIL_FROM = "noreply@graphicart.ch"
EMAIL_TO = "a.hafner@graphicart.ch"

LOW_THRESHOLD_PERCENT = 10.0

# -----------------------
# Spooler-side checks
# -----------------------
JOB_BUSY_FLAGS = (
    getattr(win32print, "JOB_STATUS_PRINTING", 0x10)
    | getattr(win32print, "JOB_STATUS_SPOOLING", 0x08)
)

PRINTER_BUSY_FLAGS = (
    getattr(win32print, "PRINTER_STATUS_PRINTING", 0x400)
    | getattr(win32print, "PRINTER_STATUS_PROCESSING", 0x4000)
    | getattr(win32print, "PRINTER_STATUS_BUSY", 0x200)
)

def is_printer_busy(printer_name: str) -> bool:
    h = win32print.OpenPrinter(printer_name)
    try:
        info = win32print.GetPrinter(h, 2)
        status_flags = info.get("Status", 0) or 0
        if status_flags & PRINTER_BUSY_FLAGS:
            return True

        total_jobs = info.get("cJobs", 0) or 0
        if total_jobs > 0:
            jobs = win32print.EnumJobs(h, 0, total_jobs, 1)
            for j in jobs:
                jstatus = j.get("Status", 0) or 0
                if jstatus & JOB_BUSY_FLAGS:
                    return True
        return False
    finally:
        win32print.ClosePrinter(h)

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

        # If your relay doesn’t require login, skip this
        # server.login("USERNAME", "PASSWORD")

        server.sendmail(EMAIL_FROM, [EMAIL_TO], msg.as_string())

# -----------------------
# Main check logic
# -----------------------
def check_remaining():
    print("Checking printer status...")

    if is_printer_busy(PRINTER_NAME):
        print("Printer busy → skip this round.")
        return

    api = bind_dnp_functions(DLL_PATH)
    if api["GetMediaCounter"] is None:
        print("DLL missing required export.")
        return

    port = find_vendor_port(api)
    if port is None:
        print("No DS620 found.")
        return

    remaining = api["GetMediaCounter"](port)
    total = api["GetInitialMediaCount"](port) if api["GetInitialMediaCount"] else -1

    print(f"Remaining prints: {remaining}/{total}")

    if total > 0:
        pct = (remaining / total) * 100.0
        print(f"Remaining percent: {pct:.1f}%")
        if pct < LOW_THRESHOLD_PERCENT:
            subject = f"DNP DS620 Papierstand niedrig ({pct:.1f}%)"
            body = f"Drucker {PRINTER_NAME}: Papier fast leer – noch {remaining} von {total} Drucken übrig. Bitte bald Papier nachlegen."
            send_email(subject, body)
            print("Low-paper email sent.")

# -----------------------
# Scheduler loop
# -----------------------
if __name__ == "__main__":
    schedule.every(5).minutes.do(check_remaining)

    print("Started monitoring loop. Ctrl+C to stop.")
    while True:
        schedule.run_pending()
        time.sleep(1)
