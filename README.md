# DNP DS620 – Restdrucke per Python auslesen

Dieses Repository enthält ein Beispielskript, mit dem sich die **verbleibenden Drucke** eines  
DNP DS620 Fotodruckers auslesen lassen.  

Die im Treiber sichtbare Zahl der Restdrucke wird **nicht** über die Windows-Standard-API bereitgestellt,  
sondern über eine herstellerspezifische DLL (`CspStat.dll`), die auch von der DNP-Software *PrinterInfo* verwendet wird.  
Das hier enthaltene Python-Skript zeigt, wie man diese DLL einbindet und die Werte abfragt.  

---

## Funktionsumfang

- Auslesen der Restdrucke (`GetMediaCounter`)  
- Auslesen der Gesamtkapazität der Rolle (`GetInitialMediaCount`)  
- Abfrage nur, wenn der Drucker **im Leerlauf** ist (Vermeidung von Problemen beim Drucken)  
- Automatisierte Abfrage alle **5 Minuten**  
- Versand einer **E-Mail-Benachrichtigung**, wenn weniger als (z.B.) 10 % der Rolle verbleiben  

**Beispielausgabe:**

```
Vendor port: 0
Remaining prints: 154
Total capacity (current roll): 400
Remaining percent: 38.5%
```

Damit lässt sich z. B. in einer Fotobox-Software frühzeitig reagieren, wenn das Papier bald zu Ende geht.  

---

## Voraussetzungen

- **Windows-PC** mit angeschlossenem DNP DS620  
- Installierte **DNP PrinterInfo-Software** (darin befindet sich die Datei `CspStat.dll`)  
  - [DNP PrinterInfo Download](https://dnpphoto.com/Portals/0/Resources/PrinterInfo_1.2.1.1.zip)  
- **Python** in der Bit-Version passend zur DLL  
  - Die DLL ist meist **32-bit**, daher bitte auch **32-bit-Python** installieren  

---

## Installation

1. **Repository klonen**
   ```powershell
   git clone https://github.com/AarchiveSoft/readPrinterInfo.git
   cd readPrinterInfo
   ```
2. **Virtuelle Umgebung erstellen (32-bit Python)**

    ```powershell
    py -3.11-32 -m venv .venv32
    .\.venv32\Scripts\activate
    ```

3. **Abhängigkeiten installieren**

    ```powershell
    pip install -r requirements.txt
    ```

    *Enthalten sind u. a.:*

    * pywin32 (Zugriff auf Windows-Druckerspooler)
    * schedule (für den 5-Minuten-Timer)
   
   
4. **DLL bereitstellen**

    * Kopiere CspStat.dll aus deinem PrinterInfo-Installationsordner (z. B. C:\DNPIA\PrinterInfo\CspStat.dll)
    * Lege die Datei in den Ordner ./data/ des Projekts

## Nutzung

**Starte den Monitor mit:**

```powershell
python .\main.py
```
Das Skript überprüft alle _5 Minuten_ den Druckerstatus.

* Ist der Drucker beschäftigt → wird die Abfrage übersprungen und beim nächsten Zyklus erneut geprüft.
* Ist der Drucker frei → wird die DLL abgefragt und der Restbestand ermittelt.
* Fällt der Restbestand unter 10 % → wird automatisch eine E-Mail an a.hafner@graphicart.ch gesendet.

## E-Mail-Versand

Der E-Mail-Versand erfolgt über den SMTP-Relay-Server:

_In diesem Beispiel:_  

* **Absender:** noreply@graphicart.ch
* **Empfänger:** a.hafner@graphicart.ch
* **SMTP-Server:** smtp-relay.gmail.com (Port 587, STARTTLS)

Falls dein Relay keine Authentifizierung benötigt (z. B. IP-Whitelist), kann das Skript ohne server.login() betrieben werden.  
Ansonsten können Benutzername/Passwort in der main.py ergänzt werden.

## Hinweise

**Nur im Leerlauf abfragen:**  
    Laut DNP kann ein Statusaufruf während eines aktiven Drucks den Drucker blockieren.  
    Die angepasste Version des Beispiel Scripts enthält einen Sicherheitscheck, um diesem Problem zu entgehen.

**DLL nicht weitergeben:**  
    CspStat.dll ist Teil der DNP-Software. Bitte nutze die lokal installierte Version.

## Fehlerbehebung

* WinError 193 → Python-Bitversion passt nicht zur DLL (32-bit installieren).
* „Could not locate DS620“ → Drucker ist nicht verbunden, nicht bereit oder wird von einer anderen Software blockiert.

## Lizenz

Dieses Beispielskript steht unter der [MIT-Lizenz](https://mit-license.org).  
Die Datei CspStat.dll ist nicht Teil dieses Repositories und bleibt Eigentum von DNP.