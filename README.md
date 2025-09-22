# DNP DS620 – Restdrucke per Python auslesen

Dieses Repository enthält ein Beispielskript, mit dem sich die **verbleibenden Drucke** eines  
DNP DS620 Fotodruckers auslesen lassen.  

Die im Treiber sichtbare Zahl der Restdrucke wird **nicht** über die Windows-Standard-API bereitgestellt,  
sondern über eine herstellerspezifische DLL (`CspStat.dll`), die auch von der DNP-Software *PrinterInfo* verwendet wird.  
Das hier enthaltene Python-Skript zeigt, wie man diese DLL einbindet und die Werte abfragt.  

---

## Funktionsumfang

Das Skript gibt unter anderem folgende Informationen aus:

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
   
2. **Virtuelle Umgebung erstellen** (32-bit Python)

    ```powershell
    py -3.11-32 -m venv .venv32
    .\.venv32\Scripts\activate
    ```

3. **Abhängigkeiten installieren**

    ```powershell
    pip install -r requirements.txt
    ```

4. **DLL bereitstellen**


&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Kopiere **CspStat.dll** aus deinem PrinterInfo-Installationsordner (z. B. **C:\DNPIA\PrinterInfo\CspStat.dll**)

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Lege die Datei in den Ordner ./data/ des Projekts

## Nutzung

```powershell
python .\main.py
```

**Erwartete Ausgabe** (Beispiel):

```
Vendor port: 0
Remaining prints: 154
Total capacity (current roll): 400
Remaining percent: 38.5%
Raw status code: 0
```

## Hinweise
**Nur im Leerlauf abfragen:** Laut DNP kann ein Statusaufruf während eines aktiven Drucks das Gerät blockieren.

**DLL nicht weitergeben:** CspStat.dll ist Teil der DNP-Software. Bitte nutze die lokal installierte Version.

## Fehlerbehebung:

- **WinError 193** → Python-Bitversion passt nicht zur DLL (32-bit installieren).

- **„Could not locate DS620“** → Drucker ist nicht verbunden, nicht bereit oder wird von einer anderen Software blockiert.

## Lizenz
Dieses Beispielskript steht unter der [MIT-Lizenz](https://mit-license.org).

Die Datei CspStat.dll ist nicht Teil dieses Repositories und bleibt Eigentum von DNP.