# RS485 Servo GUI

Dieses kleine Tool verbindet sich über RS485/Modbus-ähnliche Frames mit Servo-Drives, lädt Parameter aus einer XML-Datei und öffnet pro Drive ein Tab mit Parameter-Read/Write-Funktionen.

Voraussetzungen
- Python 3.8+
- USB-RS485 Adapter (Windows COM-Port)

Installation
```powershell
python -m pip install -r requirements.txt
```

Start
```powershell
python src/rs485_gui.py
```

Hinweise
- Die Parameterdefinitionen werden aus `config/parameter_SpindleHS1.xml` geladen. Neue Parameter in der XML werden automatisch berücksichtigt.
- Das Skript implementiert Modbus-ähnliche Frames wie spezifiziert (0x03 read, 0x06 write) und benutzt CRC16 (Modbus).
- Die Wartezeit auf den ersten Byte der Antwort berechnet sich aus der aktuellen Baudrate (10 Zeichenlängen), danach wird der Rest der Nachricht gelesen.

Limitations
- Für Enable/Disable gibt es momentan nur einen UI-Button pro Tab (keine automatische Register-Write-Logik), weil das Zielregister dafür projektabhängig ist.
- RS485 DE/RE Steuerung wird nicht explizit per GPIO gesteuert; es wird erwartet, dass der verwendete Adapter automatisch DE/RE handhabt.
