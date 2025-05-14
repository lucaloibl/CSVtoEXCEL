# CSV to Excel Updater

A simple desktop application that reads category transaction data from a CSV file, completes and sorts it, and updates a specified region in an existing Excel workbook.

## Voraussetzungen

* Python 3.8 oder höher installiert (macOS, Linux oder Windows)
* Optional: Homebrew („brew install python3“) oder offizielle Python-Distribution

## Installation

1. Repository klonen

   ```bash
   git clone https://github.com/lucaloibl/CSVtoEXCEL.git
   cd CSVtoEXCEL
   ```

2. Virtuelle Umgebung anlegen und aktivieren

   ```bash
   python3 -m venv venv
   source venv/bin/activate   # macOS/Linux
   # venv\Scripts\activate  # Windows
   ```

3. Abhängigkeiten installieren

   ```bash
   pip install -r requirements.txt
   ```

## Nutzung

### Direkt ausführen (Test)

```bash
python main.py
```

### Standalone-Build erzeugen

```bash
python -m PyInstaller --onefile --windowed \
  --hidden-import=tkinter --hidden-import=_tkinter main.py
```

* Das fertige Programm befindet sich anschließend im Ordner `dist/`.

---

*Für Fragen oder Probleme öffne bitte ein Issue im Repo.*
