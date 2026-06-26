# Student Header Code Assigner

A fast desktop application for assigning header codes to student matricules from Excel spreadsheets.

## Features

- Modern Tkinter GUI (with CustomTkinter for better look)
- Fast search by partial matricule (last digits, full code, etc.)
- Intelligent year mismatch detection
- Duplicate handling
- Auto-incrementing codes with leading zeros
- Undo functionality
- Autosave after every assignment
- Voice input support (toggleable)
- Export coded file
- Session statistics

## Requirements

See `requirements.txt`

## How to Run

```bash
pip install -r requirements.txt
python src/main.py
```

## Build Executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=app.ico src/main.py
```

## Sample Data

See `sample_data/students.xlsx`