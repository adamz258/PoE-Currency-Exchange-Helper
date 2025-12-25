# Currency Exchange Helper (Python)

Lightweight Windows 11 desktop app that reads the Path of Exile market ratio panel on screen and suggests the correct listing price.

## Requirements
- Python 3.10+ (Windows 11)
- Tesseract OCR

## Install
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Install Tesseract (recommended):
```powershell
winget install --id UB-Mannheim.TesseractOCR
```

If Tesseract is not on PATH, set:
```powershell
$env:TESSERACT_PATH = "C:\Program Files\Tesseract-OCR\tesseract.exe"
```

## Run
```powershell
python app.py
```

## Notes
- Click "Pick Ratio", "Pick Left Box", and "Pick Right Box" once. Regions are saved in `config.json`.
- Use "Swap Left/Right" if your items are on the right side.
- OCR runs every ~1.2 seconds. Use "Pause OCR" to freeze updates.
- For best accuracy, keep each region tight around the digits only.
