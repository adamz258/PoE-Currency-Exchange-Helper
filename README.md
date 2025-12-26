# PoE Currency Exchange Helper

Lightweight Windows 11 desktop app that reads the Path of Exile currency exchange UI from your screen and suggests the correct side to match the market ratio.

## Demo

https://www.youtube.com/watch?v=8QfIVkGjcI0

## What it does
- Reads the market ratio and your input boxes via OCR.
- Calculates the recommended value to match the ratio.
- Supports auto-detection of which side you are changing.
- Minimal overlay mode for in-game use.
- Always on top toggle.
- Saves OCR regions to `config.json`.

## Requirements
- Windows 11
- Python 3.10+
- Tesseract OCR (installed separately)

## Install
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Install Tesseract:
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

## Using the app
1. Click `Pick Ratio`, then drag a tight rectangle around the ratio digits only.
2. Click `Pick Left Box` and `Pick Right Box` around the input digits.
3. Use `Swap Left/Right` if you selected boxes in reverse.
4. Set the mode:
   - `Auto`: detects which side you are changing.
   - `I have -> calc I want`: uses right input + ratio to recommend left input.
   - `I want -> calc I have`: uses left input + ratio to recommend right input.
5. Toggle `Always on top` if you want it above the game window.
6. Toggle `Minimal mode` for a compact overlay showing only ratio + recommendation.

## OCR accuracy tips
- Make the ratio region as tight as possible around the digits and colon.
- Avoid extra UI text, borders, or icons in the ratio selection.
- Keep game UI scale consistent after selecting regions.
- If the ratio is misread, reselect the ratio region first.

## Files
- `config.json`: saved regions and settings.
- `app.log`: OCR updates and app status logs.
- `crash.log`: crash diagnostics (if any).

## Troubleshooting
- If you see "Tesseract OCR not found", verify the install and `TESSERACT_PATH`.
- If Always on top does not stick, ensure no other app is forcing z-order.
- If OCR lags, reselect regions and keep them tight to digits.
