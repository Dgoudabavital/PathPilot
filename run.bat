@echo off
setlocal
cd /d "%~dp0"
if not exist venv\Scripts\python.exe (
  echo Creating virtual environment...
  py -3.13 -m venv venv
)
call venv\Scripts\activate.bat
python -m pip install -r requirements.txt
set "TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe"
if exist ".env" set "TESSERACT_CMD="
if exist ".env" (
  echo Using .env settings.
) else (
  echo Tesseract: %TESSERACT_CMD%
)
python app.py
