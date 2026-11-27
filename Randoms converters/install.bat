@echo off
echo Installing required Python libraries...
python -m pip install --upgrade pip
python -m pip install pywin32 pypdf pymupdf pillow pytesseract opencv-python numpy torch transformers
echo.
echo All libraries installed!
pause
