import os, sys, shutil
print('Python:', sys.executable)
print('Version:', sys.version.split()[0])
try:
    import flask; print('Flask: OK')
except Exception as e: print('Flask: ERROR', e)
try:
    import cv2; print('OpenCV:', cv2.__version__)
except Exception as e: print('OpenCV: ERROR', e)
try:
    import pytesseract
    candidates=[os.environ.get('TESSERACT_CMD'), shutil.which('tesseract'), r'C:\Program Files\Tesseract-OCR\tesseract.exe', r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe']
    path=next((x for x in candidates if x and os.path.isfile(x)), None)
    if path:
        pytesseract.pytesseract.tesseract_cmd=path
        print('Tesseract:', path)
        print('Tesseract version:', str(pytesseract.get_tesseract_version()).splitlines()[0])
    else: print('Tesseract: NOT FOUND')
except Exception as e: print('Tesseract/PyTesseract: ERROR', e)
print('OpenAI API key configured:', bool(os.environ.get('OPENAI_API_KEY','').strip()))
