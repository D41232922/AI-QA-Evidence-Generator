#!/usr/bin/env python
"""
Simple wrapper using pytesseract to OCR an image and print the result to stdout.
Requires: pip install pillow pytesseract
Requires Tesseract binary on PATH (or set TESSERACT_CMD env var to binary path).

Usage: pytesseract_ocr.py <image-path>
"""
import sys
import os

try:
    from PIL import Image
    import pytesseract
except Exception as e:
    sys.stderr.write('Missing dependencies: pip install pillow pytesseract\n')
    sys.exit(2)

if len(sys.argv) < 2:
    print('Usage: pytesseract_ocr.py <image-path>')
    sys.exit(1)

image_path = sys.argv[1]
if not os.path.exists(image_path):
    sys.stderr.write('Image not found: ' + image_path + '\n')
    sys.exit(3)

# Optional: allow overriding tesseract command location
tess_cmd = os.environ.get('TESSERACT_PATH')
if tess_cmd and tess_cmd.upper() != 'WINDOWS_OCR':
    pytesseract.pytesseract.tesseract_cmd = tess_cmd

try:
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img, lang=None)
    print(text)
except Exception as e:
    sys.stderr.write('pytesseract error: ' + str(e) + '\n')
    sys.exit(4)
