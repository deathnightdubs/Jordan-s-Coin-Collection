import sys
import os
import easyocr

# Reuse one reader instance
_reader = None

def get_reader():
    global _reader
    if _reader is None:
        _reader = easyocr.Reader(['ch_sim', 'en'], gpu=False, verbose=False)
    return _reader

def ocr_image(path):
    reader = get_reader()
    # detail=0 returns just text strings in reading order
    lines = reader.readtext(path, detail=0, paragraph=True)
    return "\n".join(lines)

if __name__ == '__main__':
    img = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    text = ocr_image(img)
    if out:
        with open(out, 'w') as f:
            f.write(text)
        print("wrote", out, len(text), "chars")
    else:
        print(text)
