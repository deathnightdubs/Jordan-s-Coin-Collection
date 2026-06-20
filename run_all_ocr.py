import os
import time
import numpy as np
import pypdfium2 as pdfium
import easyocr

PDF = 'Hartill Qing Coins Only.pdf'
OUTDIR = '_pages/txt'
LOG = '_pages/ocr_progress.log'
os.makedirs(OUTDIR, exist_ok=True)

def log(msg):
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    with open(LOG, 'a') as f:
        f.write(line + "\n")

def main():
    reader = easyocr.Reader(['ch_sim', 'en'], gpu=False, verbose=False)
    pdf = pdfium.PdfDocument(PDF)
    n = len(pdf)
    log(f"Total pages: {n}")
    for i in range(n):
        out = os.path.join(OUTDIR, f"p{i:03d}.txt")
        if os.path.exists(out):
            continue
        t0 = time.time()
        page = pdf[i]
        bmp = page.render(scale=2.0)
        pil = bmp.to_pil().convert('RGB')
        arr = np.array(pil)
        try:
            lines = reader.readtext(arr, detail=0, paragraph=True)
        except Exception as e:
            lines = [f"<OCR_ERROR: {e}>"]
        text = "\n".join(lines)
        with open(out, 'w') as f:
            f.write(text)
        log(f"page {i:03d} -> {len(text)} chars in {time.time()-t0:.1f}s")
    log("DONE")

if __name__ == '__main__':
    main()
