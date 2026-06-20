import sys
import os
import pypdfium2 as pdfium

PDF = 'Hartill Qing Coins Only.pdf'
OUT = '_pages'

def render_range(start, end, width=1100, quality=72, scale=2.0):
    pdf = pdfium.PdfDocument(PDF)
    n = len(pdf)
    end = min(end, n)
    for i in range(start, end):
        page = pdf[i]
        bmp = page.render(scale=scale)
        pil = bmp.to_pil().convert('RGB')
        w, h = pil.size
        nh = int(h * width / w)
        pil = pil.resize((width, nh))
        path = os.path.join(OUT, f'jpg_{i:03d}.jpg')
        pil.save(path, 'JPEG', quality=quality)
        print(path, os.path.getsize(path))

if __name__ == '__main__':
    start = int(sys.argv[1])
    end = int(sys.argv[2])
    width = int(sys.argv[3]) if len(sys.argv) > 3 else 1100
    render_range(start, end, width)
