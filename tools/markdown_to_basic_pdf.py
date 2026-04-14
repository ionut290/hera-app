#!/usr/bin/env python3
from pathlib import Path
import textwrap

INPUT = Path('docs/Libro_Completo_Hera_App.md')
OUTPUT = Path('docs/Libro_Completo_Hera_App.pdf')

PAGE_W = 595
PAGE_H = 842
MARGIN_X = 48
MARGIN_Y = 48
FONT_SIZE = 10
LEADING = 14
CHAR_WIDTH = 5.2  # approx Helvetica 10
MAX_CHARS = int((PAGE_W - 2 * MARGIN_X) / CHAR_WIDTH)


def escape_pdf_text(s: str) -> str:
    return s.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')


def normalize_line(line: str) -> str:
    line = line.replace('\t', '    ').replace('–', '-').replace('—', '-').replace('’', "'")
    return ''.join(ch if ord(ch) < 256 else '?' for ch in line)


raw = INPUT.read_text(encoding='utf-8')
all_lines = []
for line in raw.splitlines():
    line = normalize_line(line)
    if not line.strip():
        all_lines.append('')
        continue
    bullet_prefix = ''
    content = line
    stripped = line.lstrip()
    if stripped.startswith('- '):
        bullet_prefix = '- '
        content = stripped[2:]
    elif stripped.startswith(tuple(f'{i}. ' for i in range(1, 10))):
        # simple numbered list handling (keep as-is)
        content = stripped
    wrapped = textwrap.wrap(content, width=MAX_CHARS - len(bullet_prefix)) or ['']
    for i, w in enumerate(wrapped):
        all_lines.append((bullet_prefix if i == 0 else '  ') + w)

lines_per_page = int((PAGE_H - 2 * MARGIN_Y) / LEADING)
pages = [all_lines[i:i + lines_per_page] for i in range(0, len(all_lines), lines_per_page)]

objects = []


def add_obj(data: bytes):
    objects.append(data)
    return len(objects)

font_obj = add_obj(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

page_obj_ids = []
content_obj_ids = []

for page_lines in pages:
    stream_lines = ["BT", f"/F1 {FONT_SIZE} Tf", f"1 0 0 1 {MARGIN_X} {PAGE_H - MARGIN_Y} Tm", f"{LEADING} TL"]
    first = True
    for ln in page_lines:
        txt = escape_pdf_text(ln)
        if first:
            stream_lines.append(f"({txt}) Tj")
            first = False
        else:
            stream_lines.append(f"T* ({txt}) Tj")
    stream_lines.append("ET")
    stream = '\n'.join(stream_lines).encode('latin-1', errors='replace')
    content_obj = add_obj(f"<< /Length {len(stream)} >>\nstream\n".encode() + stream + b"\nendstream")
    content_obj_ids.append(content_obj)
    page_obj = add_obj(b"")  # placeholder
    page_obj_ids.append(page_obj)

kids = ' '.join(f'{pid} 0 R' for pid in page_obj_ids)
pages_obj = add_obj(f"<< /Type /Pages /Kids [ {kids} ] /Count {len(page_obj_ids)} /MediaBox [0 0 {PAGE_W} {PAGE_H}] >>".encode())

for i, pid in enumerate(page_obj_ids):
    cont = content_obj_ids[i]
    objects[pid - 1] = f"<< /Type /Page /Parent {pages_obj} 0 R /Resources << /Font << /F1 {font_obj} 0 R >> >> /Contents {cont} 0 R >>".encode()

catalog_obj = add_obj(f"<< /Type /Catalog /Pages {pages_obj} 0 R >>".encode())

pdf = bytearray(b"%PDF-1.4\n")
offsets = [0]
for i, obj in enumerate(objects, start=1):
    offsets.append(len(pdf))
    pdf.extend(f"{i} 0 obj\n".encode())
    pdf.extend(obj)
    pdf.extend(b"\nendobj\n")

xref_pos = len(pdf)
pdf.extend(f"xref\n0 {len(objects)+1}\n".encode())
pdf.extend(b"0000000000 65535 f \n")
for off in offsets[1:]:
    pdf.extend(f"{off:010d} 00000 n \n".encode())

pdf.extend(f"trailer\n<< /Size {len(objects)+1} /Root {catalog_obj} 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode())
OUTPUT.write_bytes(pdf)
print(f"Creato: {OUTPUT} ({OUTPUT.stat().st_size} bytes)")
