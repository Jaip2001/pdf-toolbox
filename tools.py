"""
tools.py — saare PDF operations ek jagah

  Conversion : to_xlsx (with / without images), to_docx (editable / exact), to_jpg, images_zip
  Security   : protect_pdf, unprotect_pdf
  Edit       : apply_ops  (rotate, delete, move, duplicate, reorder, extract, insert blank,
                           watermark, header/footer, page numbers, merge)
"""
import os
import re
import io
import csv
import math
import json
import shutil
import zipfile
import hashlib
from collections import Counter

import pymupdf
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.pagebreak import Break

import converter as C   # estimate engine + shared helpers

EMU_PER_PT = 12700
THIN = Side(style='thin', color='FF000000')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FN = 'Calibri'
F = Font(name=FN, size=8)
FB = Font(name=FN, size=8, bold=True)
LT = Alignment(horizontal='left', vertical='top', wrap_text=True)
LC = Alignment(horizontal='left', vertical='center', wrap_text=True)
CC = Alignment(horizontal='center', vertical='center', wrap_text=True)
R = Alignment(horizontal='right', vertical='center', shrink_to_fit=True)


# =====================================================================
#  helpers
# =====================================================================

def clean(t):
    return re.sub(r'[ \t]+', ' ', t.replace('\xa0', ' ')).strip()


def is_num(s):
    return bool(re.fullmatch(r'-?[\d,]+\.?\d*', s.strip())) and s.strip() not in ('', '.', '-')


def num(s):
    try:
        return float(s.replace(',', ''))
    except Exception:
        return s


def fix_num(txt):
    parts = [p.strip() for p in txt.split('\n') if p.strip()]
    if len(parts) > 1 and all(re.fullmatch(r'[\d,.]+', p) for p in parts):
        j = ''.join(parts)
        if is_num(j):
            return j
    return txt


def excel_w(pt):
    return max(round(((pt * 96 / 72) - 5) / 7, 2), 3.5)


def set_cell(ws, r, c, v, font=F, align=LT, fmt=None):
    cell = ws.cell(row=r, column=c)
    if isinstance(v, str):
        s = v.strip()
        if s == '':
            cell.font = font
            cell.alignment = align
            return cell
        if is_num(s):
            cell.value = num(s)
            cell.number_format = fmt or ('0.000' if '.' in s and len(s.split('.')[1]) == 3
                                         else '#,##0.00')
        else:
            cell.value = s
    else:
        cell.value = v
        if fmt:
            cell.number_format = fmt
    cell.font = font
    cell.alignment = align
    return cell


def page_lines(page, skip_rects=(), cache=None, pno=None):
    """text lines (list of dicts) that are not inside the given rectangles"""
    if cache is not None and pno is not None:
        spans = cache.spans(pno)
    else:
        spans = []
        for b in page.get_text('dict')['blocks']:
            if b['type'] != 0:
                continue
            for l in b['lines']:
                for s in l['spans']:
                    t = clean(s['text'])
                    if t:
                        spans.append({'t': t, 'bbox': list(s['bbox']),
                                      'size': s['size'], 'font': s.get('font', '')})
    out, cur, lasty = [], [], None
    for s in sorted(spans, key=lambda a: (round(a['bbox'][1], 1), a['bbox'][0])):
        x0, y0, x1, y1 = s['bbox']
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        inside = any(r[0] - 2 <= cx <= r[2] + 2 and r[1] - 2 <= cy <= r[3] + 2
                     for r in skip_rects)
        if inside:
            continue
        if lasty is not None and abs(y0 - lasty) > 3:
            out.append(_mk_line(cur))
            cur = []
        cur.append(s)
        lasty = y0
    if cur:
        out.append(_mk_line(cur))
    return [o for o in out if o]


def _mk_line(spans):
    if not spans:
        return None
    txt = ' '.join(s['t'] for s in spans).strip()
    if not txt:
        return None
    x0 = min(s['bbox'][0] for s in spans)
    y0 = min(s['bbox'][1] for s in spans)
    size = max(s['size'] for s in spans)
    bold = any('Bold' in (s.get('font') or '') for s in spans)
    return {'text': txt, 'x': x0, 'y': y0, 'size': size, 'bold': bold}


def table_rects(page):
    try:
        return [t.bbox for t in page.find_tables().tables]
    except Exception:
        return []



class Cache:
    """per-page cached spans / tables / lines — speed ke liye"""

    def __init__(self, doc):
        self.doc = doc
        self._spans, self._tables = {}, {}

    def spans(self, pno):
        if pno not in self._spans:
            out = []
            for b in self.doc[pno].get_text('dict')['blocks']:
                if b['type'] != 0:
                    continue
                for l in b['lines']:
                    for s in l['spans']:
                        t = clean(s['text'])
                        if t:
                            out.append({'t': t, 'bbox': list(s['bbox']),
                                        'size': s['size'], 'font': s.get('font', '')})
            self._spans[pno] = out
        return self._spans[pno]

    def tables(self, pno):
        if pno not in self._tables:
            try:
                self._tables[pno] = self.doc[pno].find_tables().tables
            except Exception:
                self._tables[pno] = []
        return self._tables[pno]


def near(edges, v, tol=1.5):
    """edge index jahan v aata hai"""
    for i, e in enumerate(edges):
        if abs(e - v) <= tol:
            return i
    return min(range(len(edges)), key=lambda i: abs(edges[i] - v))


def table_grid(pno, cache, table):
    """returns (grid, edges, merges) — grid[row][col] = text"""
    edges = C.edges_of(table)
    rows = table.rows
    ncol = len(edges) - 1
    grid = [[[] for _ in range(ncol)] for _ in rows]
    tb = table.bbox

    for s in cache.spans(pno):
        x0, y0, x1, y1 = s['bbox']
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        if not (tb[0] - 1.5 <= cx <= tb[2] + 1.5 and tb[1] - 1.5 <= cy <= tb[3] + 1.5):
            continue
        ri = None
        for i, tr in enumerate(rows):
            if tr.bbox[1] - 1.5 <= cy <= tr.bbox[3] + 1.5:
                ri = i
                break
        if ri is None:
            ri = min(range(len(rows)),
                     key=lambda i: abs((rows[i].bbox[1] + rows[i].bbox[3]) / 2 - cy))
        ci = near(edges, cx)
        ci = min(ci, ncol - 1)
        grid[ri][ci].append((round(cy, 1), round(cx, 1), s['t']))

    out = []
    for r in grid:
        row = []
        for cell in r:
            cell.sort()
            lines, cur, lasty = [], [], None
            for y, x, t in cell:
                if lasty is not None and abs(y - lasty) > 3.5:
                    lines.append(' '.join(cur))
                    cur = []
                cur.append(t)
                lasty = y
            if cur:
                lines.append(' '.join(cur))
            row.append(fix_num('\n'.join(lines)))
        out.append(row)

    merges = []
    for ri, tr in enumerate(rows):
        for c in tr.cells:
            if c is None:
                continue
            c0, c1 = near(edges, c[0]), near(edges, c[2])
            if c1 - c0 > 1:
                merges.append((ri, c0, c1 - 1))
    return out, edges, merges


# =====================================================================
#  XLSX  (structure preserving, any pdf)
# =====================================================================

def to_xlsx(doc, out_path, with_images=True, img_dir=None, force='auto'):
    """force: auto | estimate | table | text"""
    tpl = C.detect_template(doc)
    engine = tpl if force in ('auto', None) else force
    if engine == 'generic':
        engine = 'table'
    if engine == 'estimate' and tpl != 'estimate':
        engine = 'table' if tpl == 'generic' else 'text'
    if engine == 'text' and C.table_check(doc) == 'generic':
        engine = 'table'

    if engine == 'estimate':
        parsed = C.parse_estimate(doc, img_dir)
        parsed['header'] = C.parse_header(doc)
        C.build_estimate_xlsx(parsed, out_path, with_images=with_images)
        return {'engine': 'estimate', 'rows': None}
    if engine == 'table':
        build_table_xlsx(doc, out_path, with_images=with_images, img_dir=img_dir)
        return {'engine': 'table', 'rows': None}
    build_text_xlsx(doc, out_path)
    return {'engine': 'text', 'rows': None}


def build_table_xlsx(doc, out_path, with_images=True, img_dir=None):
    """har page ke tables ko structure ke saath ek sheet me likhta hai,
    aur table ke bahar ka text bhi saath me rakhta hai"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'PDF'
    cache = Cache(doc)
    row = 1
    colw = {}

    for pno in range(len(doc)):
        page = doc[pno]
        tables = cache.tables(pno)
        trects = [t.bbox for t in tables]
        lines = page_lines(page, skip_rects=trects, cache=cache)
        tbl_y = min([r[1] for r in trects], default=10 ** 6)
        top = [l for l in lines if l['y'] < tbl_y - 2]
        bottom = [l for l in lines if l['y'] >= tbl_y - 2]
        if pno > 0:
            row += 1

        for l in top + bottom:
            c = ws.cell(row=row, column=1, value=l['text'])
            c.font = FB if l['bold'] or l['size'] > 12 else F
            c.alignment = LC
            ws.row_dimensions[row].height = max(12, min(l['size'] * 1.4, 22))
            row += 1
        if top:
            row += 0

        for table in tables:
            grid, edges, merges = table_grid(pno, cache, table)
            if len(edges) < 3 or not grid:
                continue
            widths = [edges[i + 1] - edges[i] for i in range(len(edges) - 1)]
            for i, w in enumerate(widths, start=1):
                wid = min(excel_w(w), 42)
                colw[i] = max(colw.get(i, 0), wid)
            for ri, grow in enumerate(grid):
                tr = table.rows[ri]
                h = tr.bbox[3] - tr.bbox[1]
                ws.row_dimensions[row].height = max(11.5, min(h, 60))
                merged_cols = set()
                for (mr, c0, c1) in merges:
                    if mr == ri:
                        merged_cols.update(range(c0, c1))
                for ci, txt in enumerate(grow):
                    if ci in merged_cols:
                        continue
                    span_end = ci
                    for (mr, c0, c1) in merges:
                        if mr == ri and c0 == ci:
                            span_end = c1
                    body = '\n'.join(grow[j] for j in range(ci, span_end + 1) if grow[j])
                    if span_end > ci:
                        ws.merge_cells(start_row=row, start_column=ci + 1,
                                       end_row=row, end_column=span_end + 1)
                    headerish = (ri == 0 and body and not is_num(body))
                    set_cell(ws, row, ci + 1, body, FB if headerish else F,
                             R if is_num(body) else LT)
                for cc in range(1, len(edges)):
                    ws.cell(row=row, column=cc).border = BORDER
                row += 1
            row += 1

        if with_images and page.get_images(full=True):
            for im in C.collect_page_images(doc, pno, img_dir, ''):
                ix0, iy0, ix1, iy1 = im['bbox']
                h_pt = max(14, min(iy1 - iy0, 70))
                w_pt = (ix1 - ix0) * (h_pt / max(iy1 - iy0, .1))
                if w_pt > 150:
                    h_pt = h_pt * 150 / w_pt
                    w_pt = 150
                img = XLImage(im['path'])
                px = lambda pt: pt * 96 / 72
                img.width, img.height = px(w_pt), px(h_pt)
                need = h_pt + 3
                if (ws.row_dimensions[row].height or 12) < need:
                    ws.row_dimensions[row].height = need
                img.anchor = OneCellAnchor(
                    _from=AnchorMarker(col=0, row=row - 1,
                                       colOff=int(2 * EMU_PER_PT),
                                       rowOff=int(1.5 * EMU_PER_PT)),
                    ext=XDRPositiveSize2D(cx=int(pixels_to_EMU(img.width)),
                                          cy=int(pixels_to_EMU(img.height))))
                ws.add_image(img)
                row += 1

    for i, w in colw.items():
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    wb.save(out_path)
    return ws.max_row


def build_text_xlsx(doc, out_path):
    """table detect na ho to bhi x-position ke hisaab se columns bana kar likhta hai"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'PDF'
    cache = Cache(doc)
    row = 1

    for pno in range(len(doc)):
        spans = cache.spans(pno)
        if not spans:
            row += 1
            continue
        # column anchors: distinct x positions
        xs = sorted({round(s['bbox'][0]) for s in spans})
        anchors = []
        for x in xs:
            if not anchors or x - anchors[-1] > 9:
                anchors.append(x)
        # line grouping
        lines, cur, lasty = [], [], None
        for s in sorted(spans, key=lambda a: (round(a['bbox'][1], 1), a['bbox'][0])):
            y = s['bbox'][1]
            if lasty is not None and abs(y - lasty) > 3.5:
                lines.append(cur)
                cur = []
            cur.append(s)
            lasty = y
        if cur:
            lines.append(cur)

        used = [0] * len(anchors)
        for ln in lines:
            cells = {}
            size = 0
            bold = False
            for s in ln:
                ci = min(range(len(anchors)),
                         key=lambda i: abs(anchors[i] - s['bbox'][0]))
                cells.setdefault(ci, []).append(s['t'])
                size = max(size, s['size'])
                bold = bold or ('Bold' in (s['font'] or ''))
                used[ci] += 1
            big = size > 12 or bold
            for ci, txt in cells.items():
                c = ws.cell(row=row, column=ci + 1, value=' '.join(txt))
                c.font = FB if big else F
                c.alignment = R if is_num(' '.join(txt)) else LT
                ws.row_dimensions[row].height = max(12, min(size * 1.4, 24))
            row += 1
        # widths
        for i, a in enumerate(anchors):
            if not used[i]:
                continue
            nxt = anchors[i + 1] if i + 1 < len(anchors) else a + 100
            ws.column_dimensions[get_column_letter(i + 1)].width = \
                max(ws.column_dimensions[get_column_letter(i + 1)].width or 0,
                    min(excel_w(max(nxt - a, 30)), 42))
        row += 1

    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(out_path)
    return row


# =====================================================================
#  DOCX
# =====================================================================

def to_docx(doc, out_path, mode='editable', img_dir=None, dpi=150):
    import docx
    from docx.shared import Pt, Inches, RGBColor, Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT

    d = docx.Document()
    sec = d.sections[0]
    first_land = doc[0].rect.width > doc[0].rect.height
    try:
        sec.style.font.name = 'Calibri'
        sec.style.font.size = Pt(10)
    except Exception:
        pass
    if first_land:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = Inches(11.69), Inches(8.27)
    else:
        sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    sec.left_margin = sec.right_margin = Inches(0.4)
    sec.top_margin = sec.bottom_margin = Inches(0.4)

    if mode == 'exact':
        for pno in range(len(doc)):
            pix = doc[pno].get_pixmap(dpi=dpi)
            tmp = os.path.join(img_dir, f'wpage_{pno + 1}.jpg')
            pix.save(tmp, jpg_quality=88)
            usable_w = sec.page_width - sec.left_margin - sec.right_margin
            ratio = pix.width / pix.height
            width = usable_w
            height = int(usable_w / ratio)
            maxh = sec.page_height - sec.top_margin - sec.bottom_margin
            if height > maxh:
                height = int(maxh)
                width = int(height * ratio)
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            p.add_run().add_picture(tmp, width=Emu(int(width)), height=Emu(int(height)))
            if pno < len(doc) - 1:
                d.add_page_break()
        d.save(out_path)
        return {'mode': 'exact', 'pages': len(doc)}

    cache = Cache(doc)
    for pno in range(len(doc)):
        page = doc[pno]
        tables = cache.tables(pno)
        trects = [t.bbox for t in tables]
        lines = page_lines(page, skip_rects=trects, cache=cache)
        tbl_y = min([r[1] for r in trects], default=10 ** 6)
        top = [l for l in lines if l['y'] < tbl_y - 2]
        bottom = [l for l in lines if l['y'] >= tbl_y - 2]

        for l in top:
            para = d.add_paragraph()
            run = para.add_run(l['text'])
            run.font.size = Pt(min(20, max(7, l['size'] * 0.75)))
            if l['bold'] or l['size'] > 12:
                run.bold = True
            para.paragraph_format.space_after = Pt(2)

        for table in tables:
            grid, edges, merges = table_grid(pno, cache, table)
            if not grid or len(edges) < 3:
                continue
            ncols = len(edges) - 1
            t = d.add_table(rows=len(grid), cols=ncols)
            try:
                t.style = 'Table Grid'
            except Exception:
                pass
            usable = (sec.page_width - sec.left_margin - sec.right_margin)
            widths = [edges[i + 1] - edges[i] for i in range(len(edges) - 1)]
            scale = usable / max(sum(widths), 1)
            _Emu = Emu
            for ri, trow in enumerate(t.rows):          # ek hi baar cells nikalo
                cells = trow.cells
                grow = grid[ri] if ri < len(grid) else []
                for ci, cell in enumerate(cells):
                    txt = grow[ci] if ci < len(grow) else ''
                    cell.text = ''
                    para = cell.paragraphs[0]
                    run = para.add_run(txt.replace('\n', ' '))
                    run.font.size = Pt(8)
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    if ri == 0 and txt and not is_num(txt):
                        run.bold = True
                    if ci < len(widths):
                        try:
                            cell.width = _Emu(int(widths[ci] * scale))
                        except Exception:
                            pass
            for (mr, c0, c1) in merges:                 # merged cells
                try:
                    row_cells = t.rows[mr].cells
                    row_cells[c0].merge(row_cells[c1])
                except Exception:
                    pass
            d.add_paragraph()

        for l in bottom:
            para = d.add_paragraph()
            run = para.add_run(l['text'])
            run.font.size = Pt(min(20, max(7, l['size'] * 0.75)))
            if l['bold'] or l['size'] > 12:
                run.bold = True
            para.paragraph_format.space_after = Pt(2)

        if pno < len(doc) - 1:
            d.add_page_break()

    d.save(out_path)
    return {'mode': 'editable', 'pages': len(doc)}


def usable_width(d, section, landscape):
    from docx.shared import Inches
    return Inches(11.69 if landscape else 8.27) - Inches(0.8)


# =====================================================================
#  JPG / images
# =====================================================================

def to_jpg_zip(doc, zip_path, dpi=150, quality=88):
    names = []
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for pno in range(len(doc)):
            pix = doc[pno].get_pixmap(dpi=dpi)
            data = pix.tobytes('jpg', jpg_quality=quality)
            name = f'page_{pno + 1:03d}.jpg'
            z.writestr(name, data)
            names.append({'file': name, 'w': pix.width, 'h': pix.height, 'bytes': len(data)})
    return names


def images_zip(doc, zip_path):
    seen, rows = set(), []
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        n = 0
        for pno in range(len(doc)):
            for info in doc[pno].get_image_info(xrefs=True):
                xref = info['xref']
                if not xref or xref in seen:
                    continue
                seen.add(xref)
                ext = doc.extract_image(xref)
                if not ext or not ext.get('image'):
                    continue
                n += 1
                name = f'page{pno + 1:02d}_{n:03d}.{ext["ext"]}'
                z.writestr(name, ext['image'])
                rows.append([name, pno + 1, ext['width'], ext['height'], len(ext['image'])])
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(['file', 'page', 'width_px', 'height_px', 'size_bytes'])
        w.writerows(rows)
        z.writestr('images_list.csv', buf.getvalue())
    return rows


# =====================================================================
#  SECURITY
# =====================================================================

PERM = {
    'print': pymupdf.PDF_PERM_PRINT | pymupdf.PDF_PERM_PRINT_HQ,
    'copy': pymupdf.PDF_PERM_COPY,
    'modify': pymupdf.PDF_PERM_MODIFY,
    'annotate': pymupdf.PDF_PERM_ANNOTATE,
    'form': pymupdf.PDF_PERM_FORM,
    'accessibility': pymupdf.PDF_PERM_ACCESSIBILITY,
    'assemble': pymupdf.PDF_PERM_ASSEMBLE,
}


def protect_pdf(src, out, user_pw='', owner_pw='', allow=('print', 'copy'),
                encrypt='aes256'):
    """allow: list of allowed permissions"""
    bits = 0
    for k in allow:
        bits |= PERM.get(k, 0)
    enc = pymupdf.PDF_ENCRYPT_AES_256 if encrypt == 'aes256' else pymupdf.PDF_ENCRYPT_AES_128
    doc = pymupdf.open(src)
    if doc.needs_pass:
        raise ValueError('File pehle se password protected hai — pehle unprotect karein.')
    doc.save(out, encryption=enc, owner_pw=owner_pw or user_pw or 'owner',
             user_pw=user_pw, permissions=bits, garbage=4, deflate=True)
    doc.close()
    return {'encryption': encrypt, 'allow': list(allow)}


def unprotect_pdf(src, out, password=''):
    doc = pymupdf.open(src)
    if not doc.is_encrypted:
        # bas ek clean copy
        doc.save(out, garbage=4, deflate=True)
        doc.close()
        return {'already_open': True}
    ok = doc.authenticate(password)
    if not ok:
        doc.close()
        raise ValueError('Password galat hai.')
    doc.save(out, encryption=pymupdf.PDF_ENCRYPT_NONE, garbage=4, deflate=True)
    doc.close()
    return {'already_open': False}


# =====================================================================
#  EDIT
# =====================================================================

def page_list(pages, total):
    """'all' | [1,3] | '1-4' -> 0-based sorted list"""
    if pages in (None, '', 'all', 'All', '*'):
        return list(range(total))
    out = set()
    if isinstance(pages, str):
        for part in pages.split(','):
            part = part.strip()
            if not part:
                continue
            if '-' in part:
                a, b = part.split('-')[0], part.split('-')[1]
                a, b = int(a), int(b)
                out.update(range(min(a, b) - 1, max(a, b)))
            else:
                out.add(int(part) - 1)
    else:
        for p in pages:
            out.add(int(p) - 1)
    return sorted(p for p in out if 0 <= p < total)


def apply_ops(doc, ops, helpers=None):
    """ops: list of dicts -> naya pymupdf doc"""
    if helpers is None:
        helpers = {}
    for op in ops:
        name = op.get('op')
        total = len(doc)

        if name == 'rotate':
            ang = int(op.get('angle', 90)) % 360
            for i in page_list(op.get('pages', 'all'), total):
                doc[i].set_rotation((doc[i].rotation + ang) % 360)

        elif name == 'delete':
            for i in sorted(page_list(op.get('pages'), total), reverse=True):
                doc.delete_page(i)

        elif name == 'rotate_abs':
            ang = int(op.get('angle', 0)) % 360
            for i in page_list(op.get('pages', 'all'), total):
                doc[i].set_rotation(ang)

        elif name == 'move':
            frm = int(op.get('from', 1)) - 1
            to = int(op.get('to', 1)) - 1
            if 0 <= frm < total:
                doc.move_page(frm, max(0, min(to, total - 1)))

        elif name == 'duplicate':
            for i in sorted(page_list(op.get('pages'), total), reverse=True):
                doc.fullcopy_page(i)

        elif name == 'reorder':
            order = [int(x) - 1 for x in op.get('order', [])]
            if sorted(order) == list(range(total)):
                doc.select(order)

        elif name == 'extract':
            idx = page_list(op.get('pages'), total)
            new = pymupdf.open()
            for i in idx:
                new.insert_pdf(doc, from_page=i, to_page=i)
            helpers['extract'] = new

        elif name == 'insert_blank':
            after = int(op.get('after', total))
            count = max(1, int(op.get('count', 1)))
            src = doc[after - 1] if 0 < after <= total else doc[0]
            for _ in range(count):
                doc.new_page(pno=after, width=src.rect.width, height=src.rect.height)

        elif name == 'merge':
            other = op.get('other_doc')
            if other is not None:
                if op.get('position', 'end') == 'start':
                    other.insert_pdf(doc)
                    doc = other
                else:
                    doc.insert_pdf(other)

        elif name == 'watermark':
            txt = str(op.get('text', '')).strip()
            if txt:
                size = float(op.get('size', 70))
                opacity = float(op.get('opacity', 0.22))
                color = op.get('color', [0.62, 0.62, 0.62])
                ang = float(op.get('angle', 45))
                a = math.cos(math.radians(ang))
                b = math.sin(math.radians(ang))
                for p in doc:
                    r = p.rect
                    w = pymupdf.get_text_length(txt, fontname='hebo', fontsize=size)
                    x = max(20, (r.width - w * abs(a)) / 2)
                    y = r.height * 0.62
                    morph = (pymupdf.Point(x, y), (a, b, -b, a, 0, 0))
                    p.insert_text(pymupdf.Point(x, y), txt, fontsize=size, fontname='hebo',
                                  color=tuple(color), morph=morph,
                                  fill_opacity=opacity, stroke_opacity=opacity)

        elif name == 'header' or name == 'footer':
            txt = str(op.get('text', ''))
            if txt:
                size = float(op.get('size', 9))
                color = op.get('color', [0, 0, 0])
                for i, p in enumerate(doc):
                    r = p.rect
                    t = txt.replace('{page}', str(i + 1)).replace('{pages}', str(len(doc)))
                    w = pymupdf.get_text_length(t, fontname='helv', fontsize=size)
                    x = max(12, (r.width - w) / 2)
                    y = 22 if name == 'header' else r.height - 16
                    p.insert_text(pymupdf.Point(x, y), t, fontsize=size, fontname='helv',
                                  color=tuple(color))

        elif name == 'page_numbers':
            size = float(op.get('size', 9))
            fmt = str(op.get('format', '{page} / {pages}'))
            start = int(op.get('start', 1))
            pos = op.get('position', 'bottom-center')
            for i, p in enumerate(doc):
                r = p.rect
                t = fmt.replace('{page}', str(i + start)).replace('{pages}', str(len(doc)))
                w = pymupdf.get_text_length(t, fontname='helv', fontsize=size)
                if pos.endswith('left'):
                    x = 30
                elif pos.endswith('right'):
                    x = r.width - w - 30
                else:
                    x = (r.width - w) / 2
                y = 24 if pos.startswith('top') else r.height - 16
                p.insert_text(pymupdf.Point(x, y), t, fontsize=size, fontname='helv',
                              color=(0, 0, 0))

        elif name == 'set_metadata':
            md = {k: v for k, v in (op.get('meta') or {}).items() if v}
            if md:
                doc.set_metadata({**doc.metadata, **md})

    return doc
