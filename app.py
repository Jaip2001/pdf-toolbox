"""app.py — PDF Toolbox: Convert (Excel/Word/JPG/Images) · Secure (Protect/Unprotect) · Edit"""
import os
import re
import json
import time
import uuid
import shutil
import functools

import pymupdf
from flask import (Flask, request, jsonify, send_file, render_template,
                   abort, send_from_directory)

import converter as C
import tools

BASE = os.path.dirname(os.path.abspath(__file__))
JOBS = os.path.join(BASE, 'jobs')
os.makedirs(JOBS, exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 400 * 1024 * 1024
MAX_JOBS = 25
MAX_THUMBS = 400


# --------------------------------------------------------------------- helpers

def cleanup_old_jobs():
    dirs = []
    for d in os.listdir(JOBS):
        p = os.path.join(JOBS, d)
        if os.path.isdir(p):
            dirs.append((os.path.getmtime(p), p))
    dirs.sort()
    while len(dirs) > MAX_JOBS:
        _, p = dirs.pop(0)
        shutil.rmtree(p, ignore_errors=True)


def job_dir(job):
    if not job or not re.fullmatch(r'[a-f0-9]{8,20}', job or ''):
        abort(404)
    d = os.path.join(JOBS, job)
    if not os.path.isdir(d):
        abort(404)
    return d


def state_path(d):
    return os.path.join(d, 'state.json')


def load_state(d):
    with open(state_path(d)) as fh:
        return json.load(fh)


def save_state(d, st):
    st['rev'] = int(st.get('rev', 0)) + 1
    with open(state_path(d), 'w') as fh:
        json.dump(st, fh)
    return st


def work_pdf(d, st):
    return os.path.join(d, st.get('work', 'work.pdf'))


def open_pdf(path, password=''):
    doc = pymupdf.open(path)
    if doc.needs_pass:
        if not password or not doc.authenticate(password):
            doc.close()
            raise ValueError('Password galat hai.')
    return doc


def make_thumbs(d, st, force=False):
    """page thumbnails (small jpg) — rev ke saath refresh"""
    folder = os.path.join(d, 'thumbs')
    rev = st.get('rev', 1)
    stamp = os.path.join(folder, f'.rev{rev}')
    if not force and os.path.exists(stamp):
        return
    shutil.rmtree(folder, ignore_errors=True)
    os.makedirs(folder, exist_ok=True)
    doc = pymupdf.open(work_pdf(d, st))
    if doc.needs_pass:
        doc.authenticate(st.get('password', ''))
    n = min(len(doc), MAX_THUMBS)
    for i in range(n):
        page = doc[i]
        zoom = 120 / max(page.rect.width, 1)
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
        pix.save(os.path.join(folder, f'{i + 1}.jpg'), jpg_quality=72)
    doc.close()
    open(stamp, 'w').close()


def job_info(d, st, extra=None):
    info = {
        'job': os.path.basename(d),
        'name': st.get('name', 'file.pdf'),
        'pages': st.get('pages', 0),
        'size': st.get('size', 0),
        'rev': st.get('rev', 1),
        'encrypted': st.get('encrypted', False),
        'opened_with_password': st.get('opened_with_password', False),
        'has_tables': st.get('has_tables', False),
        'engine_guess': st.get('engine_guess', 'text'),
        'thumbs': min(st.get('pages', 0), MAX_THUMBS),
        'landscape': st.get('landscape', False),
    }
    if extra:
        info.update(extra)
    return info


def scan_pdf(path, password=''):
    """upload ke waqt ki basic info"""
    doc = open_pdf(path, password)
    pages = len(doc)
    land = doc[0].rect.width > doc[0].rect.height if pages else False
    try:
        engine = C.detect_template(doc)
    except Exception:
        engine = 'text'
    has_tables = engine in ('estimate', 'generic')
    doc.close()
    return {'pages': pages, 'landscape': land, 'engine_guess': engine,
            'has_tables': has_tables}


# --------------------------------------------------------------------- pages

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify({'error': 'File nahi mili'}), 400
    name = f.filename
    low = name.lower()
    if low.endswith('.pdf'):
        kind = 'pdf'
    elif low.endswith(('.png', '.jpg', '.jpeg', '.webp', '.bmp', '.tif', '.tiff')):
        kind = 'image'
    else:
        return jsonify({'error': 'PDF ya image file upload karein'}), 400

    password = request.form.get('password', '') or ''
    job = uuid.uuid4().hex[:14]
    d = os.path.join(JOBS, job)
    os.makedirs(d, exist_ok=True)
    src = os.path.join(d, 'source' + ('.pdf' if kind == 'pdf' else os.path.splitext(name)[1]))
    f.save(src)

    if kind == 'image':
        # image -> 1 page PDF
        pdfp = os.path.join(d, 'work.pdf')
        imgdoc = pymupdf.open(src)
        pdfdoc = pymupdf.open()
        rect = imgdoc[0].rect
        page = pdfdoc.new_page(width=max(rect.width, 200), height=max(rect.height, 200))
        page.insert_image(page.rect, filename=src)
        pdfdoc.save(pdfp)
        pdfdoc.close()
        imgdoc.close()
        src = pdfp
        name = os.path.splitext(name)[0] + '.pdf'

    try:
        doc = pymupdf.open(src)
        locked = bool(doc.needs_pass)
        if locked and (not password or doc.authenticate(password) <= 0):
            doc.close()
            st = {'name': name, 'work': 'work.pdf', 'pages': 0, 'size': os.path.getsize(src),
                  'rev': 0, 'needs_password': True, 'source': os.path.basename(src)}
            save_state(d, st)
            return jsonify({'job': job, 'needs_password': True, 'name': name})
        if locked:
            # decrypted working copy
            work = os.path.join(d, 'work.pdf')
            doc.save(work, encryption=pymupdf.PDF_ENCRYPT_NONE, garbage=4, deflate=True)
            doc.close()
            doc = pymupdf.open(work)
        doc.close()
    except Exception as e:
        return jsonify({'error': f'PDF khul nahi payi: {e}'}), 400

    work = os.path.join(d, 'work.pdf')
    if not os.path.exists(work):
        shutil.copy(src, work)

    meta = scan_pdf(work, password)
    st = {'name': name if name.lower().endswith('.pdf') else name + '.pdf',
          'work': 'work.pdf', 'pages': meta['pages'], 'size': os.path.getsize(work),
          'rev': 0, 'encrypted': locked, 'opened_with_password': bool(locked and password),
          'has_tables': meta['has_tables'], 'engine_guess': meta['engine_guess'],
          'landscape': meta['landscape'], 'password': password,
          'source': os.path.basename(src)}
    save_state(d, st)
    make_thumbs(d, st, force=True)
    cleanup_old_jobs()
    return jsonify(job_info(d, st))


@app.route('/api/unlock', methods=['POST'])
def unlock():
    data = request.get_json(force=True) or {}
    job = data.get('job')
    password = data.get('password', '')
    d = job_dir(job)
    st = load_state(d)
    src = os.path.join(d, st.get('source', 'source.pdf'))
    try:
        doc = open_pdf(src, password)
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    work = os.path.join(d, 'work.pdf')
    doc.save(work, encryption=pymupdf.PDF_ENCRYPT_NONE, garbage=4, deflate=True)
    doc.close()
    meta = scan_pdf(work, '')
    st.update({'pages': meta['pages'], 'has_tables': meta['has_tables'],
               'engine_guess': meta['engine_guess'], 'landscape': meta['landscape'],
               'encrypted': True, 'opened_with_password': True, 'password': password})
    save_state(d, st)
    make_thumbs(d, st, force=True)
    return jsonify(job_info(d, st))


@app.route('/api/info', methods=['POST'])
def info():
    data = request.get_json(force=True) or {}
    d = job_dir(data.get('job'))
    return jsonify(job_info(d, load_state(d)))


@app.route('/api/thumb/<job>/<int:n>')
def thumb(job, n):
    d = job_dir(job)
    p = os.path.join(d, 'thumbs', f'{n}.jpg')
    if not os.path.exists(p):
        abort(404)
    return send_file(p, mimetype='image/jpeg')


@app.route('/api/preview/<job>')
def preview(job):
    d = job_dir(job)
    st = load_state(d)
    doc = pymupdf.open(work_pdf(d, st))
    if doc.needs_pass:
        doc.authenticate(st.get('password', ''))
    pix = doc[0].get_pixmap(dpi=130)
    out = os.path.join(d, f'preview_{st.get("rev", 1)}.png')
    pix.save(out)
    doc.close()
    return send_file(out, mimetype='image/png')


# --------------------------------------------------------------------- convert

OUT_NAMES = {
    'xlsx_with': 'Excel (with images).xlsx',
    'xlsx_without': 'Excel (without images).xlsx',
    'docx_editable': 'Word (editable).docx',
    'docx_exact': 'Word (same look).docx',
    'jpg': 'JPG pages.zip',
    'images': 'All images.zip',
}


@app.route('/api/convert', methods=['POST'])
def convert():
    data = request.get_json(force=True) or {}
    job = data.get('job')
    fmt = data.get('format')
    d = job_dir(job)
    st = load_state(d)
    src = work_pdf(d, st)
    t0 = time.time()
    imgs = os.path.join(d, 'imgs')
    os.makedirs(imgs, exist_ok=True)

    try:
        doc = pymupdf.open(src)
        if doc.needs_pass:
            doc.authenticate(st.get('password', ''))

        if fmt in ('xlsx_with', 'xlsx_without'):
            out = os.path.join(d, OUT_NAMES[fmt])
            info = tools.to_xlsx(doc, out, with_images=(fmt == 'xlsx_with'),
                                 img_dir=imgs, force=data.get('force', 'auto'))
            extra = {'engine_used': info.get('engine')}

        elif fmt in ('docx_editable', 'docx_exact'):
            out = os.path.join(d, OUT_NAMES[fmt])
            tools.to_docx(doc, out, mode=('exact' if fmt == 'docx_exact' else 'editable'),
                          img_dir=imgs, dpi=int(data.get('dpi', 130)))
            extra = {}

        elif fmt == 'jpg':
            out = os.path.join(d, OUT_NAMES[fmt])
            pages = tools.to_jpg_zip(doc, out, dpi=int(data.get('dpi', 150)),
                                     quality=int(data.get('quality', 88)))
            extra = {'pages_rendered': len(pages)}

        elif fmt == 'images':
            out = os.path.join(d, OUT_NAMES[fmt])
            rows = tools.images_zip(doc, out)
            if not rows:
                doc.close()
                return jsonify({'error': 'Is PDF me koi image nahi mili'}), 400
            extra = {'images': len(rows)}

        else:
            doc.close()
            return jsonify({'error': 'Unknown format'}), 400
        doc.close()
    except Exception as e:
        return jsonify({'error': f'Convert fail hua: {e}'}), 500

    return jsonify({'file': os.path.basename(out), 'label': OUT_NAMES[fmt],
                    'size': os.path.getsize(out), 'seconds': round(time.time() - t0, 1),
                    **extra})


# --------------------------------------------------------------------- security

@app.route('/api/secure', methods=['POST'])
def secure():
    data = request.get_json(force=True) or {}
    job = data.get('job')
    action = data.get('action')
    d = job_dir(job)
    st = load_state(d)

    try:
        if action == 'protect':
            pw = (data.get('password') or '').strip()
            if not pw:
                return jsonify({'error': 'Password likhein'}), 400
            out = os.path.join(d, 'Protected.pdf')
            info = tools.protect_pdf(work_pdf(d, st), out,
                                     user_pw=pw,
                                     owner_pw=(data.get('owner_password') or pw),
                                     allow=data.get('allow') or ['print', 'copy'],
                                     encrypt=data.get('encrypt', 'aes256'))
            return jsonify({'file': os.path.basename(out), 'label': 'Protected PDF',
                            'size': os.path.getsize(out), **info})

        if action == 'unprotect':
            pw = data.get('password', '') or st.get('password', '')
            out = os.path.join(d, 'Unlocked.pdf')
            info = tools.unprotect_pdf(os.path.join(d, st.get('source', 'source.pdf')), out,
                                       password=pw)
            # working copy bhi unlock karke rakh do
            doc = pymupdf.open(out)
            doc.save(work_pdf(d, st), garbage=4, deflate=True)
            st['pages'] = len(doc)
            doc.close()
            save_state(d, st)
            make_thumbs(d, st, force=True)
            return jsonify({'file': os.path.basename(out), 'label': 'Unlocked PDF (no password)',
                            'size': os.path.getsize(out), **info})
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        return jsonify({'error': f'Fail hua: {e}'}), 500
    return jsonify({'error': 'Unknown action'}), 400


# --------------------------------------------------------------------- edit

@app.route('/api/edit', methods=['POST'])
def edit():
    data = request.get_json(force=True) or {}
    job = data.get('job')
    ops = data.get('ops') or []
    d = job_dir(job)
    st = load_state(d)
    if not ops:
        return jsonify({'error': 'Koi change select nahi kiya'}), 400

    try:
        doc = pymupdf.open(work_pdf(d, st))
        if doc.needs_pass:
            doc.authenticate(st.get('password', ''))
        helpers = {}
        doc = tools.apply_ops(doc, ops, helpers)
        rev = st.get('rev', 0) + 1
        newp = os.path.join(d, f'work_{rev}.pdf')
        doc.save(newp, garbage=4, deflate=True)
        doc.close()
        if helpers.get('extract') is not None:
            ex = os.path.join(d, 'Extracted pages.pdf')
            helpers['extract'].save(ex, garbage=4, deflate=True)
            helpers['extract'].close()
        shutil.move(newp, work_pdf(d, st))
        st['pages'] = len(pymupdf.open(work_pdf(d, st)))
        save_state(d, st)
        make_thumbs(d, st, force=True)
    except Exception as e:
        return jsonify({'error': f'Edit fail hua: {e}'}), 500
    return jsonify(job_info(d, st, {'extracted': helpers.get('extract') is not None}))


@app.route('/api/edit/merge', methods=['POST'])
def edit_merge():
    f = request.files.get('file')
    job = request.form.get('job')
    position = request.form.get('position', 'end')
    if not f:
        return jsonify({'error': 'PDF chunein jise merge karna hai'}), 400
    d = job_dir(job)
    st = load_state(d)
    try:
        other = pymupdf.open(stream=f.read(), filetype='pdf')
        if other.needs_pass:
            other.close()
            return jsonify({'error': 'Merge wali PDF password protected hai'}), 400
        otherp = os.path.join(d, 'merge_src.pdf')
        other.save(otherp)
        other.close()
        doc = pymupdf.open(work_pdf(d, st))
        if doc.needs_pass:
            doc.authenticate(st.get('password', ''))
        doc = tools.apply_ops(doc, [{'op': 'merge', 'position': position,
                                     'other_doc': pymupdf.open(otherp)}])
        rev = st.get('rev', 0) + 1
        newp = os.path.join(d, f'work_{rev}.pdf')
        doc.save(newp, garbage=4, deflate=True)
        doc.close()
        shutil.move(newp, work_pdf(d, st))
        st['pages'] = len(pymupdf.open(work_pdf(d, st)))
        save_state(d, st)
        make_thumbs(d, st, force=True)
    except Exception as e:
        return jsonify({'error': f'Merge fail hua: {e}'}), 500
    return jsonify(job_info(d, st))


@app.route('/api/save', methods=['POST'])
def save_edited():
    """edited PDF ko final naam se save karke download link deta hai"""
    data = request.get_json(force=True) or {}
    job = data.get('job')
    d = job_dir(job)
    st = load_state(d)
    out = os.path.join(d, 'Edited.pdf')
    shutil.copy(work_pdf(d, st), out)
    return jsonify({'file': os.path.basename(out), 'label': 'Edited PDF',
                    'size': os.path.getsize(out)})


# --------------------------------------------------------------------- download

@app.route('/api/download/<job>/<path:name>')
def download(job, name):
    d = job_dir(job)
    # sirf job folder ke andar ki files
    safe = os.path.basename(name)
    p = os.path.join(d, safe)
    if not os.path.exists(p):
        abort(404)
    mime = {
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.zip': 'application/zip', '.pdf': 'application/pdf',
        '.jpg': 'image/jpeg', '.png': 'image/png',
    }.get(os.path.splitext(safe)[1].lower(), 'application/octet-stream')
    return send_file(p, as_attachment=True, download_name=safe, mimetype=mime)


@app.route('/api/files/<job>')
def list_files(job):
    d = job_dir(job)
    out = []
    for n in sorted(os.listdir(d)):
        p = os.path.join(d, n)
        if os.path.isfile(p) and n.lower().endswith(('.xlsx', '.docx', '.zip', '.pdf')):
            if n.startswith('work') or n in ('source.pdf', 'merge_src.pdf'):
                continue
            out.append({'file': n, 'size': os.path.getsize(p)})
    return jsonify(out)


@app.errorhandler(413)
def too_big(e):
    return jsonify({'error': 'File bahut badi hai (400 MB se kam rakhein)'}), 413


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    print(f'* PDF Toolbox running on http://0.0.0.0:{port}')
    app.run(host='0.0.0.0', port=port, threaded=True, debug=False)
