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


# ---------------------------------------------------------------------
#  Embedded UI copy (fallback — site works even if templates/ is missing)
# ---------------------------------------------------------------------
_INDEX_HTML = r"""<!DOCTYPE html>
<html lang="hi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PDF Toolbox — Convert · Secure · Edit</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  :root{--ink:#131a26;--muted:#6b7687;--brand:#1f6feb;--teal:#0ea5a4;--line:#e5e9f0;
        --ok:#16a34a;--warn:#d97706;--danger:#dc2626}
  body{font-family:'Segoe UI',Roboto,-apple-system,Arial,sans-serif;color:var(--ink);
       background:linear-gradient(160deg,#eef3fb,#e7eff9 45%,#eaf6f4);min-height:100vh;padding:22px 14px 70px}
  .wrap{max-width:1040px;margin:0 auto}
  header.top{text-align:center;margin-bottom:18px}
  .logo{display:inline-flex;align-items:center;gap:9px;padding:7px 15px;border-radius:999px;background:#fff;
        box-shadow:0 2px 10px rgba(20,40,80,.08);font-weight:700;font-size:13.5px;color:var(--brand)}
  header.top h1{font-size:26px;margin:12px 0 5px;letter-spacing:-.4px}
  header.top p{color:var(--muted);font-size:14px;line-height:1.55;max-width:700px;margin:0 auto}
  .card{background:#fff;border-radius:16px;box-shadow:0 6px 28px rgba(20,40,80,.09);padding:20px;margin-bottom:16px}
  .drop{border:2.5px dashed #c3d2e8;border-radius:14px;padding:30px 18px;text-align:center;cursor:pointer;
        background:#fafcff;transition:.18s}
  .drop:hover,.drop.over{border-color:var(--brand);background:#f2f7ff;transform:translateY(-1px)}
  .drop .ico{font-size:38px}
  .drop h3{font-size:16.5px;margin:9px 0 4px}
  .drop p{color:var(--muted);font-size:13px}
  input[type=file]{display:none}
  .btn{appearance:none;border:0;border-radius:11px;padding:11px 18px;font-size:14px;font-weight:600;cursor:pointer;
       font-family:inherit;transition:.15s;display:inline-flex;align-items:center;justify-content:center;gap:7px}
  .btn.primary{background:linear-gradient(135deg,var(--brand),#2f86ff);color:#fff;box-shadow:0 4px 14px rgba(31,111,235,.25)}
  .btn.primary:hover{transform:translateY(-1px)}
  .btn.ghost{background:#fff;border:1.5px solid var(--line)}
  .btn.ghost:hover{border-color:#b9cdea;background:#f7fafe}
  .btn.sm{padding:8px 13px;font-size:13px;border-radius:9px}
  .btn.danger{background:#fef2f2;color:var(--danger);border:1.5px solid #fecaca}
  .btn:disabled{opacity:.5;cursor:not-allowed}
  .row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
  .tabs{display:flex;gap:6px;background:#eef3fa;padding:5px;border-radius:12px;margin-bottom:16px;flex-wrap:wrap}
  .tab{flex:1 1 120px;text-align:center;padding:9px 12px;border-radius:9px;cursor:pointer;font-weight:600;font-size:13.5px;color:#44506a}
  .tab.on{background:#fff;color:var(--brand);box-shadow:0 2px 8px rgba(20,40,80,.10)}
  .pane{display:none}.pane.on{display:block}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:11px}
  .act{display:flex;align-items:center;gap:11px;padding:13px 14px;border-radius:12px;border:1.5px solid var(--line);
       background:#fff;cursor:pointer;text-align:left;font-family:inherit;transition:.15s;width:100%}
  .act:hover{transform:translateY(-1px);box-shadow:0 6px 16px rgba(20,40,80,.10)}
  .act .em{font-size:22px}.act .tx{flex:1}
  .act b{display:block;font-size:13.2px}.act small{color:var(--muted);font-size:11.4px}
  .act.blue{border-color:#bcd6ff;background:#f6faff}
  .act.grey{border-color:#e2e6ee;background:#fafbfd}
  .act.green{border-color:#bfe7d1;background:#f5fdf8}
  .act.violet{border-color:#d6cbf7;background:#faf8ff}
  .act.amber{border-color:#fde3b1;background:#fffbf2}
  .act.red{border-color:#fecaca;background:#fff7f7}
  .stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:9px;margin:12px 0}
  .stat{background:#f7fafd;border:1px solid var(--line);border-radius:10px;padding:9px 12px}
  .stat b{display:block;font-size:16.5px;letter-spacing:-.3px}
  .stat span{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.35px}
  .pill{display:inline-block;padding:3px 10px;border-radius:999px;font-size:11.3px;font-weight:700}
  .pill.ok{background:#e8f8ee;color:#15803d}.pill.info{background:#eaf2ff;color:#1d4ed8}
  .pill.warn{background:#fff7e6;color:#b45309}
  .thumbs{display:grid;grid-template-columns:repeat(auto-fill,minmax(96px,1fr));gap:9px;max-height:430px;overflow:auto;
          padding:10px;background:#f8fafd;border:1px solid var(--line);border-radius:12px;margin-top:12px}
  .th{position:relative;border:2px solid transparent;border-radius:9px;overflow:hidden;background:#fff;cursor:pointer;transition:.12s}
  .th:hover{border-color:#b9cdea}
  .th.sel{border-color:var(--brand);box-shadow:0 0 0 3px rgba(31,111,235,.15)}
  .th img{display:block;width:100%;height:auto}
  .th span{position:absolute;left:4px;bottom:4px;background:rgba(15,20,32,.78);color:#fff;font-size:10.5px;
           padding:1px 6px;border-radius:6px}
  label.f{display:block;font-size:12.4px;color:#3b4759;font-weight:600;margin:11px 0 5px}
  input[type=text],input[type=password],input[type=number],select{width:100%;padding:10px 12px;border:1.5px solid var(--line);
      border-radius:10px;font-size:13.5px;font-family:inherit;background:#fff}
  input:focus,select:focus{outline:none;border-color:var(--brand);box-shadow:0 0 0 3px rgba(31,111,235,.10)}
  .chk{display:flex;align-items:center;gap:8px;font-size:12.8px;color:#3b4759;margin:7px 0}
  .chk input{width:16px;height:16px;accent-color:var(--brand)}
  .cols2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  @media(max-width:640px){.cols2{grid-template-columns:1fr}}
  .msg{border-radius:11px;padding:11px 14px;font-size:13px;margin-top:11px;display:none}
  .msg.err{background:#fef2f2;color:#b91c1c;border:1px solid #fecaca;display:block}
  .msg.ok{background:#f0fdf4;color:#166534;border:1px solid #bbf7d0;display:block}
  .msg.info{background:#eff6ff;color:#1e40af;border:1px solid #bfdbfe;display:block}
  .bar{height:6px;background:#e8eef7;border-radius:99px;overflow:hidden;margin-top:12px;display:none}
  .bar i{display:block;height:100%;width:0;background:linear-gradient(90deg,var(--brand),var(--teal));transition:width .25s}
  .hide{display:none!important}
  .spin{width:15px;height:15px;border:2.5px solid rgba(255,255,255,.4);border-top-color:#fff;border-radius:50%;
        animation:sp .7s linear infinite}
  @keyframes sp{to{transform:rotate(360deg)}}
  .foot{text-align:center;color:#93a1b5;font-size:12px;margin-top:20px;line-height:1.6}
  .note{background:#f8fafd;border:1px dashed #cbd5e1;border-radius:10px;padding:10px 13px;font-size:12.3px;color:#54606f;margin-top:12px}
  .prev{margin-top:12px;border:1px solid var(--line);border-radius:12px;overflow:hidden;background:#f4f7fb;text-align:center}
  .prev img{max-width:100%;height:auto;display:block;margin:0 auto}
  .prevcap{padding:8px 12px;font-size:12px;color:var(--muted);background:#fbfdff;border-top:1px solid var(--line);text-align:left}
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <div class="logo">🧰 PDF Toolbox</div>
    <h1>Ek PDF → Excel, Word, JPG, Edit, Lock — sab kuch</h1>
    <p>Koi bhi PDF daaliye. Hum uska <b>format pehchaankar</b> same-to-same Excel banate hain, Word banate hain,
       JPG nikalte hain, images dete hain, password lagate/hataate hain aur PDF ko edit bhi kar sakte hain.</p>
  </header>

  <!-- ============ UPLOAD ============ -->
  <div class="card" id="upCard">
    <div class="drop" id="drop">
      <div class="ico">📄</div>
      <h3>PDF yahan drop karein ya click karein</h3>
      <p>PDF · ya image (JPG/PNG) bhi chalega — wo 1-page PDF ban jayegi</p>
    </div>
    <input type="file" id="file" accept="application/pdf,.pdf,image/*">
    <div id="pwBox" class="hide" style="margin-top:14px">
      <label class="f">🔒 Ye PDF password protected hai — password daalein</label>
      <div class="row">
        <input type="password" id="pw" placeholder="PDF password" style="flex:1;min-width:180px">
        <button class="btn primary" id="pwGo">Unlock karein</button>
      </div>
    </div>
    <div class="bar" id="bar"><i></i></div>
    <div class="msg" id="err"></div>
  </div>

  <!-- ============ RESULT / TOOLBOX ============ -->
  <div class="card hide" id="toolCard">
    <div class="row" style="justify-content:space-between">
      <div>
        <h3 id="fname" style="font-size:15.5px">—</h3>
        <div class="row" style="margin-top:6px;gap:7px">
          <span class="pill ok" id="pStat">✓ ready</span>
          <span class="pill info" id="pPages">—</span>
          <span class="pill warn hide" id="pEnc">🔒 protected</span>
        </div>
      </div>
      <button class="btn ghost sm" id="again">↺ Dusri PDF</button>
    </div>

    <div class="stats" id="stats"></div>

    <div class="tabs">
      <div class="tab on" data-tab="convert">🔄 Convert</div>
      <div class="tab" data-tab="secure">🔐 Protect / Unlock</div>
      <div class="tab" data-tab="edit">✏️ PDF Edit</div>
      <div class="tab" data-tab="preview">👁️ Preview</div>
    </div>

    <div class="msg" id="tmsg"></div>

    <!-- ---------- CONVERT ---------- -->
    <div class="pane on" id="pane-convert">
      <div class="grid">
        <button class="act blue" data-fmt="xlsx_with">
          <span class="em">🖼️</span><span class="tx"><b>Excel — images ke saath</b><small>PDF jaisa same-to-same layout + photos</small></span><span>⬇</span>
        </button>
        <button class="act grey" data-fmt="xlsx_without">
          <span class="em">🚫</span><span class="tx"><b>Excel — bina images</b><small>Sirf data, halki file</small></span><span>⬇</span>
        </button>
        <button class="act violet" data-fmt="docx_editable">
          <span class="em">📝</span><span class="tx"><b>Word file (editable)</b><small>Tables + text, aap edit kar sakte hain</small></span><span>⬇</span>
        </button>
        <button class="act violet" data-fmt="docx_exact">
          <span class="em">🖨️</span><span class="tx"><b>Word — bilkul same look</b><small>Har page PDF jaisa dikhega</small></span><span>⬇</span>
        </button>
        <button class="act amber" data-fmt="jpg">
          <span class="em">🖼️</span><span class="tx"><b>JPG (har page)</b><small>Saare pages JPG me, ZIP</small></span><span>⬇</span>
        </button>
        <button class="act green" data-fmt="images">
          <span class="em">📦</span><span class="tx"><b>All images only</b><small>PDF ke andar ki saari photos</small></span><span>⬇</span>
        </button>
      </div>

      <div class="cols2" style="margin-top:14px">
        <div>
          <label class="f">Excel/Word format engine</label>
          <select id="optForce">
            <option value="auto">Auto detect (recommended)</option>
            <option value="estimate">Estimate bill format (IIJS)</option>
            <option value="table">Normal table PDF</option>
            <option value="text">Simple text lines</option>
          </select>
        </div>
        <div>
          <label class="f">JPG / image quality</label>
          <select id="optDpi">
            <option value="100">Draft — 100 DPI (chhoti file)</option>
            <option value="150" selected>Normal — 150 DPI</option>
            <option value="220">Good — 220 DPI</option>
            <option value="300">Print — 300 DPI (badi file)</option>
          </select>
        </div>
      </div>
      <div class="note">💡 Excel me har number editable hota hai (rate badlein to format wahi rahega). Print setup bhi PDF jaisa set rehta hai.</div>
    </div>

    <!-- ---------- SECURE ---------- -->
    <div class="pane" id="pane-secure">
      <div class="cols2">
        <div>
          <h3 style="font-size:14.5px">🔒 PDF ko protect karein</h3>
          <label class="f">PDF password (khulne ke liye zaroori)</label>
          <input type="password" id="secPw" placeholder="e.g. 1234">
          <label class="f">Owner password (optional)</label>
          <input type="password" id="secOwner" placeholder="khali chhod dein to same rahega">
          <label class="f">Kya allow karna hai?</label>
          <label class="chk"><input type="checkbox" id="perPrint" checked> Printing</label>
          <label class="chk"><input type="checkbox" id="perCopy" checked> Text/photo copy karna</label>
          <label class="chk"><input type="checkbox" id="perMod"> Editing / modify</label>
          <label class="chk"><input type="checkbox" id="perAnno"> Notes / annotation</label>
          <button class="btn primary" id="doProtect" style="width:100%;margin-top:12px">🔒 Protect karein (AES-256)</button>
        </div>
        <div>
          <h3 style="font-size:14.5px">🔓 Password hatayein</h3>
          <label class="f">PDF ka password</label>
          <input type="password" id="unPw" placeholder="khali chhod dein agar password nahi hai">
          <button class="btn primary" id="doUnprotect" style="width:100%;margin-top:12px">🔓 Unprotect karein</button>
          <div class="note" style="margin-top:14px">
            Password sirf tabhi hat sakta hai jab aapko pata ho (ya file bina password khulti ho).
            Owner-password wali restrictions (copy/print band) aise hi hat jaati hain.
          </div>
        </div>
      </div>
    </div>

    <!-- ---------- EDIT ---------- -->
    <div class="pane" id="pane-edit">
      <div class="row" style="justify-content:space-between">
        <div class="row">
          <button class="btn ghost sm" id="selAll">Sab select</button>
          <button class="btn ghost sm" id="selNone">Select hatayein</button>
          <button class="btn ghost sm" id="selOdd">Odd</button>
          <button class="btn ghost sm" id="selEven">Even</button>
        </div>
        <div class="row">
          <span class="pill info" id="selInfo">0 selected</span>
        </div>
      </div>
      <div class="thumbs" id="thumbs"></div>

      <div class="grid" style="margin-top:14px">
        <button class="act grey" data-op="rotate" data-arg="-90"><span class="em">↺</span><span class="tx"><b>Left rotate</b><small>Selected pages (ya sab)</small></span></button>
        <button class="act grey" data-op="rotate" data-arg="90"><span class="em">↻</span><span class="tx"><b>Right rotate</b><small>Selected pages (ya sab)</small></span></button>
        <button class="act red" data-op="delete"><span class="em">🗑️</span><span class="tx"><b>Pages delete</b><small>Selected pages hata dein</small></span></button>
        <button class="act grey" data-op="duplicate"><span class="em">⧉</span><span class="tx"><b>Duplicate</b><small>Selected pages copy</small></span></button>
        <button class="act violet" data-op="extract"><span class="em">✂️</span><span class="tx"><b>Extract pages</b><small>Selected pages ki nayi PDF</small></span></button>
        <button class="act grey" data-op="reorder"><span class="em">🔀</span><span class="tx"><b>Reverse order</b><small>Page sequence ulta karein</small></span></button>
      </div>

      <div class="cols2" style="margin-top:16px">
        <div>
          <label class="f">💧 Watermark text (saari pages par)</label>
          <div class="row">
            <input type="text" id="wmText" placeholder="e.g. SAMPLE / AISHHPRA" style="flex:1;min-width:150px">
            <button class="btn ghost sm" id="wmGo">Lagaayein</button>
          </div>
          <label class="chk" style="margin-top:4px"><input type="checkbox" id="wmLight" checked> Halka (transparent) rakhein</label>

          <label class="f">🔢 Page numbers</label>
          <div class="row">
            <input type="text" id="pnFmt" value="Page {page} / {pages}" style="flex:1;min-width:150px">
            <button class="btn ghost sm" id="pnGo">Daalein</button>
          </div>

          <label class="f">📝 Header / Footer text</label>
          <div class="row">
            <input type="text" id="hfText" placeholder="e.g. Aishhpra Gems & Jewels" style="flex:1;min-width:150px">
            <button class="btn ghost sm" id="hfHead">Header</button>
            <button class="btn ghost sm" id="hfFoot">Footer</button>
          </div>

          <label class="f">➕ Blank pages jodein (selected page ke baad)</label>
          <div class="row">
            <input type="number" id="blankCount" value="1" min="1" max="20" style="width:88px">
            <button class="btn ghost sm" id="blankGo">Add karein</button>
          </div>
        </div>
        <div>
          <label class="f">🔗 Doosri PDF merge karein</label>
          <input type="file" id="mergeFile" accept=".pdf" style="display:block;padding:9px;border:1.5px dashed #c3d2e8;border-radius:10px;width:100%">
          <div class="row" style="margin-top:9px">
            <select id="mergePos" style="flex:1;min-width:130px">
              <option value="end">Uske baad (end me)</option>
              <option value="start">Uske pehle (start me)</option>
            </select>
            <button class="btn primary sm" id="mergeGo">Merge</button>
          </div>

          <div class="note" style="margin-top:14px">
            💡 Koi bhi change karte hi nayi PDF ban jaati hai — thumbnails refresh ho jaate hain.
            Neeche se <b>edited PDF download</b> kar lein.
          </div>
          <button class="btn primary" id="savePdf" style="width:100%;margin-top:12px">⬇️ Edited PDF download karein</button>
        </div>
      </div>
    </div>

    <!-- ---------- PREVIEW ---------- -->
    <div class="pane" id="pane-preview">
      <div class="prev"><img id="prevImg" alt="preview"><div class="prevcap" id="prevCap">Page 1</div></div>
    </div>

    <div id="filesBox" class="hide" style="margin-top:14px">
      <label class="f">📁 Is PDF se banayi gayi files</label>
      <div class="grid" id="files"></div>
    </div>
  </div>

  <div class="foot">
    PDF Toolbox — Convert · Secure · Edit &nbsp;|&nbsp; sab kuch aapke device par process hota hai<br>
    Excel me har item ka <b>Wt × Rate = Amount</b> aur total verify kiya jata hai.
  </div>
</div>

<script>
(function(){
  const $ = s => document.querySelector(s);
  const $$ = s => Array.from(document.querySelectorAll(s));
  let job = null, pages = 0, selected = new Set(), state = {};

  const fmtSize = b => !b ? '0 KB' : (b < 1048576 ? Math.round(b/1024)+' KB'
                     : (b/1048576).toFixed(2)+' MB');
  const err = $('#err'), tmsg = $('#tmsg');

  function show(el, txt, cls){ el.textContent = txt; el.className = 'msg ' + cls; }
  function hide(el){ el.className = 'msg'; el.style.display='none'; }
  function tInfo(t, cls='ok'){ show(tmsg, t, cls); if(cls==='ok') setTimeout(()=>hide(tmsg), 3500); }

  /* ---------------- upload ---------------- */
  const drop = $('#drop'), fileIn = $('#file');
  drop.addEventListener('click', () => fileIn.click());
  ['dragenter','dragover'].forEach(ev => drop.addEventListener(ev, e => {
    e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave','drop'].forEach(ev => drop.addEventListener(ev, e => {
    e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => { if(e.dataTransfer.files[0]) doUpload(e.dataTransfer.files[0]); });
  fileIn.addEventListener('change', e => { if(e.target.files[0]) doUpload(e.target.files[0]); });

  function doUpload(file, password){
    hide(err);
    const fd = new FormData();
    fd.append('file', file);
    if(password) fd.append('password', password);
    $('#bar').style.display='block'; $('#bar').firstElementChild.style.width='10%';
    const xhr = new XMLHttpRequest();
    xhr.open('POST','api/upload');
    xhr.upload.onprogress = e => { if(e.lengthComputable)
      $('#bar').firstElementChild.style.width = (10 + e.loaded/e.total*80) + '%'; };
    xhr.onload = () => {
      $('#bar').style.display='none'; $('#bar').firstElementChild.style.width='0';
      let r = {}; try { r = JSON.parse(xhr.responseText); } catch(e){}
      if(xhr.status !== 200 || r.error){ show(err, r.error || ('Error '+xhr.status), 'err'); return; }
      if(r.needs_password){ window._pending = file; $('#pwBox').classList.remove('hide'); return; }
      $('#pwBox').classList.add('hide');
      window._file = file;
      job = r.job; state = r;
      renderState(r);
      $('#upCard').classList.add('hide');
      $('#toolCard').classList.remove('hide');
      window.scrollTo({top:0,behavior:'smooth'});
    };
    xhr.onerror = () => { $('#bar').style.display='none'; show(err,'Network error','err'); };
    xhr.send(fd);
  }
  $('#pwGo').addEventListener('click', () => {
    const p = $('#pw').value;
    if(!p){ show(err,'Password likhein','err'); return; }
    doUpload(window._pending, p);
  });

  $('#again').addEventListener('click', () => {
    $('#toolCard').classList.add('hide'); $('#upCard').classList.remove('hide');
    job = null; pages = 0; selected.clear();
    fileIn.value=''; $('#pwBox').classList.add('hide');
    drop.querySelector('h3').textContent='PDF yahan drop karein ya click karein';
    window.scrollTo({top:0,behavior:'smooth'});
  });

  /* ---------------- state render ---------------- */
  function renderState(r){
    state = r; job = r.job; pages = r.pages;
    $('#fname').textContent = r.name;
    $('#pPages').textContent = r.pages + ' pages';
    $('#pEnc').classList.toggle('hide', !r.opened_with_password && !r.encrypted);
    const engMap = {estimate:'💎 Estimate bill format', generic:'📊 Table PDF', text:'📝 Text PDF'};
    $('#pStat').textContent = '✓ ' + (engMap[r.engine_guess] || 'Ready');
    const st = [
      ['Pages', r.pages], ['Size', fmtSize(r.size)],
      ['Tables', r.has_tables ? 'Yes' : 'No'],
      ['Orientation', r.landscape ? 'Landscape' : 'Portrait'],
      ['Version', 'v' + r.rev]
    ];
    $('#stats').innerHTML = st.map(([k,v]) => `<div class="stat"><b>${v}</b><span>${k}</span></div>`).join('');
    renderThumbs(); refreshPreview(); listFiles();
    if(selected.size > pages) { selected.clear(); }
  }

  function renderThumbs(){
    const box = $('#thumbs');
    box.innerHTML = '';
    for(let i=1;i<=state.thumbs;i++){
      const d = document.createElement('div');
      d.className = 'th' + (selected.has(i) ? ' sel' : '');
      d.dataset.n = i;
      d.innerHTML = `<img src="api/thumb/${job}/${i}?v=${state.rev}" loading="lazy"><span>${i}</span>`;
      d.onclick = () => {
        if(selected.has(i)) selected.delete(i); else selected.add(i);
        d.classList.toggle('sel'); updateSelInfo();
      };
      box.appendChild(d);
    }
    updateSelInfo();
  }

  function updateSelInfo(){
    $('#selInfo').textContent = selected.size ? (selected.size + ' selected') : 'koi page select nahi (sab par lagega)';
  }

  /* ---------------- tabs ---------------- */
  $$('.tab').forEach(t => t.onclick = () => {
    $$('.tab').forEach(x => x.classList.remove('on'));
    t.classList.add('on');
    $$('.pane').forEach(p => p.classList.remove('on'));
    $('#pane-' + t.dataset.tab).classList.add('on');
    if(t.dataset.tab === 'preview') refreshPreview();
  });

  /* ---------------- selection helpers ---------------- */
  const selPages = () => selected.size ? Array.from(selected).join(',') : 'all';
  $('#selAll').onclick = () => { for(let i=1;i<=state.thumbs;i++) selected.add(i); renderThumbs(); };
  $('#selNone').onclick = () => { selected.clear(); renderThumbs(); };
  $('#selOdd').onclick = () => { selected.clear(); for(let i=1;i<=state.thumbs;i+=2) selected.add(i); renderThumbs(); };
  $('#selEven').onclick = () => { selected.clear(); for(let i=2;i<=state.thumbs;i+=2) selected.add(i); renderThumbs(); };

  /* ---------------- convert ---------------- */
  $$('#pane-convert .act').forEach(b => b.onclick = async () => {
    const fmt = b.dataset.fmt;
    const label = b.querySelector('b').textContent;
    b.disabled = true; const old = b.innerHTML;
    b.innerHTML = '<span class="em">⏳</span><span class="tx"><b>Ban raha hai…</b><small>' + label + '</small></span>';
    try{
      const r = await post('api/convert', {job, format: fmt, force: $('#optForce').value,
                                           dpi: parseInt($('#optDpi').value)});
      if(r.error) throw new Error(r.error);
      tInfo('✓ ' + r.label + ' ready — ' + fmtSize(r.size) + ' (' + r.seconds + 's)');
      listFiles();
      download(r.file);
    }catch(e){ tInfo('❌ ' + e.message, 'err'); }
    b.disabled = false; b.innerHTML = old;
  });

  /* ---------------- secure ---------------- */
  $('#doProtect').onclick = async () => {
    const pw = $('#secPw').value;
    if(!pw){ tInfo('Password likhein', 'err'); return; }
    const allow = [];
    if($('#perPrint').checked) allow.push('print');
    if($('#perCopy').checked) allow.push('copy');
    if($('#perMod').checked) allow.push('modify');
    if($('#perAnno').checked) allow.push('annotate');
    try{
      const r = await post('api/secure', {job, action:'protect', password: pw,
          owner_password: $('#secOwner').value, allow});
      if(r.error) throw new Error(r.error);
      tInfo('🔒 Protected PDF ready — ' + fmtSize(r.size));
      listFiles(); download(r.file);
    }catch(e){ tInfo('❌ ' + e.message, 'err'); }
  };

  $('#doUnprotect').onclick = async () => {
    try{
      const r = await post('api/secure', {job, action:'unprotect', password: $('#unPw').value});
      if(r.error) throw new Error(r.error);
      tInfo('🔓 Unlocked PDF ready — ' + fmtSize(r.size));
      renderState(await post('api/info', {job}));
      listFiles(); download(r.file);
    }catch(e){ tInfo('❌ ' + e.message, 'err'); }
  };

  /* ---------------- edit ---------------- */
  async function runOps(ops, msg){
    try{
      const r = await post('api/edit', {job, ops});
      if(r.error) throw new Error(r.error);
      tInfo('✓ ' + msg);
      renderState(r);
    }catch(e){ tInfo('❌ ' + e.message, 'err'); }
  }

  $$('#pane-edit .act').forEach(b => b.onclick = () => {
    const op = b.dataset.op;
    const sel = Array.from(selected);
    if(op === 'rotate'){
      runOps([{op:'rotate', pages: selPages(), angle: parseInt(b.dataset.arg)}], 'Page rotate ho gaye');
    } else if(op === 'delete'){
      if(!sel.length){ tInfo('Pehle pages select karein', 'err'); return; }
      if(sel.length >= pages){ tInfo('Saare pages delete nahi kar sakte', 'err'); return; }
      runOps([{op:'delete', pages: selPages()}], sel.length + ' page delete ho gaye');
    } else if(op === 'duplicate'){
      if(!sel.length){ tInfo('Pehle pages select karein', 'err'); return; }
      runOps([{op:'duplicate', pages: selPages()}], 'Pages duplicate ho gaye');
    } else if(op === 'extract'){
      if(!sel.length){ tInfo('Pehle pages select karein', 'err'); return; }
      runOps([{op:'extract', pages: selPages()}], 'Extracted PDF ban gayi');
      setTimeout(listFiles, 700);
    } else if(op === 'reorder'){
      const ord = []; for(let i=pages;i>=1;i--) ord.push(i);
      runOps([{op:'reorder', order: ord}], 'Order ulta ho gaya');
    }
  });

  $('#wmGo').onclick = () => {
    const t = $('#wmText').value.trim();
    if(!t){ tInfo('Watermark text likhein', 'err'); return; }
    runOps([{op:'watermark', text:t, size:70, opacity: $('#wmLight').checked ? 0.18 : 0.5}],
           'Watermark lag gaya');
  };
  $('#pnGo').onclick = () => {
    runOps([{op:'page_numbers', format: $('#pnFmt').value || '{page}', position:'bottom-center'}],
           'Page numbers lag gaye');
  };
  $('#hfHead').onclick = () => {
    const t = $('#hfText').value.trim();
    if(!t){ tInfo('Text likhein', 'err'); return; }
    runOps([{op:'header', text:t}], 'Header lag gaya');
  };
  $('#hfFoot').onclick = () => {
    const t = $('#hfText').value.trim();
    if(!t){ tInfo('Text likhein', 'err'); return; }
    runOps([{op:'footer', text:t}], 'Footer lag gaya');
  };
  $('#blankGo').onclick = () => {
    const after = selected.size ? Math.max.apply(null, Array.from(selected)) : pages;
    runOps([{op:'insert_blank', after: after, count: parseInt($('#blankCount').value)||1}],
           'Blank page add ho gaya');
  };
  $('#mergeGo').onclick = () => {
    const f = $('#mergeFile').files[0];
    if(!f){ tInfo('PDF chunein jise merge karna hai', 'err'); return; }
    const fd = new FormData();
    fd.append('file', f); fd.append('job', job); fd.append('position', $('#mergePos').value);
    tInfo('Merge ho raha hai…', 'info');
    fetch('api/edit/merge', {method:'POST', body: fd})
      .then(r => r.json())
      .then(r => {
        if(r.error) throw new Error(r.error);
        tInfo('✓ Merge ho gaya — ab ' + r.pages + ' pages');
        renderState(r);
      }).catch(e => tInfo('❌ ' + e.message, 'err'));
  };
  $('#savePdf').onclick = async () => {
    try{
      const r = await post('api/save', {job});
      if(r.error) throw new Error(r.error);
      tInfo('✓ Edited PDF ready — ' + fmtSize(r.size));
      listFiles(); download(r.file);
    }catch(e){ tInfo('❌ ' + e.message, 'err'); }
  };

  /* ---------------- preview / files ---------------- */
  function refreshPreview(){
    if(!job) return;
    $('#prevImg').src = 'api/preview/' + job + '?v=' + Date.now();
    $('#prevCap').textContent = 'Page 1 preview — ' + (state.name || '');
  }
  async function listFiles(){
    if(!job) return;
    const r = await fetch('api/files/' + job).then(x => x.json()).catch(() => []);
    if(!r || !r.length){ $('#filesBox').classList.add('hide'); return; }
    $('#filesBox').classList.remove('hide');
    $('#files').innerHTML = r.map(f =>
      `<button class="act grey" onclick="location='api/download/${job}/${encodeURIComponent(f.file)}'">
         <span class="em">📄</span><span class="tx"><b>${f.file}</b><small>${fmtSize(f.size)} · click to download</small></span><span>⬇</span>
       </button>`).join('');
  }

  /* ---------------- helpers ---------------- */
  function post(url, body){
    return fetch(url, {method:'POST', headers:{'Content-Type':'application/json'},
                       body: JSON.stringify(body)}).then(r => r.json());
  }
  function download(file){
    const a = document.createElement('a');
    a.href = 'api/download/' + job + '/' + encodeURIComponent(file);
    a.download = file;
    document.body.appendChild(a); a.click(); a.remove();
  }
})();
</script>
</body>
</html>
"""


# --------------------------------------------------------------------- pages

@app.route('/')
def index():
    # Robust: templates/index.html -> root index.html -> embedded copy
    here = os.path.dirname(os.path.abspath(__file__))
    for rel in ('templates/index.html', 'index.html'):
        p = os.path.join(here, rel)
        if os.path.isfile(p):
            with open(p, encoding='utf-8') as f:
                return f.read()
    return _INDEX_HTML


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
