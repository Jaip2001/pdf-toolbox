# 🧰 PDF Toolbox — Convert · Secure · Edit

Ek hi simple web app me **kal ka kuch bhi PDF** kaam ho jata hai:
PDF upload karo → same format me Excel / Word / JPG milega, images milengi,
password laga/haata sakte ho, aur PDF ko edit bhi kar sakte ho.

---

## 1) 🔄 Convert (koi bhi PDF — format khud pehchanta hai)

| Option | Kya milega | Kaha kaam aayega |
|---|---|---|
| 🖼️ **Excel — images ke saath** | PDF jaisa **same-to-same** layout + embedded photos | customer ko bhejne / record ke liye |
| 🚫 **Excel — bina images** | Sirf data, halki file | fast edit, WhatsApp, email |
| 📝 **Word file (editable)** | Asli Word tables + text — aap type kar sakte hain | bill me changes karne ke liye |
| 🖨️ **Word — bilkul same look** | Har page PDF jaisa dikhta hai | print / forward karne ke liye |
| 🖼️ **JPG (har page)** | Saare pages JPG me, ZIP | photo ki tarah share karne ke liye |
| 📦 **All images only** | PDF ke andar ki saari photos + `images_list.csv` | design catalogue banane ke liye |

**Format engine** (auto detect):
- **Estimate bill** (IIJS / jewellery: Sr · Design · Diamond · Metal · Stone & Misc · Other · Labour · Total)
  → bilkul same columns, image box, subtotal row aur SUMMARY box bhi same.
- **Table PDF** → kisi bhi ruled-table PDF ka structure jaisa hai waisa hi.
- **Text PDF** → table na ho to bhi x-position ke hisaab se columns ban jaate hain.

Har conversion ke baad data **verify** hota hai: `Wt × Rate = Amount`, `Net Wt × Labour = Labour`, `grand total`.

---

## 2) 🔐 Secure

- **Protect PDF** — apna password set karein (AES-256). Chuniye kya allow karna hai:
  printing · copy · editing · notes.
- **Unprotect PDF** — password daal kar lock hataayein
  (owner-password restrictions — copy/print band — bhi hat jaati hain).

## 3) ✏️ PDF Edit

Page thumbnails me se pages select karein aur:

- ↺ ↻ **Rotate** (selected or all)
- 🗑️ **Delete pages** · ⧉ **Duplicate** · ✂️ **Extract** (nayi PDF) · 🔀 **Reverse order**
- 💧 **Watermark** (transparent/rotated)
- 🔢 **Page numbers** (`Page {page} / {pages}`)
- 📝 **Header / Footer text** (har page par)
- ➕ **Blank page** insert
- 🔗 **Doosri PDF merge** (start ya end me)
- ⬇️ **Edited PDF download**

Har change ke baad nayi PDF ban jaati hai aur thumbnails refresh hote hain — original safe rehta hai.

---

## Chalane ka tarika (apne laptop / PC par)

Python 3.9+ chahiye.

```bash
cd pdf2excel
pip install -r requirements.txt
python app.py
```

Browser me kholein: **http://localhost:8000**

- **Windows** → `run.bat` par double-click
- **Mac / Linux** → `./start.sh`

Requirements:
```
flask · pymupdf · openpyxl · python-docx · pillow
```

> Tip: `pip install pymupdf_layout` karne se table detection aur bhi accurate ho jaati hai (optional).

---

## Files

```
pdf2excel/
├── app.py            # Flask server + saare API
├── converter.py      # PDF → Excel engine (estimate ka exact replica + table engine)
├── tools.py          # Word · JPG · images · protect/unprotect · edit operations
├── templates/
│   └── index.html    # UI (4 tabs: Convert · Protect/Unlock · Edit · Preview)
├── requirements.txt
├── run.bat / start.sh
└── jobs/             # har upload ka temporary folder (auto-purge, last 25 jobs)
```

## API (integration ke liye)

```
POST /api/upload                 file=<pdf|image> [password=…]  → job info (ya needs_password)
POST /api/unlock                 {job, password}
POST /api/info                   {job}
GET  /api/thumb/<job>/<n>        page thumbnail (jpg)
GET  /api/preview/<job>          page 1 preview (png)

POST /api/convert                {job, format: xlsx_with|xlsx_without|docx_editable|docx_exact|jpg|images,
                                  force: auto|estimate|table|text, dpi}
POST /api/secure                 {job, action: protect|unprotect, password, owner_password, allow:[…]}
POST /api/edit                   {job, ops:[{op:rotate|delete|duplicate|extract|reorder|insert_blank|
                                   watermark|page_numbers|header|footer|merge, …}]}
POST /api/edit/merge             file=<pdf> job=… position=end|start
POST /api/save                   {job}                       → Edited.pdf
GET  /api/files/<job>            is job se bani saari files
GET  /api/download/<job>/<file>  file download
```

## Notes

- Sab kuch aapke computer/server par process hota hai; koi data bahar nahi jaata.
- Jobs temporary hote hain (last 25 auto-delete).
- Excel me **shrink-to-fit** format hai — rate badalne par `###` nahi aayega; print setup A4 landscape me set rehta hai.

---

## 🚀 Chalana / Deploy karna

Pura step-by-step guide: **`DEPLOY.md`** file me hai (local run + Render + Railway + VPS + Docker + office LAN).

**Sabse short version:**
```bash
# local (apne computer par)
pip install -r requirements.txt
python app.py          # → http://localhost:8000
```
```bash
# production (server par)
gunicorn -c gunicorn_conf.py wsgi:app     # Linux/Mac
python wsgi.py                            # Windows (waitress)
```

Ready-made files already included: `Dockerfile` · `Procfile` · `render.yaml` · `railway.json` · `gunicorn_conf.py` · `wsgi.py` · `run.bat` · `start.sh`
