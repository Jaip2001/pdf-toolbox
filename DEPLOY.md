# 🚀 PDF Toolbox — Run & Deploy guide

Do tarah se chala sakte hain:
**A) apne computer par** (sabse aasan, offline chalta hai) · **B) internet par** (staff/clients ke liye, mobile se bhi)

---

# A) Apne computer par chalana (5 minute)

## Step 0 — Python install karein (ek hi baar)
- Windows/Mac: https://www.python.org/downloads/ se **Python 3.9 ya usse naya** install karein
- Windows par install karte waqt ye tick zaroor lagayein: **"Add Python to PATH"** ✅
- Check karne ke liye terminal/CMD kholein aur likhein: `python --version`

## Step 1 — App folder apne computer par le aayein
Workspace se poora **`pdf2excel`** folder download karein (ya uski ZIP file).
Maan lijiye aapne ise `C:\pdf2excel` (Windows) ya `~/pdf2excel` (Mac) me rakha hai.

## Step 2 — Chalu karein

**Windows:**
```
C:\pdf2excel folder kholein  →  run.bat par double-click
```
(pehli baar packages install honge — 1-2 minute lag sakte hain)

**Mac / Linux:** Terminal me —
```bash
cd ~/pdf2excel
./start.sh
```

**Ya manually (koi bhi OS):**
```bash
cd pdf2excel
pip install -r requirements.txt
python app.py
```

## Step 3 — Browser me kholein
```
http://localhost:8000
```
Bas! PDF daaliye aur convert/edit/protect karna shuru kariye.

> **Band karne ke liye:** black window / terminal me `Ctrl + C` dabayein.
> **Dobara chalane ke liye:** sirf Step 2 dobara karein (packages dobara install nahi honge).

### Agar problem aaye
| Problem | Solution |
|---|---|
| `python: command not found` | PATH me python nahi hai → Python dobara install karein, "Add to PATH" tick laga kar |
| `pip install` fail ho raha hai | `pip install --upgrade pip` chala kar dobara try karein |
| Port 8000 busy hai | `PORT=8080 python app.py` (Windows: `set PORT=8080` phir `python app.py`) |
| Table detect kamzor hai | `pip install pymupdf_layout` (optional, accuracy badhti hai) |
| Badi PDF par timeout | `gunicorn` / `waitress` wale start commands use karein (neeche diye hain), timeout 600s set hai |

### Word/JPG ke liye extra kuch nahi chahiye
Saare packages `requirements.txt` me hain: flask · pymupdf · openpyxl · python-docx · pillow

---

# B0) 🌐 Bina domain ke FREE public link (AppScript jaisa) — sabse pehle ye try karein

Aapke paas domain nahi hai? **Koi dikkat nahi.** Neeche diye tareekon se aapko ek
free `https://...` link mil jayega — bilkul Google AppScript ki tarah, bina kisi
account, bina kisi domain ke. Browser me khulega, sabke saath share kar sakte hain.

| Tarika | Link milega | Account chahiye? | Kab tak chalega |
|---|---|---|---|
| **1. Cloudflare Tunnel (tunnel.bat / tunnel.sh)** ⭐ sabse aasan | `https://xyz-abc.trycloudflare.com` | ❌ Nahi | Jab tak aapka PC + script chal raha hai |
| **2. ngrok** | `https://xyz.ngrok-free.app` | ✅ Free signup (authtoken) | Jab tak PC chal raha hai |
| **3. localtunnel** | `https://xyz.loca.lt` | ❌ Nahi | Jab tak PC chal raha hai |
| **4. Render free** (neeche Method 1) | `https://pdf-toolbox.onrender.com` | ✅ Free signup (GitHub) | Hamesha (15 min idle ke baad pehla load slow) |

> 💡 **AppScript jaisa permanent URL chahiye (PC band bhi ho jaye)?** → Neeche **Method 1 (Render)** use karein. Wahan `xyz.onrender.com` ek asli permanent free URL hai — domain khareedne ki zarurat nahi.

---

## Tarika 1 — Cloudflare Tunnel (sabse aasan, 100% free, no account) ⭐ RECOMMENDED

Ye ek chhota sa free tool hai jo aapke PC ke app ko internet par ek `https://` link de deta hai.

### Apne PC par (Windows):
```
1. PDF-Toolbox-App.zip extract karein
2. tunnel.bat par double-click karein
3. Ek https://....trycloudflare.com link dikhega — woh copy kar lein
4. Kisi bhi browser (phone/laptop) me woh link kholein — app chal raha hoga!
```

### Apne PC par (Mac / Linux):
```bash
cd pdf2excel
./tunnel.sh
# ek https://....trycloudflare.com link dikhega — browser me kholein
```

### Manually (agar samajhna ho to):
```bash
# terminal 1 — app chalayein
cd pdf2excel && python app.py

# terminal 2 — free tunnel banayein
# pehli baar cloudflared download karein:
#   Windows: https://github.com/cloudflare/cloudflared/releases/latest  (cloudflared-windows-amd64.exe)
#   Mac:     brew install cloudflared   (ya .tgz download karke extract)
#   Linux:   wget https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64
cloudflared tunnel --url http://localhost:8000
```
Terminal me ek link dikhega:
```
|  Your quick Tunnel has been created! Visit it at:  |
|  https://connecting-responding-hydrocodone-treo.trycloudflare.com   |
```
**Yahi aapki free public link hai.** ✅ Test kiya gaya hai — upload, convert, download sab is link se kaam karta hai.

> ⚠️ **Dhyan dein:** ye link har baar naya banta hai (restart karne par badal jata hai).
> Permanent link ke liye neeche **Method 1 (Render)** dekhein.

---

## Tarika 2 — ngrok (alternative)

```bash
# 1. https://ngrok.com se free account banayein, authtoken copy karein
# 2. Download karein: https://ngrok.com/download
ngrok config add-authtoken <aapka-token>

# 3. app chalayein (dusre terminal me):  python app.py
# 4. tunnel banayein:
ngrok http 8000
```
Link milega: `https://abcd-1234.ngrok-free.app` — browser me khol dein.

---

## Tarika 3 — localtunnel (sabse halka, no signup)

```bash
# Node.js hona chahiye (https://nodejs.org)
npx localtunnel --port 8000
```
Link milega: `https://xyz.loca.lt` + ek password (aapka public IP) dikhega — usko daal kar page khulega.

---

# B) Internet par deploy karna (site ban jayegi)

Bahut options hain. Neeche 4 aasan tarike — jo aapke kaam ke hisaab se chuniye:

| Platform | Kharcha | Sabse achha kiske liye | Difficulty |
|---|---|---|---|
| **1. Render** | Free (thoda slow) / $7/mo | Turant live, GitHub se auto-deploy | ⭐ Sabse aasan |
| **2. Railway** | ~$5/mo trial | Simple, tez | ⭐ Aasan |
| **3. VPS (DigitalOcean/Hetzner/AWS)** | $5-6/mo | Poora control, apna domain, unlimited | ⭐⭐⭐ |
| **4. Docker (kisi bhi server par)** | Server ka kharcha | Company ke apne server me | ⭐⭐ |
| **5. apna PC hi server** (LAN office) | Free | Sirf office ke andar use | ⭐ |

---

## Method 1 — Render (sabse aasan, free se shuru)

**Step 1:** Code ko GitHub par daalein
```bash
cd pdf2excel
git init
git add .
git commit -m "PDF Toolbox"
# GitHub par naya repo banayein, phir:
git remote add origin https://github.com/<aapka-naam>/pdf-toolbox.git
git branch -M main
git push -u origin main
```

**Step 2:** https://render.com par sign up karein (GitHub se login)

**Step 3:** `New +` → **Web Service** → apna repo chunein
Render `render.yaml` file khud padh leta hai. Agar manually bharein to:

| Field | Value |
|---|---|
| Runtime | Python 3 |
| Build Command | `pip install -r requirements.txt gunicorn` |
| Start Command | `gunicorn -c gunicorn_conf.py wsgi:app` |
| Instance Type | Free (ya Starter $7) |

**Step 4:** Environment variables (optional, free plan ke liye recommended):
```
WEB_CONCURRENCY = 1
WEB_THREADS = 8
```
**Step 5:** `Create Web Service` dabayein → 3-5 minute me live!
Aapka URL milega: `https://pdf-toolbox-xxxx.onrender.com` — yahi aapki website hai. 🎉

> ⚠️ **Free plan ki 2 baatein:** (1) 15 min idle ke baad app "so jaati hai" — pehla request 30-50 sec leta hai. (2) Free disks temporary hain — purani converted files restart par mit jaati hain (app khud bhi 25 jobs ke baad saaf kar deti hai, to koi dikkat nahi).

---

## Method 2 — Railway

**Step 1:** GitHub par code push karein (Method 1, Step 1)
**Step 2:** https://railway.app → **New Project** → **Deploy from GitHub repo**
**Step 3:** Repo chunein — Railway `railway.json` + `Procfile` khud use kar lega.
**Step 4:** **Settings → Networking → Generate Domain** dabayein → live URL mil jayega.
**Step 5 (optional):** Variables me `WEB_CONCURRENCY=1`, `WEB_THREADS=8` daal dein.

---

## Method 3 — VPS (DigitalOcean / Hetzner / AWS) — apna domain + unlimited

Yeh sabse professional setup hai. Ubuntu 22.04 server maan kar steps:

**Step 1 — Server taiyar karein**
```bash
ssh root@<aapka-server-ip>
apt update && apt install -y python3-pip python3-venv nginx git
```

**Step 2 — Code le aayein**
```bash
mkdir -p /opt/pdf-toolbox && cd /opt/pdf-toolbox
# GitHub se clone karein:
git clone https://github.com/<aapka-naam>/pdf-toolbox.git .
# (ya scp se apne computer se files bhej dein)
```

**Step 3 — Virtual env + packages**
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt gunicorn
```

**Step 4 — Systemd service banayein** (`/etc/systemd/system/pdf-toolbox.service`)
```ini
[Unit]
Description=PDF Toolbox
After=network.target

[Service]
User=www-data
WorkingDirectory=/opt/pdf-toolbox
Environment="PORT=8000"
Environment="WEB_CONCURRENCY=2"
Environment="WEB_THREADS=8"
ExecStart=/opt/pdf-toolbox/venv/bin/gunicorn -c gunicorn_conf.py wsgi:app
Restart=always

[Install]
WantedBy=multi-user.target
```
```bash
systemctl daemon-reload
systemctl enable --now pdf-toolbox
systemctl status pdf-toolbox      # chal raha hai?
```

**Step 5 — Nginx reverse proxy** (`/etc/nginx/sites-available/pdf-toolbox`)
```nginx
server {
    listen 80;
    server_name toolbox.aapkadomain.com;

    client_max_body_size 500M;      # badi PDF ke liye zaroori

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_read_timeout 900s;    # bade conversion ke liye
        proxy_send_timeout 900s;
    }
}
```
```bash
ln -s /etc/nginx/sites-available/pdf-toolbox /etc/nginx/sites-enabled/
nginx -t && systemctl reload nginx
```

**Step 6 — Free HTTPS (SSL)**
```bash
apt install -y certbot python3-certbot-nginx
certbot --nginx -d toolbox.aapkadomain.com
```
Bas — `https://toolbox.aapkadomain.com` live! 🔒

**Update karne ke liye:** `cd /opt/pdf-toolbox && git pull && systemctl restart pdf-toolbox`

---

## Method 4 — Docker (jahan bhi Docker ho)

Dockerfile ready hai. Build & run:
```bash
cd pdf2excel
docker build -t pdf-toolbox .
docker run -d --name pdf-toolbox -p 8080:8080 --restart unless-stopped \
  -e WEB_CONCURRENCY=2 -e WEB_THREADS=8 pdf-toolbox
```
Kholein: `http://localhost:8080`

**docker-compose se (recommended):**
```yaml
version: "3.9"
services:
  pdf-toolbox:
    build: .
    ports: ["8080:8080"]
    environment:
      WEB_CONCURRENCY: "2"
      WEB_THREADS: "8"
    restart: unless-stopped
```
```bash
docker compose up -d --build
```

---

## Method 5 — Office LAN me apne PC par (free, sabse practical)

Agar sirf office ke computers/mobiles se use karna hai:

1. Us PC par app chalayein (Method A).
2. PC ka local IP nikaalein:
   - Windows: CMD me `ipconfig` → "IPv4 Address" (jaise `192.168.1.7`)
   - Mac/Linux: `ifconfig | grep inet`
3. Server ko sab interfaces par chalayein (default already `0.0.0.0` hai):
   ```bash
   python app.py          # ya:  python -m waitress --host=0.0.0.0 --port=8000 wsgi:app
   ```
4. Baaki computers/mobiles browser me kholein: `http://192.168.1.7:8000`
5. Windows Firewall me port 8000 allow kar dein (pehli baar popup aayega → Allow).

---

## Production ke liye zaroori settings (already set hain)

| Setting | Value | Kahan |
|---|---|---|
| Bind address | `0.0.0.0` (all interfaces) | `app.py` / `gunicorn_conf.py` |
| Port | `PORT` env var (default 8000/8080) | dono |
| Worker timeout | **600 seconds** (bade PDF ke liye) | `gunicorn_conf.py` |
| Max upload | 400 MB | `app.py` |
| Job cleanup | last 25 jobs, auto-delete | `app.py` |
| File types | xlsx · docx · zip · pdf · jpg · png | app API |

**Production start commands:**
```bash
# Linux/Mac (recommended)
gunicorn -c gunicorn_conf.py wsgi:app

# Windows (gunicorn Windows par nahi chalta)
python wsgi.py            # waitress server use karta hai
```

---

## Customisation (aap khud badal sakte hain)

| Kya badalna hai | File | Line |
|---|---|---|
| Branding/title, rang, text | `templates/index.html` | top me `<title>` / `<h1>` / CSS `:root` |
| Max upload size | `app.py` | `MAX_CONTENT_LENGTH = 400 * 1024 * 1024` |
| Kitne purane jobs rakhein | `app.py` | `MAX_JOBS = 25` |
| Thumbnails ki limit | `app.py` | `MAX_THUMBS = 400` |
| Watermark default size/opacity | `tools.py` | `watermark` block |
| Excel ke column widths | `converter.py` | `build_estimate_xlsx` ke shuru me |

---

## Troubleshooting (deploy ke baad)

| Dikkat | Wajah / Fix |
|---|---|
| Page 502 / timeout | Badi PDF par worker timeout → `gunicorn_conf.py` me `timeout = 900` |
| Upload fail (413) | Nginx me `client_max_body_size 500M;` + `app.py` me `MAX_CONTENT_LENGTH` |
| Free Render par slow pehla load | 15 min idle ke baad "sleep" hota hai — paid plan ya VPS lijiye |
| Word/Excel file khulti nahi | Aadhi download hui file — dobara download karein; bade file par thoda wait |
| Hindi text gum ho gaya | Saare PDFs ka text hi aisa hai (us PDF me grayscale/matras), source PDF me hi nahi hota |
| `pymupdf` install fail (Windows) | `pip install --upgrade pip` phir `pip install pymupdf` |

---

Made with ❤️ for Aishhpra Gems & Jewels — sab kuch aapke server par process hota hai, koi data bahar nahi jaata.
