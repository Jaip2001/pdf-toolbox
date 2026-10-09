#!/bin/bash
# ============================================================
#  PDF Toolbox — FREE public URL (no domain needed)
#  Ye script aapke PC se ek free https link bana dega,
#  jaise Google AppScript deta hai. Koi account nahi chahiye.
#
#  Use:   ./tunnel.sh
# ============================================================
cd "$(dirname "$0")"

echo "=============================================="
echo "  PDF Toolbox — Free public link ban rahi hai"
echo "=============================================="

# --- Step 1: packages check ---
if ! python3 -c "import flask, pymupdf, openpyxl, docx, PIL" 2>/dev/null; then
  echo "[*] Pehli baar: packages install ho rahe hain (1-2 min)..."
  pip3 install -r requirements.txt || pip install -r requirements.txt
fi

# --- Step 2: cloudflared download (free tunnel tool, no account) ---
if [ ! -f ./cloudflared ]; then
  echo "[*] cloudflared download ho raha hai..."
  OS=$(uname -s)
  case "$OS" in
    Darwin)  # Mac
      curl -sL -o /tmp/cloudflared.tgz "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-darwin-amd64.tgz"
      tar -xzf /tmp/cloudflared.tgz -C . cloudflared && rm /tmp/cloudflared.tgz
      ;;
    Linux)
      curl -sL -o cloudflared "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64"
      ;;
    *)
      echo "[!] Mac/Linux hi supported hai. Windows ke liye tunnel.bat use karein."
      exit 1
      ;;
  esac
  chmod +x ./cloudflared
fi

# --- Step 3: app chalu karein (background me) ---
echo "[*] App chalu ho rahi hai..."
python3 app.py > /tmp/pdf-toolbox.log 2>&1 &
APP_PID=$!
sleep 5

# --- Step 4: free tunnel — ye URL dikhayega ---
echo ""
echo "=============================================="
echo "  Ab niche jo https://....trycloudflare.com"
echo "  link dikhega — woh aapki FREE public link hai!"
echo "  Browser me kholein, ya kisi ko bhi bhej dein."
echo "  (Band karne ke liye Ctrl+C)"
echo "=============================================="
echo ""
./cloudflared tunnel --url http://localhost:8000 --no-autoupdate

# cleanup jab band karein
kill $APP_PID 2>/dev/null
