import sys
import webbrowser
import threading
import time
import uvicorn
from app.config import HOST, PORT

def open_browser():
    time.sleep(1.2)
    url = f"http://{HOST}:{PORT}"
    print(f"[*] Membuka PureDoc PDF di browser: {url}")
    webbrowser.open(url)

if __name__ == "__main__":
    print("=" * 60)
    print("  PureDoc PDF - Konverter Word ke PDF Tanpa Kompresi Gambar")
    print(f"  Server berjalan di http://{HOST}:{PORT}")
    print("=" * 60)
    
    # Launch browser automatically
    threading.Thread(target=open_browser, daemon=True).start()
    
    # Start server
    uvicorn.run("app.main:app", host=HOST, port=PORT, reload=False, log_level="info")
