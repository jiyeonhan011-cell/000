import sys, os, threading, time, subprocess, pathlib
import webview

PORT = 8501
URL  = f"http://localhost:{PORT}"
BASE = pathlib.Path(__file__).parent
LOG  = BASE / "launcher.log"

LOADING_HTML = """<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body {
    background: #111827;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    height: 100vh; font-family: 'Malgun Gothic', sans-serif; color: white;
  }
  .icon { font-size: 64px; margin-bottom: 24px; }
  h1 { font-size: 24px; font-weight: 700; margin-bottom: 8px; }
  p  { font-size: 14px; color: #9ca3af; margin-bottom: 36px; }
  .bar-wrap {
    width: 280px; height: 6px;
    background: #1f2937; border-radius: 9999px; overflow: hidden;
  }
  .bar {
    height: 100%; width: 0%;
    background: linear-gradient(90deg, #f97316, #fb923c);
    border-radius: 9999px;
    animation: fill 8s ease-in-out forwards;
  }
  @keyframes fill { 0%{width:0%} 80%{width:88%} 100%{width:92%} }
</style></head>
<body>
  <div class="icon">🏭</div>
  <h1>창고이동 검수 시스템</h1>
  <p>잠시만 기다려주세요...</p>
  <div class="bar-wrap"><div class="bar"></div></div>
</body></html>"""

def find_python():
    exe = sys.executable
    if exe.lower().endswith("py.exe"):
        try:
            r = subprocess.run([exe, "-c", "import sys;print(sys.executable)"],
                               capture_output=True, text=True)
            real = r.stdout.strip()
            if real and pathlib.Path(real).exists():
                return real
        except Exception:
            pass
    return exe

def start_streamlit():
    python = find_python()
    with open(LOG, "w", encoding="utf-8") as lf:
        lf.write(f"Python: {python}\nCWD: {BASE}\n")
        proc = subprocess.Popen(
            [python, "-m", "streamlit", "run", str(BASE / "app.py"),
             "--server.port", str(PORT),
             "--server.headless", "true",
             "--browser.gatherUsageStats", "false",
             "--server.enableCORS", "false",
             "--server.enableXsrfProtection", "false"],
            stdout=lf, stderr=lf, cwd=str(BASE),
        )
        proc.wait()

def wait_and_open():
    import urllib.request
    for _ in range(120):
        try:
            urllib.request.urlopen(URL, timeout=1)
            window.load_url(URL)
            return
        except Exception:
            time.sleep(0.5)
    try:
        log = LOG.read_text(encoding="utf-8")
    except Exception:
        log = "로그 없음"
    window.load_html(f"<pre style='color:red;background:#111;padding:20px;font-size:13px'>Streamlit 시작 실패:\n\n{log}</pre>")

if __name__ == "__main__":
    os.chdir(str(BASE))
    threading.Thread(target=start_streamlit, daemon=True).start()

    window = webview.create_window(
        title="창고이동 검수 시스템",
        html=LOADING_HTML,
        width=1400, height=900,
        min_size=(900, 600),
        text_select=True,
    )

    threading.Thread(target=wait_and_open, daemon=True).start()
    webview.start()
