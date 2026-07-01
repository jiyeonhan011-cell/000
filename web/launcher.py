import sys, os, threading, time, subprocess, pathlib
import webview

PORT = 8501
URL  = f"http://localhost:{PORT}"
BASE = pathlib.Path(__file__).parent
LOG  = BASE / "launcher.log"

def find_python():
    exe = sys.executable
    # Windows Store py.exe → 실제 python.exe 경로 찾기
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
        lf.write(f"Python: {python}\n")
        lf.write(f"CWD: {BASE}\n")
        proc = subprocess.Popen(
            [python, "-m", "streamlit", "run", str(BASE / "app.py"),
             "--server.port", str(PORT),
             "--server.headless", "true",
             "--browser.gatherUsageStats", "false",
             "--server.enableCORS", "false",
             "--server.enableXsrfProtection", "false"],
            stdout=lf,
            stderr=lf,
            cwd=str(BASE),
        )
        proc.wait()

def wait_and_open():
    import urllib.request
    for _ in range(120):   # 최대 60초 대기
        try:
            urllib.request.urlopen(URL, timeout=1)
            window.load_url(URL)
            return
        except Exception:
            time.sleep(0.5)
    # 60초 지나도 안 되면 로그 내용을 창에 표시
    try:
        log = LOG.read_text(encoding="utf-8")
    except Exception:
        log = "로그 없음"
    window.load_html(f"<pre style='color:red;font-size:14px'>Streamlit 시작 실패:\n\n{log}</pre>")

if __name__ == "__main__":
    os.chdir(str(BASE))
    threading.Thread(target=start_streamlit, daemon=True).start()

    window = webview.create_window(
        title="창고이동 검수 시스템",
        url="about:blank",
        width=1400,
        height=900,
        min_size=(900, 600),
        text_select=True,
    )

    threading.Thread(target=wait_and_open, daemon=True).start()
    webview.start()
