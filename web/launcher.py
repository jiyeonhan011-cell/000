"""
창고이동 검수 - 데스크탑 앱 런처
Streamlit을 백그라운드에서 실행하고 pywebview 창으로 띄움
"""
import sys, threading, time, subprocess
import webview

PORT = 8501
URL  = f"http://localhost:{PORT}"

def start_streamlit():
    subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", "app.py",
         "--server.port", str(PORT),
         "--server.headless", "true",
         "--browser.gatherUsageStats", "false",
         "--server.enableCORS", "false",
         "--server.enableXsrfProtection", "false"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

def wait_and_open():
    import urllib.request
    # Streamlit 준비될 때까지 대기 (최대 30초)
    for _ in range(60):
        try:
            urllib.request.urlopen(URL, timeout=1)
            break
        except:
            time.sleep(0.5)
    window.load_url(URL)

if __name__ == "__main__":
    # Streamlit 먼저 시작
    t = threading.Thread(target=start_streamlit, daemon=True)
    t.start()

    # pywebview 창 생성 (로딩 중 표시)
    window = webview.create_window(
        title="창고이동 검수 시스템",
        url="about:blank",
        width=1400,
        height=900,
        min_size=(900, 600),
        text_select=True,
    )

    # 준비되면 URL 로드
    threading.Thread(target=wait_and_open, daemon=True).start()

    webview.start()
