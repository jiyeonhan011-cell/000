"""
ALPS(a.alpsys.co.kr) 출고 내역(작업내역) 자동 다운로드
사용자 컴퓨터에서 실행됨 — 사내 인트라넷 접근은 이 스크립트를 실행하는
PC 기준이므로, 앱이 설치된 PC에서 정상적으로 접속 가능해야 합니다.

화면 구조를 스크린샷으로만 확인했기 때문에, 실제 사이트의 HTML 구조와
완전히 일치하지 않을 수 있습니다. 처음 실행 시 오류가 나면 log 폴더의
alps_download_error.png 스크린샷을 보고 알려주세요 — 선택자를 바로 고칠 수 있습니다.
"""
import time
import datetime
import pathlib
import glob
import os

LOGIN_URL = "http://a.alpsys.co.kr/login/login.html"
SALES_URL = "http://a.alpsys.co.kr/sales.html"

BASE = pathlib.Path(__file__).parent
DOWNLOAD_DIR = BASE / "_alps_temp_download"
DEST_DIR = BASE / "검수파일" / "작업내역"
ERROR_SHOT = BASE / "alps_download_error.png"


def _make_driver():
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    opts = Options()
    opts.add_argument("--window-size=1400,1000")
    # 필요하면 헤드리스로 바꿀 수 있음 (opts.add_argument("--headless=new"))
    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    opts.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=opts)


def download_yesterday_workreport(alps_id, alps_pw, target_date=None):
    """
    ALPS에 로그인해서 전일자(또는 target_date) 작업일 기준 출고내역을
    엑셀로 받아 검수파일/작업내역 폴더에 저장.
    target_date: datetime.date, 생략시 오늘의 전날.
    반환값: 저장된 파일 경로 (str) 또는 실패 시 예외 발생.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC

    if target_date is None:
        target_date = datetime.date.today() - datetime.timedelta(days=1)
    date_str = target_date.strftime("%Y-%m-%d")

    try:
        driver = _make_driver()
    except Exception as e:
        raise RuntimeError(
            "Chrome 브라우저를 자동으로 띄우지 못했습니다. "
            "selenium을 최신 버전으로 업데이트해야 할 수 있습니다 "
            "(cmd에서 'py -m pip install --upgrade selenium' 실행 후 다시 시도). "
            f"원본 오류: {e}"
        ) from e
    wait = WebDriverWait(driver, 20)
    try:
        # 1) 로그인
        driver.get(LOGIN_URL)
        inputs = wait.until(lambda d: d.find_elements(By.CSS_SELECTOR, "input") or None)
        id_box = driver.find_element(By.CSS_SELECTOR, "input[type='text'], input:not([type])")
        pw_box = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        id_box.clear(); id_box.send_keys(alps_id)
        pw_box.clear(); pw_box.send_keys(alps_pw)
        login_btn = driver.find_element(
            By.XPATH, "//button[contains(text(),'로그인')] | //div[contains(text(),'로그인')] | //a[contains(text(),'로그인')]"
        )
        login_btn.click()
        wait.until(EC.url_contains("user_main"))

        # 2) 출고내역 화면으로 바로 이동 (세션 쿠키로 접근됨)
        driver.get(SALES_URL)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "select")))

        # 3) 배송일 → 작업일 드롭다운 변경
        # 페이지 안의 select들 중 "배송일/작업일/조회일" 옵션을 가진 것을 찾음
        target_select = None
        for sel_el in driver.find_elements(By.TAG_NAME, "select"):
            opt_texts = [o.text.strip() for o in sel_el.find_elements(By.TAG_NAME, "option")]
            if "작업일" in opt_texts and "배송일" in opt_texts:
                target_select = sel_el
                break
        if target_select is None:
            raise RuntimeError("배송일/작업일 드롭다운을 찾지 못했습니다.")
        Select(target_select).select_by_visible_text("작업일")
        time.sleep(0.5)

        # 4) 날짜 입력 — 화면의 date/text 입력 중 날짜 형식(YYYY-MM-DD)을 받는 첫 필드에 설정
        date_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='date'], input.date, input[name*='date' i]")
        if not date_inputs:
            # 대체: 값이 YYYY-MM-DD 패턴인 text input 찾기
            import re
            for el in driver.find_elements(By.CSS_SELECTOR, "input[type='text']"):
                v = el.get_attribute("value") or ""
                if re.match(r"\d{4}-\d{2}-\d{2}", v):
                    date_inputs = [el]
                    break
        if not date_inputs:
            raise RuntimeError("날짜 입력창을 찾지 못했습니다.")
        for el in date_inputs[:1]:
            driver.execute_script(
                "arguments[0].removeAttribute('readonly'); arguments[0].value = arguments[1];"
                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));",
                el, date_str,
            )

        # 5) 전체보기 클릭
        show_all_btn = driver.find_element(By.XPATH, "//*[contains(text(),'전체보기')]")
        show_all_btn.click()
        time.sleep(1.5)

        # 6) 버튼 텍스트가 "엑셀"로 바뀔 때까지 대기 후 클릭
        excel_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'엑셀')]"))
        )
        before_files = set(glob.glob(str(DOWNLOAD_DIR / "*")))
        excel_btn.click()

        # 7) 다운로드 완료 대기
        downloaded = None
        for _ in range(60):
            time.sleep(1)
            now_files = set(glob.glob(str(DOWNLOAD_DIR / "*")))
            new_files = [f for f in (now_files - before_files) if not f.endswith(".crdownload")]
            if new_files:
                downloaded = new_files[0]
                break
        if not downloaded:
            raise RuntimeError("엑셀 다운로드가 완료되지 않았습니다 (60초 초과).")

        # 8) 검수파일/작업내역 폴더로 이동
        DEST_DIR.mkdir(parents=True, exist_ok=True)
        dest_name = f"작업내역_{date_str}{pathlib.Path(downloaded).suffix}"
        dest_path = DEST_DIR / dest_name
        os.replace(downloaded, dest_path)
        return str(dest_path)

    except Exception:
        try:
            driver.save_screenshot(str(ERROR_SHOT))
        except Exception:
            pass
        raise
    finally:
        driver.quit()


if __name__ == "__main__":
    import sys
    aid = os.environ.get("ALPS_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
    apw = os.environ.get("ALPS_PW") or (sys.argv[2] if len(sys.argv) > 2 else None)
    if not aid or not apw:
        print("사용법: python alps_downloader.py <아이디> <비밀번호>")
        sys.exit(1)
    path = download_yesterday_workreport(aid, apw)
    print(f"저장됨: {path}")
