# -*- coding: utf-8 -*-
"""
panda_dm_sender.py
- PandaLive 자동 쪽지 발송 (상태파일 실시간 갱신 + 5명마다 '줄 끝 스페이스' 변형)
- 필요 파일:
    .env (PANDA_ID=..., PANDA_PW=...)
    recipients_preview.csv (열: '후원아이디' [필수], '후원하트'[선택])
    message.txt (기본 메시지, 여러 줄 가능)
- 실행 예:
    python panda_dm_sender.py --headless --status-file send_status.json --reset
"""

import os, sys, time, json, argparse
from pathlib import Path
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import random

LOGIN_URL = "https://www.pandalive.co.kr/my/post/received"


# ===================== 공통 유틸 =====================

# ----- 전각 공백(U+3000) 사용 -----
FULLWIDTH_SPACE = "\u3000"  # 한글 IME에서 'ㄱ + 한자 + 1'로 입력되는 전각 스페이스
from selenium.common.exceptions import TimeoutException

def wait_result_h3_text(wait: WebDriverWait, timeout: float = 6.0) -> str:
    """
    '전송하겠습니까' 확인을 누른 뒤 나타나는 결과 모달의 <h3> 텍스트를 기다려서 반환.
    - 성공/실패/제한 등의 메시지가 모두 동일한 h3 클래스를 사용하므로 해당 요소를 명시적으로 대기.
    - 없으면 "" 반환.
    """
    driver = wait._driver
    XPATHS = [
        # 모달 내부 h3 (가장 정확)
        "//div[@role='dialog']//h3[contains(@class,'text-center')]",
        # 혹시 role 속성이 없는 케이스 대비
        "//h3[contains(@class,'text-center')]",
    ]
    end = time.time() + timeout
    while time.time() < end:
        for xp in XPATHS:
            try:
                elems = driver.find_elements(By.XPATH, xp)
                for el in elems:
                    s = (el.text or "").strip()
                    if s:
                        return s
            except Exception:
                pass
        time.sleep(0.15)
    return ""


def contains_any(s: str, needles):
    s_norm = " ".join(s.split())  # 공백 정규화
    return any(n in s_norm for n in needles)

def now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def load_status(path: Path) -> dict:
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"items": [], "meta": {}}


def save_status(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ----- 5명마다 '줄 끝 스페이스' 규칙 -----
def msg_with_line_end_spaces(base_message: str, send_index: int) -> str:
    """
    5명 단위로 특정 줄 '끝'에 전각 공백(U+3000)을 추가.
      - block = send_index // 5
      - target_line = block % line_count
      - spaces = 1 + (block // line_count)
    """
    lines = base_message.split("\n")  # 끝 공백 보존
    if not lines:
        return base_message

    line_count = len(lines)
    block = send_index // 5
    target_line = block % line_count
    spaces = 1 + (block // line_count)

    lines[target_line] = lines[target_line] + (FULLWIDTH_SPACE * spaces)
    out = "\n".join(lines)
    return out[:500]  # 팬더 최대 500자



# ===================== 셀레니움 유틸 =====================
def short_wait_click(wait: WebDriverWait, xpath: str, timeout: float = 1.2) -> bool:
    """짧게 기다렸다가 클릭. 실패 시 False."""
    try:
        elem = WebDriverWait(wait._driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        elem.click()
        return True
    except Exception:
        return False


def short_wait_present(wait: WebDriverWait, xpath: str, timeout: float = 1.2):
    """짧은 대기 내 존재 확인. 없으면 None."""
    try:
        return WebDriverWait(wait._driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
    except Exception:
        return None


def click_any_ok(wait: WebDriverWait, tries: int = 2, timeout_each: float = 1.0) -> None:
    """
    페이지에 떠 있는 일반 '확인' 모달/다이얼로그를 최대 tries회 닫는다.
    (성공/실패 알림, 비밀번호 변경 알림 등 동일 텍스트 처리)
    """
    for _ in range(tries):
        clicked = short_wait_click(wait, "//button[normalize-space()='확인']", timeout_each)
        if not clicked:
            clicked = short_wait_click(wait, "//div[@role='dialog']//button[normalize-space()='확인']", timeout_each)
        if not clicked:
            break
        time.sleep(0.2)


def login_and_open_compose(driver, wait, uid, pw):
    # 1) 접속
    driver.get(LOGIN_URL)

    # 2) 로그인 탭 클릭(회원가입이 기본일 수 있음)
    short_wait_click(wait, "//button[@role='tab']//p[normalize-space()='로그인']", timeout=3.0)

    # 3) ID/PW 입력
    id_box = short_wait_present(wait, "//*[@id='id' or @name='id']", timeout=5.0)
    if not id_box:
        raise RuntimeError("ID 입력창을 찾지 못했습니다.")
    id_box.clear()
    id_box.send_keys(uid)

    pw_box = short_wait_present(wait, "//input[@name='pw']", timeout=4.0)
    if not pw_box:
        raise RuntimeError("PW 입력창을 찾지 못했습니다.")
    pw_box.clear()
    pw_box.send_keys(pw)
    pw_box.send_keys(Keys.RETURN)

    # 4) 로그인 완료 신호: '쪽지쓰기' 버튼 등장
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='쪽지쓰기']"))
    )

    # 로그인 직후 떠 있을 수 있는 안내/확인 모달 처리
    click_any_ok(wait, tries=3, timeout_each=1.0)

    # 5) 쪽지쓰기 클릭
    short_wait_click(wait, "//button[normalize-space()='쪽지쓰기']", timeout=3.0)

    # 6) 모달의 입력창 확인
    WebDriverWait(driver, 8).until(
        EC.presence_of_element_located((By.XPATH, "//input[@placeholder='받는회원 ID']"))
    )
    WebDriverWait(driver, 8).until(
        EC.presence_of_element_located((By.XPATH, "//textarea[@placeholder='쪽지내용을 입력하세요.']"))
    )


def ensure_compose_open(driver, wait):
    """모달이 닫혔으면 다시 '쪽지쓰기'를 눌러 연다."""
    id_box = short_wait_present(wait, "//input[@placeholder='받는회원 ID']", timeout=0.6)
    msg_box = short_wait_present(wait, "//textarea[@placeholder='쪽지내용을 입력하세요.']", timeout=0.6)
    if id_box and msg_box:
        return
    short_wait_click(wait, "//button[normalize-space()='쪽지쓰기']", timeout=2.0)
    WebDriverWait(driver, 4).until(
        EC.presence_of_element_located((By.XPATH, "//input[@placeholder='받는회원 ID']"))
    )
    WebDriverWait(driver, 4).until(
        EC.presence_of_element_located((By.XPATH, "//textarea[@placeholder='쪽지내용을 입력하세요.']"))
    )


def send_one(wait: WebDriverWait, target_id: str, message: str) -> bool:
    """
    1명 전송: 받는회원 ID, 본문 입력 → [보내기] → '전송하겠습니까?' [확인] →
    결과 모달의 <h3> 텍스트로 성공/실패 판정 → '확인' 닫기 → compose 유지.
    """
    driver = wait._driver
    ensure_compose_open(driver, wait)

    # 받는회원 ID
    to_box = short_wait_present(wait, "//input[@placeholder='받는회원 ID']", timeout=2.0)
    if not to_box:
        return False
    to_box.send_keys(Keys.CONTROL, "a"); to_box.send_keys(Keys.DELETE); to_box.send_keys(target_id)

    # 본문
    msg_box = short_wait_present(wait, "//textarea[@placeholder='쪽지내용을 입력하세요.']", timeout=2.0)
    if not msg_box:
        return False
    msg_box.send_keys(Keys.CONTROL, "a"); msg_box.send_keys(Keys.DELETE); msg_box.send_keys(message)

    # 보내기 버튼
    if not short_wait_click(wait, "//button[normalize-space()='보내기']", timeout=2.5):
        # 모달이 닫혔거나 렌더 상태 꼬였을 때 1회 재시도
        ensure_compose_open(driver, wait)
        if not short_wait_click(wait, "//button[normalize-space()='보내기']", timeout=2.5):
            return False

    # '전송하겠습니까?' 확인
    short_wait_click(wait, "//button[normalize-space()='확인']", timeout=2.0)

    # 결과 모달의 <h3> 텍스트 대기/판독
    txt = wait_result_h3_text(wait, timeout=6.0)
    # 판별 키워드(정확 매칭/부분매칭 혼용)
    OK_TEXT = "성공적으로 보냈습니다."
    FAIL_VIP = "VIP 등급 아래인 경우 쪽지 수신을 제한"         # 1) 수신제한
    FAIL_RATE = "쪽지 전송이 제한되었습니다"                    # 2) 전송제한(동일문구/속도제한 등)

    ok = None
    s_norm = " ".join(txt.split())
    if txt == OK_TEXT or OK_TEXT in s_norm:
        ok = True
    elif (FAIL_VIP in s_norm) or (FAIL_RATE in s_norm):
        ok = False
    else:
        # 혹시 다른 안내문이거나 텍스트 못 잡은 경우 보수적으로 실패 처리
        ok = False

    # 남은 '확인' 닫기 (성공/실패 안내 모달)
    click_any_ok(wait, tries=3, timeout_each=1.2)

    # 다음 대상 대비 compose 열어두기
    try:
        ensure_compose_open(driver, wait)
    except Exception:
        pass

    return bool(ok)


# ===================== 메인 =====================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--status-file", type=str, default=str(Path(__file__).with_name("send_status.json")))
    ap.add_argument("--reset", action="store_true")
    ap.add_argument("--start", type=int, default=0)   # 시작 인덱스 (0-base)
    ap.add_argument("--limit", type=int, default=0)   # 최대 전송 수 (0=전체)
    args = ap.parse_args()

    base = Path(__file__).parent
    recipients_csv = base / "recipients_preview.csv"
    message_txt     = base / "message.txt"
    env_file        = base / ".env"
    status_path     = Path(args.status_file)

    if not recipients_csv.exists():
        print("recipients_preview.csv 없음"); sys.exit(1)
    if not message_txt.exists():
        print("message.txt 없음"); sys.exit(1)

    df = pd.read_csv(recipients_csv)
    if "후원아이디" not in df.columns:
        print("CSV에 '후원아이디' 열 없음"); sys.exit(1)

    base_message = Path(message_txt).read_text(encoding="utf-8")

    # 상태파일 초기화/로드
    st = load_status(status_path)
    need_reinit = args.reset or (not st.get("items")) or (len(st.get("items", [])) != len(df))
    if need_reinit:
        st = {"items": [], "meta": {"created": now_ts()}}
        for i, row in df.iterrows():
            st["items"].append({
                "index": int(i),
                "id": str(row["후원아이디"]),
                "hearts": int(row.get("후원하트", 0)) if "후원하트" in df.columns else 0,
                "status": "pending",
                "updated": now_ts()
            })
        save_status(status_path, st)
        print(f"[init] status 초기화: {len(st['items'])}건")

    # 로그인 정보
    load_dotenv(env_file)
    uid = os.getenv("PANDA_ID", "")
    pw  = os.getenv("PANDA_PW", "")
    if not uid or not pw:
        print(".env에 PANDA_ID/PANDA_PW 필요"); sys.exit(1)

    # 브라우저 옵션
    opts = Options()
    if args.headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--lang=ko-KR")
    opts.add_argument("--start-maximized")
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "profile.password_manager_leak_detection": False,
        "profile.default_content_setting_values.notifications": 2,
    })
    # 크롬 자체 비번 경고/알림류 억제
    opts.add_argument("--disable-features=PasswordLeakDetection,PasswordCheck,PasswordManagerOnboarding,NotificationTriggers,PushMessaging,PermissionPromptFilter")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-popup-blocking")

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)
    wait = WebDriverWait(driver, 8)

    try:
        # 로그인 + '쪽지쓰기' 모달 열기
        login_and_open_compose(driver, wait, uid, pw)

        success, fail, sent = 0, 0, 0
        for i, row in df.iterrows():
            # 범위 제어
            if args.start and i < args.start:
                continue
            if args.limit and sent >= args.limit:
                break

            tid = str(row["후원아이디"]).strip()
            if not tid:
                st["items"][i]["status"]  = "fail"
                st["items"][i]["updated"] = now_ts()
                save_status(status_path, st)
                continue

            # 5명마다 '줄 끝 스페이스' 적용
            message = msg_with_line_end_spaces(base_message, sent)

            ok = send_one(wait, tid, message)

            st["items"][i]["status"]  = "success" if ok else "fail"
            st["items"][i]["updated"] = now_ts()
            save_status(status_path, st)

            if ok:
                success += 1
            else:
                fail += 1
            sent += 1

            # 사람이 직접 보내는 것처럼 0.2~2초 랜덤 대기
            delay = random.uniform(0.2, 2)  # 0.2초 ~ 2초 사이 부동소수
            time.sleep(delay)

        print(f"[done] 성공 {success} / 실패 {fail}")
        sys.exit(0)

    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()