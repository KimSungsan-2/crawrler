"""SIX 더 뮤지컬 캐스팅표 -> 포도알 스케줄 엑셀 양식 변환.

양식은 crawrling.py 가 만드는 뮤지컬 스케줄 파일과 동일하다.
    id | 시즌 | 공연장명 | 날짜 | 시간(HH:MM:00) | 배우([a,b,c,d,e,f])
"""
import argparse
import json
import os
import time

import pandas as pd

YEAR = 2026
SEARCH_KEYWORD = "식스"   # 티켓오픈 목록에서 찾을 공연명 키워드
START_ID = None      # 어드민 스케줄 목록 최신 ID + 1 (None이면 어드민에서 자동 조회)
SEASON_ID = None     # performance_season ID (None이면 자동 조회)
PLACE_ID = None      # performance_place ID (None이면 자동 조회)

PLACES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "places.json")
BASE_URL = "https://podor.co.kr/admin/performance"
PLACE_URL = "https://podor.co.kr/admin/performance/performance_place/"

# 배우 순서: 아라곤, 불린, 시모어, 클레페, 하워드, 파
CASTING = """
12-15 20:00 손승연 배수정 한재아 김지선 김려원 유주혜
12-16 20:00 장보람 김지우 이보람 최현선 효정 주다온
12-17 20:00 손승연 김지우 이보람 최현선 김려원 유주혜
12-18 17:00 장보람 배수정 한재아 김지선 효정 주다온
12-18 20:30 손승연 배수정 한재아 최현선 효정 주다온
12-19 17:00 장보람 김지우 이보람 김지선 김려원 유주혜
12-19 20:30 장보람 김지우 이보람 김지선 효정 주다온
12-20 15:00 손승연 배수정 한재아 최현선 김려원 주다온
12-22 20:00 장보람 김지우 이보람 최현선 김려원 유주혜
12-23 20:00 손승연 배수정 한재아 김지선 효정 주다온
12-24 17:00 장보람 김지우 한재아 김지선 효정 유주혜
12-24 20:30 손승연 김지우 한재아 김지선 김려원 주다온
12-25 17:00 손승연 배수정 이보람 최현선 효정 유주혜
12-25 20:30 장보람 배수정 이보람 김지선 효정 유주혜
12-26 17:00 장보람 김지우 한재아 최현선 김려원 주다온
12-26 20:30 손승연 배수정 한재아 최현선 김려원 유주혜
12-27 15:00 장보람 김지우 이보람 김지선 효정 주다온
"""


def fetch_admin_ids(keyword=SEARCH_KEYWORD):
    """포도알 어드민에서 (시작 id, 시즌 id, 공연장 id)를 조회한다. crawrling.py와 같은 방식."""
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager

    try:
        from crawrling import PODOAL_ID, PODOAL_PW
    except ImportError:
        PODOAL_ID = PODOAL_PW = None
    user = os.environ.get("PODOAL_ID", PODOAL_ID)
    pw = os.environ.get("PODOAL_PW", PODOAL_PW)
    if not (user and pw):
        raise RuntimeError("PODOAL_ID / PODOAL_PW 환경변수를 설정하세요.")

    opts = webdriver.ChromeOptions()
    for a in ("--headless", "--no-sandbox", "--disable-dev-shm-usage", "--window-size=1920,1080"):
        opts.add_argument(a)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    wait = WebDriverWait(driver, 15)
    first_row = "#result_list tbody tr:first-child"
    try:
        driver.get("https://podor.co.kr/admin/login/")
        driver.find_element(By.NAME, "username").send_keys(user)
        driver.find_element(By.NAME, "password").send_keys(pw + Keys.ENTER)
        time.sleep(2)

        # 1) 스케줄 목록 최신 id + 1
        driver.get(f"{BASE_URL}/performanceschedule/")
        row = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, first_row)))
        last_id = int(row.find_element(By.CSS_SELECTOR, "th.field-id, td.field-id").text.strip())

        # 2) 티켓오픈 목록에서 공연 검색 -> 시즌 id (cells[1]=공연명, cells[2]=시즌 id)
        driver.get(f"{BASE_URL}/performance_open/?q={keyword}")
        row = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, first_row)))
        cells = row.find_elements(By.CSS_SELECTOR, "td, th")
        title, season_id = cells[1].text.strip(), cells[2].text.strip()
        print(f"공연: {title} / 시즌 id: {season_id}")

        # 3) 시즌 -> 공연장명(td[4]) -> 공연장 id
        driver.get(f"{BASE_URL}/performance_season/?q={season_id}")
        row = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, first_row)))
        place_name = row.find_elements(By.TAG_NAME, "td")[4].text.strip()
        driver.get(f"{PLACE_URL}?q={place_name}")
        place_id = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "th.field-id, td.field-id"))).text.strip()
        print(f"공연장: {place_name} / 공연장 id: {place_id}")
        return last_id + 1, int(season_id), int(place_id)
    finally:
        driver.quit()


def place_id_from_name(name):
    """places.json(공연장명 -> id)에서 공연장 id를 찾는다. 부분 일치는 후보가 하나일 때만 허용."""
    with open(PLACES_FILE, encoding="utf-8") as f:
        places = json.load(f)
    if name in places:
        return places[name]
    hits = [n for n in places if name in n]
    if len(hits) == 1:
        return places[hits[0]]
    raise SystemExit(f"공연장 '{name}' 일치 항목 {len(hits)}개: {hits[:10]}")


def build_rows(start_id=START_ID, season_id=SEASON_ID, place_id=PLACE_ID):
    rows = []
    for i, line in enumerate(l for l in CASTING.strip().splitlines() if l.strip()):
        md, t, *actors = line.split()
        assert len(actors) == 6, line
        rows.append({
            "id": start_id + i if start_id is not None else None,
            "시즌": season_id,
            "공연장명": place_id,
            "날짜": f"{YEAR}-{md}",
            "시간": f"{t}:00",
            "배우": "[" + ",".join(actors) + "]",
        })
    return rows


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--start-id", type=int, default=START_ID)
    ap.add_argument("--season-id", type=int, default=SEASON_ID)
    ap.add_argument("--place-id", type=int, default=PLACE_ID)
    ap.add_argument("--place", help="공연장명 (places.json에서 id 조회)")
    ap.add_argument("--keyword", default=SEARCH_KEYWORD)
    ap.add_argument("--offline", action="store_true", help="어드민 조회 없이 빈 칸으로 생성")
    a = ap.parse_args()

    if a.place and a.place_id is None:
        a.place_id = place_id_from_name(a.place)
    ids = (a.start_id, a.season_id, a.place_id)
    if not a.offline and None in ids:
        fetched = fetch_admin_ids(a.keyword)
        ids = tuple(g if g is not None else f for g, f in zip(ids, fetched))

    out = "podoal_musical_SIX.xlsx"
    pd.DataFrame(build_rows(*ids)).to_excel(out, index=False)
    print("saved", out, "(id/시즌/공연장:", ids, ")")
