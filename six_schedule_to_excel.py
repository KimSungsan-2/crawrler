"""캐스팅표(casting/*.txt) -> 포도알 스케줄 엑셀 양식 변환. 예: python six_schedule_to_excel.py janhok --after 2026-11-01

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
# crawrling.py CONFIG와 동일: 카테고리별 어드민 경로 (공연장 목록은 뮤지컬/연극 공용)
CATEGORIES = {
    "musical": {"base": "https://podor.co.kr/admin/performance", "open": "/performance_open/",
                "schedule": "/performanceschedule/", "season": "/performance_season/"},
    "play": {"base": "https://podor.co.kr/admin/plays", "open": "/play_open/",
             "schedule": "/playschedule/", "season": "/play_season/"},
}
PLACE_URL = "https://podor.co.kr/admin/performance/performance_place/"

CASTING_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "casting")


def load_casting(show):
    """casting/<show>.txt: '# 주석' 또는 'MM-DD HH:MM 배우1 배우2 ...' 형식."""
    with open(os.path.join(CASTING_DIR, f"{show}.txt"), encoding="utf-8") as f:
        return [l.strip() for l in f if l.strip() and not l.startswith("#")]


def fetch_admin_ids(keyword=SEARCH_KEYWORD, category="musical"):
    """포도알 어드민에서 (시작 id, 시즌 id, 공연장 id)를 조회한다. crawrling.py와 같은 방식."""
    c = CATEGORIES[category]
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
        driver.get(c["base"] + c["schedule"])
        row = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, first_row)))
        last_id = int(row.find_element(By.CSS_SELECTOR, "th.field-id, td.field-id").text.strip())

        # 2) 티켓오픈 목록에서 공연 검색 -> 시즌 id (cells[1]=공연명, cells[2]=시즌 id)
        driver.get(f'{c["base"]}{c["open"]}?q={keyword}')
        row = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, first_row)))
        cells = row.find_elements(By.CSS_SELECTOR, "td, th")
        title, season_id = cells[1].text.strip(), cells[2].text.strip()
        print(f"공연: {title} / 시즌 id: {season_id}")

        # 3) 시즌 -> 공연장명(td[4]) -> 공연장 id
        driver.get(f'{c["base"]}{c["season"]}?q={season_id}')
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


def status_text(show, season_id, existing_xlsx):
    """복사용 상태 문구를 돌려준다.

    - 캡처의 모든 회차가 어드민에 같은 캐스팅으로 이미 있으면 '최신화 필요'
    - 어드민에 없는 회차가 있으면 캐스팅표의 마지막 날짜로 '2026-mm-dd 반영'
    - 같은 날짜·시간인데 캐스팅이 다른 행은 별도로 알려준다(문구 판정에는 '있음'으로 취급하지 않음).
    """
    df = pd.read_excel(existing_xlsx)
    df = df[df["시즌"] == season_id]
    admin = {(str(r["날짜"])[:10], str(r["시간"])[:5]): str(r["배우"]).replace(" ", "")
             for _, r in df.iterrows()}
    new, diff, last = 0, [], None
    for line in load_casting(show):
        md, t, *actors = line.split()
        key = (f"{YEAR}-{md}", t)
        last = key[0]
        cast = "[" + ",".join(actors) + "]"
        if key not in admin:
            new += 1
        elif admin[key] != cast:
            diff.append(key)
    text = f"{last} 반영" if new else "최신화 필요"
    return text, new, diff


def build_rows(show, start_id=START_ID, season_id=SEASON_ID, place_id=PLACE_ID, after=None):
    rows = []
    lines = [l for l in load_casting(show) if after is None or f"{YEAR}-{l.split()[0]}" > after]
    for i, line in enumerate(lines):
        md, t, *actors = line.split()
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
    ap.add_argument("show", help="casting/<show>.txt 이름 (예: six, janhok)")
    ap.add_argument("--after", help="YYYY-MM-DD 이후 날짜만 (어드민에 이미 있는 날짜 제외)")
    ap.add_argument("--start-id", type=int, default=START_ID)
    ap.add_argument("--season-id", type=int, default=SEASON_ID)
    ap.add_argument("--place-id", type=int, default=PLACE_ID)
    ap.add_argument("--place", help="공연장명 (places.json에서 id 조회)")
    ap.add_argument("--category", choices=CATEGORIES, default="musical")
    ap.add_argument("--existing", help="어드민 스케줄 내보내기 xlsx (상태 문구 판정용, --season-id 필요)")
    ap.add_argument("--keyword", default=SEARCH_KEYWORD)
    ap.add_argument("--offline", action="store_true", help="어드민 조회 없이 빈 칸으로 생성")
    a = ap.parse_args()

    if a.place and a.place_id is None:
        a.place_id = place_id_from_name(a.place)
    ids = (a.start_id, a.season_id, a.place_id)
    if not a.offline and None in ids:
        fetched = fetch_admin_ids(a.keyword, a.category)
        ids = tuple(g if g is not None else f for g, f in zip(ids, fetched))

    if a.existing:
        text, new, diff = status_text(a.show, ids[1], a.existing)
        print(f"[상태] {text}   (신규 {new}회차)")
        for k in diff:
            print("  ! 어드민과 캐스팅이 다른 회차:", k)
        if not new:
            raise SystemExit(0)   # 올릴 행이 없으면 엑셀을 만들지 않는다

    out = f"podoal_{a.category}_{a.show}.xlsx"
    pd.DataFrame(build_rows(a.show, *ids, after=a.after)).to_excel(out, index=False)
    print("saved", out, "(id/시즌/공연장:", ids, ")")
