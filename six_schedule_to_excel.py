"""SIX 더 뮤지컬 캐스팅표 -> 포도알 스케줄 엑셀 양식 변환.

양식은 crawrling.py 가 만드는 뮤지컬 스케줄 파일과 동일하다.
    id | 시즌 | 공연장명 | 날짜 | 시간(HH:MM:00) | 배우([a,b,c,d,e,f])
"""
import pandas as pd

YEAR = 2026
START_ID = None      # 어드민 스케줄 목록 최신 ID + 1 (모르면 None -> 빈 칸)
SEASON_ID = None     # performance_season ID
PLACE_ID = None      # performance_place ID

# 배우 순서: 아라곤, 불린, 시모어, 클레페, 하워드, 파
CASTING = """
12-15 20:00 손승연 배수정 한재아 김지선 김려원 유주혜
12-16 20:00 장보람 김지우 이보람 최현선 효정 주다은
12-17 20:00 손승연 김지우 이보람 최현선 김려원 유주혜
12-18 17:00 장보람 배수정 한재아 김지선 효정 주다은
12-18 20:30 손승연 배수정 한재아 최현선 효정 주다은
12-19 17:00 장보람 김지우 이보람 김지선 김려원 유주혜
12-19 20:30 장보람 김지우 이보람 김지선 효정 주다은
12-20 15:00 손승연 배수정 한재아 최현선 김려원 주다은
12-22 20:00 장보람 김지우 이보람 최현선 김려원 유주혜
12-23 20:00 손승연 배수정 한재아 김지선 효정 주다은
12-24 17:00 장보람 김지우 한재아 김지선 효정 유주혜
12-24 20:30 손승연 김지우 한재아 김지선 김려원 주다은
12-25 17:00 손승연 배수정 이보람 최현선 효정 유주혜
12-25 20:30 장보람 배수정 이보람 김지선 효정 유주혜
12-26 17:00 장보람 김지우 한재아 최현선 김려원 주다은
12-26 20:30 손승연 배수정 한재아 최현선 김려원 유주혜
12-27 15:00 장보람 김지우 이보람 김지선 효정 주다은
"""


def build_rows():
    rows = []
    for i, line in enumerate(l for l in CASTING.strip().splitlines() if l.strip()):
        md, t, *actors = line.split()
        assert len(actors) == 6, line
        rows.append({
            "id": START_ID + i if START_ID is not None else None,
            "시즌": SEASON_ID,
            "공연장명": PLACE_ID,
            "날짜": f"{YEAR}-{md}",
            "시간": f"{t}:00",
            "배우": "[" + ",".join(actors) + "]",
        })
    return rows


if __name__ == "__main__":
    out = "podoal_musical_SIX.xlsx"
    pd.DataFrame(build_rows()).to_excel(out, index=False)
    print("saved", out)
