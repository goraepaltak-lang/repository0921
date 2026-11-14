"""
울산항도선사회 도선예보 크롤러  (매일 23:55 실행용)
결과 엑셀 시트 3개: 오늘_완료 / 오늘_미완료 / 내일_예정
- 기본 화면 수집 -> '완료' 탭 클릭 후 다시 수집
- 엑셀 저장 후 Build_db.save_forecast_xlsx()로 tug_agent.db의 forecast 테이블에도 저장
- 설치: pip install playwright pandas openpyxl lxml && playwright install chromium
"""
import re
from datetime import datetime, timedelta
from io import StringIO
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

URL = "http://www.ulsanpilot.co.kr/main/pilot_forecast.php"
OUT_DIR = Path(__file__).resolve().parent / "output"   # 실행 위치와 무관하게 스크립트 옆 output 폴더
COLUMNS = ["NO", "STATUS", "C/F", "TIME", "SHIP'S NAME", "PILOT", "C/SIGN", "G/T",
           "LOA", "DFT", "FROM", "TO", "L/A", "CA", "S/A", "B/T", "T", "L", "Q", "REMARKS"]
KEY_COLS = ["TIME", "SHIP'S NAME", "FROM", "TO"]  # 같은 건 판별용


def click_done_tab(page) -> bool:
    """'완료' 탭/버튼 클릭. 성공 여부 반환"""
    candidates = [
        page.get_by_role("tab", name="완료", exact = True),
        page.get_by_role("button", name="완료", exact = True),
        page.get_by_label("완료", exact = True),
        page.get_by_text("완료", exact = True),
    ]
    for loc in candidates:
        try:
            if loc.count() > 0:
                loc.first.click(timeout = 5000)
                page.wait_for_load_state("networkidle", timeout = 30000)
                page.wait_for_timeout(2000)
                return True
        except Exception:
            continue
    return False


def fetch_html():
    """(기본화면 HTML, 완료탭 HTML) 반환"""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless = True)
        page = browser.new_page(viewport = {"width": 1600, "height": 1000})
        page.goto(URL, wait_until = "networkidle", timeout = 60000)
        page.wait_for_timeout(3000)
        html_default = page.content()

        if not click_done_tab(page):
            page.screenshot(path = "debug_page.png", full_page = True)
            browser.close()
            raise RuntimeError("'완료' 탭을 찾지 못했습니다. debug_page.png 확인 후 선택자를 수정하세요.")
        html_done = page.content()
        browser.close()
    return html_default, html_done


def parse(html: str) -> dict:
    """{'YYYY-MM-DD': DataFrame}"""
    dates = [f"{y}-{m}-{d}" for y, m, d in re.findall(r"(\d{4})년\s*(\d{2})월\s*(\d{2})일", html)]
    try:
        tables = pd.read_html(StringIO(html), match = "SHIP'S NAME")
    except ValueError:
        return {}
    result = {}
    for i, t in enumerate(tables):
        if i >= len(dates):
            break
        t = t.dropna(how = "all").fillna("").reindex(columns=COLUMNS, fill_value = "")
        result[dates[i]] = t.astype(str).replace({r"\.0$": ""}, regex = True)
    return result


def style_sheet(ws):
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left = thin, right = thin, top = thin, bottom = thin)
    head_fill = PatternFill("solid", fgColor = "3B6B5C")
    widths = [5, 13, 6, 8, 26, 16, 11, 9, 6, 6, 12, 12, 12, 5, 6, 6, 4, 8, 4, 40]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal = "center", vertical = "center", wrap_text = True)
            if cell.row == 1:
                cell.fill = head_fill
                cell.font = Font(bold = True, color = "FFFFFF")
            elif cell.column_letter == "E":
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal = "left", vertical = "center", wrap_text = True)
            elif cell.column_letter == "T":
                cell.alignment = Alignment(horizontal = "left", vertical = "center", wrap_text = True)
            elif cell.column_letter == "B":
                cell.font = Font(color = "1F3FBF")
            elif cell.column_letter == "O":
                cell.font = Font(color = "C00000")
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"


def main(out_dir: Path = OUT_DIR) -> Path:
    """크롤링 후 엑셀 저장. 저장된 파일 경로를 반환 (Tug_Agent에서 import해서 호출)"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents = True, exist_ok = True)
    now = datetime.now()
    d_today = now.strftime("%Y-%m-%d")
    d_tomorrow = (now + timedelta(days = 1)).strftime("%Y-%m-%d")

    html_default, html_done = fetch_html()
    default, done = parse(html_default), parse(html_done)
    empty = pd.DataFrame(columns = COLUMNS)

    today_done = done.get(d_today, empty)
    today_all = default.get(d_today, empty)
    done_keys = set(map(tuple, today_done[KEY_COLS].values))
    not_done = today_all[[tuple(r) not in done_keys for r in today_all[KEY_COLS].values]]

    sheets = {
        f"{d_today} 완료": today_done,
        f"{d_today} 미완료": not_done,
        f"{d_tomorrow} 예정": default.get(d_tomorrow, empty),
    }
    path = out_dir / f"도선예보_{d_today}.xlsx"
    with pd.ExcelWriter(path, engine = "openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name = name, index = False)
            style_sheet(w.sheets[name])
    print(f"저장: {path} | " + ", ".join(f"{k} {len(v)}건" for k, v in sheets.items()))

    from Build_db import save_forecast_xlsx, DB        # 같은 폴더의 Build_db.py
    save_forecast_xlsx(path)
    print(f"DB 저장: {DB.name} (forecast 테이블)")
    return path


if __name__ == "__main__":
    main()