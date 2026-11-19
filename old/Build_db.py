# -*- coding: utf-8 -*-
"""data 폴더의 엑셀/CSV -> tug_agent.db  (excel_rows, tug_names 테이블)

실제 업무 엑셀(작업일지, 예보, 건수일지, 당직표, 순번표) 구조에 맞춘 버전.

시트 유형별 처리
 1) '당직표' 시트: 2단 헤더(상위 그룹 + 마력)를 합쳐서 열 이름을 만들고, 날짜 한 줄을 한 레코드로 저장
      예) '6/24 | 온산정계지대기예선: 흥5, 선101 | 온산1300HP: 해101, 조101 | ... | 본항1300HP: 조107, 해9'
 2) 헤더 행이 있는 시트(작업일지, 건수일지, 순번표, LNG일지 ...): 열 이름을 붙여 저장
      예) '2026-07-01 | 예정시간: 00:20 | 선명: WOO DONG | G/T: 8581 | 사용예선: 선진6,흥진5'
      - 좌우로 같은 표가 여러 개 나란히 있으면(날짜별 블록) 블록마다 한 레코드로 나눈다.
      - '날짜' 칸 옆의 날짜를 레코드 앞에 붙인다.
 3) 헤더가 없는 시트(예보, 설명문 등): 값을 ' | '로 이어서 저장 (날짜가 있으면 앞에 붙임)

tug_names 테이블: 에이전트가 질문에서 예선 이름을 찾을 때 쓰는 이름 목록 (3회 이상 나온 이름)

forecast 테이블: Crawling.py가 수집한 도선예보 (열 단위 정형 데이터, Tug_Agent가 배정에 사용)
 - 구분(kind): 완료 / 미완료 / 예정,  같은 (날짜, 구분)은 최신 크롤링 결과로 덮어쓴다.
 - forecast_log: (날짜, 구분)별 마지막 저장 시각과 건수. '미완료 0건'(전부 끝남)과 '아직 수집 안 함'을 구분하는 용도.
 - Crawling.py가 실행될 때마다 save_forecast_xlsx()로 자동 저장된다.

실행: python build_db.py   (excel_rows/tug_names는 새로 만들고, forecast는 유지하면서 output 폴더 엑셀로 갱신)
필요: pip install openpyxl
"""
import csv
import datetime
import re
import sqlite3
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA = ROOT / 'data'
DB = ROOT / 'tug_agent.db'
OUTPUT = ROOT / 'output'          # Crawling.py가 도선예보 엑셀을 저장하는 폴더

# 도선예보 엑셀 열 이름 -> DB 열 이름
FORECAST_COLS = {
    'NO': 'no', 'STATUS': 'status', 'C/F': 'cf', 'TIME': 'time', "SHIP'S NAME": 'ship',
    'PILOT': 'pilot', 'C/SIGN': 'callsign', 'G/T': 'gt', 'LOA': 'loa', 'DFT': 'dft',
    'FROM': 'frm', 'TO': 'dest', 'L/A': 'la', 'CA': 'ca', 'S/A': 'sa', 'B/T': 'bt',
    'T': 'tugs', 'L': 'lines', 'Q': 'q', 'REMARKS': 'remarks',
}
FORECAST_SHEET_RE = re.compile(r'(\d{4}-\d{2}-\d{2})\s*(미완료|완료|예정)')

HEADER_WORDS = {
    '선명', 'FROM', 'TO', '톤수', 'G/T', '작업TUG', '사용예선', '예정시간', '출발시간', '종료시간',
    '일시', 'TUG', '시간', '부두', '대리점', '비고', '입항', '출항', '구분', '총톤수', '작업',
    '총작업시간', '내용',
}
DATE_LABELS = {'날짜', '날짜:', '일자'}
TUG_LABELS = {'사용예선', '작업TUG'}
DUTY_DATE_RE = re.compile(r'^\d{1,2}/\d{1,2}$')


def norm(s):
    return re.sub(r'\s+', '', s or '')


def cell_text(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime.time):
        return v.strftime('%H:%M')
    if hasattr(v, 'strftime'):
        if getattr(v, 'hour', 0) or getattr(v, 'minute', 0):
            return v.strftime('%Y-%m-%d %H:%M')
        return v.strftime('%Y-%m-%d')
    return str(v).replace('\n', ' ').strip()


# 예선 이름이 아닌데 '사용예선' 칸에 자주 적히는 말
NOT_TUG = {'터그', '선박', '예선', '출항선박', '인도출항', '출항터그', '중공업작업', 'T작업', 'CANCEL',
           '정리완료', '취소', '대기', '없음'}


# 예선 이름처럼 생긴 형태: 선진8, 해강101, K601, 109청룡, 505해룡 ...
TUG_SHAPE = re.compile(r'[가-힣]{1,3}\d{1,3}|[Kk]\d{2,3}|\d{3}[가-힣]{2}|[가-힣]{1,3}\d{1,3}[가-힣]{1,2}')


def is_tug_token(t):
    return (2 <= len(t) <= 8 and t not in NOT_TUG and not re.search(r'[<>#()\[\]*~:]', t)
            and not re.fullmatch(r'[\d\-:.]+', t))


def split_tugs(text):
    return [t for t in re.split(r'[,/\s.+]+', text) if is_tug_token(t)]


# ----------------------------------------------------------------- 당직표
def duty_records(rows):
    """2단 헤더 당직표 -> (행번호, 텍스트, 예선토큰) 생성"""
    h = next((i for i, r in enumerate(rows) if r and norm(cell_text(r[0])).startswith('날짜')), None)
    if h is None or h + 1 >= len(rows):
        yield from plain_records(rows)
        return
    top = [norm(cell_text(v)) for v in rows[h]]
    sub = [norm(cell_text(v)) for v in rows[h + 1]]
    width = max(len(top), len(sub))
    labels, cur_top, cur_sub = [], '', ''
    for j in range(width):
        t = top[j] if j < len(top) else ''
        s = sub[j] if j < len(sub) else ''
        if t:
            cur_top, cur_sub = t, ''      # 새 상위 그룹이 시작되면 하위 라벨 초기화
        if s:
            cur_sub = s
        labels.append(cur_top + cur_sub)
    for i in range(h + 2, len(rows)):
        vals = [cell_text(v) for v in rows[i]]
        if not any(vals):
            continue
        date = vals[0]
        if not DUTY_DATE_RE.match(date):
            yield i + 1, ' | '.join(v for v in vals if v), []
            continue
        groups, tugs = {}, []
        for j in range(1, len(vals)):
            if vals[j]:
                lab = labels[j] if j < len(labels) and labels[j] else f'열{j}'
                groups.setdefault(lab, []).append(vals[j])
                if lab != '비고':
                    tugs.append(vals[j])
        text = date + ' | ' + ' | '.join(f'{k}: {", ".join(v)}' for k, v in groups.items())
        yield i + 1, text, tugs


# ----------------------------------------------------------------- 일반 시트
def plain_records(rows, date=''):
    for i, raw in enumerate(rows, 1):
        vals = [cell_text(v) for v in raw]
        if any(vals):
            text = ' | '.join(v for v in vals if v)
            yield i, (f'{date} | {text}' if date else text), []


def find_date_labels(raw, vals):
    """행에서 '날짜' 라벨과 그 옆의 날짜 값을 찾는다. -> {라벨열: 'YYYY-MM-DD'}"""
    found = {}
    for j, v in enumerate(vals):
        if norm(v) in DATE_LABELS:
            for k in range(j + 1, min(j + 4, len(raw))):
                x = raw[k]
                if isinstance(x, (datetime.datetime, datetime.date)):
                    found[j] = x.strftime('%Y-%m-%d')
                    break
    return found


def sheet_records(title, rows):
    rows = [list(r) for r in rows]
    if '당직표' in title:
        yield from duty_records(rows)
        return

    header, starts = None, []
    date_cols, sheet_date = {}, ''
    for n, raw in enumerate(rows, 1):
        vals = [cell_text(v) for v in raw]
        if not any(vals):
            continue

        dl = find_date_labels(raw, vals)
        if dl:                                          # 날짜 행: 문맥을 갱신하고, 날짜별로 한 줄씩 저장
            date_cols.update(dl)
            cols = sorted(dl)
            sheet_date = dl[cols[0]]
            for i, jc in enumerate(cols):
                end = cols[i + 1] if i + 1 < len(cols) else len(vals)
                yield n, ' | '.join(v for v in vals[jc:end] if v), []
            continue

        hits = sum(1 for v in vals if v and norm(v).upper() in HEADER_WORDS)
        if hits >= 3:                                   # 헤더 행: 열 이름만 기억하고 따로 저장하지 않는다
            header = {j: norm(v) for j, v in enumerate(vals) if v}   # '사  용  예  선' -> '사용예선'
            first = header[min(header)].upper()
            starts = [j for j, v in header.items() if v.upper() == first]
            continue

        if header is None:
            yield n, (f'{sheet_date} | ' if sheet_date else '') + ' | '.join(v for v in vals if v), []
            continue

        covered = set()
        for bi, s in enumerate(starts):
            e = starts[bi + 1] if bi + 1 < len(starts) else len(vals)
            cells = [(j, vals[j]) for j in range(s, min(e, len(vals))) if vals[j]]
            if not cells:
                continue
            covered.update(j for j, _ in cells)
            d = next((dt for jc, dt in date_cols.items() if s <= jc < e), sheet_date)
            parts, tugs = [], []
            for j, v in cells:
                lab = header.get(j)
                parts.append(f'{lab}: {v}' if lab else v)
                if lab and lab.upper() in TUG_LABELS:
                    tugs += split_tugs(v)
            yield n, (f'{d} | ' if d else '') + ' | '.join(parts), tugs
        rest = [vals[j] for j in range(len(vals)) if vals[j] and j not in covered]
        if rest:                                        # 표 바깥에 있는 값(제목, 순번 규칙 등)
            yield n, ' | '.join(rest), []


# ----------------------------------------------------------------- 파일 읽기
def read_xlsx(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
        for n, text, tugs in sheet_records(ws.title, rows):
            yield ws.title, n, text, tugs
    wb.close()


def read_csv(path):
    for enc in ('utf-8-sig', 'cp949'):
        try:
            with open(path, newline='', encoding=enc) as f:
                rows = list(csv.reader(f))
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(f'{path.name}: 인코딩을 읽을 수 없습니다(utf-8/cp949 아님)')
    for n, text, tugs in sheet_records('csv', rows):
        yield 'csv', n, text, tugs


# ----------------------------------------------------------------- 도선예보 (Crawling.py 결과)
def init_forecast(conn):
    cols = ', '.join(f'{c} TEXT' for c in FORECAST_COLS.values())
    conn.execute(f'CREATE TABLE IF NOT EXISTS forecast '
                 f'(forecast_date TEXT, kind TEXT, crawled_at TEXT, {cols})')
    conn.execute('CREATE INDEX IF NOT EXISTS ix_forecast ON forecast(forecast_date, kind)')
    conn.execute('CREATE TABLE IF NOT EXISTS forecast_log '
                 '(forecast_date TEXT, kind TEXT, crawled_at TEXT, n INTEGER, source_file TEXT, '
                 'PRIMARY KEY (forecast_date, kind))')


def save_forecast_xlsx(path, conn=None) -> dict:
    """Crawling.py가 만든 도선예보 엑셀(시트 'YYYY-MM-DD 완료/미완료/예정')을 forecast 테이블에 저장.
    같은 (날짜, 구분)은 지우고 새로 넣는다. {시트명: 건수} 반환"""
    from openpyxl import load_workbook
    path = Path(path)
    own = conn is None
    if own:
        conn = sqlite3.connect(DB)
    init_forecast(conn)
    crawled = datetime.datetime.fromtimestamp(path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')
    marks = ','.join('?' * (3 + len(FORECAST_COLS)))
    counts = {}
    wb = load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        m = FORECAST_SHEET_RE.search(ws.title)
        if not m:
            continue
        d, kind = m.groups()
        rows = list(ws.iter_rows(values_only=True))
        idx = {str(h).strip(): i for i, h in enumerate(rows[0]) if h is not None} if rows else {}
        data = []
        for r in rows[1:]:
            vals = tuple(cell_text(r[idx[c]]) if c in idx and idx[c] < len(r) else '' for c in FORECAST_COLS)
            if any(vals):
                data.append((d, kind, crawled) + vals)
        conn.execute('DELETE FROM forecast WHERE forecast_date=? AND kind=?', (d, kind))
        conn.executemany(f'INSERT INTO forecast VALUES ({marks})', data)
        conn.execute('INSERT OR REPLACE INTO forecast_log VALUES (?,?,?,?,?)',
                     (d, kind, crawled, len(data), path.name))
        counts[ws.title] = len(data)
    wb.close()
    conn.commit()
    if own:
        conn.close()
    return counts


def has_forecast(forecast_date: str, kind: str) -> bool:
    """(날짜, 구분)이 한 번이라도 저장됐는지 (0건 저장도 True)"""
    if not DB.exists():
        return False
    conn = sqlite3.connect(DB)
    try:
        init_forecast(conn)
        return conn.execute('SELECT 1 FROM forecast_log WHERE forecast_date=? AND kind=?',
                            (forecast_date, kind)).fetchone() is not None
    finally:
        conn.close()


def load_forecast(forecast_date: str, kinds=('미완료', '예정')):
    """kinds 순서대로 저장된 첫 구분의 행을 원래 엑셀 열 이름 dict로 반환 -> (구분, 저장시각, rows)
    하나도 없으면 (None, None, [])"""
    if not DB.exists():
        return None, None, []
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    try:
        init_forecast(conn)
        for k in kinds:
            log = conn.execute('SELECT crawled_at FROM forecast_log WHERE forecast_date=? AND kind=?',
                               (forecast_date, k)).fetchone()
            if log is None:
                continue
            rs = conn.execute('SELECT * FROM forecast WHERE forecast_date=? AND kind=? ORDER BY rowid',
                              (forecast_date, k)).fetchall()
            return k, log[0], [{orig: r[col] for orig, col in FORECAST_COLS.items()} for r in rs]
        return None, None, []
    finally:
        conn.close()


def main():
    DATA.mkdir(exist_ok=True)
    files = sorted(p for p in DATA.iterdir()
                   if p.suffix.lower() in ('.xlsx', '.xlsm', '.csv', '.xls') and not p.name.startswith('~$'))
    if not files:
        print(f'{DATA} 폴더에 엑셀(.xlsx)/CSV 파일이 없습니다. (도선예보만 갱신합니다)')

    conn = sqlite3.connect(DB)            # DB 파일은 지우지 않는다: forecast(크롤링 이력) 보존
    conn.execute('DROP TABLE IF EXISTS excel_rows')
    conn.execute('DROP TABLE IF EXISTS tug_names')
    conn.execute('CREATE TABLE excel_rows (source_file TEXT, sheet_name TEXT, '
                 'row_number INTEGER, row_text TEXT)')
    conn.execute('CREATE TABLE tug_names (name TEXT PRIMARY KEY, cnt INTEGER)')
    tug_count = Counter()
    total = 0
    for p in files:
        if p.suffix.lower() == '.xls':
            print(f'건너뜀: {p.name} (.xls는 지원하지 않습니다. 엑셀에서 .xlsx로 저장하세요)')
            continue
        reader = read_csv if p.suffix.lower() == '.csv' else read_xlsx
        n = 0
        for sheet, row_no, text, tugs in reader(p):
            conn.execute('INSERT INTO excel_rows VALUES (?,?,?,?)', (p.name, sheet, row_no, text))
            tug_count.update(tugs)
            n += 1
        total += n
        print(f'{p.name}: {n}행')
    names = [(k, c) for k, c in tug_count.items()
             if not k.isdigit() and (c >= 30 or (c >= 3 and TUG_SHAPE.fullmatch(k)))]
    conn.executemany('INSERT INTO tug_names VALUES (?,?)', names)
    conn.commit()

    fc = sorted(p for p in OUTPUT.glob('*.xlsx') if not p.name.startswith('~$')) if OUTPUT.is_dir() else []
    for p in fc:                                   # 날짜순 → 같은 (날짜, 구분)은 나중 파일이 덮어씀
        counts = save_forecast_xlsx(p, conn)
        print(f'도선예보 {p.name}: ' + ', '.join(f'{k} {v}건' for k, v in counts.items()))
    conn.close()
    print(f'완료: 총 {total}행, 예선 이름 {len(names)}개, 도선예보 파일 {len(fc)}개 -> {DB.name}')


if __name__ == '__main__':
    main()