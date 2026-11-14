# -*- coding: utf-8 -*-
"""TUG v12 — conservative simultaneous-slot candidate preview, NOT for navigation.
Requires openpyxl. Does not modify the source workbook or the v10/v11 agents.
"""
import argparse
import re
from datetime import date, datetime
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

ROOT = Path(__file__).resolve().parent
DEFAULT_EXCEL = ROOT / 'TUG 당직표(2).xlsx'


def normalize(s):
    return re.sub(r'\s+', '', str(s or '')).casefold()


def csv_names(s):
    return [v.strip() for v in str(s or '').split(',') if v.strip()]


def parse_date(v, year, month):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    m = re.fullmatch(r'\s*(\d{1,2})/(\d{1,2})\s*', str(v or ''))
    if m:
        try:
            return date(year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            return None
    return None


def read_day(path, day):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            m = re.search(r'(20\d{2})년\s*(\d{1,2})월', ws.title)
            if not m:
                continue
            year, month = map(int, m.groups())
            for row in ws.iter_rows(min_row=1, max_col=20):
                if parse_date(row[0].value, year, month) != day:
                    continue
                trio = [(str(row[i].value).strip() if row[i].value is not None else '', row[i].coordinate) for i in (9, 10, 11)]
                return {'sheet': ws.title, 'row': row[0].row, 'trio': trio,
                        'note': str(row[18].value or '').strip()}
    finally:
        wb.close()
    raise ValueError(f'원본 당직표에서 {day} 날짜를 찾지 못했습니다.')


def repair_mentions(note, name):
    """Only exact ship-name adjacent to repair/breakdown wording; not fuzzy matching."""
    if not name:
        return False
    escaped = re.escape(name)
    # '선900 수리', '선900:고장', '선900(수리)' etc.
    return bool(re.search(r'(?<![\w가-힣])' + escaped +
                          r'\s*[:：(（-]?\s*(?:수리|고장|정비|작업불가)', note, re.I))


def generate(info, port, count, waiting, unavailable, busy):
    pool = ([(n, '당일 당직선', c) for n, c in info['trio'] if n]
            if port == '온산' else [(n, '사용자 입력 대기선', '사용자 입력') for n in waiting])
    excluded, available, seen = [], [], set()
    blocked = {normalize(n) for n in unavailable}
    occupied = {normalize(n) for n in busy}
    for name, category, cell in pool:
        key = normalize(name)
        if key in seen:
            excluded.append((name, '원본 후보 목록 중복'))
            continue
        seen.add(key)
        if repair_mentions(info['note'], name):
            excluded.append((name, '비고란 수리·고장·정비 기록'))
        elif key in blocked:
            excluded.append((name, '사용자 지정 작업 불가'))
        elif key in occupied:
            excluded.append((name, '이미 다른 작업에 투입 중'))
        else:
            available.append((name, category, cell))
    result = []
    for i in range(count):
        if i < len(available):
            name, category, cell = available[i]
            result.append((i+1, name, category, cell, '검토 필요', '동시 작업 가정: 중복 없이 후보 1회만 사용'))
        else:
            result.append((i+1, '미배정', '', '', '추가 확인 필요', '중복 없이 배정할 후보 부족; 다른 선박/시간 검토 필요'))
    return result, excluded


def report(path, day, port, info, result, excluded, busy, unavailable):
    wb = Workbook()
    ws = wb.active
    ws.title = '동시작업_검토용'
    ws.append(['날짜', '항구', '동시작업 번호', '후보 선박', '후보 출처', '원본 셀', '상태', '설명'])
    for num, name, kind, cell, status, desc in result:
        ws.append([day.isoformat(), port, num, name, kind, cell, status, desc])
    meta = [('항목','내용'), ('원본 시트', info['sheet']), ('원본 행', info['row']),
            ('원본 비고', info['note']), ('사용자 작업불가', ', '.join(unavailable) or '없음'),
            ('사용자 작업중', ', '.join(busy) or '없음'),
            ('잠정 규칙', '3,000HP 변경안: 온산 당직선 순서 / 본항 대기선 순서'),
            ('작업 가정', '동일 시간대의 서로 다른 작업 슬롯; 한 선박은 1회만'),
            ('주의', '운항 가능성, 실제 마력, 실시간 위치, 작업시간, 도선사 조건 미검증'),
            ('사용 제한', '교육·시뮬레이션 전용. 실운항 사용 금지')]
    for i, (k,v) in enumerate(meta, 1):
        ws.cell(i, 10, k)
        ws.cell(i, 11, v)
    ex = wb.create_sheet('제외_기록')
    ex.append(['선박명', '제외 이유'])
    for item in excluded:
        ex.append(list(item))
    for sheet in (ws, ex):
        for c in sheet[1]:
            if c.value is None: continue
            c.fill = PatternFill('solid', fgColor='17324F')
            c.font = Font(color='FFFFFF', bold=True)
        sheet.freeze_panes = 'A2'
    for col, width in {'A':16,'B':12,'C':18,'D':17,'E':21,'F':15,'G':18,'H':55,'J':20,'K':70}.items():
        ws.column_dimensions[col].width = width
    ex.column_dimensions['A'].width = 20
    ex.column_dimensions['B'].width = 40
    ws.auto_filter.ref = f'A1:H{len(result)+1}'
    ex.auto_filter.ref = f'A1:B{max(2, len(excluded)+1)}'
    for row in ws.iter_rows(min_row=2, max_row=len(result)+1, min_col=1, max_col=8):
        for c in row:
            c.alignment = Alignment(vertical='center', wrap_text=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main():
    p = argparse.ArgumentParser(description='v12: 3,000HP 동시작업 후보 검토 (실운항 금지)')
    p.add_argument('--excel', type=Path, default=DEFAULT_EXCEL)
    p.add_argument('--date', required=True, help='YYYY-MM-DD')
    p.add_argument('--port', required=True, choices=['온산','본항'])
    p.add_argument('--count', type=int, default=3, help='동시 작업 수: 1~100')
    p.add_argument('--waiting', default='', help='본항 대기선, 쉼표 구분 (수동 확인 필수)')
    p.add_argument('--unavailable', default='', help='추가 수리·고장·작업불가 선박')
    p.add_argument('--busy', default='', help='이미 다른 작업에 투입 중인 선박')
    p.add_argument('--output', type=Path)
    args = p.parse_args()
    try:
        day = date.fromisoformat(args.date)
    except ValueError:
        p.error('날짜는 YYYY-MM-DD 형식이어야 합니다.')
    if not 1 <= args.count <= 100:
        p.error('--count는 1~100이어야 합니다.')
    if not args.excel.is_file():
        p.error(f'원본 엑셀을 찾을 수 없습니다: {args.excel}')
    waiting, unavailable, busy = map(csv_names, (args.waiting, args.unavailable, args.busy))
    if args.port == '본항' and not waiting:
        p.error('본항은 --waiting "바다1,제우스,선606"처럼 대기선 입력이 필요합니다.')
    try:
        info = read_day(args.excel, day)
        result, excluded = generate(info, args.port, args.count, waiting, unavailable, busy)
        output = args.output or ROOT / f'v12_동시작업_검토_{day}_{args.port}.xlsx'
        report(output, day, args.port, info, result, excluded, busy, unavailable)
    except (OSError, ValueError) as e:
        p.error(str(e))
    print(f'원본 시트: {info["sheet"]}, {info["row"]}행')
    print('당직선(J/K/L):', ', '.join(f'{n}({c})' for n,c in info['trio']))
    print('비고:', info['note'] or '없음')
    for name, reason in excluded:
        print(f'제외: {name} — {reason}')
    for num, name, *_ in result:
        print(f'{num:>2}. {name}')
    print('검토용 엑셀 저장:', output)
    print('주의: 동일 시간대 가정. 잠정 규칙·실시간 가용성 미확인. 실운항 사용 금지.')


if __name__ == '__main__':
    main()
