# -*- coding: utf-8 -*-
"""TUG v13: v12 후보 생성 + 검토용 평가 지표. 실운항 사용 금지.
필요: openpyxl, tug_v12_safe_preview.py, TUG 당직표(2).xlsx
"""
import argparse
from collections import Counter
from datetime import date
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
import tug_v12_safe_preview as v12

ROOT = Path(__file__).resolve().parent


def evaluate(info, result, excluded, busy, unavailable):
    total = len(result)
    assigned = [r for r in result if r[1] != '미배정']
    names = [v12.normalize(r[1]) for r in assigned]
    duplicates = sum(n - 1 for n in Counter(names).values() if n > 1)
    forbidden = {v12.normalize(n) for n in busy + unavailable}
    forbidden.update(v12.normalize(n) for n, _ in info['trio'] if v12.repair_mentions(info['note'], n))
    violations = sum(name in forbidden for name in names)
    blocked_pool = {v12.normalize(n) for n, _ in excluded if '중복' not in _}
    blocked_assigned = sum(name in blocked_pool for name in names)
    return {
        '요청 작업 수': total,
        '후보 배정 수': len(assigned),
        '미배정 수': total - len(assigned),
        '후보 충족률 (%)': round(100 * len(assigned) / total, 1) if total else 0,
        '중복 후보 배정 횟수': duplicates,
        '명시적 제외 조건 위반 횟수': violations + blocked_assigned,
        '제외 후보 수': len(excluded),
    }


def make_report(output, day, port, info, result, excluded, metrics, busy, unavailable):
    wb = Workbook()
    summary = wb.active
    summary.title = '평가_요약'
    summary.append(['평가 항목', '결과', '해석 / 제한'])
    explanations = {
        '요청 작업 수': '동시에 진행한다고 가정한 작업 슬롯 수',
        '후보 배정 수': '원본 당직선 중 조건에 걸리지 않은 후보를 한 번씩 선택',
        '미배정 수': '다른 선박 또는 시간 조정 검토 필요',
        '후보 충족률 (%)': '후보 배정 수 ÷ 요청 작업 수 × 100; 실제 배정 정확도가 아님',
        '중복 후보 배정 횟수': '동시 작업 내 같은 선박 반복 선택 횟수',
        '명시적 제외 조건 위반 횟수': '입력된 작업 중/작업 불가 및 비고 수리 기록에 대한 검사',
        '제외 후보 수': '후보 목록에서 제외된 항목 수 (중복 후보 항목 포함)',
    }
    for k, v in metrics.items():
        summary.append([k, v, explanations[k]])
    summary.append(['규칙 버전', '잠정 변경안', '매뉴얼 마지막 3,000HP 수정안; 현행 적용 여부 미확인'])
    summary.append(['검증 범위', '제한적', '마력 실측, 위치, 기상, 도선사 조건, 실시간 운항 정보 미확인'])
    summary.append(['실운항', '사용 금지', '교육·시뮬레이션 전용'])

    detail = wb.create_sheet('배정_검토')
    detail.append(['날짜','항구','작업 슬롯','후보 선박','후보 출처','원본 셀','상태','사유'])
    for n, name, category, cell, status, desc in result:
        detail.append([day.isoformat(), port, n, name, category, cell, status, desc])

    excluded_sheet = wb.create_sheet('제외_사유')
    excluded_sheet.append(['선박', '사유'])
    for name, reason in excluded:
        excluded_sheet.append([name, reason])

    provenance = wb.create_sheet('입력_출처')
    for key, value in [
        ('원본 파일', 'TUG 당직표(2).xlsx'), ('원본 시트', info['sheet']),
        ('원본 행', info['row']), ('당직선 J/K/L', ', '.join(f'{n}({c})' for n,c in info['trio'])),
        ('비고', info['note']), ('작업 중 선박', ', '.join(busy) or '없음'),
        ('사용자 지정 작업 불가', ', '.join(unavailable) or '없음'),
        ('작업 가정', '모든 슬롯 동시 진행'),
        ('주의', '실제 배정 결과/정답 라벨 없음: 정확도, 재현율, 실제 규칙 준수율은 계산 불가'),
    ]:
        provenance.append([key, value])

    for ws in wb.worksheets:
        ws.freeze_panes = 'A2'
        ws.sheet_view.showGridLines = False
        for cell in ws[1]:
            cell.fill = PatternFill('solid', fgColor='17324F')
            cell.font = Font(color='FFFFFF', bold=True)
        for row in ws:
            for cell in row:
                cell.alignment = Alignment(vertical='center', wrap_text=True)
        ws.auto_filter.ref = ws.dimensions
    for col, width in {'A':30,'B':23,'C':88}.items():
        summary.column_dimensions[col].width = width
    for col, width in {'A':17,'B':13,'C':15,'D':18,'E':22,'F':14,'G':19,'H':62}.items():
        detail.column_dimensions[col].width = width
    excluded_sheet.column_dimensions['A'].width = 23
    excluded_sheet.column_dimensions['B'].width = 52
    provenance.column_dimensions['A'].width = 28
    provenance.column_dimensions['B'].width = 95
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def main():
    p = argparse.ArgumentParser(description='v13: 동시작업 후보 및 평가 보고서 (실운항 금지)')
    p.add_argument('--excel', type=Path, default=v12.DEFAULT_EXCEL)
    p.add_argument('--date', required=True)
    p.add_argument('--port', choices=['온산', '본항'], required=True)
    p.add_argument('--count', type=int, default=3)
    p.add_argument('--waiting', default='')
    p.add_argument('--busy', default='')
    p.add_argument('--unavailable', default='')
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
    waiting, busy, unavailable = map(v12.csv_names, (args.waiting, args.busy, args.unavailable))
    if args.port == '본항' and not waiting:
        p.error('본항은 --waiting "바다1,제우스" 등 대기선 순서 입력이 필요합니다.')
    try:
        info = v12.read_day(args.excel, day)
        result, excluded = v12.generate(info, args.port, args.count, waiting, unavailable, busy)
        metrics = evaluate(info, result, excluded, busy, unavailable)
        output = args.output or ROOT / f'v13_배정평가_{day}_{args.port}.xlsx'
        make_report(output, day, args.port, info, result, excluded, metrics, busy, unavailable)
    except (ValueError, OSError) as e:
        p.error(str(e))
    print(f'원본: {info["sheet"]} / {info["row"]}행')
    for n, name, *_ in result:
        print(f'{n:>2}. {name}')
    print('평가지표:')
    for key, value in metrics.items():
        print(f' - {key}: {value}')
    print('평가 엑셀 저장:', output)
    print('주의: 후보 충족률은 실제 배정 정확도가 아닙니다. 잠정 규칙, 실운항 금지.')


if __name__ == '__main__':
    main()
