# -*- coding: utf-8 -*-
"""TUG v14 교육용 통합 실행기: 기존 YES/NO 에이전트 + v13 후보/평가.
기존 원본/DB/FAISS를 변경하지 않습니다. 실운항 사용 금지.
"""
from pathlib import Path
from datetime import date
import subprocess
import sys

ROOT = Path(__file__).resolve().parent


def ask(prompt, default=''):
    value = input(f'{prompt}' + (f' [{default}]' if default else '') + ': ').strip()
    return value or default


def run_agent():
    script = ROOT / 'agent_v9_fixed.py'
    if not script.is_file():
        print('agent_v9_fixed.py가 없습니다. 같은 폴더에 넣어 주세요.')
        return
    print('\n기존 v9 YES/NO 에이전트를 별도로 실행합니다. 종료 후 이 메뉴로 돌아옵니다.')
    try:
        subprocess.run([sys.executable, str(script)], cwd=ROOT, check=False)
    except OSError as exc:
        print('에이전트 실행 실패:', exc)


def run_evaluation():
    try:
        import tug_v13_evaluation as v13
        import tug_v12_safe_preview as v12
    except ImportError as exc:
        print('필수 파일 또는 라이브러리 누락:', exc)
        print('tug_v13_evaluation.py, tug_v12_safe_preview.py, openpyxl을 확인해 주세요.')
        return
    excel = ROOT / 'TUG 당직표(2).xlsx'
    if not excel.is_file():
        print('원본 엑셀을 찾을 수 없습니다:', excel)
        return
    try:
        day = date.fromisoformat(ask('날짜 YYYY-MM-DD', '2026-06-01'))
        port = ask('항구 (온산/본항)', '온산')
        if port not in ('온산', '본항'):
            raise ValueError('항구는 온산 또는 본항이어야 합니다.')
        count = int(ask('동시 작업 슬롯 수', '5'))
        if not 1 <= count <= 100:
            raise ValueError('작업 슬롯 수는 1~100이어야 합니다.')
        waiting = v12.csv_names(ask('본항 대기선 순서 (쉼표 구분; 온산이면 Enter)'))
        if port == '본항' and not waiting:
            raise ValueError('본항은 대기선 순서를 직접 입력해야 합니다.')
        busy = v12.csv_names(ask('작업 중 선박 (쉼표 구분; 없으면 Enter)'))
        unavailable = v12.csv_names(ask('작업 불가 선박 (쉼표 구분; 없으면 Enter)'))
        info = v12.read_day(excel, day)
        result, excluded = v12.generate(info, port, count, waiting, unavailable, busy)
        metrics = v13.evaluate(info, result, excluded, busy, unavailable)
        output = ROOT / f'v14_통합평가_{day}_{port}.xlsx'
        v13.make_report(output, day, port, info, result, excluded, metrics, busy, unavailable)
    except (ValueError, OSError) as exc:
        print('입력/처리 오류:', exc)
        return
    print(f'\n원본: {info["sheet"]} / {info["row"]}행')
    print('당직선:', ', '.join(name for name, _ in info['trio']))
    print('비고:', info['note'] or '없음')
    for num, name, *_ in result:
        print(f'{num:>2}. {name}')
    print('\n평가지표 (실제 정확도가 아님):')
    for key, value in metrics.items():
        print(f' - {key}: {value}')
    print('보고서 저장:', output)
    print('※ 매뉴얼 수정안 잠정 적용 / 실시간 운항 조건 미확인 / 실운항 금지')


def main():
    print('TUG v14 — 통합 메뉴 (교육·검토용, 실운항 금지)')
    while True:
        print('\n1. YES/NO 질문 판단 (기존 v9 실행)')
        print('2. 당직표 후보 생성 + 평가지표 + 엑셀 저장 (v12/v13 연동)')
        print('0. 종료')
        try:
            choice = input('선택: ').strip()
            if choice == '1':
                run_agent()
            elif choice == '2':
                run_evaluation()
            elif choice in ('0', 'q', 'Q'):
                print('종료합니다.')
                break
            else:
                print('0, 1, 2 중 선택해 주세요.')
        except (KeyboardInterrupt, EOFError):
            print('\n종료합니다.')
            break


if __name__ == '__main__':
    main()
