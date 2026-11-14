# -*- coding: utf-8 -*-
"""TUG YES/NO agent: SQLite + FAISS + Ollama, with compact grounded evidence."""
import json
import re
import sqlite3
from pathlib import Path

import faiss
import numpy as np
from ollama import chat, embed

ROOT = Path(__file__).resolve().parent
DB = ROOT / 'tug_agent.db'
INDEX = ROOT / 'manual.index'
CHUNKS = ROOT / 'manual_chunks.json'


def db_search(question):
    terms = re.findall(r'[가-힣A-Za-z]*\d+[가-힣A-Za-z0-9]*', question)
    if not terms:
        return []
    with sqlite3.connect(DB) as conn:
        found = {}
        for term in terms[:3]:
            for row in conn.execute(
                'SELECT source_file,sheet_name,row_number,row_text '
                'FROM excel_rows WHERE row_text LIKE ? LIMIT 10',
                (f'%{term}%',),
            ):
                found[row[:3]] = row
        return list(found.values())[:5]


def manual_search(question, index, chunks):
    result = embed(model='embeddinggemma', input=question)
    v = np.array([result.embeddings[0]], dtype=np.float32)
    faiss.normalize_L2(v)
    scores, ids = index.search(v, min(10, index.ntotal))
    return [
        (float(score), {'source': '매뉴얼', 'text': str(chunks[int(i)])})
        for score, i in zip(scores[0], ids[0])
        if 0 <= int(i) < len(chunks)
    ]


def question_terms(question):
    """Basic lexical hints; semantic search remains the primary retrieval method."""
    words = re.findall(r'[가-힣A-Za-z0-9]+', question.lower())
    stop = {'에는', '에서', '으로', '하는', '하나요', '인가요', '있나요',
            '맞나요', '선박을', '선박', '여부', '대한', '따르면', '매뉴얼',
            '규칙', '어떻게', '배정', '하나요'}
    return [w for w in words if len(w) >= 2 and w not in stop]


def compact_evidence(question, matches, rows):
    """Pick short, question-related passages instead of entire long tables."""
    terms = question_terms(question)
    snippets = []
    seen = set()
    for score, item in matches[:8]:
        lines = [line.strip() for line in item['text'].splitlines() if line.strip()]
        ranked = []
        for pos, line in enumerate(lines):
            low = line.lower().replace(' ', '')
            hits = sum(1 for word in terms if word.replace(' ', '') in low)
            # Keep a nearby heading for context; don't treat headings alone as proof.
            if hits:
                context = ' / '.join(lines[max(0, pos - 1):min(len(lines), pos + 2)])
                ranked.append((hits, pos, context[:380]))
        ranked.sort(key=lambda x: (-x[0], x[1]))
        for _, _, snippet in ranked[:3]:
            if snippet not in seen:
                seen.add(snippet)
                snippets.append((score, snippet))
    snippets.sort(key=lambda x: -x[0])
    # Preserve an explicit heading + rule when both are present in retrieved chunks.
    # This avoids losing a short rule buried among long tabular passages.
    qnorm = question.replace(' ', '')
    if all(x in qnorm for x in ('평일오후', '주말', '당직')):
        for _, item in matches:
            lines = [x.strip() for x in item['text'].splitlines() if x.strip()]
            for i, line in enumerate(lines):
                norm = line.replace(' ', '')
                if '평일오후' in norm and '주말' in norm:
                    near = lines[i:i+4]
                    if any('당직순번' in x.replace(' ', '') and '배정' in x for x in near):
                        snippet = ' / '.join(near[:3])[:300]
                        snippets.insert(0, (2.0, snippet))
                        break
            else:
                continue
            break
    evidence_parts = [f'[엑셀 {f}/{s}/{n}행] {str(t)[:200]}' for f, s, n, t in rows[:3]]
    evidence_parts += [f'[매뉴얼 검색근거] {s}' for _, s in snippets[:4]]
    # When no lexical match exists, provide one short retrieved passage rather than invent evidence.
    if not evidence_parts and matches:
        evidence_parts.append('[매뉴얼 검색근거] ' + matches[0][1]['text'][:450])
    return '\n'.join(evidence_parts)[:1500]



def unsupported_live_data(question, evidence):
    """Avoid answering live environmental measurements from unrelated tug manuals."""
    live_terms = ('파고', '파도 높이', '풍속', '기온', '수온', '현재 날씨', '실시간 날씨')
    normalized_evidence = evidence.replace(' ', '')
    return any(term.replace(' ', '') in question.replace(' ', '')
               and term.replace(' ', '') not in normalized_evidence
               for term in live_terms)


def unsupported_no_reason(verdict, reason):
    """A missing fact is not evidence of its negation."""
    if verdict != 'NO':
        return False
    normalized = reason.replace(' ', '')
    missing = ('정보없', '근거없', '자료없', '증거없', '확인불가', '명시되지않',
               '언급되지않', '확인되지않', '찾을수없', '판단불가', '불충분')
    return any(word in normalized for word in missing)

def verified_duty_order_rule(question, matches):
    """Resolve only the explicitly documented weekday-afternoon/weekend duty-order rule.

    Require the actual retrieved manual text, not just an LLM-generated reason.
    Other questions still go through the LLM.
    """
    q = re.sub(r"\s+", "", question)
    if not ("평일오후" in q and "주말" in q and "당직순번" in q):
        return None
    if not ("배정" in q or "사용" in q):
        return None

    verified = False
    for _, item in matches:
        lines = [line.strip() for line in item['text'].splitlines() if line.strip()]
        for i, line in enumerate(lines):
            window = re.sub(r"\s+", "", " ".join(lines[i:i+4]))
            if ("평일오후" in window and "주말" in window
                    and "당직순번" in window and "배정" in window
                    and ("기준" in window or "따라" in window)):
                verified = True
                break
        if verified:
            break
    if not verified:
        return None

    # This checks the polarity of this specific proposition, not arbitrary negations.
    negative_forms = (
        "사용하지않", "사용안", "사용하지않고", "기준으로하지않",
        "기준으로하지않고", "따르지않", "따르지않고", "적용하지않",
        "배정하지않", "무관하게", "관계없이", "기준이아닌",
    )
    if any(form in q for form in negative_forms):
        return ('NO', '매뉴얼은 평일 오후와 주말에 당직순번 기준 배정을 명시합니다.')
    # Avoid claiming YES for other complex or negative propositions.
    positive_forms = ("당직순번으로", "당직순번을기준으로", "당직순번기준", "당직순번에따라")
    if any(form in q for form in positive_forms) and not any(
            form in q for form in ("아닌가", "없나", "못하", "금지", "제외", "다른")):
        return ('YES', '매뉴얼에 평일 오후와 주말의 당직순번 기준 배정이 명시되어 있습니다.')
    return None



def dated_record_check(question):
    """Check an explicit YYYY-MM-DD + vessel status against one dated duty-table row.

    Returns None for questions outside this narrow, verifiable pattern.
    Missing records are UNKNOWN, not automatically NO.
    """
    date = re.search(r'(20\d{2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일', question)
    record = re.search(r'([가-힣A-Za-z]+\d+)\s*(수리|복귀)', question)
    if not date or not record or not any(x in question for x in ('기록', '적혀', '있나요', '있습니까')):
        return None
    year, month, day = map(int, date.groups())
    vessel, status = record.groups()
    date_label = f'{month}/{day}'
    with sqlite3.connect(DB) as conn:
        candidates = conn.execute(
            'SELECT source_file,sheet_name,row_number,row_text FROM excel_rows '
            'WHERE sheet_name LIKE ? AND row_text LIKE ?',
            (f'%{year}년 {month}월%', f'{date_label}%'),
        ).fetchall()
    rows = [r for r in candidates if re.match(r'^\s*' + re.escape(date_label) + r'\s*(?:\||$)', r[3])]
    if not rows:
        return ('UNKNOWN', f'{year}년 {month}월 {day}일의 당직표 행을 확인하지 못했습니다.', [])
    target = re.compile(re.escape(vessel) + r'\s*' + re.escape(status) + r'(?![가-힣A-Za-z0-9])')
    matched = [r for r in rows if target.search(r[3])]
    if matched:
        return ('YES', f'{date_label} 당직표에 {vessel} {status} 기록이 있습니다.', matched)
    return ('UNKNOWN', f'{date_label} 당직표에서 {vessel} {status} 기록을 확인하지 못했습니다.', rows)


def main():
    print('=== TUG YES/NO 에이전트 v9 ===')
    missing = [p.name for p in (DB, INDEX, CHUNKS) if not p.exists()]
    if missing:
        print('필요한 파일이 없습니다:', ', '.join(missing))
        print('이 파일을 기존 YESNO_AGENT 폴더에 넣어 주세요.')
        return
    index = faiss.read_index(str(INDEX))
    chunks = json.loads(CHUNKS.read_text(encoding='utf-8'))
    if index.ntotal != len(chunks):
        print('인덱스와 규칙 개수가 다릅니다. build_vectors.py를 다시 실행하세요.')
        return
    print(f'준비 완료: DB + FAISS 규칙 {len(chunks)}개 (전체 매뉴얼 아님)')
    schema = {
        'type': 'object',
        'properties': {
            'verdict': {'type': 'string', 'enum': ['YES', 'NO', 'UNKNOWN']},
            'reason': {'type': 'string'},
        },
        'required': ['verdict', 'reason'],
        'additionalProperties': False,
    }
    while True:
        try:
            question = input('\n질문 (종료: q): ').strip()
        except (EOFError, KeyboardInterrupt):
            print('\n종료합니다.')
            break
        if question.lower() == 'q':
            break
        if not question:
            continue
        try:
            dated = dated_record_check(question)
            if dated is not None:
                verdict, reason, source_rows = dated
                print(f'\n[RDB 날짜별 정확 검색] {len(source_rows)}건')
                for f, sheet, n, row_text in source_rows[:2]:
                    print(f'[엑셀 {f}/{sheet}/{n}행] {row_text[:350]}')
                print(f'\n판정: {verdict}\n근거: {reason}')
                print('※ 날짜별 당직표의 기록 존재 여부만 확인한 결과입니다.')
                continue
            rows = db_search(question)
            matches = manual_search(question, index, chunks)
            print(f'\n[RDB 검색] {len(rows)}건 / [FAISS 검색] {len(matches)}건')
            evidence = compact_evidence(question, matches, rows)
            if not evidence:
                print('판정: UNKNOWN\n근거: 검색된 자료가 없습니다.')
                continue
            print('\n[AI에게 전달되는 근거]\n' + evidence)
            if unsupported_live_data(question, evidence):
                print('\n판정: UNKNOWN\n근거: 검색 자료에 해당 실시간 관측 정보가 없습니다.')
                continue
            verified_rule = verified_duty_order_rule(question, matches)
            if verified_rule is not None:
                verdict, reason = verified_rule
                print(f'\n판정: {verdict}\n근거: {reason}')
                print('※ 매뉴얼에서 확인한 규칙에 대한 판정입니다. 원본과 대조하세요.')
                continue
            print('\n[Qwen3 4B 판단 중...]', flush=True)
            response = chat(
                model='qwen3:4b',
                messages=[
                    {'role': 'system', 'content': (
                        '당신은 선박 배정 매뉴얼의 사실 확인 담당자입니다. '
                        '오직 아래 제공된 증거에만 근거해 YES, NO, UNKNOWN 중 하나로 답하세요. '
                        '질문의 전체 의미(부정문 포함)를 먼저 파악하세요. 질문의 주장과 증거를 비교하세요. '
                        '질문을 직접 뒷받침하는 규정이 있으면 YES, 직접 반박하면 NO, '
                        '둘 다 없거나 모호하면 UNKNOWN입니다. '
                        '특히 자료에 정보가 없다는 것은 NO의 근거가 아닙니다. '
                        'NO는 질문을 명시적으로 부정하는 증거가 있을 때만 선택하세요. '
                        '예: 증거에 "평일 오후 / 주말 - 당직순번 기준 배정"이라고 적혀 있고 '
                        '질문이 "평일 오후와 주말에는 당직 순번으로 배정하나요?"라면 YES입니다. '
                        '같은 증거에 질문이 \"당직순번을 사용하지 않고 배정하나요?\"라면 NO입니다. '
                        '단, 이 예시를 다른 질문의 근거로 사용하지 마세요. '
                        '선박 이름이 언급됐다는 사실만으로 당직표 등록을 단정하지 마세요. '
                        'reason은 35자 이내의 한국어 한 문장으로 쓰고 JSON만 출력하세요.'
                    )},
                    {'role': 'user', 'content': f'질문: {question}\n\n증거:\n{evidence}'},
                ],
                format=schema,
                think=False,
                options={'temperature': 0, 'num_predict': 320, 'num_ctx': 3072},
            )
            raw = response.message.content or '{}'
            data = json.loads(raw)
            verdict = data.get('verdict')
            reason = str(data.get('reason', '')).strip()
            if verdict not in ('YES', 'NO', 'UNKNOWN') or not reason:
                raise ValueError('모델 출력이 불완전합니다.')
            if unsupported_no_reason(verdict, reason):
                verdict = 'UNKNOWN'
                reason = '제공된 자료만으로 질문의 참거짓을 판단할 수 없습니다.'
            print(f'\n판정: {verdict}\n근거: {reason[:240]}')
            print('※ AI 판정은 반드시 원본 매뉴얼 및 실제 당직표와 대조하세요.')
        except (ValueError, TypeError) as exc:
            print('판정: UNKNOWN\n근거: 모델 출력 형식 확인 실패:', exc)
        except Exception as exc:
            print('실행 오류:', exc)


if __name__ == '__main__':
    main()
