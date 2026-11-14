# -*- coding: utf-8 -*-
"""data 폴더의 매뉴얼(docx/pdf/txt/md) -> manual_chunks.json + manual.index (FAISS)

- 매뉴얼을 제목/조문 단위로 나누고, 한 덩어리가 MAX_CHARS를 넘으면 더 쪼갠다. (쪼갠 덩어리마다 제목을 다시 붙임)
- 표(docx)는 행 단위로 '칸1 | 칸2 | 칸3' 텍스트가 된다.
- 임베딩 모델은 agent의 manual_search와 반드시 같아야 한다 (embeddinggemma).
- 실행 전: ollama pull embeddinggemma   (Ollama 앱이 켜져 있어야 함)
- 필요: pip install faiss-cpu numpy ollama python-docx pypdf
- 실행: python build_vectors.py
"""
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA = ROOT / 'data'
INDEX = ROOT / 'manual.index'
CHUNKS = ROOT / 'manual_chunks.json'

EMBED_MODEL = 'embeddinggemma'
MAX_CHARS = 500
BATCH = 16

# txt/pdf에서 제목으로 볼 줄: '제3조', '1.', '가.', '■', '▶', '[제목]' 등으로 시작
HEAD_RE = re.compile(r'^(제\s*\d+\s*[조장절항]|\d+[.)]\s|[가-하][.)]\s|[■▶●◆]|\[[^\]]{1,30}\]|#{1,6}\s)')


def read_text_file(path):
    for enc in ('utf-8-sig', 'cp949'):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f'{path.name}: 인코딩을 읽을 수 없습니다(utf-8/cp949 아님)')


def lines_from_text(text):
    out = []
    for raw in text.splitlines():
        line = raw.strip()
        if line:
            out.append((line.lstrip('#').strip() if line.startswith('#') else line,
                        bool(HEAD_RE.match(line))))
    return out


def lines_from_docx(path):
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    doc = Document(str(path))
    out = []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith('}p'):
            p = Paragraph(child, doc)
            t = p.text.strip()
            if t:
                style = (p.style.name or '') if p.style is not None else ''
                out.append((t, style.startswith(('Heading', '제목')) or bool(HEAD_RE.match(t))))
        elif child.tag.endswith('}tbl'):
            for row in Table(child, doc).rows:
                cells, seen = [], set()
                for c in row.cells:            # 병합 셀 중복 제거
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    t = c.text.strip().replace('\n', ' ')
                    if t:
                        cells.append(t)
                if cells:
                    out.append((' | '.join(cells), False))
    return out


def lines_from_pdf(path):
    from pypdf import PdfReader
    text = '\n'.join((page.extract_text() or '') for page in PdfReader(str(path)).pages)
    return lines_from_text(text)


# 엑셀 중 '규칙/설명'이 적힌 시트만 매뉴얼로 읽는다 (일지·당직표 같은 데이터 시트는 build_db.py가 처리)
RULE_SHEET_KEYS = ('준수사항', '이용방법', '규정', '매뉴얼', '사용방법')


def lines_from_xlsx_rules(path):
    from openpyxl import load_workbook
    out = []
    wb = load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        if not any(k in ws.title for k in RULE_SHEET_KEYS):
            continue
        out.append((ws.title, True))
        for row in ws.iter_rows(values_only=True):
            cells = [str(v).replace('\n', ' ').strip() for v in row if v not in (None, '')]
            if not cells:
                continue
            # 첫 칸이 '1', '2' 같은 큰 항목 번호면 제목으로 본다 ('1.1' 같은 하위 항목은 본문)
            out.append((' '.join(cells), bool(re.fullmatch(r'\d+', cells[0])) and len(cells) > 1))
    wb.close()
    return out


def read_manual(path):
    ext = path.suffix.lower()
    if ext == '.xlsx':
        return lines_from_xlsx_rules(path)
    if ext == '.docx':
        return lines_from_docx(path)
    if ext == '.pdf':
        return lines_from_pdf(path)
    return lines_from_text(read_text_file(path))


def make_chunks(lines, max_chars=MAX_CHARS):
    """[(텍스트, 제목여부)] -> 청크 문자열 리스트. 제목 줄에서 새 청크를 시작한다."""
    chunks, head, buf, size = [], '', [], 0

    def flush():
        nonlocal buf, size
        if buf and not (len(buf) == 1 and buf[0] == head):     # 제목만 있는 청크는 버림
            chunks.append('\n'.join(([head] if head and buf[0] != head else []) + buf))
        buf, size = [], 0

    for text, is_head in lines:
        if is_head:
            flush()
            head, buf, size = text, [text], len(text)
            continue
        if buf and size + len(text) > max_chars:
            flush()
        buf.append(text)
        size += len(text)
    flush()

    seen, unique = set(), []
    for c in chunks:                                            # 중복 제거
        if c not in seen:
            seen.add(c)
            unique.append(c)
    return unique


def main():
    import faiss
    import numpy as np
    from ollama import embed

    DATA.mkdir(exist_ok=True)
    files = sorted(p for p in DATA.iterdir()
                   if p.suffix.lower() in ('.docx', '.pdf', '.txt', '.md', '.xlsx') and not p.name.startswith('~$'))
    hwp = sorted(p.name for p in DATA.iterdir() if p.suffix.lower() in ('.hwp', '.hwpx'))
    if hwp:                                # 한글 파일은 읽지 못함 → 조용히 빠지지 않도록 알림
        print(f'건너뜀(한글 파일 미지원): {", ".join(hwp)} → 한글에서 .docx 또는 .pdf로 저장해 data 폴더에 넣으세요.')
    if not files:
        print(f'{DATA} 폴더에 매뉴얼(.docx/.pdf/.txt/.md)이 없습니다. 파일을 넣고 다시 실행하세요.')
        return

    chunks = []
    for p in files:
        c = make_chunks(read_manual(p))
        if p.suffix.lower() == '.xlsx' and not c:
            continue                       # 규칙 시트가 없는 엑셀(일지, 당직표)은 건너뜀
        print(f'{p.name}: 청크 {len(c)}개' + ('  ※ 글자가 거의 없음 (스캔 PDF?)' if p.suffix.lower() == '.pdf' and len(c) < 3 else ''))
        chunks += c
    if not chunks:
        print('읽을 수 있는 매뉴얼이 없습니다. (.hwp는 .docx로 저장 / 스캔 PDF는 글자 추출 불가)')
        return

    vecs = []
    try:
        for i in range(0, len(chunks), BATCH):
            vecs += embed(model=EMBED_MODEL, input=chunks[i:i + BATCH]).embeddings
    except Exception as exc:
        print(f'임베딩 실패: {exc}\nOllama 앱이 켜져 있는지, `ollama pull {EMBED_MODEL}`을 했는지 확인하세요.')
        return

    mat = np.array(vecs, dtype=np.float32)
    faiss.normalize_L2(mat)
    index = faiss.IndexFlatIP(mat.shape[1])
    index.add(mat)
    faiss.write_index(index, str(INDEX))
    CHUNKS.write_text(json.dumps(chunks, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'완료: 청크 {len(chunks)}개 -> {INDEX.name}, {CHUNKS.name}')


if __name__ == '__main__':
    main()