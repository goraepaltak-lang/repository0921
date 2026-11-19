from pathlib import Path
import json
import numpy as np
import faiss
from ollama import embed

ROOT = Path(__file__).resolve().parent
# 확인된 규칙으로 시작하는 시제품. 전체 매뉴얼은 추후 추출하여 추가.
rules = [
    {"source": "매뉴얼 규칙 요약: 평일 오전", "text": "평일 오전에는 작업 건수가 적은 순서로 선박을 우선 배정한다."},
    {"source": "매뉴얼 규칙 요약: 평일 오후 및 주말", "text": "평일 오후와 주말에는 당직 순번을 기준으로 선박을 배정한다."},
]

vectors = []
for i, rule in enumerate(rules, 1):
    print(f"임베딩 중 {i}/{len(rules)}: {rule['source']}", flush=True)
    result = embed(model="embeddinggemma", input=rule["text"])
    vectors.append(result.embeddings[0])

matrix = np.asarray(vectors, dtype="float32")
faiss.normalize_L2(matrix)
index = faiss.IndexFlatIP(matrix.shape[1])
index.add(matrix)
faiss.write_index(index, str(ROOT / "manual.index"))
(ROOT / "manual_chunks.json").write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"벡터DB 생성 완료! {index.ntotal}개 규칙, {matrix.shape[1]}차원")

while True:
    question = input("\n검색 질문 (종료 q): ").strip()
    if question.lower() == "q":
        break
    if not question:
        continue
    q = np.asarray([embed(model="embeddinggemma", input=question).embeddings[0]], dtype="float32")
    faiss.normalize_L2(q)
    scores, ids = index.search(q, min(2, index.ntotal))
    for score, idx in zip(scores[0], ids[0]):
        print(f"유사도 {score:.3f} | {rules[int(idx)]['source']} | {rules[int(idx)]['text']}")
