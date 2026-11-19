"""예선 배정 순서 자동화 v2
 - 도선예보 엑셀의 정형 컬럼(TIME, G/T, FROM, TO, T …)은 Python이 직접 읽는다.
 - REMARKS 자유 텍스트(마력 요청, 선후 작업 조건 등)만 Qwen이 구조화한다.
 - 배정 순서는 Python 규칙 엔진이 결정한다(근거 출력).

흐름:  Crawling.py → output/도선예보_날짜.xlsx + tug_agent.db(forecast 테이블) → Tug_Agent.py가 DB에서 읽음
사용:  python tug_agent.py                           (DB에서 오늘 '미완료' 작업 읽기, 없으면 Crawling 먼저 실행, Qwen 사용)
       python tug_agent.py --crawl                  (DB에 있어도 새로 크롤링해서 최신 예보로 배정)
       python tug_agent.py --no-crawl               (크롤링 안 함. 미완료가 없으면 전날 수집한 '예정' 사용)
       python tug_agent.py --no-llm                 (REMARKS 해석 없이 규칙만)
       python tug_agent.py --brief                  (배정 근거 설명 생략, 결과만)
       python tug_agent.py --no-manual              (매뉴얼 근거 검색 끔. 켜려면 Build_vectors.py로 manual.index 먼저 생성)
       python tug_agent.py --date 2026-11-06        (DB에서 특정 날짜 읽기, 크롤링은 오늘만 가능)
       python tug_agent.py 도선예보.xlsx            (DB 대신 엑셀 파일 직접 지정)
필요:  pip install -U ollama pydantic openpyxl faiss-cpu numpy   (매뉴얼 근거: ollama pull embeddinggemma)
"""

"""
캡스톤 디자인 프로젝트 폴더 내에 output 폴더에서 오늘 날짜 도선 예보 excel을 읽는다.

배정 판단 흐름
① 배가 어디서 어디로 가나? (FROM / TO)
FROM이 P/S면 → 입항(접안), 기준 부두는 TO입니다.
TO가 P/S면 → 출항(이안), 기준 부두는 FROM입니다.
둘 다 아니면 → 이선입니다.
TO가 묘박지(E-1~E-9)면 기준 부두는 FROM입니다.
아니면 기준 부두는 TO입니다.

↓

② 기준 부두는 어느 항인가?
부두가 SILO, COAL로 시작하면 → 미분류입니다. 배정하지 않고 "항 확인 필요"로 표시한 뒤 끝냅니다.
P#, GP, HCP, YMP, YEP, H#, HMD로 시작하면 → 본항 보드입니다.
그 밖의 부두는 모두 → 온산 보드입니다(장생포 포함).

↓

③ 톤수(G/T)는 얼마인가?
1만 톤 미만이면 → 1000HP급 예선입니다.
1만 톤 이상 5만 톤 미만이면 → 2000HP급 예선입니다.
5만 톤 이상이면 → 3000HP급인데, 아직 미구현이라 배정하지 않고 끝냅니다.

②와 ③을 합치면 쓸 보드가 정해집니다. 예: 온산/1000HP.

↓

④ REMARKS에 마력 요청이 있나?
있으면(예: “1500HP 이상”) → 그 마력에 못 미치는 예선은 후보에서 뺍니다.
없으면 → 그대로 다음 질문으로 갑니다.
“OO 이안 후”, "재연락" 같은 문구는 메모로만 남기고 배정에는 영향을 주지 않습니다.

↓

⑤ 예선이 지금 다른 작업 중인가?
이전 작업이 끝나는 시각이 이번 작업 시작 시각보다 늦으면 → 작업 중이라 후보에서 뺍니다.
끝나는 시각과 시작 시각이 같으면 → 투입할 수 있습니다.

↓

⑥ 언제 하는 작업인가? (매뉴얼 Ⅲ-2)
평일 06:00~11:59 → 처진순: 당직+대기 중 건수 적은 순 → 전날 당직 처진순 → 상위마력 하향지원
평일 12:00~17:29 / 주말 06:00~15:29 → 당직순: 당일 당직(직전 배정 다음 순번부터 순환) → 당일 대기
    → 전날 당직(13:30 정리 전까지만) → 동일 항내 상위마력 → 타항 지원(다른 항의 같은 마력 보드)
평일 17:30~05:59 / 주말 15:30~05:59 → 야간: 당일 당직선 번갈아 배정
    부족 시 대응은 매뉴얼에 없음 → 주간 규칙(대기 → 상위마력 → 타항)을 준용하고 '사람 확인' 표시

↓

⑦ 필요한 척수(T)를 다 채웠나?
채웠으면 → 배정을 확정합니다.
위 단계를 다 써도 모자라면 → 뽑힌 만큼만 배정하고 "가용 예선 부족(n/m척)" 경고를 붙입니다.

↓

⑧ 배정 후 기록
배정된 예선의 건수를 1 올립니다.
작업이 끝나는 시각을 기록합니다. 시작 시각에 아래 예상 시간을 더한 값입니다.
"""
import re, sys
from dataclasses import dataclass, field
from datetime import datetime, time, date
from pathlib import Path
from typing import Optional
from pydantic import BaseModel

MODEL = "qwen3:8b"

# 날짜별 도선예보 엑셀이 모이는 폴더 / 파일명 규칙 (예: 도선예보_2026-11-09.xlsx)
OUTPUT_DIR = Path(r"C:\Users\gusgu\PycharmProjects\Capstone Design\output")

# ═════════ 설정 (현장 기준에 맞게 수정) ═════════
# 본항 부두 접두어. 여기에 없으면 온산 보드(장생포 포함)로 본다. SILO/COAL 등 모호한 건 UNKNOWN_PREFIX.
BONHANG_PREFIX = ("P#", "GP", "HCP", "YMP", "YEP", "H#", "HMD")
UNKNOWN_PREFIX = ("SILO", "COAL")          # 소속 항이 문서에 없음 → 사람이 확인
ANCHORAGE = re.compile(r"\bE-\d\b")        # 묘박지(E-1~E-9). 'P/S(E-2)'처럼 괄호 안에 있어도 인식


def hp_group(gt: int) -> str:
    """[가정] G/T → 마력군. 문서에 기준이 없어 임시값임. 현장 기준으로 반드시 교체."""
    if gt < 10_000: return "1000"
    if gt < 50_000: return "2000"
    return "3000"      # 3천마력 규칙은 아직 미구현


def find_forecast_file(target: date, folder: Path = OUTPUT_DIR) -> Path:
    """folder 안에서 target 날짜(YYYY-MM-DD)가 이름에 들어간 .xlsx를 찾는다."""
    if not folder.is_dir():
        raise FileNotFoundError(f"폴더를 찾을 수 없음: {folder}")
    stamp = target.strftime("%Y-%m-%d")
    hits = sorted(p for p in folder.glob("*.xlsx")
                  if stamp in p.name and not p.name.startswith("~$"))   # ~$ = 엑셀 임시 잠금파일
    if not hits:
        avail = sorted(p.name for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))[-5:]
        raise FileNotFoundError(f"{stamp} 날짜 엑셀이 없음: {folder}\n  최근 파일: {avail or '없음'}")
    if len(hits) > 1:
        print(f"[주의] {stamp} 파일이 {len(hits)}개 → 가장 최근 수정본 사용: "
              f"{[p.name for p in hits]}", file=sys.stderr)
        hits.sort(key=lambda p: p.stat().st_mtime)
    return hits[-1]




# ═════════ 1. 도선예보 엑셀 → Job ═════════
@dataclass
class Job:
    no: int; time: time; ship: str; gt: int; frm: str; to: str
    tugs: int; ca: str; remarks: str
    kind: str = ""; port: str = ""; berth: str = ""
    requested_hp: Optional[int] = None      # REMARKS에서 추출 (Qwen)
    after_ship: Optional[str] = None        # "~ 이안 후 접안" 같은 선행 조건 (Qwen)
    notes: list = field(default_factory=list)


def classify(j: Job):
    ps = lambda x: x.upper().startswith("P/S")      # 'P/S', 'P/S(E-2)' 모두 도선점으로 본다
    if ps(j.frm):        j.kind, j.berth = "입항(접안)", j.to
    elif ps(j.to):       j.kind, j.berth = "출항(이안)", j.frm
    else:                j.kind, j.berth = "이선", (j.to if not ANCHORAGE.search(j.to) else j.frm)
    b = j.berth.upper()
    if b.startswith(UNKNOWN_PREFIX): j.port = "미분류"
    elif b.startswith(BONHANG_PREFIX): j.port = "본항"
    else: j.port = "온산"


def _int(v) -> int:
    s = str(v or "").replace(",", "").strip()
    try: return int(float(s)) if s else 0
    except ValueError: return 0


def make_job(rec: dict) -> Optional[Job]:
    """도선예보 한 행(엑셀 열 이름 → 값) → Job. NO/TIME이 없으면 None"""
    s = lambda k: str(rec.get(k) or "").strip()
    if not _int(rec.get("NO")) or ":" not in s("TIME"): return None
    h, m = map(int, s("TIME").split(":")[:2])
    j = Job(no=_int(rec["NO"]), time=time(h, m), ship=s("SHIP'S NAME"), gt=_int(rec.get("G/T")),
            frm=s("FROM"), to=s("TO"), tugs=_int(rec.get("T")), ca=s("CA"), remarks=s("REMARKS"))
    classify(j)
    return j


def load_jobs(path: str, sheet: Optional[str] = None) -> list[Job]:
    """엑셀 파일에서 직접 읽기 (파일을 인자로 지정했을 때만 사용)"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else next(w for w in wb if "미완료" in w.title)
    rows = list(ws.iter_rows(values_only=True))
    head = [str(h).strip() for h in rows[0]]
    return [j for r in rows[1:] if (j := make_job(dict(zip(head, r))))]


def load_jobs_db(target: date, force: bool = False, allow: bool = True) -> list[Job]:
    """tug_agent.db의 forecast 테이블에서 target 날짜 작업을 읽는다.
    - 오늘인데 '미완료'가 아직 저장 안 됐으면(또는 --crawl) Crawling.py를 실행 → DB에 저장된 뒤 읽음
    - '미완료'가 없으면 전날 크롤링의 '예정'을 대신 사용 (완료 여부 미반영)"""
    import Build_db as bdb
    stamp = target.strftime("%Y-%m-%d")
    if force or (allow and not bdb.has_forecast(stamp, "미완료")):
        if target != date.today():
            if force: raise ValueError(f"크롤러는 오늘({date.today()}) 예보만 수집 가능 (요청: {stamp})")
        else:
            print(f"[크롤링] {stamp} 도선예보 수집 → DB 저장 중…", file=sys.stderr)
            from Crawling import main as crawl      # 지연 import: 크롤링 안 할 땐 playwright 불필요
            crawl()
    kind, crawled_at, recs = bdb.load_forecast(stamp)
    if kind is None:
        raise LookupError(f"DB({bdb.DB})에 {stamp} 도선예보가 없음. Crawling.py를 실행하거나 --crawl 사용")
    print(f"[DB] {stamp} '{kind}' {len(recs)}건 (수집 {crawled_at})"
          + ("  ※ 전날 수집한 예정 목록 → 이미 끝난 작업이 섞여 있을 수 있음" if kind == "예정" else ""))
    return [j for r in recs if (j := make_job(r))]


# ═════════ 2. Qwen: REMARKS 해석만 담당 ═════════
class RemarkInfo(BaseModel):
    requested_hp: Optional[int] = None   # 예: "5,000HP 이상" → 5000
    after_ship: Optional[str] = None     # 예: "PRINCESS SANGO 이안 후 접안" → "PRINCESS SANGO"
    needs_confirm: bool = False          # "재연락", "대리점 확인 후" 등 미확정 표현


SYSTEM_PROMPT = """너는 울산항 도선예보 REMARKS 해석기다. REMARKS 문장에서 아래 필드만 뽑아 JSON으로 답하라.
- requested_hp: 'N HP 이상' 같은 마력 요청 숫자. 없으면 null
- after_ship: 특정 선박의 작업 '이후'에 해야 한다는 조건이 있으면 그 선박 이름. 없으면 null
- needs_confirm: '재연락', '대리점 확인 후', '협의 후' 등 확정되지 않은 표현이 있으면 true
예선 이름을 정하거나 배정을 판단하지 마라. 문장에 없는 값은 null/false. 전화번호는 무시."""


def parse_remarks(text: str) -> RemarkInfo:
    import ollama
    r = ollama.chat(model=MODEL, think=False, options={"temperature": 0},
                    format=RemarkInfo.model_json_schema(),
                    messages=[{"role": "system", "content": SYSTEM_PROMPT},
                              {"role": "user", "content": text}])
    return RemarkInfo.model_validate_json(r["message"]["content"])


def enrich(jobs: list[Job], use_llm: bool = True):
    for j in jobs:
        if not (use_llm and j.remarks.strip()): continue
        info = parse_remarks(j.remarks)
        j.requested_hp, j.after_ship = info.requested_hp, info.after_ship
        if info.needs_confirm: j.notes.append("미확정(재연락/확인 필요)")
        if info.after_ship: j.notes.append(f"선행작업: {info.after_ship}")


# ═════════ 2-1. 매뉴얼 근거 검색 (Build_vectors.py가 만든 FAISS 인덱스) ═════════
# 배정 결정에는 쓰지 않는다. 규칙 엔진이 적용한 규칙마다 매뉴얼 원문을 찾아 '근거'로 붙이기만 한다.
ROOT = Path(__file__).resolve().parent
MANUAL_INDEX, MANUAL_CHUNKS = ROOT / "manual.index", ROOT / "manual_chunks.json"
EMBED_MODEL = "embeddinggemma"     # Build_vectors.py와 반드시 같은 모델
MIN_SCORE = 0.35                   # 코사인 유사도가 이보다 낮으면 '근거 못 찾음' 처리 (결과 보고 조정)
SNIPPET = 300                      # 맨 끝 근거 목록에 보여줄 원문 길이

TOP_K = 10                         # 상위 몇 개 후보 중에서 핵심어가 들어간 조각을 고를지

# 규칙 종류 → 근거 조각에 반드시 들어 있어야 할 핵심어 (하나라도). 엉뚱한 조각이 근거로 붙는 것을 막는다.
RULE_KEYWORDS = {
    "보드":     ("부두명", "부두 현황"),
    "처진순":   ("처진", "쳐진"),          # 현장 문서에 '쳐지다' 표기도 섞여 있음
    "당직순":   ("당직 순번", "당직순번"),
    "마력요청": ("마력 요청",),
    "상위마력": ("상위마력",),
    "작업시간": ("TON",),                   # 'Ⅲ-6 작업 예상 시간표'의 G/T 구간 행
    "예선부족": ("타항",),
}

# 규칙 종류 → 매뉴얼 검색 문장 (매뉴얼 표현에 가깝게 고칠수록 정확해짐)
RULE_QUERIES = {
    "보드":     "울산항 부두 현황 본항 장생포항 온산항 부두명 표기 호칭",
    "처진순":   "평일 오전 06:00 ~ 11:59 실시간 건수일지를 통해 처진순 배정",
    "당직순":   "평일 오후 및 주말 당직 순번 배정 당일 당직선 번갈아 가면서 배정",
    "마력요청": "특정 마력 요청 시 해당 마력 지원 및 청구",
    "상위마력": "당일 당직선박 당일 대기선박 동일 항내 상위마력 투입",
    "작업시간": "작업 예상 시간표 G/T 접안 출항 9,999TON 1시간 15~30분",
    "예선부족": "동일 항내 상위마력 투입 타항 지원",
}


class ManualRefs:
    """규칙별 매뉴얼 근거를 한 번씩만 검색해 각주 번호([M1], [M2] …)로 관리한다."""
    def __init__(self, enabled: bool = True):
        self.enabled = enabled
        self.by_key: dict = {}        # 규칙 key → 각주 번호(없으면 None)
        self.notes: list = []         # [(번호, 원문, 점수, 규칙 key들)]
        self._index = self._chunks = None
        if enabled and not (MANUAL_INDEX.exists() and MANUAL_CHUNKS.exists()):
            self._off(f"{MANUAL_INDEX.name}/{MANUAL_CHUNKS.name} 없음 → Build_vectors.py 먼저 실행")

    def _off(self, why):
        print(f"[매뉴얼 근거 끔] {why}", file=sys.stderr)
        self.enabled = False

    def search(self, query: str, k: int = 1) -> list[tuple[str, float]]:
        import json, faiss, numpy as np
        from ollama import embed
        if self._index is None:
            self._index = faiss.read_index(str(MANUAL_INDEX))
            self._chunks = json.loads(MANUAL_CHUNKS.read_text(encoding="utf-8"))
        v = np.array(embed(model=EMBED_MODEL, input=[query]).embeddings, dtype=np.float32)
        faiss.normalize_L2(v)
        D, I = self._index.search(v, k)
        return [(self._chunks[i], float(d)) for d, i in zip(D[0], I[0]) if i >= 0]

    def cite(self, key: str) -> str:
        """규칙 key의 각주 표시(' 📖[M1]'). 근거가 없거나 꺼져 있으면 ''"""
        if not self.enabled: return ""
        if key not in self.by_key:
            try:
                hits = self.search(RULE_QUERIES[key], k=TOP_K)
            except Exception as e:                    # faiss/ollama 미설치, Ollama 꺼짐 등
                self._off(f"{type(e).__name__}: {e}"); return ""
            n = None
            kws = RULE_KEYWORDS.get(key, ())
            hit = next(((t, sc) for t, sc in hits                 # 유사도 + 핵심어 포함 둘 다 만족하는 첫 조각
                        if sc >= MIN_SCORE and (not kws or any(k in t for k in kws))), None)
            if hit:
                text, score = hit
                same = next((x for x in self.notes if x[1] == text), None)   # 같은 조각이면 번호 공유
                if same: same[3].append(key); n = same[0]
                else: n = len(self.notes) + 1; self.notes.append((n, text, score, [key]))
            self.by_key[key] = n
        n = self.by_key[key]
        return f" 📖[M{n}]" if n else ""

    def report(self) -> list[str]:
        out = []
        for n, text, score, keys in self.notes:
            body = " ".join(text.split())
            out.append(f"[M{n}] ({', '.join(keys)} / 유사도 {score:.2f}) "
                       + (body[:SNIPPET] + "…" if len(body) > SNIPPET else body))
        miss = [k for k, n in self.by_key.items() if n is None]
        if miss: out.append(f"근거 못 찾음(유사도 {MIN_SCORE} 이상 + 핵심어 포함 조각 없음): {', '.join(miss)}")
        return out


# ═════════ 3. 규칙 엔진 (매뉴얼 Ⅲ-2) ═════════
@dataclass
class Tug:
    name: str; hp: int; count: int = 0; available: bool = True
    free_at: int = 0        # 이 시각(분, 0~1439)까지 작업 중
    busy: bool = False      # 현재 배정 대상 작업 시각에 작업 중인지

@dataclass
class Board:
    on_duty: list[Tug]
    standby: list[Tug] = field(default_factory=list)
    prev_duty: list[Tug] = field(default_factory=list)
    last_duty_idx: int = -1

def _ok(ts): return [t for t in ts if t.available and not t.busy]
def _lag(ts): return sorted(_ok(ts), key=lambda t: t.count)      # 처진순 = 건수 적은 순

# 매뉴얼 Ⅲ-2 시간대 구분
MORNING, DAY, NIGHT = "오전", "주간", "야간"

def time_slot(t: time, weekday: bool) -> str:
    """오전: 평일 06:00~11:59 (처진순)
    주간: 평일 12:00~17:29 / 주말 06:00~15:29 (당직순 → 대기 → 상위마력 → 타항)
    야간: 평일 17:30~05:59 / 주말 15:30~05:59 (당직선 번갈아 배정)"""
    if weekday:
        if time(6, 0) <= t < time(12, 0):  return MORNING
        if time(12, 0) <= t < time(17, 30): return DAY
        return NIGHT
    return DAY if time(6, 0) <= t < time(15, 30) else NIGHT

def is_morning(j_time: time, weekday: bool) -> bool:          # (하위 호환)
    return time_slot(j_time, weekday) == MORNING

PREV_DUTY_RELEASE = time(13, 30)   # 매뉴얼: 전날 당직선은 13:30 정리 → 그 전까지는 투입 가능

def _duty_rotation(board: Board) -> list:
    """당직선을 '직전 배정 다음 순번'부터 순환한 순서 (전체 당직 목록 기준 인덱스)"""
    d = board.on_duty
    if not d: return []
    s = (board.last_duty_idx + 1) % len(d)
    return d[s:] + d[:s]

def rank(job: Job, board: Board, weekday: bool, upper: Optional[Board] = None,
         other: Optional[Board] = None):
    """후보 예선을 (사유, 예선) 순서로 반환. 앞에서부터 필요 척수만큼 쓴다.
    upper: 같은 항 상위마력 보드(1000HP → 2000HP), other: 타항의 같은 마력 보드"""
    out, seen = [], set()
    pool = (lambda ts: [t for t in ts if t.hp >= job.requested_hp]) if job.requested_hp else (lambda ts: ts)
    def add(why, ts):
        for t in ts:
            if t.name not in seen: seen.add(t.name); out.append((why, t))
    slot = time_slot(job.time, weekday)
    if slot == MORNING:                          # Ⅲ-2 ① 평일 오전 처진순
        add("당직+대기 처진순", _lag(pool(board.on_duty + board.standby)))
        add("전날 당직 처진순", _lag(pool(board.prev_duty)))
        if upper: add("상위마력 하향지원", _lag(pool(upper.on_duty + upper.standby)))
    elif slot == DAY:                            # Ⅲ-2 ② 평일 오후 / 주말 주간 당직순
        add("당직순", _ok(pool(_duty_rotation(board))))
        add("당일 대기", _ok(pool(board.standby)))
        if job.time < PREV_DUTY_RELEASE:
            add("전날 당직(13:30 정리 전)", _ok(pool(board.prev_duty)))
        if upper: add("상위마력 투입", _lag(pool(upper.on_duty + upper.standby)))
        if other: add("타항 지원", _lag(pool(other.on_duty + other.standby)))
    else:                                        # Ⅲ-2 ③ 야간: 당일 당직선 번갈아 배정
        add("야간 당직순", _ok(pool(_duty_rotation(board))))
        # 야간 부족 시 대응은 매뉴얼에 없음 → 주간 규칙(대기 → 상위마력 → 타항)을 준용하고 표시한다
        add("야간 보충: 대기", _ok(pool(board.standby)))
        if upper: add("야간 보충: 상위마력", _lag(pool(upper.on_duty + upper.standby)))
        if other: add("야간 보충: 타항 지원", _lag(pool(other.on_duty + other.standby)))
    return out

def duration_min(j: Job) -> int:
    """매뉴얼 Ⅲ-6 작업 예상 시간표(분). 접안/이안은 상한 쪽 값, 이선은 90분."""
    if j.kind == "이선": return 90
    if j.gt < 10_000:  d = (60, 30)
    elif j.gt < 30_000: d = (90, 60)
    elif j.gt < 50_000: d = (120, 60)
    else:               d = (150, 90)
    return d[0] if j.kind.startswith("입항") else d[1]

def explain(j: Job, grp: str, b: Board, weekday: bool, chosen: list, start: int,
            man: Optional[ManualRefs] = None, extra: tuple = ()) -> list[str]:
    """배정 근거를 사람이 읽을 문장 리스트로 만든다. 건수/작업중 상태를 갱신하기 *전에* 호출할 것.
    man이 있으면 각 규칙 줄 끝에 매뉴얼 각주(📖[Mn])를 붙인다."""
    c = man.cite if man else (lambda k: "")
    lines = [f"보드: {j.port}/{grp}HP  (부두 '{j.berth}' → {j.port}, G/T {j.gt:,} → {grp}HP급)" + c("보드")]
    slot = time_slot(j.time, weekday)
    if slot == MORNING:
        lines.append("규칙: 평일 오전(06:00~11:59) → 처진순: 당직+대기 → 전날 당직 → 상위마력 하향지원"
                     + c("처진순"))
    elif slot == DAY:
        span = "평일 오후(12:00~17:30)" if weekday else "주말 주간(06:00~15:30)"
        lines.append(f"규칙: {span} → 당직순: 당일 당직(순환) → 당일 대기 → 상위마력 → 타항 지원" + c("당직순"))
    else:
        span = "평일 야간(17:30~06:00)" if weekday else "주말 야간(15:30~06:00)"
        lines.append(f"규칙: {span} → 당일 당직선 번갈아 배정 (부족 시 주간 규칙 준용)" + c("당직순"))
    allb = b.on_duty + b.standby + b.prev_duty
    for eb in extra:                                   # 상위마력/타항 보드도 '작업 중' 표시에 포함
        allb += [t for t in eb.on_duty + eb.standby + eb.prev_duty if t not in allb]
    if j.requested_hp:
        low = [f"{t.name}({t.hp})" for t in b.on_duty + b.standby + b.prev_duty if t.hp < j.requested_hp]
        lines.append(f"마력 요청: REMARKS '{j.requested_hp}HP 이상' → 미달 예선 제외"
                     + (f" [{', '.join(low)}]" if low else "") + c("마력요청"))
    busy = [f"{t.name}(~{t.free_at // 60 % 24:02d}:{t.free_at % 60:02d})" for t in allb if t.busy]
    if busy: lines.append(f"작업 중 제외: {', '.join(busy)}" + c("작업시간"))
    for i, (why, t) in enumerate(chosen, 1):
        lines.append(f"{i}번선 {t.name}({t.hp}HP): {why} — 누적 {t.count}건 → {t.count + 1}건"
                     + (c("상위마력") if "상위마력" in why else c("예선부족") if "타항" in why else ""))
        if why.startswith("야간 보충"):
            lines[-1] += "  ⚠ 야간 부족 대응은 매뉴얼에 없음 → 사람 확인"
    if chosen:
        end = start + duration_min(j)
        lines.append(f"작업 예상 {duration_min(j)}분 ({j.kind}, G/T {j.gt:,}) → 배정 예선은 "
                     f"{end // 60 % 24:02d}:{end % 60:02d}까지 작업 중 처리" + c("작업시간"))
    if len(chosen) < j.tugs:
        lines.append(f"⚠ 필요 {j.tugs}척 중 {len(chosen)}척만 배정 가능 (가용 예선 부족) → 사람 확인" + c("예선부족"))
    if j.notes: lines.append("REMARKS 참고: " + " / ".join(j.notes))
    return lines

def assign_all(jobs: list[Job], boards: dict, weekday: bool, man: Optional[ManualRefs] = None):
    """시간순으로 배정하며 건수를 갱신한다. boards 키: (항, 마력군)"""
    results = []
    for j in sorted(jobs, key=lambda x: (x.time, x.no)):
        grp = hp_group(j.gt)
        start = j.time.hour * 60 + j.time.minute
        for bd in boards.values():
            for t in bd.on_duty + bd.standby + bd.prev_duty: t.busy = t.free_at > start
        if j.port == "미분류" or grp == "3000" or (j.port, grp) not in boards:
            why = "항 확인 필요" if j.port == "미분류" else "3천마력 규칙/보드 미구현"
            results.append((j, [], why, [f"미배정 사유: {why}"] + (["REMARKS 참고: " + " / ".join(j.notes)] if j.notes else []))); continue
        b = boards[(j.port, grp)]
        other_port = "본항" if j.port == "온산" else "온산"
        upper = boards.get((j.port, "2000")) if grp == "1000" else None      # 3000HP 보드는 미구현
        other = boards.get((other_port, grp))
        cands = rank(j, b, weekday, upper=upper, other=other)
        chosen = cands[:max(j.tugs, 1)]
        reasons = explain(j, grp, b, weekday, chosen, start, man,    # 건수 갱신 전에 근거 기록
                          extra=tuple(x for x in (upper, other) if x))
        for _, t in chosen:
            t.count += 1; t.free_at = start + duration_min(j)
            if t in b.on_duty: b.last_duty_idx = b.on_duty.index(t)   # 다음 당직순은 이 예선 다음부터
        results.append((j, chosen, f"가용 예선 부족({len(chosen)}/{j.tugs}척)" if len(chosen) < j.tugs else "", reasons))
    return results


# ═════════ 4. 실행 ═════════
def demo_boards() -> dict:
    """[샘플] 매뉴얼 예시의 당직표. 건수는 임의값 → 실제로는 건수일지에서 읽어 올 것."""
    T = Tug
    return {
        ("온산", "1000"): Board(on_duty=[T("조광101",1300,2), T("111청룡",1500,3), T("선진6",1500,1), T("흥진6",1500,3)]),
        ("본항", "1000"): Board(on_duty=[T("조101",1300,1), T("K101",1500,2), T("조105",1500,0)]),
        ("온산", "2000"): Board(on_duty=[T("선802",2000,2), T("강산",2000,3)]),
        ("본항", "2000"): Board(on_duty=[T("쥬피터",2000,1), T("선7",2000,2)]),
    }

if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if "--date" in sys.argv:                                   # --date 2026-11-06
        target = datetime.strptime(sys.argv[sys.argv.index("--date") + 1], "%Y-%m-%d").date()
        args = [a for a in args if a != sys.argv[sys.argv.index("--date") + 1]]
    else:
        target = date.today()
    day = "평일" if target.weekday() < 5 else "주말"
    if args:                                                   # 엑셀 파일 직접 지정
        path = Path(args[0])
        m = re.search(r"\d{4}-\d{2}-\d{2}", path.name)
        if m: target = datetime.strptime(m.group(), "%Y-%m-%d").date()
        day = "평일" if target.weekday() < 5 else "주말"
        print(f"[입력] {path}  (기준일 {target}, {day})")
        jobs = load_jobs(str(path))
    else:                                                      # 기본: DB에서 읽고, 오늘 데이터 없으면 크롤링
        print(f"[입력] tug_agent.db  (기준일 {target}, {day})")
        try:
            jobs = load_jobs_db(target, force="--crawl" in sys.argv, allow="--no-crawl" not in sys.argv)
        except (LookupError, ValueError) as e:
            sys.exit(f"[오류] {e}")
    enrich(jobs, use_llm="--no-llm" not in sys.argv)
    weekday = target.weekday() < 5
    brief = "--brief" in sys.argv                               # --brief: 근거 설명 생략
    man = ManualRefs(enabled=not brief and "--no-manual" not in sys.argv)
    for j, chosen, why, reasons in assign_all(jobs, demo_boards(), weekday, man):
        names = ", ".join(f"{t.name}({t.hp})" for _, t in chosen) or f"[미배정: {why}]"
        if chosen and why: names += f" ⚠{why}"
        print(f"{j.time:%H:%M} #{j.no:<2} {j.ship[:22]:<22} {j.kind:<7} {j.port}/{hp_group(j.gt)}HP 필요{j.tugs}척 → {names}")
        if not brief:
            for r in reasons: print(f"        └ {r}")
            print()
    if man.notes or man.by_key:
        print("═════ 📖 매뉴얼 근거 (manual.index 검색) ═════")
        for r in man.report(): print(r)