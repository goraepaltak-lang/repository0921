# -*- coding: utf-8 -*-
"""평가지표(후보 예선 점수표) 근거 계산

무엇을 하는가
  1) 마력급: 톤수표(매뉴얼)대로 정한 마력군과 실제 배정이 얼마나 같은지 센다. (2025-03-01 ~ 2026-07-27)
  2) 예선 선택: 2026년 7월 작업일지에서 "어느 예선을 골랐는가"를
       A. 규칙 그대로(가중치 없음)
       B. 규칙 항목만, 가중치는 기록에서 맞춤
       C. 규칙 항목 + 경험 항목, 가중치는 기록에서 맞춤
     세 방식으로 추천해 보고 실제 배정과 비교한다. (7/1~7/20으로 가중치를 맞추고 7/21~7/27로 확인)

실행
  python 평가지표_근거계산.py [자료 폴더] [가중치를 맞출 마지막 날, 기본 20]
  - 자료 폴더 기본값: 이 파일 옆의 .narnia-ai-uploads
  - 필요: pip install openpyxl numpy

주의 (계산에 들어간 가정)
  - 척수(1척/2척)는 입력으로 주어진 것으로 본다(실제 사용 척수를 그대로 씀).
  - 마력군은 당직표의 열 위치로 정했다(1,300 6척 / 1,500 8척 / 2,000 10척 / 3,000 24척).
  - 톤수표의 "1,900HP급"은 1,500마력군으로 읽었다(확인 필요).
  - 3,000마력군은 3K급 8척과 4K 이상 16척(건수일지의 T/L 순서에 적힌 16척)으로 나눠 본다.
  - 건수(처진순)는 건수일지 '★전체 건수★(2026)'의 7/27 값에서 7월 작업 수를 빼서 7/1 값으로 되돌린 뒤
    작업마다 1건씩 더한 값이다(7월 중 보존건수·페널티 조정은 반영하지 못함).
  - 작업 중인지 여부는 매뉴얼의 작업 예상 시간표 하한으로 추정했다.
  - 부두가 본항인지 온산인지는 7/1~7/20 야간에 어느 쪽 당직선이 맡았는지로 추정했다.
"""
import datetime as dt
import re
import sys
import warnings
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np
from openpyxl import load_workbook

warnings.filterwarnings("ignore")

# ※ 실행 전에 고칠 곳: 아래 세 파일 이름을 자료 폴더에 있는 실제 엑셀 파일 이름으로 바꾼다.
F_DUTY = "20261115061428-001-TUG.xlsx"      # 당직표
F_COUNT = "20261115061428-002-file.xlsx"    # 건수일지 (날짜별 작업 기록 포함)
F_LOG = "20261115061428-004-file.xlsx"      # 2026년 7월 작업일지 + 도선예보

TRAIN_LAST_DAY = 20          # 7/1~7/20: 가중치 맞추기, 7/21~: 확인 (두 번째 인자로 바꿀 수 있음)
RULE_START = dt.date(2025, 3, 1)   # 톤수표 시행일
SPLIT = r"[,/\s.]+"


# ---------------------------------------------------------------- 공통 도구
def norm(name):
    """예선 이름 표기를 하나로 맞춘다 (해강9=해9, 109청룡=109 ...)."""
    s = str(name).strip().upper().replace(" ", "")
    for a, b in (("해강", "해"), ("조광", "조"), ("선진", "선"), ("흥진", "흥"), ("청룡", "")):
        s = s.replace(a, b)
    return {"SBM111": "SBM11"}.get(s, s)


def t2min(v):
    if isinstance(v, (dt.time, dt.datetime)):
        return v.hour * 60 + v.minute
    m = re.match(r"(\d{1,2}):(\d{2})", str(v or ""))
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def allowed(gt, k):
    """톤수표: 총톤수와 척수 -> 예선 자리별로 허용되는 마력군"""
    if gt < 5000:
        return [{1300, 1500}] * k
    if gt < 7000:
        return [{1500}] + [{1300, 1500}] * (k - 1)
    if gt < 10000:
        return [{1500}] * k
    if gt < 25000:
        return [{2000}] * k
    if gt < 30000:
        return [{3000}] + [{2000}] * (k - 1)
    return [{3000}] * k


K3 = {"해7", "처용", "선909", "제우스", "조13", "SBM11", "선11", "흥3"}   # 3,000마력군 중 T/L 순서(4K 이상 16척)에 없는 8척


def tier_of(name, grp):
    """예선 선택 계산에 쓰는 등급: 1300 / 1500 / 2000 / 3000(3K급) / 4000(4K 이상)"""
    g = grp[name]
    return g if g < 3000 or name in K3 else 4000


def allowed_tier(gt, k):
    """톤수표를 3K급/4K 이상까지 나눠 읽은 것 (예선 선택 계산용)"""
    if gt < 25000:
        return allowed(gt, k)
    if gt < 30000:
        return [{3000}] + [{2000}] * (k - 1)
    if gt < 50000:
        return [{3000, 4000}] * k
    return [{4000}] * k


BANDS = [(0, 5000, "5천 톤 미만"), (5000, 7000, "5천~7천"), (7000, 10000, "7천~1만"),
         (10000, 25000, "1만~2만5천"), (25000, 30000, "2만5천~3만"), (30000, 10 ** 9, "3만 이상")]


def band(gt):
    for lo, hi, name in BANDS:
        if lo <= gt < hi:
            return name


def judge(gt, groups):
    up = down = False
    for a, s in zip(sorted(groups, reverse=True), allowed(gt, len(groups))):
        if a in s:
            continue
        if a > max(s):
            up = True
        else:
            down = True
    return "일치" if not (up or down) else "상위" if up and not down else "하위" if down and not up else "혼합"


def work_minutes(gt, kind):
    """작업 예상 시간표의 하한(분)"""
    if kind == "U":
        return 15 if gt < 10000 else 30 if gt < 30000 else 60
    return 60 if gt < 30000 else 90 if gt < 50000 else 120


# ---------------------------------------------------------------- 자료 읽기
def load_groups(folder):
    """당직표 열 위치 -> 예선별 마력군, 날짜별 당직 계획"""
    wb = load_workbook(folder / F_DUTY, read_only=True, data_only=True)
    cols = {3: 1300, 4: 1300, 5: 1500, 6: 1500, 7: 2000, 8: 2000, 9: 3000, 10: 3000, 11: 3000,
            12: 1300, 13: 1300, 14: 1500, 15: 1500, 16: 2000, 17: 2000}
    votes, plan = defaultdict(Counter), {}
    for ws in wb.worksheets:
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i < 3 or row[0] in (None, ""):
                continue
            r = [("" if v is None else norm(v)) for v in row[:20]]
            for c, g in cols.items():
                if c < len(r) and r[c]:
                    votes[r[c]][g] += 1
            if "7월" in ws.title:
                plan[str(row[0]).strip()] = dict(onsan=[x for x in r[3:10] if x], cross=[x for x in r[10:11] if x],
                                                 bon=[x for x in r[11:18] if x])
    wb.close()
    grp = {k: v.most_common(1)[0][0] for k, v in votes.items()}
    grp.update({"선808": 3000, "바다1": 3000})      # 2025년 건수일지에만 나오는 3,000마력군 이름
    return grp, plan


def load_history(folder):
    """건수일지의 '20xx년 작업(월)' 시트 -> 날짜별 작업 기록"""
    wb = load_workbook(folder / F_COUNT, read_only=True, data_only=True)
    jobs = []
    for name in wb.sheetnames:
        if not re.match(r"\d{4}년 작업\(\d+\)", name):
            continue
        rows = list(wb[name].iter_rows(values_only=True))
        if len(rows) < 3:
            continue
        r0 = rows[0]
        for c in [c for c, v in enumerate(r0) if str(v or "").strip() == "날짜"]:
            d = r0[c + 1] if c + 1 < len(r0) else None
            if not isinstance(d, (dt.datetime, dt.date)):
                continue
            d = d.date() if isinstance(d, dt.datetime) else d
            for r in rows[2:]:
                cell = lambda k: (r[c + k] if c + k < len(r) and r[c + k] is not None else "")
                if not str(cell(1)).strip() and not str(cell(8)).strip():
                    continue
                jobs.append(dict(date=d, time=t2min(cell(0)), ship=str(cell(1)).strip(), gt=cell(3),
                                 frm=str(cell(4)).strip(), to=str(cell(5)).strip(), kind=str(cell(7)).strip(),
                                 tugs=str(cell(8)).strip()))
    wb.close()
    return jobs


def load_official_counts(folder):
    """건수일지 '★전체 건수★(2026)' 시트의 예선별 건수 (시트에 적힌 날짜 기준)"""
    wb = load_workbook(folder / F_COUNT, read_only=True, data_only=True)
    rows = list(wb["★전체 건수★(2026)"].iter_rows(values_only=True))
    wb.close()
    out = {}
    for r in rows[4:40]:
        for c in (5, 9, 13, 17):
            if c + 1 < len(r) and r[c] not in (None, "") and isinstance(r[c + 1], (int, float)):
                out[norm(r[c])] = int(r[c + 1])
    return out, rows[2][1]


def load_july(folder, grp):
    """7월 작업일지: 작업 행 + 날짜별 머리글(당직, 조기퇴근, 수리선)"""
    wb = load_workbook(folder / F_LOG, read_only=True, data_only=True)
    keys = ("본항당직", "교차당직", "온산당직", "조기퇴근", "T/L", "수리선")
    jobs, heads = [], {}
    for d in range(1, 32):
        name = f"26년 7월일지({d})"
        if name not in wb.sheetnames:
            continue
        head, started, n = {}, False, 0
        for row in wb[name].iter_rows(values_only=True):
            cells = [("" if v is None else v) for v in row]
            txt = [str(c).strip() for c in cells]
            if not started:
                for k in keys:
                    if k in txt:
                        nxt = [t for t in txt[txt.index(k) + 1:] if t]
                        head[k] = nxt[0] if nxt and nxt[0] not in keys else ""
                started = bool(txt) and txt[0].replace(" ", "") == "예정시간"
                continue
            if not any(txt[:12]):
                continue
            n += 1
            jobs.append(dict(day=d, date=dt.date(2026, 7, d), time=t2min(cells[0]), ship=txt[1], gt=cells[3],
                             frm=txt[4], to=txt[5], ca=txt[6], kind=txt[7], tugs=txt[8], la=txt[9], t=txt[10],
                             note=txt[11], seq=n))
        if n:
            heads[d] = head
    wb.close()
    names = lambda s: [norm(x) for x in re.split(SPLIT, s) if x and norm(x) in grp]
    info = {}
    for d, h in heads.items():
        left = []
        for m in re.finditer(r"\((\d{1,2}):(\d{2})\)([^()]*)", h.get("조기퇴근", "")):
            left += [(int(m.group(1)) * 60 + int(m.group(2)), x) for x in names(m.group(3))]
        info[d] = dict(bon=names(h.get("본항당직", "")), onsan=names(h.get("온산당직", "")),
                       cross=names(h.get("교차당직", "")), left=left, repair=names(h.get("수리선", "")))
    return jobs, info


def tug_names(job, grp):
    """사용예선 칸 -> 이름 목록. 모르는 이름이 섞여 있으면 None"""
    ns = [norm(x) for x in re.split(SPLIT, job["tugs"]) if x]
    return ns if ns and all(n in grp for n in ns) else None


# ---------------------------------------------------------------- 1) 마력급 일치
def part1_hp(hist, grp):
    print("=" * 78)
    print("1) 마력급: 톤수표대로 정한 마력군 vs 실제 (2025-03-01 이후, 구분 M/U)")
    rows, skip = [], Counter()
    for j in hist:
        if j["date"] < RULE_START:
            skip["시행일 이전"] += 1
            continue
        if j["kind"] not in ("M", "U"):
            skip["구분이 M/U 아님"] += 1
            continue
        try:
            gt = float(j["gt"])
        except (TypeError, ValueError):
            skip["톤수가 숫자 아님"] += 1
            continue
        ns = tug_names(j, grp)
        if not ns:
            skip["예선 이름 인식 불가"] += 1
            continue
        rows.append(dict(j, gtn=gt, res=judge(gt, [grp[n] for n in ns]), band=band(gt),
                         berth=j["to"] if j["kind"] == "M" else j["frm"]))
    dates = sorted({r["date"] for r in rows})
    c = Counter(r["res"] for r in rows)
    print(f"   기간 {dates[0]} ~ {dates[-1]} ({len(dates)}일), 대상 {len(rows):,}건, 제외 {dict(skip)}")
    print(f"   일치 {c['일치']:,} / 상위마력 {c['상위']:,} / 하위마력 {c['하위']:,} / 혼합 {c['혼합']:,}"
          f"  -> 일치율 {c['일치'] / len(rows):.1%}")
    for _, _, b in BANDS:
        rs = [r for r in rows if r["band"] == b]
        cc = Counter(r["res"] for r in rs)
        print(f"   {b:<10s} {len(rs):>6,}건  일치 {cc['일치']:>6,}  상위 {cc['상위']:>5,}  하위 {cc['하위']:>3,}"
              f"  일치율 {cc['일치'] / len(rs):.1%}")
    for key, minn, label in (("ship", 10, "선박"), ("berth", 20, "부두")):
        d = defaultdict(Counter)
        for r in rows:
            d[r[key]][r["res"]] += 1
        big = {k: v for k, v in d.items() if sum(v.values()) >= minn}
        rate = sorted(((1 - v["일치"] / sum(v.values()), sum(v.values())) for v in big.values()), reverse=True)
        print(f"   {label}: {minn}건 이상 {len(big)}개 중 이탈률 50% 이상 {sum(1 for a, _ in rate if a >= .5)}개, "
              f"25% 이상 {sum(1 for a, _ in rate if a >= .25)}개, 최고 {rate[0][0]:.0%}({rate[0][1]}건)")


# ---------------------------------------------------------------- 2) 예선 선택
RULE_F = ["마력급 일치", "상위 마력", "당직선(평일 오전)", "당직선(오후·주말 주간)", "당직선(야간)",
          "같은 항(평일 오전)", "같은 항(오후·주말 주간)", "같은 항(야간)", "건수 적음(평일 오전)", "건수 적음(그 외)",
          "쉰 시간 김", "작업 중 추정", "조기퇴근함", "수리선"]
EXP_F = ["방금 작업을 마침", "같은 부두에서 이어서", "그 배를 맡았던 이력"]
ALL_F = RULE_F + EXP_F


def slot_of(day, t):
    we = dt.date(2026, 7, day).weekday() >= 5
    if t < 360 or t >= (930 if we else 1050):
        return "야간"
    return "주말 주간" if we else ("평일 오전" if t < 720 else "평일 오후")


def build_instances(july, info, plan, grp, hist, official):
    """작업을 시간순으로 따라가며 (작업, 예선 자리)마다 후보별 항목값을 만든다."""
    tugs = sorted(t for t in grp if t in official)      # 후보는 현재 건수 시트에 있는 예선만(48척)
    tix = {t: i for i, t in enumerate(tugs)}
    tier = {t: tier_of(t, grp) for t in tugs}
    gvec = np.array([tier[t] for t in tugs])
    peers = {t: [m for m in tugs if tier[m] == tier[t] and m != t and m in official] for t in tugs}
    # 선박별 예선 이력: 6/30까지의 기록
    ship_hist = defaultdict(Counter)
    for j in hist:
        if j["date"] >= dt.date(2026, 7, 1):
            continue
        for n in tug_names(j, grp) or []:
            ship_hist[j["ship"]][n] += 1
    # 건수: 공식 건수(7/27)에서 7월 작업 수를 빼서 7/1 시작값으로 되돌린다
    count = Counter(official)
    for j in july:
        for n in tug_names(j, grp) or []:
            count[n] -= 1
    p = plan.get("6/30")
    info = dict(info)
    info[0] = dict(bon=p["bon"], onsan=p["onsan"], cross=p["cross"], left=[], repair=[])

    jobs = [j for j in july if j["time"] is not None]
    jobs.sort(key=lambda j: (j["day"], j["time"], j["seq"]))
    # 부두 -> 항: 학습 구간의 야간 작업만으로 추정
    bp = defaultdict(Counter)
    for j in jobs:
        ns = tug_names(j, grp)
        if j["day"] > TRAIN_LAST_DAY or not ns or j["kind"] not in ("M", "U") or slot_of(j["day"], j["time"]) != "야간":
            continue
        dd = info[j["day"] - 1 if j["time"] < 360 else j["day"]]
        for n in ns:
            for port in ("bon", "onsan"):
                if n in dd[port]:
                    bp[j["to"] if j["kind"] == "M" else j["frm"]][port] += 1
    berth_port = {b: c.most_common(1)[0][0] for b, c in bp.items()}

    last = {}      # 예선 -> (시작 시각(누적 분), 끝 추정 시각, 부두)
    inst, stat = [], Counter()
    for j in jobs:
        ns = tug_names(j, grp)
        now = (j["day"] - 1) * 1440 + j["time"]
        try:
            gt = float(j["gt"])
        except (TypeError, ValueError):
            gt = None
        berth = j["to"] if j["kind"] == "M" else j["frm"]
        if ns and gt is not None and j["kind"] in ("M", "U") and all(n in tix for n in ns):
            slot = slot_of(j["day"], j["time"])
            dday = j["day"] - 1 if j["time"] < 360 else j["day"]
            dd = info[dday]
            duty = set(dd["bon"]) | set(dd["onsan"]) | set(dd["cross"])
            port = berth_port.get(berth)
            same = (set(dd[port]) | set(dd["cross"])) if port else set()
            gone = {n for t, n in dd["left"] if (dday < j["day"] or t <= j["time"])} - duty
            repair = set(dd["repair"])
            need = allowed_tier(gt, len(ns))
            actual = sorted(ns, key=lambda n: -tier[n])
            am, pm, night = slot == "평일 오전", slot in ("평일 오후", "주말 주간"), slot == "야간"
            sh = ship_hist[j["ship"]]
            sh_tot = sum(sh.values())
            X = np.zeros((len(tugs), len(ALL_F)))
            for n, i in tix.items():
                ps = peers[n]      # 같은 등급 안에서 건수가 적을수록 1에 가깝다(처진 예선)
                cnt = (sum(count[m] > count[n] for m in ps) + 0.5 * sum(count[m] == count[n] for m in ps)) / max(1, len(ps))
                ls = last.get(n)
                since = now - ls[0] if ls else 10 ** 6
                busy = bool(ls) and now < ls[1]
                X[i, 2], X[i, 3], X[i, 4] = (n in duty) * am, (n in duty) * pm, (n in duty) * night
                X[i, 5], X[i, 6], X[i, 7] = (n in same) * am, (n in same) * pm, (n in same) * night
                X[i, 8], X[i, 9] = cnt * am, cnt * (not am)
                X[i, 10] = min(since, 720) / 720
                X[i, 11], X[i, 12], X[i, 13] = busy, n in gone, n in repair
                X[i, 14] = bool(ls) and not busy and now - ls[1] <= 60
                X[i, 15] = bool(ls) and since <= 180 and ls[2] == berth and bool(berth)
                X[i, 16] = sh[n] / (sh_tot + 5)
            for a in actual:
                stat["실제 예선(전체)"] += 1
                stat["실제 예선이 '작업 중 추정'"] += int(X[tix[a], 11])
                stat["실제 예선이 '조기퇴근함'"] += int(X[tix[a], 12])
                stat["실제 예선이 '수리선'"] += int(X[tix[a], 13])
                stat["실제 예선이 당직선"] += int(a in duty)
            inst.append(dict(day=j["day"], slot=slot, need=need, actual=actual, X=X, gvec=gvec, tix=tix, tugs=tugs,
                             ship=j["ship"], time=j["time"], berth=berth, gt=gt, kind=j["kind"]))
        if ns:      # 상태 갱신은 실제 배정으로 한다
            dur = work_minutes(gt or 0, "U" if j["kind"] == "U" else "M")
            for n in ns:
                count[n] += 1
                ship_hist[j["ship"]][n] += 1
                last[n] = (now, now + dur, berth)
    return inst, stat, berth_port


def slot_matrix(it, s, picked):
    """예선 자리 s의 후보 마스크와 항목값(마력급 항목은 자리마다 다름)"""
    need = it["need"][s]
    g = it["gvec"]
    X = it["X"].copy()
    X[:, 0] = np.isin(g, list(need))
    X[:, 1] = g > max(need)
    mask = g >= min(need)                      # 걸러내기: 최소 마력 미달은 후보에서 뺀다
    for p in picked:
        mask[it["tix"][p]] = False
    return X, mask


def score_rule_as_written(it, X):
    """A. 규칙 그대로: 매뉴얼의 시간대별 순서를 사전식으로 옮긴 것 (가중치를 기록에 맞추지 않음)"""
    base = 1e6 * X[:, 0] - 1e7 * (X[:, 12] + X[:, 13]) - 1e5 * X[:, 11]
    if it["slot"] == "평일 오전":            # 처진순: 건수 적은 예선부터
        return base + 1e3 * X[:, 8] + X[:, 10]
    duty = X[:, 3] + X[:, 4]
    same = X[:, 6] + X[:, 7]                  # 당직선 -> 같은 항 -> 오래 쉰 예선(번갈아)
    return base + 1e4 * duty + 1e3 * same + 10 * X[:, 10]


def predict(it, scorer):
    picked, ranks = [], []
    for s in range(len(it["need"])):
        X, mask = slot_matrix(it, s, picked)
        sc = np.where(mask, scorer(it, X), -np.inf)
        order = [it["tugs"][i] for i in np.argsort(-sc, kind="stable") if mask[i]]
        ranks.append(order)
        picked.append(order[0])
    return picked, ranks


def fit(train, cols, l2=1.0, iters=600, lr=0.5):
    """조건부 로짓: 후보 중 실제로 뽑힌 예선의 점수가 가장 높아지도록 가중치를 맞춘다."""
    Xs, Ms, Ys = [], [], []
    for it in train:
        picked = []
        for s, a in enumerate(it["actual"]):
            X, mask = slot_matrix(it, s, picked)
            if mask[it["tix"][a]]:
                Xs.append(X[:, cols]); Ms.append(mask); Ys.append(it["tix"][a])
            picked.append(a)
    X, M, Y = np.array(Xs), np.array(Ms), np.array(Ys)
    w = np.zeros(len(cols))
    idx = np.arange(len(Y))
    for _ in range(iters):
        z = np.where(M, X @ w, -1e9)
        z -= z.max(axis=1, keepdims=True)
        p = np.exp(z); p /= p.sum(axis=1, keepdims=True)
        grad = (X[idx, Y] - (p[:, :, None] * X).sum(axis=1)).mean(axis=0) - l2 * w / len(Y)
        w += lr * grad
    return w


def evaluate(items, scorer):
    out = Counter()
    for it in items:
        pred, ranks = predict(it, scorer)
        k = len(it["actual"])
        out["건"] += 1
        out["건 완전 일치"] += set(pred) == set(it["actual"])
        out["예선"] += k
        out["예선 일치"] += len(set(pred) & set(it["actual"]))
        top = set(ranks[0][:k + 2])
        out["상위 후보 안"] += sum(a in top for a in it["actual"])
    return out


def random_baseline(items):
    """마력급만 맞추고 그 안에서 아무나 고를 때 기대되는 예선 일치 수"""
    exp = tot = 0
    for it in items:
        picked = []
        for s, a in enumerate(it["actual"]):
            X, mask = slot_matrix(it, s, picked)
            ok = mask & (X[:, 0] > 0)
            exp += (1 / ok.sum()) if ok[it["tix"][a]] else 0
            tot += 1
            picked.append(a)
    return exp, tot


def part2_tug(july, info, plan, grp, hist, official):
    print()
    print("=" * 78)
    print("2) 예선 선택: 어느 예선을 골랐는가 (2026년 7월 작업일지, 구분 M/U)")
    inst, stat, berth_port = build_instances(july, info, plan, grp, hist, official)
    train = [it for it in inst if it["day"] <= TRAIN_LAST_DAY]
    test = [it for it in inst if it["day"] > TRAIN_LAST_DAY]
    days = sorted({it["day"] for it in inst})
    print(f"   대상 {len(inst):,}건 (7/{days[0]}~7/{days[-1]}) = 가중치 맞추기 {len(train):,}건(~7/{TRAIN_LAST_DAY})"
          f" + 확인 {len(test):,}건")
    n = stat["실제 예선(전체)"]
    for k in ("실제 예선이 당직선", "실제 예선이 '작업 중 추정'", "실제 예선이 '조기퇴근함'", "실제 예선이 '수리선'"):
        print(f"   {k}: {stat[k]:,}/{n:,} = {stat[k] / n:.1%}")
    print(f"   항(본항/온산)을 추정한 부두 {len(berth_port)}개")

    rule_cols = [ALL_F.index(f) for f in RULE_F]
    all_cols = list(range(len(ALL_F)))
    wB, wC = fit(train, rule_cols), fit(train, all_cols)
    full = lambda cols, w: (lambda _it, X: X[:, cols] @ w)
    models = [("A. 규칙 그대로 (가중치 없음)", score_rule_as_written),
              ("B. 규칙 항목만 + 가중치", full(rule_cols, wB)),
              ("C. 규칙 + 경험 항목 + 가중치", full(all_cols, wC))]

    for label, items in ((f"확인 구간 (7/{TRAIN_LAST_DAY + 1}~, 가중치를 맞출 때 쓰지 않은 날)", test),
                         (f"참고: 가중치를 맞춘 구간 (~7/{TRAIN_LAST_DAY})", train)):
        print(f"\n   [{label}] {len(items):,}건")
        e, t = random_baseline(items)
        print(f"   {'(참고) 마력급만 맞추고 무작위':<30s} 예선 일치 {e / t:.1%}")
        for name, sc in models:
            r = evaluate(items, sc)
            print(f"   {name:<30s} 건 완전 일치 {r['건 완전 일치']:>4}/{r['건']} = {r['건 완전 일치'] / r['건']:.1%}"
                  f" | 예선 일치 {r['예선 일치']:>4}/{r['예선']} = {r['예선 일치'] / r['예선']:.1%}"
                  f" | 상위 후보(척수+2) 안 {r['상위 후보 안'] / r['예선']:.1%}")
    hitB = [set(predict(it, models[1][1])[0]) == set(it["actual"]) for it in test]
    hitC = [set(predict(it, models[2][1])[0]) == set(it["actual"]) for it in test]
    only_c = sum(c and not b for b, c in zip(hitB, hitC))
    only_b = sum(b and not c for b, c in zip(hitB, hitC))
    print(f"\n   [확인 구간, B와 C의 건별 비교] C만 맞힌 건 {only_c} / B만 맞힌 건 {only_b} / 둘 다 맞힘 "
          f"{sum(b and c for b, c in zip(hitB, hitC))} / 둘 다 틀림 {sum(not b and not c for b, c in zip(hitB, hitC))}")
    print("\n   [확인 구간, 시간대별 예선 일치율]")
    for s in ("평일 오전", "평일 오후", "주말 주간", "야간"):
        items = [it for it in test if it["slot"] == s]
        line = f"   {s:<6s} {len(items):>4}건 "
        for name, sc in models:
            r = evaluate(items, sc)
            line += f"| {name[:2]} {r['예선 일치'] / max(1, r['예선']):.1%} (건 {r['건 완전 일치'] / max(1, r['건']):.1%}) "
        print(line)
    print("\n   [C의 가중치] 값이 클수록 그 항목이 맞는 예선이 뽑히는 쪽으로 작용 (음수는 반대)")
    for f, w in zip(ALL_F, wC):
        print(f"   {'경험' if f in EXP_F else '규칙'}  {f:<18s} {w:+.2f}")
    return inst


def main():
    global TRAIN_LAST_DAY
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent / ".narnia-ai-uploads"
    if len(sys.argv) > 2:
        TRAIN_LAST_DAY = int(sys.argv[2])
    grp, plan = load_groups(folder)
    by = Counter(v for k, v in grp.items() if k not in ("선808", "바다1"))
    print("마력군별 척수(당직표 기준):", dict(sorted(by.items())))
    hist = load_history(folder)
    july, info = load_july(folder, grp)
    part1_hp(hist, grp)
    official, asof = load_official_counts(folder)
    print(f"공식 건수 시트 기준일: {str(asof)[:10]}, 예선 {len(official)}척")
    part2_tug(july, info, plan, grp, hist, official)


if __name__ == "__main__":
    main()
