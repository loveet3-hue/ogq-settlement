# -*- coding: utf-8 -*-
"""스타즈 & 유니즈 월별 매출 로우 데이터 -> 정산 요약.

요약 규칙
    정지형 = 콘텐츠타입 '스티커'(단가 2,000), 동작형 = '애니메이션 스티커'(단가 3,000)
    판매 개수 = 판매 금액 / 단가          (환불은 음수 금액으로 자동 차감)
    결제 수수료 = ROUND(판매 금액 합계 * 결제율, 0)
    마켓 수수료 = ROUNDDOWN(판매 금액 합계 * 마켓율, 0)   ← 일반 마켓 수수료의 2배 요율
    크리에이터 정산 금액 = 판매 금액 합계 - 결제 수수료 - 마켓 수수료

부동소수점 오차로 ROUNDDOWN이 1원 어긋나는 경우가 있어(42,000 × 0.2835) 전 계산을
Decimal로 처리한다.
"""
import io
from copy import copy
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

SUMMARY_SHEET = "요약"

# 콘텐츠타입 -> (구분, 단가)
CONTENT_TYPES = {
    "스티커":            ("정지형", 2000),
    "애니메이션 스티커": ("동작형", 3000),
}

# 요약 마켓명 -> (기술 수수료율, 결제 수수료율, 마켓 수수료율)
MARKET_RATES = {
    "SOOP OGQ이모티콘":   ("0", "0.055", "0.2835"),
    "NAVER OGQ마켓":      ("0", "0.077", "0.277"),
    "채팅+ 원스토어마켓": ("0", "0",     "0.58"),
    "채팅+ OGQ마켓":      ("0", "0.055", "0.4"),
}

# 로우 데이터의 '판매마켓' -> 요약 마켓명
RAW_MARKET_MAP = {
    "SOOP OGQ 마켓":  "SOOP OGQ이모티콘",
    "SOOP OGQ마켓":   "SOOP OGQ이모티콘",
    "OGQ 마켓":       "NAVER OGQ마켓",
    "NAVER OGQ 마켓": "NAVER OGQ마켓",
    "NAVER OGQ마켓":  "NAVER OGQ마켓",
    "채팅+ 원스토어": "채팅+ 원스토어마켓",
    "채팅+ OGQ 마켓": "채팅+ OGQ마켓",
    "채팅+ OGQ마켓":  "채팅+ OGQ마켓",
}

COLUMNS = ["계정", "마켓", "정지형 판매 개수", "동작형 판매 개수", "판매 개수(합계)",
           "정지형 판매 금액", "동작형 판매 금액", "판매 금액(합계)",
           "기술 수수료", "결제 수수료", "마켓 수수료", "크리에이터 정산 금액"]


def _round(x):
    """엑셀 ROUND(x, 0)."""
    return int(Decimal(x).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _rounddown(x):
    """엑셀 ROUNDDOWN(x, 0) — 0 방향으로 버림."""
    return int(Decimal(x).quantize(Decimal("1"), rounding=ROUND_DOWN))


def read_account(data):
    """매출 로우 데이터 워크북 1개 -> (거래 리스트, 원본 행, 미등록 타입, 미등록 마켓)

    '요약' 시트는 건너뛰고, 필수 컬럼을 가진 시트의 데이터를 모두 모은다.
    파일 하나가 계정 하나(스타즈 / 유니즈)에 해당한다.
    """
    wb = openpyxl.load_workbook(io.BytesIO(data) if isinstance(data, (bytes, bytearray))
                                else data, read_only=True, data_only=True)
    recs, raw_header, raw_rows = [], None, []
    unknown_types, unknown_markets = set(), set()
    for ws in wb.worksheets:
        if ws.title.strip() == SUMMARY_SHEET:
            continue
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = [str(c).strip() if c is not None else "" for c in rows[0]]
        need = ("콘텐츠타입", "판매마켓", "판매금액")
        if not all(n in header for n in need):
            continue
        ti, mi, pi = (header.index(n) for n in need)
        if raw_header is None:
            raw_header = list(rows[0])
        for r in rows[1:]:
            if len(r) <= max(ti, mi, pi) or r[pi] is None:
                continue
            raw_rows.append(list(r))
            ctype = str(r[ti]).strip()
            market = RAW_MARKET_MAP.get(str(r[mi]).strip())
            if ctype not in CONTENT_TYPES:
                unknown_types.add(ctype)
                continue
            if market is None:
                unknown_markets.add(str(r[mi]).strip())
                continue
            recs.append((market, ctype, int(float(r[pi]))))
    return recs, (raw_header, raw_rows), sorted(unknown_types), sorted(unknown_markets)


def summarize(accounts, units=None):
    """{계정: [(마켓, 타입, 금액)]} -> 요약 행 리스트 + 합계 행.

    units 로 콘텐츠타입별 단가를 덮어쓸 수 있다 (기본 스티커 2,000 / 애니메이션 3,000).
    """
    units = units or {}
    agg = {}
    for account, recs in accounts.items():
        for market, ctype, amount in recs:
            key = (account, market)
            slot = agg.setdefault(key, {"정지형": [0, 0], "동작형": [0, 0]})
            group, unit = CONTENT_TYPES[ctype]
            unit = units.get(ctype, unit)
            slot[group][0] += amount // unit if amount % unit == 0 else 0
            slot[group][1] += amount
            if amount % unit:                      # 단가로 나뉘지 않는 금액은 개수 추정 불가
                slot.setdefault("_odd", []).append((ctype, amount))

    rows = []
    for (account, market), slot in agg.items():
        s_cnt, s_amt = slot["정지형"]
        m_cnt, m_amt = slot["동작형"]
        total = s_amt + m_amt
        tech, pay_rate, market_rate = MARKET_RATES[market]
        pay = _round(Decimal(total) * Decimal(pay_rate))
        mk = _rounddown(Decimal(total) * Decimal(market_rate))
        rows.append(dict(
            account=account, market=market,
            static_cnt=s_cnt, motion_cnt=m_cnt, cnt=s_cnt + m_cnt,
            static_amt=s_amt, motion_amt=m_amt, amount=total,
            tech=_round(Decimal(total) * Decimal(tech)),
            pay=pay, market_fee=mk, creator=total - pay - mk,
            pay_rate=float(Decimal(pay_rate)), market_rate=float(Decimal(market_rate)),
            odd=slot.get("_odd", []),
        ))
    rows.sort(key=lambda r: (r["account"], r["market"]))

    total_row = dict(
        account="합계", market="",
        **{k: sum(r[k] for r in rows) for k in
           ("static_cnt", "motion_cnt", "cnt", "static_amt", "motion_amt",
            "amount", "tech", "pay", "market_fee", "creator")})
    return rows, total_row


# ------------------------------------------------------------------ 엑셀 출력

_HEAD_FILL = PatternFill("solid", fgColor="FF44546A")
_YELLOW = PatternFill("solid", fgColor="FFFFFF00")
_ACC = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)'


def build_workbook(raws, rows, total_row):
    """요약 시트 + 계정별 로우 데이터 시트로 새 워크북을 만들어 bytes로 반환.

    raws: {계정명: (헤더, 행 리스트)}
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(SUMMARY_SHEET)

    for col, w in zip("BCDEFGHIJKLM",
                      (10, 21.3, 17.8, 17.8, 16.5, 17.8, 17.8, 16.5, 12.7, 14, 14, 22.3)):
        ws.column_dimensions[col].width = w

    for i, name in enumerate(COLUMNS):
        c = ws.cell(row=3, column=2 + i, value=name)
        c.fill = _HEAD_FILL
        c.font = Font(bold=True, color="FFFFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center")

    def put(r, values, bold=False, fmt="#,##0"):
        for i, v in enumerate(values):
            c = ws.cell(row=r, column=2 + i, value=v)
            if i >= 2:
                c.number_format = fmt
                c.alignment = Alignment(horizontal="right")
            if bold:
                c.font = Font(bold=True)
            if i == 11:
                c.fill = _YELLOW

    dash = "-"
    r = 4
    for row in rows:
        put(r, [row["account"], row["market"],
                row["static_cnt"] or dash, row["motion_cnt"] or dash, row["cnt"],
                row["static_amt"] or dash, row["motion_amt"] or dash, row["amount"],
                dash, row["pay"], row["market_fee"], row["creator"]])
        r += 1
    put(r, [total_row["account"], None,
            total_row["static_cnt"], total_row["motion_cnt"], total_row["cnt"],
            total_row["static_amt"], total_row["motion_amt"], total_row["amount"],
            dash, total_row["pay"], total_row["market_fee"], total_row["creator"]],
        bold=True, fmt=_ACC)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")

    # 우측 요율표
    for i, name in enumerate(["마켓 구분", "기술 수수료", "결제 수수료", "마켓 수수료"]):
        ws.cell(row=3, column=15 + i, value=name).font = Font(bold=True)
    ws.column_dimensions["O"].width = 18.5
    for i, (market, (tech, pay, mk)) in enumerate(MARKET_RATES.items()):
        ws.cell(row=4 + i, column=15, value=market)
        for j, v in enumerate((tech, pay, mk)):
            c = ws.cell(row=4 + i, column=16 + j, value=float(Decimal(v)))
            c.number_format = "0.00%"

    # 계정별 로우 데이터 시트
    for account, (header, data_rows) in raws.items():
        sh = wb.create_sheet(account)
        if header:
            sh.append(header)
            for c in sh[1]:
                c.font = Font(bold=True)
        for row in data_rows:
            sh.append(row)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
