# -*- coding: utf-8 -*-
"""네이버 출신 크리에이터 타 마켓 판매 로우 데이터 -> 마켓별 분류 + 네이버 수수료 추합본 생성.

수수료 구조
    정산 대상 금액 = 판매액 - 결제 수수료 - 마켓 수수료
    네이버 수수료  = 정산 대상 금액 * 15%
따라서 판매액 대비 실효율 = (1 - 결제수수료율 - 마켓수수료율) * 0.15
"""
import csv
import io
import re
import collections
from copy import copy
from datetime import datetime

import openpyxl
from decimal import Decimal, ROUND_HALF_UP

NAVER_SHARE = 0.15


def won(x):
    """엑셀 ROUND와 동일한 사사오입(0.5는 올림)."""
    return int(Decimal(repr(float(x))).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
HEADERS = ["거래시간", "거래 ID", "거래유형", "크리에이터ID", "판매마켓 ID",
           "상품 코드", "상품명", "가격"]
COLS = "ABCDEFGHI"
FULL_HEADERS = ["거래시간", "거래 ID", "거래유형", "크리에이터ID", "판매마켓 ID",
                "판매마켓", "상품 코드", "상품명", "가격"]


class Market:
    """마켓 1개의 수수료 구조와 추합본 시트 규칙."""

    def __init__(self, key, label, pay_fee, market_fee, book, cycle,
                 offset=1, month_fmt="{m}월분"):
        self.key = key                # CSV의 '판매마켓' 값
        self.label = label            # 화면 표시명
        self.pay_fee = pay_fee        # 결제 수수료율
        self.market_fee = market_fee  # 마켓 수수료율
        self.book = book              # 추합본 파일 구분명
        self.cycle = cycle            # "month" | "quarter"
        self.offset = offset          # 정산월 = 판매월 + offset
        self.month_fmt = month_fmt

    @property
    def settle_ratio(self):
        return 1 - self.pay_fee - self.market_fee

    @property
    def rate(self):
        """판매액 대비 네이버 수수료 실효율."""
        return round(self.settle_ratio * NAVER_SHARE, 10)

    @property
    def filename(self):
        return "추합본_%s_네이버수수료.xlsx" % self.book

    def sheet_name(self, dates):
        months = sorted({d.month for d in dates})
        year = min(dates).year
        if self.cycle == "quarter":
            q = (months[0] - 1) // 3 + 1
            return "%d분기 정산(%d.%d월~%d월 분)" % (q, year, (q - 1) * 3 + 1, q * 3)
        m = months[0]
        settle, sy = m + self.offset, year
        if settle > 12:
            settle, sy = settle - 12, year + 1
        return "%d월정산(%d.%s)" % (settle, sy, self.month_fmt.format(m=m))


MARKETS = [
    Market("SOOP OGQ 마켓",  "SOOP OGQ마켓", 0.055, 0.14175, "SOOP",       "month",   offset=1),
    Market("채팅+ 원스토어",  "원스토어 (RCS)", 0.0,   0.40,    "RCS",        "month",   offset=2),
    Market("연합뉴스",        "연합뉴스",     0.0,   0.40,    "연합",       "month",   offset=1),
    Market("채팅+ OGQ 마켓",  "채팅+ OGQ마켓", 0.055, 0.20,    "채팅플러스", "quarter"),
    Market("OGQ 그라폴리오",  "그라폴리오",   0.055, 0.0,     "그라폴리오", "month",   offset=1, month_fmt="{m:02d}월분"),
    Market("네이버 밴드",     "밴드",         0.0,   0.30,    "밴드",       "month",   offset=2),
    Market("프리미엄뉴스",    "프리미엄뉴스", 0.0,   0.0,     "프리미엄뉴스", "month", offset=2),
]
BY_KEY = {m.key: m for m in MARKETS}

DATE_FORMATS = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
                "%Y.%m.%d %H:%M:%S", "%Y.%m.%d %H:%M",
                "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M")


def parse_dt(value):
    if isinstance(value, datetime):
        return value
    s = str(value).strip()
    for f in DATE_FORMATS:
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    return s


def read_csv(data):
    """bytes 또는 파일 경로를 받아 [{컬럼: 값}] 리스트로."""
    if isinstance(data, (bytes, bytearray)):
        text = data.decode("utf-8-sig")
    else:
        text = open(data, encoding="utf-8-sig").read()
    rows = list(csv.DictReader(io.StringIO(text)))
    missing = [c for c in FULL_HEADERS if c not in (rows[0].keys() if rows else [])]
    if missing:
        raise ValueError("CSV에 다음 컬럼이 없습니다: %s" % ", ".join(missing))
    out = []
    for r in rows:
        out.append([
            parse_dt(r["거래시간"]), int(float(r["거래 ID"])), r["거래유형"],
            r["크리에이터ID"], r["판매마켓 ID"], r["판매마켓"],
            r["상품 코드"], r["상품명"], int(float(r["가격"])),
        ])
    return out


def classify(records):
    """판매마켓별로 묶어 {마켓명: [row,...]} 반환. 미등록 마켓도 그대로 포함."""
    buckets = collections.defaultdict(list)
    for r in records:
        buckets[r[5]].append(r)
    for rows in buckets.values():
        rows.sort(key=lambda r: (r[0] if isinstance(r[0], datetime) else datetime.max, r[1]))
    return dict(buckets)


# ---------------------------------------------------------------- 시트 서식

_ORANGE = "FFFFC000"
_BLUE = "FF305496"


def _style_from(ws_src, ws_dst):
    """기존 추합본 시트의 서식을 그대로 복제."""
    for col, dim in ws_src.column_dimensions.items():
        ws_dst.column_dimensions[col].width = dim.width
    for col in COLS:
        for row in (1, 2, 3):
            src = ws_src["%s%d" % (col, row)]
            ws_dst["%s%d" % (col, row)]._style = copy(src._style)


def _style_default(ws):
    """참조할 기존 파일이 없을 때 쓰는 기본 서식."""
    from openpyxl.styles import Font, PatternFill, Alignment
    widths = dict(A=18.56, B=9.56, C=11.44, D=14.33, E=13.56, F=13.33, G=12.67, H=23.22, I=6.44)
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for col in COLS:
        c1, c2 = ws["%s1" % col], ws["%s2" % col]
        c1.fill = PatternFill("solid", fgColor=_ORANGE)
        c1.font = Font(bold=True)
        c1.alignment = Alignment(horizontal="center")
        c2.fill = PatternFill("solid", fgColor=_BLUE)
        c2.font = Font(bold=True, color="FFFFFFFF")
        c2.alignment = Alignment(horizontal="center")
    ws["B1"].number_format = "#,##0"
    ws["A3"].number_format = "yyyy\\-mm\\-dd\\ hh:mm:ss"
    ws["B3"].number_format = "0"
    ws["I3"].number_format = "#,##0;\\-#,##0"


def _read_sheet_rows(ws):
    rows = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if len(r) < 9 or r[1] is None:
            continue
        rows.append([parse_dt(r[0]), int(r[1]), r[2], r[3], r[4], r[5], r[6], r[7], int(r[8] or 0)])
    return rows


def write_sheet(wb, market, rows, tmpl=None, merge=True):
    """wb 맨 앞에 정산 시트를 만든다. 같은 이름 시트가 있으면 거래 ID 기준 병합."""
    name = market.sheet_name([r[0] for r in rows if isinstance(r[0], datetime)])
    merged_from = 0
    if name in wb.sheetnames:
        old = wb[name]
        if merge:
            existing = _read_sheet_rows(old)
            seen = {r[1] for r in rows}
            keep = [r for r in existing if r[1] not in seen]
            merged_from = len(keep)
            rows = keep + rows
            rows.sort(key=lambda r: (r[0] if isinstance(r[0], datetime) else datetime.max, r[1]))
        if tmpl is None:
            tmpl = old
        del wb[name]

    ws = wb.create_sheet(name, 0)
    last = len(rows) + 2
    ws["A1"] = "네이버 수수료"
    ws["B1"] = "=SUM(I3:I%d)*%s" % (last, repr(market.rate))
    for col, h in zip(COLS, FULL_HEADERS):
        ws["%s2" % col] = h
    for i, r in enumerate(rows, start=3):
        for j, col in enumerate(COLS):
            ws["%s%d" % (col, i)] = r[j]

    if tmpl is not None:
        _style_from(tmpl, ws)
    else:
        _style_default(ws)
    # 3행 서식을 데이터 전체에 확장
    for i in range(4, last + 1):
        for col in COLS:
            ws["%s%d" % (col, i)]._style = copy(ws["%s3" % col]._style)

    total = sum(r[8] for r in rows)
    return dict(sheet=name, rows=len(rows), merged=merged_from,
                total=total, fee=won(total * market.rate))


def detect_market(wb):
    """추합본 워크북의 첫 시트를 보고 어느 마켓 파일인지 판별."""
    counts = collections.Counter()
    for ws in wb.worksheets:
        for r in ws.iter_rows(min_row=3, max_row=40, values_only=True):
            if len(r) > 5 and r[5]:
                counts[r[5]] += 1
        if counts:
            break
    if counts:
        key = counts.most_common(1)[0][0]
        if key in BY_KEY:
            return BY_KEY[key]
    return None


def build(records, existing=None, merge=True):
    """records + {마켓명: 기존 워크북 bytes} -> [{market, filename, bytes, ...}]

    existing 에 없는 마켓은 동일 양식으로 새 파일을 만든다.
    """
    existing = existing or {}
    buckets = classify(records)
    results, unknown = [], []

    for name, rows in sorted(buckets.items(), key=lambda kv: -sum(r[8] for r in kv[1])):
        market = BY_KEY.get(name)
        if market is None:
            unknown.append(dict(market=name, rows=len(rows),
                                total=sum(r[8] for r in rows)))
            continue
        blob = existing.get(name)
        if blob:
            wb = openpyxl.load_workbook(io.BytesIO(blob))
            tmpl = wb.worksheets[0]
        else:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            tmpl = None
        info = write_sheet(wb, market, rows, tmpl=tmpl, merge=merge)
        buf = io.BytesIO()
        wb.save(buf)
        info.update(market=market.label, key=market.key, rate=market.rate,
                    filename=market.filename, bytes=buf.getvalue(),
                    is_new=not blob, sheets=wb.sheetnames)
        results.append(info)

    return results, unknown


def breakdown(amount, market):
    """계산기용: 판매액 -> 각 수수료 금액."""
    pay = amount * market.pay_fee
    mk = amount * market.market_fee
    base = amount - pay - mk
    return dict(pay=pay, market=mk, base=base, naver=base * NAVER_SHARE)
