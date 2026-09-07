# -*- coding: utf-8 -*-
"""OGQ 정산 — 마켓별 정산 자동화 웹앱."""
import io
import zipfile
from datetime import datetime

import streamlit as st

import fee_core as core
import stars_core as stars

st.set_page_config(page_title="OGQ 정산", page_icon="🧾", layout="wide")

st.markdown("""
<style>
  .block-container {padding-top: 2.2rem; max-width: 1240px;}
  [data-testid="stMetricValue"] {font-size: 1.5rem;}
  .hint {color:#6b7280; font-size:0.86rem; line-height:1.6;}
  .warnbox {background:#fef3c7; border-left:4px solid #f59e0b;
            padding:12px 16px; border-radius:6px; font-size:0.9rem;}
  .tblwrap {overflow-x:auto; width:100%; padding-bottom:4px;}
  table.ftbl {border-collapse:collapse; width:100%; font-size:0.9rem;
              font-variant-numeric:tabular-nums;}
  table.ftbl th {background:#f1f5f9; color:#0f172a; font-weight:600;
                 padding:9px 12px; border-bottom:2px solid #cbd5e1; text-align:right;
                 white-space:nowrap;}
  table.ftbl th:first-child, table.ftbl td:first-child {text-align:left;}
  table.ftbl td {padding:8px 12px; border-bottom:1px solid #e2e8f0; text-align:right;
                 white-space:nowrap;}
  table.ftbl td.hi {font-weight:700; color:#b91c1c;}
  table.ftbl td.mut {color:#94a3b8;}
  table.rtbl {border-collapse:collapse; width:100%; font-size:0.95rem;
              font-variant-numeric:tabular-nums;}
  table.rtbl th, table.rtbl td {border:1px solid #a9b2bf; padding:9px 12px;
                                white-space:nowrap; text-align:right;}
  table.rtbl th {background:#e8eef7; color:#1f2937; font-weight:700;}
  table.rtbl th.fee {background:#dfe3e8;}
  table.rtbl th.name, table.rtbl td.name {text-align:left;}
  table.rtbl td.fee {color:#c00000; font-weight:700;}
  table.rtbl td.creator {background:#ffff99; font-weight:700;}
  table.rtbl tr.sum td {background:#ffff99; font-weight:700;}
  table.rtbl tr.sum td.creator {background:#ffe14d;}
</style>
""", unsafe_allow_html=True)


def html_table(headers, rows, cls="ftbl"):
    head = "".join("<th>%s</th>" % h for h in headers)
    body = "".join("<tr>%s</tr>" % "".join(
        c if c.startswith("<td") else "<td>%s</td>" % c for c in r) for r in rows)
    st.markdown('<div class="tblwrap"><table class="%s"><thead><tr>%s</tr></thead>'
                '<tbody>%s</tbody></table></div>' % (cls, head, body),
                unsafe_allow_html=True)


def pct(v, digits=2):
    return "-" if not v else "%.*f%%" % (digits, v * 100)


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def num(v, dash_if_zero=False):
    return "-" if (dash_if_zero and not v) else format(v, ",")


def money(v):
    """소수점이 있으면 살려서 보여준다 (마켓 수수료 1,417.5 처럼). 0은 '-'."""
    f = float(v)
    if f == 0:
        return "-"
    return format(int(f), ",") if f == int(f) else format(f, ",.2f")


st.title("🧾 OGQ 정산")

PAGES = ["네이버 출신 작가 정산", "SOOP 이모티콘 (스타즈&유니즈) 정산"]
with st.sidebar:
    st.markdown("### 정산 구분")
    page = st.radio("정산 구분", PAGES, label_visibility="collapsed")
    st.divider()
    st.markdown('<p class="hint">각 정산의 로우 데이터를 올리면 마켓별로 분류하고 '
                '수수료를 계산해 엑셀로 내려받을 수 있습니다.</p>', unsafe_allow_html=True)


# ════════════════════════════════════════════════ 1. 네이버 출신 작가 정산
def page_naver():
    st.caption("타 마켓 판매 로우 데이터 CSV → 마켓별 분류 → 네이버 수수료 계산 → 추합본 엑셀")
    tab_build, tab_calc = st.tabs(["추합본 생성", "수수료 계산기"])

    with tab_build:
        left, right = st.columns([1, 1], gap="large")
        with left:
            st.subheader("1. 로우 데이터 CSV")
            csv_file = st.file_uploader(
                "네이버 출신 크리에이터 타 마켓 판매 로우 데이터",
                type=["csv", "xlsx", "numbers"], key="csv")
            st.markdown('<p class="hint">CSV · 엑셀 · 넘버스 모두 됩니다. 필수 컬럼: 거래시간 · '
                        '거래 ID · 거래유형 · 크리에이터ID · 판매마켓 ID · 판매마켓 · '
                        '상품 코드 · 상품명 · 가격</p>', unsafe_allow_html=True)
        with right:
            st.subheader("2. 기존 추합본 (선택)")
            xlsx_files = st.file_uploader(
                "이어붙일 추합본 엑셀 — 여러 개 한 번에 올려도 됩니다",
                type=["xlsx"], accept_multiple_files=True, key="xlsx")
            st.markdown('<p class="hint">올리면 <b>맨 앞에 새 정산 시트를 추가</b>하고 과거 시트는 '
                        '그대로 둡니다. 올리지 않은 마켓은 같은 양식으로 새 파일을 만듭니다.</p>',
                        unsafe_allow_html=True)
            merge = st.checkbox("같은 이름의 시트가 있으면 병합 (거래 ID 기준 중복 제거)", value=True,
                                help="채팅+ OGQ마켓처럼 분기 단위 시트에 다음 달치를 이어붙일 때 사용합니다.")

        st.divider()

        records = None
        if not csv_file:
            st.info("CSV를 올리면 마켓별 집계와 다운로드 버튼이 나타납니다.")
        else:
            try:
                records = core.read_csv(csv_file.getvalue(), csv_file.name)
            except Exception as e:                               # noqa: BLE001
                st.error("CSV를 읽지 못했습니다 — %s" % e)

        if not records:
            return

        existing, unmatched = {}, []
        for f in xlsx_files or []:
            blob = f.getvalue()
            try:
                import openpyxl
                market = core.detect_market(
                    openpyxl.load_workbook(io.BytesIO(blob), read_only=True))
            except Exception:                                    # noqa: BLE001
                market = None
            if market:
                existing[market.key] = blob
            else:
                unmatched.append(f.name)
        if unmatched:
            st.warning("어느 마켓인지 판별하지 못한 파일 (무시하고 새로 만듭니다): %s"
                       % ", ".join(unmatched))

        results, unknown = core.build(records, existing=existing, merge=merge)
        total_sales = sum(r["total"] for r in results)
        total_fee = sum(r["fee"] for r in results)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("거래 건수", "%s건" % format(len(records), ","))
        c2.metric("마켓 수", "%d개" % len(results))
        c3.metric("판매액 합계", "%s원" % format(total_sales, ","))
        c4.metric("정산금 합계", "%s원" % format(total_fee, ","))

        st.subheader("마켓별 결과")
        body = "".join(
            '<tr><td class="name">%s</td><td>%s</td><td>%s</td><td class="fee">%s</td></tr>'
            % (r["market"], r["sheet"], num(r["total"]), num(r["fee"])) for r in results)
        body += ('<tr class="sum"><td class="name">합계</td><td></td>'
                 '<td>%s</td><td class="fee">%s</td></tr>'
                 % (num(total_sales), num(total_fee)))
        st.markdown('<div class="tblwrap"><table class="rtbl"><thead><tr>'
                    '<th class="name">판매마켓</th><th>시트</th><th>판매액</th>'
                    '<th class="fee">정산금</th></tr></thead>'
                    '<tbody>%s</tbody></table></div>' % body, unsafe_allow_html=True)
        st.write("")

        with st.expander("상세 — 수수료를 어떻게 뗐는지 · 건수 · 처리 내역"):
            st.markdown('<p class="hint">판매액 − ① 결제 수수료 − ② 마켓 수수료 = 정산 대상 금액 · '
                        '정산금 = 정산 대상 금액 × 15%</p>', unsafe_allow_html=True)
            html_table(["판매마켓", "건수", "판매액", "① 결제", "② 마켓",
                        "정산 대상 금액", "정산금", "처리"],
                       [[
                           "<td><b>%s</b></td>" % r["market"], num(r["rows"]),
                           num(r["total"]),
                           "%s<br><small>%s</small>" % (money(r["calc"]["pay"]),
                                                        pct(r["obj"].pay_fee)),
                           "%s<br><small>%s</small>" % (money(r["calc"]["market_fee"]),
                                                        pct(r["obj"].market_fee, 3)),
                           money(r["calc"]["base"]),
                           '<td class="hi">%s</td>' % num(r["fee"]),
                           "<td>%s</td>" % ("신규 파일" if r["is_new"]
                                            else ("기존 %d건과 병합" % r["merged"]
                                                  if r["merged"] else "시트 추가")),
                       ] for r in results])
            st.write("")

        if unknown:
            items = "".join("<li><b>%s</b> — %s건 / %s원</li>"
                            % (u["market"], num(u["rows"]), num(u["total"])) for u in unknown)
            st.markdown('<div class="warnbox">⚠️ <b>수수료율이 등록되지 않은 마켓</b>이 있어 '
                        '제외했습니다.<ul>%s</ul>결제 수수료·마켓 수수료율을 확인해 '
                        '<code>fee_core.py</code>의 <code>MARKETS</code>에 추가해 주세요.</div>'
                        % items, unsafe_allow_html=True)

        st.subheader("다운로드")
        stamp = datetime.now().strftime("%Y%m")
        combined = core.build_combined(results)
        combined_name = "네이버출신작가_정산_전체합본_%s.xlsx" % stamp
        st.download_button(
            "⬇ 전체 합본 파일 받기 (정산 요약 + 마켓별 전체 내역 + 수수료 안내)",
            combined, combined_name, XLSX_MIME, type="primary", use_container_width=True)
        st.markdown('<p class="hint">합본에는 마켓별 전체 거래 내역이 시트로 들어가고, '
                    '요약 시트가 그 내역을 수식으로 합산합니다. 아래 개별 파일은 기존 '
                    '추합본에 새 정산 시트를 붙인 것입니다.</p>', unsafe_allow_html=True)

        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr(combined_name, combined)
            for r in results:
                z.writestr(r["filename"], r["bytes"])
        st.download_button("📦 합본 + 마켓별 %d개 ZIP으로 받기" % len(results), zbuf.getvalue(),
                           "추합본_%s.zip" % stamp,
                           "application/zip", use_container_width=True)
        cols = st.columns(min(3, len(results)) or 1)
        for i, r in enumerate(results):
            with cols[i % len(cols)]:
                st.download_button("⬇ %s (%s건)" % (r["market"], num(r["rows"])),
                                   r["bytes"], r["filename"],
                                   "application/vnd.openxmlformats-officedocument."
                                   "spreadsheetml.sheet",
                                   key="dl_%s" % r["key"], use_container_width=True)

    with tab_calc:
        st.subheader("네이버 출신 작가 계산기")
        amount = st.number_input("판매액 (원)", min_value=0, value=10000, step=1000, key="na")
        rows = []
        for m in core.MARKETS:
            b = core.breakdown(amount, m)
            rows.append(["<td><b>%s</b></td>" % m.label,
                         pct(m.pay_fee), money(b["pay"]),
                         pct(m.market_fee, 3), money(b["market"]),
                         money(b["base"]),
                         '<td class="hi">%s</td>' % num(b["naver"])])
        html_table(["판매마켓", "결제 수수료율", "결제 수수료", "마켓 수수료율", "마켓 수수료",
                    "정산 대상 금액", "네이버 수수료 (15%)"], rows)
        st.markdown('<p class="hint">정산 대상 금액 = 판매액 − 결제 수수료 − 마켓 수수료 · '
                    '네이버 수수료 = 정산 대상 금액 × 15%. 중간 단계는 반올림하지 않고 '
                    '마지막 네이버 수수료만 원 단위로 반올림합니다.</p>',
                    unsafe_allow_html=True)


# ════════════════════════════════════ 2. SOOP 이모티콘 (스타즈&유니즈) 정산
def page_stars():
    st.caption("스타즈 · 유니즈 월별 매출 로우 데이터 → 계정 · 마켓별 판매 집계 → 정산 요약 엑셀")
    with st.container():
        ACCOUNTS = ["스타즈", "유니즈"]
        cols = st.columns(len(ACCOUNTS), gap="large")
        uploads = {}
        for col, name in zip(cols, ACCOUNTS):
            with col:
                st.subheader("%s 파일" % name)
                uploads[name] = st.file_uploader(
                    "%s 월별 매출 로우 데이터" % name,
                    type=["xlsx", "csv", "numbers"], key="up_%s" % name)
        st.markdown('<p class="hint">한쪽만 올려도 됩니다. 필수 컬럼: 콘텐츠타입 · 판매마켓 · 판매금액 · '
                    '단가는 정지형(스티커) 2,000원 / 동작형(애니메이션 스티커) 3,000원 고정입니다.</p>',
                    unsafe_allow_html=True)
        st.divider()

        if not any(uploads.values()):
            st.info("스타즈 · 유니즈 파일을 올리면 정산 요약이 나타납니다.")
            return

        accounts, raws, bad_types, bad_markets = {}, {}, set(), set()
        for name, up in uploads.items():
            if not up:
                continue
            try:
                recs, raw, bt, bm = stars.read_account(up.getvalue(), up.name)
            except Exception as e:                               # noqa: BLE001
                st.error("%s 파일을 읽지 못했습니다 — %s" % (name, e))
                return
            if not recs and not raw["rows"]:
                st.error("%s 파일에서 콘텐츠타입 · 판매마켓 · 판매금액 컬럼을 가진 시트를 "
                         "찾지 못했습니다." % name)
                return
            accounts[name] = recs
            raws[name] = raw
            bad_types |= set(bt)
            bad_markets |= set(bm)

        notes = []
        if bad_types:
            notes.append("등록되지 않은 콘텐츠타입 (제외됨): %s" % ", ".join(sorted(bad_types)))
        if bad_markets:
            notes.append("수수료율이 등록되지 않은 판매마켓 (제외됨): %s"
                         % ", ".join(sorted(bad_markets)))
        if notes:
            st.markdown('<div class="warnbox">⚠️ %s<br><code>stars_core.py</code>의 '
                        '<code>CONTENT_TYPES</code> / <code>MARKET_RATES</code>에 추가해 '
                        '주세요.</div>' % "<br>".join(notes), unsafe_allow_html=True)

        rows, total = stars.summarize(accounts)
        if not rows:
            st.error("집계할 수 있는 거래가 없습니다.")
            return

        odd = [(r["account"], r["market"], r["odd"]) for r in rows if r["odd"]]
        if odd:
            st.markdown('<div class="warnbox">⚠️ 단가(2,000 / 3,000원)로 나누어떨어지지 않는 '
                        '판매 금액이 있어 해당 건의 <b>판매 개수만</b> 빠졌습니다 (금액은 반영). '
                        '%s</div>'
                        % ", ".join("%s / %s %d건" % (a, m, len(o)) for a, m, o in odd),
                        unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("계정 수", "%d개" % len({r["account"] for r in rows}))
        c2.metric("판매 개수", "%s개" % num(total["cnt"]))
        c3.metric("판매 금액", "%s원" % num(total["amount"]))
        c4.metric("크리에이터 정산 금액", "%s원" % num(total["creator"]))

        st.subheader("정산 요약")
        body = ""
        for r in rows:
            body += ('<tr><td class="name">%s</td><td class="name">%s</td>'
                     '<td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td>'
                     '<td>%s</td><td>%s</td><td>%s</td><td class="creator">%s</td></tr>'
                     % (r["account"], r["market"],
                        num(r["static_cnt"], True), num(r["motion_cnt"], True), num(r["cnt"]),
                        num(r["static_amt"], True), num(r["motion_amt"], True), num(r["amount"]),
                        "-", num(r["pay"]), num(r["market_fee"]), num(r["creator"])))
        body += ('<tr class="sum"><td class="name">합계</td><td></td>'
                 '<td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td>'
                 '<td>%s</td><td>%s</td><td>%s</td><td class="creator">%s</td></tr>'
                 % (num(total["static_cnt"], True), num(total["motion_cnt"], True),
                    num(total["cnt"]), num(total["static_amt"], True),
                    num(total["motion_amt"], True), num(total["amount"]),
                    "-", num(total["pay"]), num(total["market_fee"]), num(total["creator"])))
        st.markdown('<div class="tblwrap"><table class="rtbl"><thead><tr>'
                    '<th class="name">계정</th><th class="name">마켓</th>'
                    '<th>정지형 판매 개수</th><th>동작형 판매 개수</th><th>판매 개수(합계)</th>'
                    '<th>정지형 판매 금액</th><th>동작형 판매 금액</th><th>판매 금액(합계)</th>'
                    '<th>기술 수수료</th><th>결제 수수료</th><th>마켓 수수료</th>'
                    '<th class="fee">크리에이터 정산 금액</th></tr></thead>'
                    '<tbody>%s</tbody></table></div>' % body, unsafe_allow_html=True)
        st.write("")

        period = next((r["period"] for r in raws.values() if r["period"]),
                      datetime.now().strftime("%y%m"))

        st.subheader("다운로드")
        st.markdown('<p class="hint">모든 파일의 요약 시트는 수식으로 들어갑니다. '
                    '로우 데이터를 고치면 요약이 따라 바뀝니다.</p>', unsafe_allow_html=True)

        files = [("전체", raws, rows)]
        if len(raws) > 1:
            for name in raws:
                sub = [r for r in rows if r["account"] == name]
                if sub:
                    files.append((name, {name: raws[name]}, sub))

        blobs = []
        for label, sub_raws, sub_rows in files:
            blobs.append((label,
                          "정산요약_%s_%s.xlsx" % (label, period),
                          stars.build_workbook(sub_raws, sub_rows,
                                               stars.make_total(sub_rows))))

        label, fname, blob = blobs[0]
        st.download_button("⬇ 전체 요약 파일 받기 (요약 + %s 시트)" % " + ".join(raws),
                           blob, fname, XLSX_MIME, type="primary",
                           use_container_width=True)

        if len(blobs) > 1:
            cols = st.columns(len(blobs) - 1)
            for col, (label, fname, blob) in zip(cols, blobs[1:]):
                with col:
                    st.download_button("⬇ %s 개별 파일" % label, blob, fname, XLSX_MIME,
                                       key="dl_stars_%s" % label,
                                       use_container_width=True)

            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
                for _l, fn, b in blobs:
                    z.writestr(fn, b)
            st.download_button("📦 전체 + 개별 %d개 ZIP으로 받기" % len(blobs),
                               zbuf.getvalue(), "정산요약_%s.zip" % period,
                               "application/zip", use_container_width=True)


(page_naver if page == PAGES[0] else page_stars)()
