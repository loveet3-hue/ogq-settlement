# -*- coding: utf-8 -*-
"""CLI: CSV -> 마켓별 추합본 xlsx.

    python3 make_fee.py <csv> [기존추합본디렉토리] [-o 출력디렉토리]
"""
import argparse
import glob
import io
import os

import openpyxl

import fee_core as core


def main():
    ap = argparse.ArgumentParser(description="네이버 수수료 추합본 생성")
    ap.add_argument("csv", help="타 마켓 판매 로우 데이터 CSV")
    ap.add_argument("src", nargs="?", help="기존 추합본 xlsx가 있는 디렉토리 (선택)")
    ap.add_argument("-o", "--out", default="추합본_출력", help="출력 디렉토리")
    ap.add_argument("--no-merge", action="store_true", help="같은 이름 시트를 병합하지 않고 덮어씀")
    a = ap.parse_args()

    records = core.read_csv(a.csv)

    existing = {}
    if a.src:
        for p in sorted(glob.glob(os.path.join(a.src, "*.xlsx"))):
            if os.path.basename(p).startswith("~$"):
                continue
            blob = open(p, "rb").read()
            m = core.detect_market(openpyxl.load_workbook(io.BytesIO(blob), read_only=True))
            if m:
                existing[m.key] = blob
            else:
                print("  [skip] 마켓 판별 실패: %s" % os.path.basename(p))

    results, unknown = core.build(records, existing=existing, merge=not a.no_merge)

    os.makedirs(a.out, exist_ok=True)
    print("%-14s %-24s %6s %13s %11s %13s  %s"
          % ("판매마켓", "시트", "건수", "판매액", "수수료율", "네이버수수료", "비고"))
    for r in results:
        open(os.path.join(a.out, r["filename"]), "wb").write(r["bytes"])
        note = "신규 파일" if r["is_new"] else ("기존 %d건과 병합" % r["merged"] if r["merged"] else "시트 추가")
        print("%-14s %-24s %6d %13s %10.5f%% %13s  %s"
              % (r["market"], r["sheet"], r["rows"], format(r["total"], ","),
                 r["rate"] * 100, format(r["fee"], ","), note))
    print("%-14s %-24s %6d %13s %11s %13s"
          % ("합계", "", sum(r["rows"] for r in results), format(sum(r["total"] for r in results), ","),
             "", format(sum(r["fee"] for r in results), ",")))

    if unknown:
        print("\n[수수료율 미등록 마켓 - 제외됨]")
        for u in unknown:
            print("  %s: %s건 / %s원" % (u["market"], format(u["rows"], ","), format(u["total"], ",")))
    print("\n출력: %s" % os.path.abspath(a.out))


if __name__ == "__main__":
    main()
