"""
Utility script to dump SMPP 여성/중소기업 정보를 CSV로 저장합니다.

예)
python smpp_dump.py ^
    --excel C:\\path\\companies.xlsx ^
    --output women_result.csv ^
    --user SMPP_ID
"""

from __future__ import annotations

import argparse
import csv
import getpass
from typing import List

from smpp_excel import extract_biz_numbers_from_excel
from smpp_service import check_corps


def dump_results_to_csv(rows: List[dict], output_path: str) -> None:
    header = [
        "사업자등록번호",
        "여성기업_확인일",
        "여성기업_만료일",
        "중소기업_확인일",
        "중소기업_만료일",
        "비고",
    ]
    with open(output_path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=header)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def main():
    parser = argparse.ArgumentParser(description="SMPP 여성/중소기업 유효기간 조회 덤프")
    parser.add_argument("--excel", required=True, help="사업자번호가 들어있는 Excel 파일 경로")
    parser.add_argument("--output", required=True, help="저장할 CSV 경로")
    parser.add_argument("--user", required=True, help="SMPP 로그인 ID")
    parser.add_argument("--password", help="SMPP 로그인 PW (생략 시 입력 요청)")
    parser.add_argument("--delay", type=float, default=0.0, help="요청 사이 지연(초)")
    args = parser.parse_args()

    password = args.password or getpass.getpass("SMPP PW: ")

    biz_nos = extract_biz_numbers_from_excel(args.excel)
    print(f"[+] 사업자번호 {len(biz_nos)}건 조회 시작")

    rows = []

    def _progress(idx, total, biz_no):
        print(f"  - {idx}/{total}: {biz_no}", end="\r", flush=True)

    results = check_corps(
        user_id=args.user,
        password=password,
        biz_nos=biz_nos,
        delay_seconds=args.delay,
        progress_callback=_progress,
    )
    print()

    for result in results:
        if result.features:
            row = {
                "사업자등록번호": result.biz_no,
                "여성기업_확인일": result.features.women_confirm_date or "",
                "여성기업_만료일": result.features.women_expire_date or "",
                "중소기업_확인일": result.features.small_confirm_date or "",
                "중소기업_만료일": result.features.small_expire_date or "",
                "비고": "",
            }
        else:
            row = {
                "사업자등록번호": result.biz_no,
                "여성기업_확인일": "",
                "여성기업_만료일": "",
                "중소기업_확인일": "",
                "중소기업_만료일": "",
                "비고": result.error or "조회 실패",
            }
        rows.append(row)

    dump_results_to_csv(rows, args.output)
    print(f"[+] CSV 저장 완료: {args.output}")


if __name__ == "__main__":
    main()
