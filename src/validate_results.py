"""Independently compare the saved Excel results with the pre-existing answer key.

Does not import the analyzer or its implementation. Reads files only.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
SHEETS = ["분석요약", "월별계정집계", "전표대차검증", "중복의심전표", "중복의심행"]


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def project_path(value: str) -> Path:
    path = (ROOT / value).resolve()
    require(path.is_relative_to(ROOT), "프로젝트 외부 경로는 허용하지 않습니다.")
    return path


def records(sheet) -> list[dict]:
    values = list(sheet.values)
    require(bool(values), f"빈 시트: {sheet.title}")
    header = values[0]
    require(len(set(header)) == len(header), f"중복 헤더: {sheet.title}")
    return [dict(zip(header, row)) for row in values[1:]]


def integer(value) -> int:
    require(isinstance(value, int) and not isinstance(value, bool), f"정수 셀이 아닙니다: {value!r}")
    return value


def valid_header_style(cell) -> bool:
    color = cell.fill.fgColor
    argb = color.rgb
    return (
        cell.font.bold is True
        and cell.fill.fill_type == "solid"
        and color.type == "rgb"
        and isinstance(argb, str)
        and re.fullmatch(r"[0-9A-Fa-f]{8}", argb) is not None
        and argb[-6:].upper() == "17365D"
    )


def validate(workbook: Path, answer_key: Path) -> dict:
    with answer_key.open(encoding="utf-8-sig", newline="") as stream:
        key = list(csv.DictReader(stream))
    book = load_workbook(workbook, data_only=False)
    try:
        require(book.sheetnames == SHEETS, "시트명 또는 순서 불일치")
        tables = {name: records(book[name]) for name in SHEETS}
        for sheet in book:
            require(sheet.freeze_panes == "A2", f"틀 고정 누락: {sheet.title}")
            require(sheet.auto_filter.ref == sheet.dimensions, f"자동 필터 범위 오류: {sheet.title}")
            for cell in sheet[1]:
                require(valid_header_style(cell), f"헤더 스타일 오류: {sheet.title}")
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    require(cell.data_type != "f", "분석 보고서에 예상하지 않은 수식 존재")
                    if isinstance(cell.value, int):
                        number_format = cell.number_format.replace(r"\(", "(").replace(r"\)", ")")
                        require(number_format == '#,##0;[Red](#,##0);"-"', "숫자 표시 형식 누락")
            require(
                all(
                    c.column_letter in sheet.column_dimensions
                    and sheet.column_dimensions[c.column_letter].width is not None
                    and sheet.column_dimensions[c.column_letter].width >= 17
                    for c in sheet[1]
                ),
                "열 너비 부족 또는 설정 누락",
            )

        expected_monthly = {}
        expected_balance = {}
        for row in key:
            if row["기록유형"] == "월별계정집계":
                k = (row["월"], row["계정코드"], row["계정명"])
                require(k not in expected_monthly, "정답표 집계 키 중복")
                expected_monthly[k] = (int(row["예상차변합계"]), int(row["예상대변합계"]))
            elif row["기록유형"] == "전표검증":
                require(row["전표번호"] not in expected_balance, "정답표 전표 키 중복")
                expected_balance[row["전표번호"]] = int(row["예상대차차액"])
        actual_monthly = {}
        for row in tables["월별계정집계"]:
            k = (row["월"], row["계정코드"], row["계정명"])
            require(k not in actual_monthly, "Excel 집계 키 중복")
            actual_monthly[k] = (integer(row["차변합계"]), integer(row["대변합계"]))
        require(len(expected_monthly) == 33, "정답표의 집계 조합 수가 33이 아닙니다.")
        require(actual_monthly == expected_monthly, "월별·계정별 합계 불일치")

        actual_balance = {}
        for row in tables["전표대차검증"]:
            require(row["전표번호"] not in actual_balance, "Excel 전표 키 중복")
            debit, credit, delta = (integer(row[c]) for c in ("차변합계", "대변합계", "차액"))
            require(debit - credit == delta, "전표 차액 계산 불일치")
            require(row["대차상태"] == ("일치" if delta == 0 else "불일치"), "대차 상태 불일치")
            actual_balance[row["전표번호"]] = delta
        require(actual_balance == expected_balance, "전체 전표의 대차 기대값 불일치")
        imbalances = {v: d for v, d in actual_balance.items() if d != 0}
        require(len(imbalances) == 3, "대차 불일치 건수가 3이 아닙니다.")

        cases = [r for r in key if r["기록유형"] == "검증사례"]
        expected_pairs = {tuple(sorted((r["전표번호"], r["관련전표번호"])))
                          for r in cases if r["사례유형"] == "전표전체복제"}
        actual_pairs = [tuple(sorted((r["전표번호1"], r["전표번호2"]))) for r in tables["중복의심전표"]]
        require(len(actual_pairs) == len(set(actual_pairs)), "중복 후보 쌍이 중복 출력됨")
        require(len(expected_pairs) == 2 and set(actual_pairs) == expected_pairs, "전체 복제 쌍 불일치")
        expected_rows = {(r["전표번호"], tuple(sorted(map(int, r["분개순번목록"].split("|")))))
                         for r in cases if r["사례유형"] == "전표내행복제"}
        actual_rows = []
        for r in tables["중복의심행"]:
            seqs = tuple(sorted(map(int, r["분개순번목록"].split("|"))))
            require(len(seqs) == len(set(seqs)), "중복행 목록의 식별자가 중복됨")
            require(integer(r["반복횟수"]) == len(seqs), "반복횟수 불일치")
            require(integer(r["추가중복행수"]) == len(seqs) - 1, "추가중복행수 불일치")
            actual_rows.append((r["전표번호"], seqs))
        require(len(expected_rows) == 1 and len(actual_rows) == 1 and set(actual_rows) == expected_rows,
                "전표 내 행 복제 불일치")
        require(all(r["판정"] == "검토 필요" for name in ("중복의심전표", "중복의심행") for r in tables[name]),
                "중복 후보가 오류로 자동 확정됨")

        candidate_vouchers = {v for pair in actual_pairs for v in pair} | {v for v, _ in actual_rows}
        normal = [r for r in cases if r["정상오류구분"] == "정상"]
        for row in normal:
            targets = {row["전표번호"], row["관련전표번호"]} - {""}
            require(not (targets & candidate_vouchers), f"정상 사례 중복 오탐: {row['사례ID']}")
            require(all(actual_balance[v] == 0 for v in targets), "정상 사례 대차 오탐")
        normal_counts = dict(Counter(r["사례유형"] for r in normal))
        require(normal_counts == {"정상동일금액": 3, "정상월별반복": 9}, "정상 사례 수 불일치")

        # Cross-foot workbook totals and summary rather than trusting summary labels.
        monthly_totals = defaultdict(lambda: [0, 0])
        for (month, _, _), (debit, credit) in actual_monthly.items():
            monthly_totals[month][0] += debit
            monthly_totals[month][1] += credit
        total_debit = sum(v[0] for v in monthly_totals.values())
        total_credit = sum(v[1] for v in monthly_totals.values())
        require(total_debit == sum(r["차변합계"] for r in tables["전표대차검증"]), "차변 교차 합계 불일치")
        require(total_credit == sum(r["대변합계"] for r in tables["전표대차검증"]), "대변 교차 합계 불일치")
        summary = {r["항목"]: r["값"] for r in tables["분석요약"]}
        required_summary = {"입력행수": sum(integer(r["분개행수"]) for r in tables["전표대차검증"]),
                            "전표수": len(actual_balance), "차변총액": total_debit, "대변총액": total_credit,
                            "대차불일치전표수": len(imbalances), "중복의심전표쌍수": len(actual_pairs),
                            "전표내중복그룹수": len(actual_rows), "추가중복행수": 1, "월별계정조합수": len(actual_monthly)}
        require(all(summary.get(k) == v for k, v in required_summary.items()), "분석요약 교차 검증 실패")
        require(summary["입력행수"] == 131 and summary["전표수"] == 62, "자료 규모 불일치")
        require(summary["시작일"] == "2026-01-03" and summary["종료일"] == "2026-03-23", "기간 불일치")
        require(summary["중복판정"] == "검토 필요: 중복 후보를 실제 오류로 자동 확정하지 않음", "중복 판정 안내 누락")
        return {"status": "PASS", "sheet_rows": {name: len(r) for name, r in tables.items()},
                "monthly_account_matches": len(actual_monthly), "imbalances": imbalances,
                "duplicate_voucher_pairs": sorted(actual_pairs),
                "duplicate_row_groups": [[v, list(s)] for v, s in actual_rows],
                "normal_case_checks": normal_counts, "false_positives": 0,
                "monthly_totals": dict(sorted(monthly_totals.items())),
                "total_debit": total_debit, "total_credit": total_credit,
                "styles_and_saved_values": "PASS"}
    finally:
        book.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", default="outputs/abc_accounting_analysis.xlsx")
    parser.add_argument("--answer-key", default="data/validation_answer_key.csv")
    args = parser.parse_args(argv)
    try:
        result = validate(project_path(args.workbook), project_path(args.answer_key))
        print(json.dumps(result, ensure_ascii=True, indent=2))
        return 0
    except (OSError, ValueError, KeyError, TypeError, csv.Error) as exc:
        print("VALIDATION FAILED: " + str(exc).replace(str(ROOT), "<PROJECT_ROOT>"), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
