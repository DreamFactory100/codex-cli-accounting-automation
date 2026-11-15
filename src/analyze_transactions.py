"""Analyze a transaction CSV only; never consult the validation answer key.

Run from the project root with .venv/Scripts/python.exe -B src/analyze_transactions.py.
All command-line paths are project-relative (absolute paths must also be inside it).
Existing outputs are never replaced. Amounts are integer KRW, not floating point.
"""

from __future__ import annotations

import argparse
import csv
import io
import itertools
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
COLUMNS = ["거래일자", "전표번호", "분개순번", "계정코드", "계정명", "거래처", "차변금액", "대변금액", "적요"]
BUSINESS = [c for c in COLUMNS if c not in ("전표번호", "분개순번")]
REVIEW = "검토 필요"
SHEETS = ["분석요약", "월별계정집계", "전표대차검증", "중복의심전표", "중복의심행"]
# Excel stores numeric cells as doubles. Stay in its exact integer range.
MAX_EXACT_INTEGER = 2**53 - 1


def project_path(value: str) -> Path:
    path = (ROOT / value).resolve()
    if not path.is_relative_to(ROOT):
        raise ValueError("경로는 프로젝트 폴더 안에 있어야 합니다.")
    return path


def load_transactions(path: Path) -> pd.DataFrame:
    """Reject malformed input instead of silently cleaning or dropping records."""
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.reader(stream, strict=True)
        header = next(reader, None)
        if header is None or len(set(header)) != len(header) or set(header) != set(COLUMNS):
            raise ValueError("필수 9개 열이 정확히 한 번씩 필요합니다.")
        records = []
        for number, values in enumerate(reader, 2):
            if len(values) != len(header):
                raise ValueError(f"CSV 레코드 {number}: 열 개수가 다릅니다.")
            record = dict(zip(header, values))
            for column, value in record.items():
                if not value.strip():
                    raise ValueError(f"CSV 레코드 {number}: {column} 누락값입니다.")
                if value != value.strip():
                    raise ValueError(f"CSV 레코드 {number}: {column} 앞뒤 공백을 확인하세요.")
            day = record["거래일자"]
            if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", day):
                raise ValueError(f"CSV 레코드 {number}: 날짜는 YYYY-MM-DD 형식이어야 합니다.")
            try:
                datetime.strptime(day, "%Y-%m-%d")
            except ValueError as exc:
                raise ValueError(f"CSV 레코드 {number}: 유효하지 않은 날짜 {day}") from exc
            for column in ("분개순번", "차변금액", "대변금액"):
                if not re.fullmatch(r"[0-9]+", record[column]):
                    raise ValueError(f"CSV 레코드 {number}: {column}은 음수 없는 정수여야 합니다.")
                record[column] = int(record[column])
                if record[column] > MAX_EXACT_INTEGER:
                    raise ValueError(f"CSV 레코드 {number}: Excel 정수 정밀도 범위 초과입니다.")
            if record["분개순번"] < 1:
                raise ValueError(f"CSV 레코드 {number}: 분개순번은 양의 정수여야 합니다.")
            if (record["차변금액"] > 0) == (record["대변금액"] > 0):
                raise ValueError(f"CSV 레코드 {number}: 차변/대변 한쪽만 양수여야 합니다.")
            records.append(record)
    if not records:
        raise ValueError("거래자료가 비어 있습니다.")
    frame = pd.DataFrame(records, columns=COLUMNS)
    # Python integers prevent fixed-width aggregation overflow.
    for column in ("차변금액", "대변금액"):
        frame[column] = frame[column].astype(object)
    duplicated = frame.duplicated(["전표번호", "분개순번"], keep=False)
    if duplicated.any():
        keys = frame.loc[duplicated, ["전표번호", "분개순번"]].to_dict("records")
        raise ValueError(f"전표번호+분개순번이 중복됩니다: {keys}")
    if (frame.groupby("계정코드")["계정명"].nunique() != 1).any():
        raise ValueError("동일 계정코드의 계정명이 일치하지 않습니다.")
    if (frame.groupby("전표번호")["거래일자"].nunique() != 1).any():
        raise ValueError("하나의 전표에 여러 거래일자가 있습니다.")
    if any(sum(frame[c]) > MAX_EXACT_INTEGER for c in ("차변금액", "대변금액")):
        raise ValueError("합계가 Excel 정수 정밀도 범위를 초과합니다.")
    return frame


def analyze(frame: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Compute totals, balances and strict duplicate candidates from input alone."""
    frame = frame.copy()
    frame["월"] = frame["거래일자"].str[:7]
    monthly = (frame.groupby(["월", "계정코드", "계정명"], as_index=False, sort=True)
               [["차변금액", "대변금액"]].sum()
               .rename(columns={"차변금액": "차변합계", "대변금액": "대변합계"}))
    balance_rows, pair_rows, row_rows = [], [], []
    signature_groups = defaultdict(list)
    for voucher, group in frame.groupby("전표번호", sort=True):
        debit, credit = sum(group["차변금액"]), sum(group["대변금액"])
        delta = debit - credit
        balance_rows.append([voucher, group["거래일자"].iloc[0], len(group), debit, credit,
                             delta, "일치" if delta == 0 else "불일치"])
        signatures = Counter(tuple(r[c] for c in BUSINESS) for r in group.to_dict("records"))
        signature_groups[tuple(sorted(signatures.items()))].append(voucher)
        for signature, count in sorted(signatures.items()):
            if count > 1:
                matching = [r["분개순번"] for r in group.to_dict("records")
                            if tuple(r[c] for c in BUSINESS) == signature]
                row_rows.append([voucher, "|".join(map(str, sorted(matching))), count, count - 1,
                                 *signature, REVIEW])
    for signature, vouchers in signature_groups.items():
        for left, right in itertools.combinations(sorted(vouchers), 2):
            pair_rows.append([left, right, sum(count for _, count in signature),
                              "식별자 제외 7개 필드의 행 구성·반복 횟수 일치", REVIEW])
    balances = pd.DataFrame(balance_rows, columns=["전표번호", "거래일자", "분개행수", "차변합계", "대변합계", "차액", "대차상태"])
    pairs = pd.DataFrame(sorted(pair_rows), columns=["전표번호1", "전표번호2", "분개행수", "탐지근거", "판정"])
    duplicate_rows = pd.DataFrame(row_rows, columns=["전표번호", "분개순번목록", "반복횟수", "추가중복행수", *BUSINESS, "판정"])
    summary = pd.DataFrame([
        ["기업", "ABC Company (가상자료)"], ["입력행수", len(frame)],
        ["전표수", len(balances)], ["시작일", frame["거래일자"].min()],
        ["종료일", frame["거래일자"].max()], ["통화", "KRW"],
        ["차변총액", sum(frame["차변금액"])], ["대변총액", sum(frame["대변금액"])],
        ["대차불일치전표수", int((balances["차액"] != 0).sum())],
        ["중복의심전표쌍수", len(pairs)], ["전표내중복그룹수", len(duplicate_rows)],
        ["추가중복행수", sum(duplicate_rows["추가중복행수"])],
        ["월별계정조합수", len(monthly)],
        ["중복판정", "검토 필요: 중복 후보를 실제 오류로 자동 확정하지 않음"],
        ["중복기준", "날짜·계정·거래처·차변·대변·적요를 비교, 전표번호·분개순번 제외"],
        ["집계기준", "거래일자 월/계정별 원본 합계; 불일치 및 중복 후보 포함"],
        ["입력출처", "실행 시 --input으로 지정한 거래 CSV만 사용"],
    ], columns=["항목", "값"])
    return dict(zip(SHEETS, [summary, monthly, balances, pairs, duplicate_rows]))


def write_workbook(tables: dict[str, pd.DataFrame], path: Path) -> None:
    """Build in memory, then exclusively create the final path without replacing it."""
    if path.exists():
        raise FileExistsError("출력 파일이 이미 있습니다. 덮어쓰지 않습니다.")
    book = Workbook()
    book.remove(book.active)
    for title, table in tables.items():
        sheet = book.create_sheet(title)
        sheet.append(list(table.columns))
        for record in table.itertuples(index=False, name=None):
            sheet.append(list(record))
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor="17365D")
            cell.font = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=11)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        sheet.row_dimensions[1].height = 32
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.font = Font(name="맑은 고딕", size=11)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if isinstance(cell.value, str):
                    cell.data_type = "s"  # Treat external text literally, including leading '='.
                elif isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0;[Red](#,##0);"-"'
                if cell.value in ("불일치", REVIEW):
                    cell.fill = PatternFill("solid", fgColor="FFF2CC")
            sheet.row_dimensions[row[0].row].height = 32
        for index, column in enumerate(table.columns, 1):
            if title == "분석요약":
                width = 28 if index == 1 else 90
            elif column in ("적요", "탐지근거"):
                width = 62
            elif column.startswith("전표번호"):
                width = 25
            elif column in ("계정코드", "분개행수", "반복횟수", "추가중복행수"):
                width = 17
            else:
                width = 23
            sheet.column_dimensions[get_column_letter(index)].width = width
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        sheet.sheet_view.showGridLines = False
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.print_title_rows = "1:1"
    buffer = io.BytesIO()
    book.save(buffer)
    book.close()
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write(buffer.getvalue())


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", default="data/abc_transactions_2026Q1.csv")
    parser.add_argument("--output", default="outputs/abc_accounting_analysis.xlsx")
    args = parser.parse_args(argv)
    try:
        source, output = project_path(args.input), project_path(args.output)
        if output.exists():
            raise FileExistsError("출력 파일이 이미 있습니다. 다른 --output 이름을 지정하세요.")
        if output.suffix.lower() != ".xlsx":
            raise ValueError("출력 확장자는 .xlsx여야 합니다.")
        tables = analyze(load_transactions(source))
        write_workbook(tables, output)
        print(json.dumps({"status": "PASS", "output": output.relative_to(ROOT).as_posix(),
                          "sheet_rows": {name: len(table) for name, table in tables.items()}}, ensure_ascii=True, indent=2))
        return 0
    except (OSError, ValueError, csv.Error) as exc:
        # Do not expose user-specific absolute paths in shareable terminal output.
        message = str(exc).replace(str(ROOT), "<PROJECT_ROOT>")
        print(f"ERROR: {message}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
