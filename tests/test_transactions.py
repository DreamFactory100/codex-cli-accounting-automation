"""Regression tests use memory fixtures and test-owned temporary Excel files."""

import contextlib
import copy
import csv
import hashlib
import io
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
import analyze_transactions as analyzer
import validate_results as validator
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Color, Font, PatternFill


class TransactionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._temporary_directory = tempfile.TemporaryDirectory(
            prefix="transaction-tests-", dir=ROOT
        )
        cls.temporary_path = Path(cls._temporary_directory.name)
        cls.valid_workbook = cls.temporary_path / "valid-analysis.xlsx"
        source = ROOT / "data/abc_transactions_2026Q1.csv"
        tables = analyzer.analyze(analyzer.load_transactions(source))
        analyzer.write_workbook(tables, cls.valid_workbook)

    @classmethod
    def tearDownClass(cls):
        cls._temporary_directory.cleanup()

    def fixture(self):
        return [
            ["2026-01-01", "TEST-001", "1", "1020", "보통예금", "가상거래처", "100", "0", "시험"],
            ["2026-01-01", "TEST-001", "2", "4010", "매출", "가상거래처", "0", "100", "시험"],
        ]

    def load(self, records, header=None):
        buffer = io.StringIO(newline="")
        writer = csv.writer(buffer)
        writer.writerow(analyzer.COLUMNS if header is None else header)
        writer.writerows(records)
        with patch.object(Path, "open", return_value=io.StringIO(buffer.getvalue())):
            return analyzer.load_transactions(ROOT / "memory-only.csv")

    def test_valid_fixture(self):
        result = analyzer.analyze(self.load(self.fixture()))
        self.assertEqual(result["전표대차검증"].iloc[0]["차액"], 0)

    def test_invalid_field_values(self):
        for column, bad_values in {
            0: ["2026-02-30", "2026-1-1", ""],
            2: ["0", "-1", "1.5"],
            6: ["-100", "1.5", "1,000", "", str(2**53)],
            4: ["", " 계정"],
        }.items():
            for value in bad_values:
                with self.subTest(column=column, value=value):
                    rows = self.fixture()
                    rows[0][column] = value
                    with self.assertRaises(ValueError):
                        self.load(rows)

    def test_invalid_headers_and_shape(self):
        for header in [analyzer.COLUMNS[:-1], [analyzer.COLUMNS[0]] * 9]:
            with self.subTest(header=header), self.assertRaises(ValueError):
                self.load(self.fixture(), header)
        rows = self.fixture()
        rows[0].append("EXTRA")
        with self.assertRaises(ValueError):
            self.load(rows)

    def test_empty_input(self):
        with self.assertRaises(ValueError):
            self.load([])

    def test_duplicate_identifier(self):
        rows = self.fixture()
        rows.append(rows[0][:])
        with self.assertRaisesRegex(ValueError, "전표번호\\+분개순번"):
            self.load(rows)

    def test_two_sided_and_zero_entries(self):
        for debit, credit in [("100", "100"), ("0", "0")]:
            rows = self.fixture()
            rows[0][6:8] = [debit, credit]
            with self.subTest(debit=debit, credit=credit), self.assertRaises(ValueError):
                self.load(rows)

    def test_account_name_consistency(self):
        rows = self.fixture()
        rows[1][3] = "1020"
        with self.assertRaises(ValueError):
            self.load(rows)

    def test_voucher_date_consistency(self):
        rows = self.fixture()
        rows[1][0] = "2026-01-02"
        with self.assertRaises(ValueError):
            self.load(rows)

    def test_aggregate_precision_limit(self):
        rows = self.fixture()
        rows[0][6] = str(2**53 - 1)
        rows[1][7] = str(2**53 - 1)
        extra = copy.deepcopy(rows)
        for row in extra:
            row[1] = "TEST-002"
        with self.assertRaisesRegex(ValueError, "합계"):
            self.load(rows + extra)

    def test_duplicate_signature_ignores_order_and_ids(self):
        rows = self.fixture()
        clone = copy.deepcopy(rows[::-1])
        for seq, row in enumerate(clone, 1):
            row[1], row[2] = "TEST-002", str(seq)
        tables = analyzer.analyze(self.load(rows + clone))
        self.assertEqual(len(tables["중복의심전표"]), 1)
        self.assertEqual(tables["중복의심전표"].iloc[0]["판정"], "검토 필요")

    def test_duplicate_signature_preserves_multiplicity(self):
        rows = self.fixture()
        clone = copy.deepcopy(rows)
        for row in clone:
            row[1] = "TEST-002"
        extra = clone[0][:]
        extra[2] = "3"
        tables = analyzer.analyze(self.load(rows + clone + [extra]))
        self.assertEqual(len(tables["중복의심전표"]), 0)
        self.assertEqual(tables["중복의심행"].iloc[0]["분개순번목록"], "1|3")

    def test_normal_same_amount_distinct_memo(self):
        rows = self.fixture()
        second = copy.deepcopy(rows)
        for row in second:
            row[1], row[8] = "TEST-002", "별도 주문"
        self.assertTrue(analyzer.analyze(self.load(rows + second))["중복의심전표"].empty)

    def test_project_path_boundary(self):
        with self.assertRaises(ValueError):
            analyzer.project_path("../outside.xlsx")

    def test_existing_output_not_replaced(self):
        path = self.temporary_path / "existing-output.xlsx"
        path.write_bytes(b"existing output must be preserved")
        before = hashlib.sha256(path.read_bytes()).digest()
        with contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(
                analyzer.main(["--output", str(path.relative_to(ROOT))]), 1
            )
        self.assertEqual(before, hashlib.sha256(path.read_bytes()).digest())

    def test_validator_detects_saved_value_mismatch(self):
        original = validator.records

        def altered(sheet):
            rows = original(sheet)
            if sheet.title == "월별계정집계":
                rows[0]["차변합계"] += 1
            return rows

        with patch.object(validator, "records", side_effect=altered):
            with self.assertRaisesRegex(ValueError, "합계 불일치"):
                validator.validate(self.valid_workbook,
                                   ROOT / "data/validation_answer_key.csv")

    def test_validator_rejects_narrow_or_missing_column_width(self):
        for case in ("narrow", "missing"):
            with self.subTest(case=case):
                workbook = self.temporary_path / f"invalid-width-{case}.xlsx"
                shutil.copy2(self.valid_workbook, workbook)
                book = load_workbook(workbook)
                sheet = book[book.sheetnames[0]]
                if case == "narrow":
                    sheet.column_dimensions["A"].width = 16
                else:
                    del sheet.column_dimensions["A"]
                book.save(workbook)
                book.close()

                with self.assertRaisesRegex(ValueError, "열 너비"):
                    validator.validate(
                        workbook, ROOT / "data/validation_answer_key.csv"
                    )


    def test_validator_header_style_accepts_argb_and_rejects_invalid_styles(self):
        cell = Workbook().active["A1"]
        cell.font = Font(bold=True)

        for argb in ("FF17365D", "0017365D"):
            with self.subTest(valid_argb=argb):
                cell.fill = PatternFill(fill_type="solid", fgColor=argb)
                self.assertTrue(validator.valid_header_style(cell))

        invalid_fills = (
            PatternFill(fill_type="solid", fgColor="FF17365E"),
            PatternFill(fill_type="darkGrid", fgColor="FF17365D"),
            PatternFill(fill_type="solid", fgColor=Color(indexed=1)),
        )
        for fill in invalid_fills:
            with self.subTest(fill=fill):
                cell.fill = fill
                self.assertFalse(validator.valid_header_style(cell))

        cell.fill = PatternFill(fill_type="solid", fgColor="FF17365D")
        cell.font = Font(bold=False)
        self.assertFalse(validator.valid_header_style(cell))


if __name__ == "__main__":
    unittest.main(verbosity=2)
