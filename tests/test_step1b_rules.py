"""Unit tests for Step 1B rule behavior."""

from __future__ import annotations

from datetime import date, datetime
import unittest

from grader.agents.xlsx_agent import (
    TableData,
    check_date_format,
    check_duplicates,
    check_notes,
    check_null_values,
    check_total_value,
    check_transaction_type,
    has_unsplit_name_issue,
)
from grader.constants import REQUIRED_COLUMNS


def base_clean_row() -> dict[str, object]:
    return {
        "TransactionID": "T0001",
        "AccountNumber": "A1001",
        "CustomerID": "C2001",
        "CustomerFName": "Tony",
        "CustomerLName": "Stark",
        "TransactionDate": "01/31/2024",
        "TransactionType": "Deposit",
        "Merchant": "Oscorp Capital",
        "City": "New York",
        "State": "NY",
        "TotalValue": 1200.50,
        "BalanceAfter": 3000.25,
        "Notes": "",
    }


def col_map() -> dict[str, str]:
    return {name: name for name in REQUIRED_COLUMNS}


class FakeCell:
    def __init__(self, number_format: str) -> None:
        self.number_format = number_format


class FakeWorksheet:
    def __init__(self, formats: dict[tuple[int, int], str]) -> None:
        self.formats = formats

    def cell(self, row: int, column: int) -> FakeCell:
        return FakeCell(self.formats.get((row, column), "General"))


def single_column_table(header: str, values: list[object]) -> TableData:
    rows = [{header: value} for value in values]
    row_numbers = list(range(2, 2 + len(values)))
    return TableData(
        headers=[header],
        rows=rows,
        row_numbers=row_numbers,
        column_index_by_header={header: 1},
    )


class Step1BRuleTests(unittest.TestCase):
    def test_name_split_accepts_case_variation(self) -> None:
        self.assertFalse(has_unsplit_name_issue("tony", "stark"))
        self.assertFalse(has_unsplit_name_issue("TONY", "STARK"))

    def test_name_split_flags_unsplit_full_name(self) -> None:
        self.assertTrue(has_unsplit_name_issue("Tony Stark", ""))
        self.assertTrue(has_unsplit_name_issue("Stark, Tony", ""))

    def test_transaction_type_accepts_full_words_case_insensitive(self) -> None:
        rows = [
            {"TransactionType": "deposit"},
            {"TransactionType": "Withdrawal"},
            {"TransactionType": "FEE"},
            {"TransactionType": "transfer"},
            {"TransactionType": "PAYMENT"},
            {"TransactionType": "Transfer."},
        ]
        issue, _, _ = check_transaction_type(rows, {"TransactionType": "TransactionType"})
        self.assertFalse(issue)

    def test_transaction_type_flags_abbreviations(self) -> None:
        rows = [{"TransactionType": "DEP"}, {"TransactionType": "TRA"}, {"TransactionType": "TRA."}]
        issue, note, _ = check_transaction_type(rows, {"TransactionType": "TransactionType"})
        self.assertTrue(issue)
        self.assertIn("Detected 3 invalid TransactionType row(s).", note)
        self.assertIn("abbreviations:", note)

    def test_transaction_type_mixed_values_only_counts_invalid_rows(self) -> None:
        rows = [
            {"TransactionType": "Transfer"},
            {"TransactionType": "Payment"},
            {"TransactionType": "TRA"},
        ]
        issue, note, _ = check_transaction_type(rows, {"TransactionType": "TransactionType"})
        self.assertTrue(issue)
        self.assertIn("Detected 1 invalid TransactionType row(s).", note)
        self.assertIn("Full-word values are accepted", note)

    def test_null_rule_allows_empty_cells(self) -> None:
        row = base_clean_row()
        row["Notes"] = ""
        row["Merchant"] = ""
        row["City"] = ""
        row["State"] = ""
        row["TransactionType"] = "Transfer."

        issue, _, _ = check_null_values([row], col_map())
        self.assertFalse(issue)

    def test_null_rule_flags_explicit_null_markers(self) -> None:
        row = base_clean_row()
        row["CustomerID"] = "null"

        issue, _, _ = check_null_values([row], col_map())
        self.assertTrue(issue)

    def test_notes_rule_flags_qerr_and_special_chars(self) -> None:
        row1 = base_clean_row()
        row1["Notes"] = "possible ?err on transfer"
        row2 = base_clean_row()
        row2["Notes"] = "follow-up required @ branch"

        issue, _, _ = check_notes([row1, row2], {"Notes": "Notes"})
        self.assertTrue(issue)

    def test_duplicates_ignore_sparse_rows(self) -> None:
        rows = [
            {
                "TransactionID": "",
                "AccountNumber": "",
                "CustomerID": "",
                "CustomerFName": "",
                "CustomerLName": "",
            },
            {
                "TransactionID": "",
                "AccountNumber": "",
                "CustomerID": "",
                "CustomerFName": "",
                "CustomerLName": "",
            },
        ]
        mapping = {
            "TransactionID": "TransactionID",
            "AccountNumber": "AccountNumber",
            "CustomerID": "CustomerID",
            "CustomerFName": "CustomerFName",
            "CustomerLName": "CustomerLName",
        }

        issue, _ = check_duplicates(rows, mapping)
        self.assertFalse(issue)

    def test_duplicates_flag_exact_row_match(self) -> None:
        row = base_clean_row()
        rows = [row.copy(), row.copy()]

        issue, note = check_duplicates(rows, col_map())
        self.assertTrue(issue)
        self.assertIn("Found 1 duplicate row(s).", note)

    def test_duplicates_allow_same_transaction_id_with_different_type(self) -> None:
        row1 = base_clean_row()
        row2 = base_clean_row()
        row2["TransactionType"] = "Withdrawal"

        issue, _ = check_duplicates([row1, row2], col_map())
        self.assertFalse(issue)

    def test_date_rule_accepts_date_cells_without_format(self) -> None:
        ws = FakeWorksheet({})
        table = single_column_table("TransactionDate", [date(2024, 1, 15), datetime(2024, 1, 16, 0, 0)])

        issue, _, review = check_date_format(ws, table, {"TransactionDate": "TransactionDate"})
        self.assertFalse(issue)
        self.assertEqual(review, [])

    def test_date_rule_accepts_excel_serial_dates_without_format(self) -> None:
        ws = FakeWorksheet({})
        table = single_column_table("TransactionDate", [45205, 45206.0])

        issue, _, review = check_date_format(ws, table, {"TransactionDate": "TransactionDate"})
        self.assertFalse(issue)
        self.assertEqual(review, [])

    def test_date_rule_flags_excel_serial_with_fractional_time(self) -> None:
        ws = FakeWorksheet({})
        table = single_column_table("TransactionDate", [45205.5])

        issue, _, _ = check_date_format(ws, table, {"TransactionDate": "TransactionDate"})
        self.assertTrue(issue)

    def test_date_rule_flags_datetime_with_time_component(self) -> None:
        ws = FakeWorksheet({})
        table = single_column_table("TransactionDate", [datetime(2024, 1, 15, 13, 45, 0)])

        issue, _, _ = check_date_format(ws, table, {"TransactionDate": "TransactionDate"})
        self.assertTrue(issue)

    def test_currency_rule_no_review_for_two_decimal_general_numeric(self) -> None:
        ws = FakeWorksheet({})
        table = single_column_table("TotalValue", [10.25, -3.25, 0.0])

        issue, _, review = check_total_value(ws, table, {"TotalValue": "TotalValue"})
        self.assertFalse(issue)
        self.assertEqual(review, [])

    def test_currency_rule_flags_general_nonzero_integer(self) -> None:
        ws = FakeWorksheet({})
        table = single_column_table("TotalValue", [14, 0])

        issue, _, review = check_total_value(ws, table, {"TotalValue": "TotalValue"})
        self.assertFalse(issue)
        self.assertEqual(len(review), 1)

    def test_currency_rule_flags_one_decimal_general_numeric(self) -> None:
        ws = FakeWorksheet({})
        table = single_column_table("TotalValue", [10.5, 12.25])

        issue, _, review = check_total_value(ws, table, {"TotalValue": "TotalValue"})
        self.assertFalse(issue)
        self.assertEqual(len(review), 1)

    def test_currency_rule_flags_ambiguous_general_numeric(self) -> None:
        ws = FakeWorksheet({})
        table = single_column_table("TotalValue", [10.505, 12.0])

        issue, _, review = check_total_value(ws, table, {"TotalValue": "TotalValue"})
        self.assertFalse(issue)
        self.assertEqual(len(review), 1)


if __name__ == "__main__":
    unittest.main()
