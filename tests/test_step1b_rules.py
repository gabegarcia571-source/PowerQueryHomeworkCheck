"""Unit tests for Step 1B rule behavior."""

from __future__ import annotations

from datetime import date, datetime
import unittest
from unittest.mock import patch

from grader.agents.xlsx_agent import (
    TableData,
    check_city_state,
    check_open_date_time,
    check_duplicates,
    evaluate_1b,
    check_notes,
    check_null_values,
    check_total_value,
    check_transaction_type,
    has_unsplit_name_issue,
    resolve_step1b_sheet,
)
from grader.constants import REQUIRED_COLUMNS, STEP1B_REQUIRED_COLUMNS


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
        "OpenDate": datetime(2022, 1, 1, 9, 30, 0),
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
            {"TransactionType": "Transaction"},
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

    def test_city_state_rule_allows_blank_values(self) -> None:
        row = base_clean_row()
        row["City"] = ""
        row["State"] = ""

        issue, _, _ = check_city_state([row], col_map())
        self.assertFalse(issue)

    def test_city_state_rule_flags_unsplit_city_when_state_blank(self) -> None:
        row = base_clean_row()
        row["City"] = "New York, NY"
        row["State"] = ""

        issue, note, _ = check_city_state([row], col_map())
        self.assertTrue(issue)
        self.assertIn("city/state split issues", note)

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

    def test_open_date_rule_skips_when_column_missing(self) -> None:
        table = single_column_table("TransactionDate", [date(2024, 1, 15)])

        issue, note, review = check_open_date_time(FakeWorksheet({}), table)
        self.assertFalse(issue)
        self.assertEqual(note, "")
        self.assertEqual(review, [])

    def test_open_date_rule_flags_date_cells_without_time(self) -> None:
        table = single_column_table("OpenDate", [date(2024, 1, 15), datetime(2024, 1, 16, 0, 0)])

        issue, note, review = check_open_date_time(FakeWorksheet({}), table)
        self.assertTrue(issue)
        self.assertIn("without a time component", note)
        self.assertEqual(review, [])

    def test_open_date_rule_flags_excel_serial_dates_without_time(self) -> None:
        table = single_column_table("OpenDate", [45205, 45206.0])

        issue, _, review = check_open_date_time(FakeWorksheet({}), table)
        self.assertTrue(issue)
        self.assertEqual(review, [])

    def test_open_date_rule_accepts_excel_serial_with_fractional_time(self) -> None:
        table = single_column_table("OpenDate", [45205.5])

        issue, _, _ = check_open_date_time(FakeWorksheet({}), table)
        self.assertFalse(issue)

    def test_open_date_rule_accepts_datetime_with_time_component(self) -> None:
        table = single_column_table("OpenDate", [datetime(2024, 1, 15, 13, 45, 0)])

        issue, _, _ = check_open_date_time(FakeWorksheet({}), table)
        self.assertFalse(issue)

    def test_open_date_rule_accepts_date_alias_column(self) -> None:
        table = single_column_table("Date", [datetime(2024, 1, 15, 8, 15, 0)])

        issue, _, _ = check_open_date_time(FakeWorksheet({}), table)
        self.assertFalse(issue)

    def test_open_date_rule_accepts_datetime_number_format_when_value_lacks_time(self) -> None:
        table = single_column_table("OpenDate", [date(2024, 1, 15), datetime(2024, 1, 16, 0, 0), 45205.0])
        ws = FakeWorksheet({
            (2, 1): "m/d/yyyy h:mm AM/PM",
            (3, 1): "yyyy-mm-dd hh:mm",
            (4, 1): "mm/dd/yyyy hh:mm:ss",
        })

        issue, _, _ = check_open_date_time(ws, table)
        self.assertFalse(issue)

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

    def test_evaluate_1b_accepts_required_headers_with_trailing_chars(self) -> None:
        suffix_headers = [f"{name}x" for name in STEP1B_REQUIRED_COLUMNS]
        row = {
            "TransactionIDx": "T0001",
            "AccountNumberx": "A1001",
            "CustomerIDx": "C2001",
            "CustomerFNamex": "Tony",
            "CustomerLNamex": "Stark",
            "TransactionDatex": datetime(2024, 1, 31, 14, 30, 0),
            "TransactionTypex": "Deposit",
            "Merchantx": "Oscorp Capital",
            "Cityx": "New York",
            "Statex": "NY",
            "TotalValuex": 1200.25,
            "BalanceAfterx": 3000.25,
            "Notesx": "",
            "AccountStatusLevelx": "Gold",
            "Branchx": "NY-001",
            "Teamx": "North",
            "OpenDatex": datetime(2022, 1, 1, 10, 0, 0),
        }
        table = TableData(
            headers=suffix_headers,
            rows=[row],
            row_numbers=[2],
            column_index_by_header={header: idx + 1 for idx, header in enumerate(suffix_headers)},
        )

        with patch("grader.agents.xlsx_agent.extract_table", return_value=table):
            score, errors, _ = evaluate_1b(FakeWorksheet({}))

        self.assertEqual(score, 4.0)
        self.assertEqual(errors, ["No errors found."])

    def test_evaluate_1b_allows_extra_columns_without_penalty(self) -> None:
        headers = list(STEP1B_REQUIRED_COLUMNS) + ["SomeNewColumn"]
        row = base_clean_row()
        row["AccountStatusLevel"] = "Gold"
        row["Branch"] = "NY-001"
        row["Team"] = "North"
        row["SomeNewColumn"] = "accepted"

        table = TableData(
            headers=headers,
            rows=[row],
            row_numbers=[2],
            column_index_by_header={header: idx + 1 for idx, header in enumerate(headers)},
        )

        with patch("grader.agents.xlsx_agent.extract_table", return_value=table):
            score, errors, _ = evaluate_1b(FakeWorksheet({}))

        self.assertEqual(score, 4.0)
        self.assertEqual(errors, ["No errors found."])

    def test_evaluate_1b_accepts_prefixed_required_columns_via_aliases(self) -> None:
        headers = list(REQUIRED_COLUMNS) + [
            "Oscorp-CustomerList.AccountStatusLevel",
            "Oscorp-CustomerList.Branch",
            "Oscorp-CustomerList.Team",
            "Oscorp-CustomerList.OpenDate",
        ]
        row = base_clean_row()
        row["Oscorp-CustomerList.AccountStatusLevel"] = "Gold"
        row["Oscorp-CustomerList.Branch"] = "NY-001"
        row["Oscorp-CustomerList.Team"] = "North"
        row["Oscorp-CustomerList.OpenDate"] = datetime(2022, 1, 1, 9, 30, 0)

        table = TableData(
            headers=headers,
            rows=[row],
            row_numbers=[2],
            column_index_by_header={header: idx + 1 for idx, header in enumerate(headers)},
        )

        with patch("grader.agents.xlsx_agent.extract_table", return_value=table):
            score, errors, _ = evaluate_1b(FakeWorksheet({}))

        self.assertEqual(score, 4.0)
        self.assertEqual(errors, ["No errors found."])

    def test_evaluate_1b_allows_variant_c_reordered_columns(self) -> None:
        headers = [
            "TransactionID",
            "AccountNumber",
            "CustomerID",
            "CustomerFName",
            "CustomerLName",
            "TransactionDate",
            "TransactionType",
            "TotalValue",
            "Merchant",
            "City",
            "State",
            "BalanceAfter",
            "Notes",
            "AccountStatusLevel",
            "Branch",
            "Team",
            "OpenDate",
        ]
        row = base_clean_row()
        row["AccountStatusLevel"] = "Gold"
        row["Branch"] = "NY-001"
        row["Team"] = "North"

        table = TableData(
            headers=headers,
            rows=[row],
            row_numbers=[2],
            column_index_by_header={header: idx + 1 for idx, header in enumerate(headers)},
        )

        with patch("grader.agents.xlsx_agent.extract_table", return_value=table):
            score, errors, _ = evaluate_1b(FakeWorksheet({}))

        self.assertEqual(score, 4.0)
        self.assertEqual(errors, ["No errors found."])

    def test_evaluate_1b_missing_open_date_deducts_in_column_check(self) -> None:
        headers = list(REQUIRED_COLUMNS) + ["AccountStatusLevel", "Branch", "Team"]
        row = base_clean_row()

        table = TableData(
            headers=headers,
            rows=[row],
            row_numbers=[2],
            column_index_by_header={header: idx + 1 for idx, header in enumerate(headers)},
        )

        with patch("grader.agents.xlsx_agent.extract_table", return_value=table):
            score, errors, _ = evaluate_1b(FakeWorksheet({}))

        self.assertEqual(score, 3.5)
        self.assertEqual(len(errors), 1)
        self.assertIn("Missing: OpenDate", errors[0])

    def test_evaluate_1b_missing_required_column_still_deducts(self) -> None:
        headers = [name for name in STEP1B_REQUIRED_COLUMNS if name != "Notes"]
        row = base_clean_row()

        table = TableData(
            headers=headers,
            rows=[row],
            row_numbers=[2],
            column_index_by_header={header: idx + 1 for idx, header in enumerate(headers)},
        )

        with patch("grader.agents.xlsx_agent.extract_table", return_value=table):
            score, errors, _ = evaluate_1b(FakeWorksheet({}))

        self.assertEqual(score, 3.5)
        self.assertEqual(len(errors), 1)
        self.assertIn("Missing: Notes", errors[0])

    def test_evaluate_1b_totalvalue_deducts_once_for_multiple_bad_rows(self) -> None:
        row1 = base_clean_row()
        row1["TransactionDate"] = datetime(2024, 1, 15, 10, 30, 0)
        row1["TotalValue"] = "bad"

        row2 = base_clean_row()
        row2["TransactionID"] = "T0002"
        row2["AccountNumber"] = "A1002"
        row2["CustomerID"] = "C2002"
        row2["TransactionDate"] = datetime(2024, 1, 15, 11, 30, 0)
        row2["TotalValue"] = "worse"

        headers = list(STEP1B_REQUIRED_COLUMNS)
        table = TableData(
            headers=headers,
            rows=[row1, row2],
            row_numbers=[2, 3],
            column_index_by_header={header: idx + 1 for idx, header in enumerate(headers)},
        )

        with patch("grader.agents.xlsx_agent.extract_table", return_value=table):
            score, errors, _ = evaluate_1b(FakeWorksheet({}))

        self.assertEqual(score, 3.5)
        self.assertEqual(len(errors), 1)
        self.assertIn("non-numeric", errors[0])

    def test_resolve_step1b_sheet_prefers_customer_data_identifier(self) -> None:
        class Sheet:
            def __init__(self, title: str) -> None:
                self.title = title

        class Workbook:
            def __init__(self, titles: list[str]) -> None:
                self.worksheets = [Sheet(title) for title in titles]

        workbook = Workbook([
            "AllTransactions",
            "AllTransactionsAndCustomerD (2)",
            "AllTransactionsAndCustomerData",
            "Report 1",
        ])

        target = resolve_step1b_sheet(workbook)
        self.assertIsNotNone(target)
        self.assertEqual(target.title, "AllTransactionsAndCustomerData")

    def test_resolve_step1b_sheet_returns_none_when_identifier_missing(self) -> None:
        class Sheet:
            def __init__(self, title: str) -> None:
                self.title = title

        class Workbook:
            def __init__(self, titles: list[str]) -> None:
                self.worksheets = [Sheet(title) for title in titles]

        workbook = Workbook([
            "AllTransactions",
            "Report 1",
            "Report 2",
        ])

        target = resolve_step1b_sheet(workbook)
        self.assertIsNone(target)


if __name__ == "__main__":
    unittest.main()
