"""Constants and rubric definitions for COMM2003 HW4 grading."""

from __future__ import annotations

FOR_REVIEW = "FOR REVIEW"
NEEDS_MANUAL_REVIEW = "NEEDS MANUAL REVIEW"

REQUIRED_COLUMNS = [
    "TransactionID",
    "AccountNumber",
    "CustomerID",
    "CustomerFName",
    "CustomerLName",
    "TransactionDate",
    "TransactionType",
    "Merchant",
    "City",
    "State",
    "TotalValue",
    "BalanceAfter",
    "Notes",
]

SCORE_MAX = {
    "1A": 1.0,
    "1B": 4.0,
    "1C": 0.5,
    "1D": 0.5,
    "Q1": 2.0,
    "Q2": 2.0,
    "TOTAL": 10.0,
}

TRANSACTION_TYPE_ALLOWED = {
    "Fee",
    "Deposit",
    "Withdrawal",
    "Transfer",
    "Payment",
}

TRANSACTION_TYPE_RAW = {
    "FEE",
    "DEP",
    "WIT",
    "TRA",
    "PAY",
}

TRANSACTION_TYPE_ALLOWED_NORMALIZED = {
    "fee",
    "deposit",
    "withdrawal",
    "transfer",
    "payment",
}

# Abbreviations that are clearly not full words and must be replaced.
TRANSACTION_TYPE_ABBREVIATION_NORMALIZED = {
    "dep",
    "wit",
    "tra",
    "pay",
}

WORKSHEET_INTENTS = {
    "all_transactions_customers": [
        "alltransactionsandcustomers",
        "alltransactionsandcustomerdata",
        "alltransactionscustomers",
        "alltransactionscustomer",
    ],
    "report1": [
        "report1",
        "report 1",
        "r1",
        "pivot1",
    ],
    "report2": [
        "report2",
        "report 2",
        "r2",
        "pivot2",
    ],
    "all_transactions_cleaned": [
        "alltransactions",
        "transactions",
        "combineddata",
        "combinedtransactions",
    ],
    "customer_list_raw": [
        "oscorpcustomerlist",
        "customerlist",
        "customers",
        "customer",
    ],
    "jan_june_raw": [
        "oscorpjantojune",
        "jantojune",
        "janjune",
        "h1",
    ],
    "jul_dec_raw": [
        "oscorpjulytodec",
        "julytodec",
        "juldec",
        "h2",
    ],
}

ERROR_DESCRIPTIONS = {
    "COLUMNS": "Required 13-column structure is incorrect (missing and/or extra columns).",
    "DUPLICATES": "Duplicate transaction rows remain in the cleaned dataset.",
    "TRANSACTION_TYPE_VALUES": "TransactionType still contains invalid or raw abbreviation values.",
    "TRANSACTION_ID_FORMAT": "TransactionID values are not consistently formatted as T####.",
    "NAME_FORMAT": "CustomerFName and CustomerLName are not properly split into separate columns.",
    "DATE_FORMAT": "TransactionDate still includes a time component in at least one row.",
    "MERCHANT_COLUMN": "Merchant still contains city/state text instead of merchant-only values.",
    "CITY_STATE_COLUMNS": "City and State columns are not consistently separated and populated for merchant rows.",
    "TOTAL_VALUE": "TotalValue is not consistently stored as numeric currency values.",
    "NOTES_COLUMN": "Notes still contains special characters and/or unresolved ?err-style text.",
    "NULL_VALUES": "Explicit null-style markers remain where cleaned values are expected.",
}
