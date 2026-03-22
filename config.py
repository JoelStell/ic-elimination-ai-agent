"""
Configuration for the Intercompany Elimination & Reconciliation Agent.

To adapt this for a different company:
  1. Update ENTITIES with your legal entity IDs, names, currencies, and segments.
  2. Update AUTHORIZED_IC_PAIRS to reflect your IC agreement schedule.
  3. Update EXPECTED_ACCOUNTS to match your chart of accounts.
  4. Adjust SEVERITY_THRESHOLDS to your materiality levels.
  5. Update FX_RATES with the period's rates from Treasury.
"""

from decimal import Decimal

# --- Report Period ---
REPORT_PERIOD = "Q1 2025"
REPORT_DATE = "March 31, 2025"

# --- Materiality Thresholds ---
SEVERITY_THRESHOLDS = {
    "low_max": Decimal("10000"),
    "medium_max": Decimal("100000"),
    "high_max": Decimal("500000"),
}

MATCHING_TOLERANCE = Decimal("1.00")  # $1 tolerance for rounding

# --- FX Rates ---
FX_RATES = {
    "GBP": {
        "transaction_date": Decimal("1.2750"),
        "month_end": Decimal("1.2600"),
    },
    "EUR": {
        "transaction_date": Decimal("1.1500"),
        "month_end": Decimal("1.1400"),
    },
    "USD": {
        "transaction_date": Decimal("1.0000"),
        "month_end": Decimal("1.0000"),
    },
}

# --- Entity Registry ---
ENTITIES = {
    "APX-TECH": {
        "name": "Apex Technologies Inc.",
        "currency": "USD",
        "segment": "Enterprise Solutions",
        "country": "US",
    },
    "APX-CS": {
        "name": "Apex Cloud Services LLC",
        "currency": "USD",
        "segment": "Cloud & Infrastructure",
        "country": "US",
    },
    "APX-UK": {
        "name": "Apex UK Ltd",
        "currency": "GBP",
        "segment": "Enterprise Solutions",
        "country": "UK",
    },
    "APX-DE": {
        "name": "Apex Deutschland GmbH",
        "currency": "EUR",
        "segment": "Cloud & Infrastructure",
        "country": "DE",
    },
    "APX-SS": {
        "name": "Apex Shared Services Corp",
        "currency": "USD",
        "segment": "Corporate",
        "country": "US",
    },
}

# --- Authorized IC Pairs (from IC Agreement Schedule) ---
# Keys are frozensets so order doesn't matter: APX-TECH <-> APX-UK == APX-UK <-> APX-TECH
AUTHORIZED_IC_PAIRS = {
    frozenset({"APX-TECH", "APX-UK"}): ["Trade", "Dividend"],
    frozenset({"APX-TECH", "APX-SS"}): ["MgmtFee"],
    frozenset({"APX-CS", "APX-DE"}): ["Trade", "Loan"],
    frozenset({"APX-CS", "APX-SS"}): ["MgmtFee"],
    frozenset({"APX-UK", "APX-SS"}): ["MgmtFee"],
    frozenset({"APX-DE", "APX-SS"}): ["MgmtFee"],
}

# --- Expected Account Codes by Transaction Type ---
# Used to validate that entities booked IC activity to the correct accounts
EXPECTED_ACCOUNTS = {
    "Trade": {
        "revenue": "4100",
        "cogs": "5100",
        "receivable": "1310",
        "payable": "2310",
    },
    "Loan": {
        "receivable": "1320",
        "payable": "2320",
    },
    "MgmtFee": {
        "revenue": "4200",
        "expense": "6100",
        "receivable": "1330",
        "payable": "2330",
    },
    "Dividend": {
        "equity_paid": "3510",
        "equity_received": "3520",
        "receivable": "1340",
        "payable": "2340",
    },
    "Allocation": {
        "expense": "6200",
        "expense_alt": "6300",
    },
}

# --- Chart of Accounts (IC-relevant + context accounts) ---
CHART_OF_ACCOUNTS = {
    "1010": "Cash & Equivalents",
    "1200": "Trade Receivables - Third Party",
    "1310": "IC Receivable - Trade",
    "1320": "IC Receivable - Loan",
    "1330": "IC Receivable - Mgmt Fee",
    "1340": "IC Receivable - Dividends",
    "1500": "Property & Equipment, Net",
    "2010": "Accounts Payable - Third Party",
    "2100": "Accrued Liabilities",
    "2310": "IC Payable - Trade",
    "2320": "IC Payable - Loan",
    "2330": "IC Payable - Mgmt Fee",
    "2340": "IC Payable - Dividends",
    "2500": "Long-Term Debt",
    "3100": "Common Stock",
    "3200": "Retained Earnings",
    "3510": "IC Equity - Dividend Paid",
    "3520": "IC Equity - Dividend Received",
    "4000": "Revenue - Third Party",
    "4100": "IC Revenue - Product Sales",
    "4200": "IC Revenue - Services",
    "5000": "COGS - Third Party",
    "5100": "IC COGS - Product Purchases",
    "6000": "SG&A - Third Party",
    "6100": "IC Expense - Mgmt Fee",
    "6200": "IC Expense - IT Allocation",
    "6300": "IC Expense - Cost Allocation",
    "7000": "Depreciation & Amortization",
    "8000": "Interest Expense",
    "9000": "Income Tax Expense",
}

# --- AI Configuration ---
AI_MODEL = "claude-sonnet-4-20250514"
AI_MAX_TOKENS = 4096
AI_TEMPERATURE = 0.2  # Low temperature for factual analysis

# --- File Paths ---
INPUT_FILE = "ic_input_workbook.xlsx"
OUTPUT_EXCEL = "ic_elimination_report.xlsx"
OUTPUT_MARKDOWN = "ic_elimination_summary.md"
