# Intercompany Elimination & Reconciliation Agent

An AI-powered tool that automates intercompany reconciliation, mismatch detection, and elimination journal entry generation for public company quarterly close consolidations.

Built with Python and the Anthropic Claude API.

## What It Does

Public company consolidation teams spend significant time each quarter reconciling intercompany balances across legal entities. When balances don't match — due to timing differences, FX rate discrepancies, settlement disputes, or posting errors — someone manually investigates each mismatch, determines root cause, and prepares elimination entries.

This agent automates that workflow:

1. **Reads** intercompany trial balance data from multiple legal entities (Excel-based input)
2. **Matches** IC pairs and validates against an authorized IC Agreement Schedule
3. **Reconciles** each pair — netting USD amounts, comparing FX rates, validating account codes
4. **Classifies** mismatches by error type and severity (Clean → Low → Medium → High → Critical)
5. **Generates** elimination journal entries for clean matches; blocks or flags entries requiring review
6. **Analyzes** each finding using the Claude API — root cause analysis, financial statement impact assessment, and resolution recommendations with entity attribution
7. **Produces** a formatted Excel report (3 tabs) and markdown executive summary

## Architecture

The tool separates deterministic logic from AI analysis:

- **Python handles**: parsing, matching, netting, FX rate comparison, account code validation, authorization checks, severity classification, and JE generation
- **Claude handles**: root cause analysis, impact assessment, resolution recommendations, and executive summary narrative

This separation makes the matching logic fast, testable, and reliable, while leveraging AI for the judgment-intensive work that consolidation systems don't automate.

## Demo Data

The included input workbook contains 5 entities across 2 segments with 7 intercompany transaction sets. Five deliberate errors are planted for the agent to detect:

| Error | Type | Detection Method |
|-------|------|-----------------|
| One-sided posting | Entity 2 books IC receivable; Entity 4 has no offsetting payable | Orphan detection — no matching partner transactions |
| FX rate mismatch | Both entities show same USD, but used different exchange rates (transaction date vs month-end) | Rate comparison per ASC 830-10-45 |
| Settlement mismatch | Partial loan repayment: $300K recorded by lender, $250K by borrower | Net difference on partially settled IC balance |
| Account misclassification | Correct dollars posted to wrong IC account code (6200 vs 6100) | Account code validation against expected mappings |
| Unauthorized IC pair | Matched transaction between entities with no authorized IC relationship | Validation against IC Agreement Schedule |

## Entity Structure

| Entity | Name | Currency | Segment |
|--------|------|----------|---------|
| APX-TECH | Apex Technologies Inc. | USD | Enterprise Solutions |
| APX-CS | Apex Cloud Services LLC | USD | Cloud & Infrastructure |
| APX-UK | Apex UK Ltd | GBP | Enterprise Solutions |
| APX-DE | Apex Deutschland GmbH | EUR | Cloud & Infrastructure |
| APX-SS | Apex Shared Services Corp | USD | Corporate |

## Output

**Excel Report** (3 tabs):
- **Recon Summary** — Dashboard view of all IC pairs with match status, severity, and AI commentary
- **Elimination JEs** — Generated journal entries with debit/credit lines, status (Auto/Review/Blocked), and notes
- **Mismatch Detail** — Deep-dive findings with root cause analysis, impact assessment, and resolution recommendations

**Markdown Summary** — Executive-level narrative suitable for CFO review

## ASC References

- **ASC 810-10** — Consolidation: requirement to eliminate all IC balances and transactions
- **ASC 830-10-45** — Foreign Currency: consistent exchange rate application for IC transactions
- **ASC 850-10-50** — Related Party Disclosures: unauthorized IC transaction disclosure requirements

## Setup

```bash
pip install -r requirements.txt
```

Create an `API_Key.txt` file in the project directory with your Anthropic API key. The agent also checks for the `ANTHROPIC_API_KEY` environment variable. Without an API key, the agent uses deterministic fallback analysis.

## Usage

```bash
python ic_elimination_agent.py
```

The agent reads `ic_input_workbook.xlsx` and produces `ic_elimination_report.xlsx` and `ic_elimination_summary.md`.

## File Structure

```
├── ic_elimination_agent.py    # Main orchestrator
├── config.py                  # Thresholds, entity registry, account mappings
├── models.py                  # Data classes (ICTransaction, ICPair, EliminationJE, Finding)
├── input_parser.py            # Excel input reader and validator
├── matching_engine.py         # IC pair matching and reconciliation logic
├── je_generator.py            # Elimination journal entry generator
├── ai_analyzer.py             # Claude API integration for analysis
├── output_writer.py           # Excel and markdown output producer
├── build_input_workbook.py    # Script to regenerate the demo input file
├── ic_input_workbook.xlsx     # Demo input data (5 entities, 7 transaction sets)
├── requirements.txt
└── README.md
```

## Customization

To adapt for a different company, edit `config.py`:
- `ENTITIES` — your legal entity IDs, names, currencies, and segments
- `AUTHORIZED_IC_PAIRS` — your IC agreement schedule
- `EXPECTED_ACCOUNTS` — your chart of accounts mappings
- `SEVERITY_THRESHOLDS` — your materiality levels
- `FX_RATES` — the period's exchange rates

## Author

Joel Stell, CPA MBA
