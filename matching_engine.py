"""
IC Matching Engine.
Groups transactions into IC pairs, reconciles, and classifies errors.
All logic here is deterministic — no AI calls.
"""

from decimal import Decimal
from typing import List, Dict, Tuple
from collections import defaultdict
from models import ICTransaction, ICPair
import config


def match_and_reconcile(
    transactions: List[ICTransaction],
    agreement_schedule: Dict[frozenset, List[str]],
) -> List[ICPair]:

    # Step 1: Group transactions into IC pairs
    grouped = _group_into_pairs(transactions)

    # Step 2-6: Process each pair
    ic_pairs = []
    for (pair_key, txn_type), sides in grouped.items():
        entity_a, entity_b = sorted(pair_key)
        pair = ICPair(
            entity_a_id=entity_a,
            entity_b_id=entity_b,
            transaction_type=txn_type,
            entity_a_transactions=sides.get(entity_a, []),
            entity_b_transactions=sides.get(entity_b, []),
        )

        # Calculate totals
        pair.entity_a_total_usd = _sum_usd(pair.entity_a_transactions)
        pair.entity_b_total_usd = _sum_usd(pair.entity_b_transactions)
        pair.net_difference = abs(pair.entity_a_total_usd - pair.entity_b_total_usd)

        # Store FX rates and account codes for deep validation
        pair.fx_rate_a = _get_primary_fx_rate(pair.entity_a_transactions)
        pair.fx_rate_b = _get_primary_fx_rate(pair.entity_b_transactions)
        pair.account_code_a = _get_primary_expense_account(pair.entity_a_transactions)
        pair.account_code_b = _get_primary_expense_account(pair.entity_b_transactions)

        # Step 2: Check authorization
        _check_authorization(pair, agreement_schedule)

        # Step 3: Check for one-sided postings
        if pair.match_status == "Pending":
            _check_one_sided(pair)

        # Step 4 & 5: Reconcile and deep validate
        if pair.match_status == "Pending":
            _reconcile_pair(pair)

        # Step 6: Assign severity
        _assign_severity(pair)

        ic_pairs.append(pair)

    return ic_pairs


def _group_into_pairs(
    transactions: List[ICTransaction],
) -> Dict[Tuple[frozenset, str], Dict[str, List[ICTransaction]]]:
    """
    Group transactions by (unordered entity pair, transaction type).
    Within each group, separate by which entity booked the transaction.
    """
    grouped = defaultdict(lambda: defaultdict(list))

    for txn in transactions:
        pair_key = frozenset({txn.entity_id, txn.partner_entity_id})
        group_key = (pair_key, txn.transaction_type)
        grouped[group_key][txn.entity_id].append(txn)

    return grouped


def _sum_usd(transactions: List[ICTransaction]) -> Decimal:
    """Compute the IC transaction amount from one entity's perspective.
    Each IC transaction has two legs: a P&L entry and a BS entry.
    (e.g., Revenue + Receivable, or COGS + Payable).
    We want the unique transaction amount, not double-counted legs.

    Strategy: group by reference number to identify unique transactions,
    then take one leg's amount per transaction. If no reference numbers,
    deduplicate by taking only P&L or only BS accounts."""
    if not transactions:
        return Decimal("0")

    # P&L accounts (the "substance" of the IC transaction)
    pnl_accounts = {"4100", "4200", "5100", "6100", "6200", "6300", "3510", "3520"}
    # BS accounts (the settlement mechanism)
    bs_accounts = {"1310", "1320", "1330", "1340", "2310", "2320", "2330", "2340"}

    # Try grouping by reference number first
    by_ref = defaultdict(list)
    for txn in transactions:
        by_ref[txn.reference_number].append(txn)

    total = Decimal("0")
    for ref, ref_txns in by_ref.items():
        # Take the P&L leg if available (it's the economic substance)
        pnl_txns = [t for t in ref_txns if t.account_code in pnl_accounts]
        if pnl_txns:
            total += pnl_txns[0].usd_amount
        else:
            # No P&L leg (e.g., loan — BS only), take one BS leg
            bs_txns = [t for t in ref_txns if t.account_code in bs_accounts]
            if bs_txns:
                total += bs_txns[0].usd_amount
            elif ref_txns:
                total += ref_txns[0].usd_amount

    return total


def _get_primary_fx_rate(transactions: List[ICTransaction]) -> Decimal:
    """Get the FX rate from the first non-USD transaction in the list."""
    for txn in transactions:
        if txn.local_currency != "USD" and txn.fx_rate != Decimal("1.0000"):
            return txn.fx_rate
    if transactions:
        return transactions[0].fx_rate
    return Decimal("0")


def _get_primary_expense_account(transactions: List[ICTransaction]) -> str:
    """Get the expense/COGS/revenue account code (not the receivable/payable)."""
    pnl_accounts = {"4100", "4200", "5100", "6100", "6200", "6300", "3510", "3520"}
    for txn in transactions:
        if txn.account_code in pnl_accounts:
            return txn.account_code
    return transactions[0].account_code if transactions else ""


def _check_authorization(pair: ICPair, schedule: Dict[frozenset, List[str]]):
    """Check if this IC pair and transaction type is authorized."""
    pair_key = frozenset({pair.entity_a_id, pair.entity_b_id})
    authorized_types = schedule.get(pair_key, [])

    if not authorized_types:
        pair.is_authorized = False
        pair.match_status = "Unauthorized"
        pair.error_category = "Unauthorized"
        return

    if pair.transaction_type not in authorized_types:
        pair.is_authorized = False
        pair.match_status = "Unauthorized"
        pair.error_category = "Unauthorized"


def _check_one_sided(pair: ICPair):
    """Check if only one side has booked the transaction."""
    a_has_data = len(pair.entity_a_transactions) > 0
    b_has_data = len(pair.entity_b_transactions) > 0

    if a_has_data and not b_has_data:
        pair.match_status = "OneSided"
        pair.error_category = "OneSided"
        pair.net_difference = pair.entity_a_total_usd
    elif b_has_data and not a_has_data:
        pair.match_status = "OneSided"
        pair.error_category = "OneSided"
        pair.net_difference = pair.entity_b_total_usd


def _reconcile_pair(pair: ICPair):
    """Full reconciliation: net amounts, check FX rates, check account codes."""

    # Check net difference first
    if pair.net_difference > config.MATCHING_TOLERANCE:
        pair.match_status = "Mismatch"
        # Determine error category based on context
        if _has_partial_settlement(pair):
            pair.error_category = "Settlement"
        else:
            pair.error_category = "Amount"
        return

    # USD amounts match — now deep validate

    # Check FX rate consistency (Error #2 pattern)
    if _has_fx_rate_mismatch(pair):
        pair.match_status = "Mismatch"
        pair.error_category = "FX"
        return

    # Check account code consistency (Error #4 pattern)
    if _has_account_mismatch(pair):
        pair.match_status = "Mismatch"
        pair.error_category = "Classification"
        return

    # All checks passed
    pair.match_status = "Matched"
    pair.error_category = "None"


def _has_partial_settlement(pair: ICPair) -> bool:
    """Check if either side has partial settlement status."""
    for txn in pair.entity_a_transactions + pair.entity_b_transactions:
        if txn.settlement_status == "Partial":
            return True
    return False


def _has_fx_rate_mismatch(pair: ICPair) -> bool:
    """Check if the two sides used different FX rates for the same transaction.
    
    Detection strategy:
    1. If both sides are foreign currency, compare rates directly.
    2. If one side is USD and the other foreign, match transactions by reference
       number and check whether the foreign entity's rate matches the expected
       transaction-date rate from config. If they used month-end instead,
       that's a mismatch even though USD amounts appear to tie.
    """
    # Collect FX rates by reference number for each side
    a_rates_by_ref = {}
    b_rates_by_ref = {}

    for txn in pair.entity_a_transactions:
        if txn.local_currency != "USD" and txn.reference_number:
            a_rates_by_ref[txn.reference_number] = (txn.fx_rate, txn.local_currency)
    for txn in pair.entity_b_transactions:
        if txn.local_currency != "USD" and txn.reference_number:
            b_rates_by_ref[txn.reference_number] = (txn.fx_rate, txn.local_currency)

    # Case 1: Both sides have foreign currency rates — compare per reference
    common_refs = set(a_rates_by_ref.keys()) & set(b_rates_by_ref.keys())
    for ref in common_refs:
        rate_a, _ = a_rates_by_ref[ref]
        rate_b, _ = b_rates_by_ref[ref]
        if rate_a != rate_b:
            return True

    # Case 2: One side is USD, other is foreign
    # Check if the foreign side used a rate that differs from the transaction-date rate
    foreign_refs = a_rates_by_ref or b_rates_by_ref
    if foreign_refs and not common_refs:
        for ref, (rate, currency) in foreign_refs.items():
            currency_config = config.FX_RATES.get(currency, {})
            txn_date_rate = currency_config.get("transaction_date")
            month_end_rate = currency_config.get("month_end")
            if txn_date_rate and month_end_rate and txn_date_rate != month_end_rate:
                # If the entity used the month-end rate instead of transaction-date,
                # flag it — the USD side implicitly used the transaction-date rate
                if rate == month_end_rate and rate != txn_date_rate:
                    return True

    return False


def _has_account_mismatch(pair: ICPair) -> bool:
    """Check if the accounts used match expected mappings for the transaction type.
    Each side of an IC pair has a different expected role:
    - The 'charger' books revenue + receivable
    - The 'recipient' books expense/COGS + payable
    We need to check each side against its expected role, not lump all codes together."""
    txn_type = pair.transaction_type
    expected = config.EXPECTED_ACCOUNTS.get(txn_type, {})
    if not expected:
        return False

    # Determine which accounts are acceptable for EITHER side
    all_expected = set(expected.values())

    # Check each entity's accounts
    for txn in pair.entity_a_transactions:
        if txn.account_code not in all_expected:
            return True
    for txn in pair.entity_b_transactions:
        if txn.account_code not in all_expected:
            return True

    return False


def _assign_severity(pair: ICPair):
    """Assign severity based on match status, error category, and dollar amounts."""
    if pair.match_status == "Matched":
        pair.severity = "Clean"
        return

    if pair.match_status == "Unauthorized":
        pair.severity = "Critical"
        return

    diff = pair.net_difference

    if pair.match_status == "OneSided":
        if diff > config.SEVERITY_THRESHOLDS["high_max"]:
            pair.severity = "Critical"
        else:
            pair.severity = "High"
        return

    # Mismatch cases
    if pair.error_category == "Classification":
        # Account misclassification is at least Medium regardless of dollars
        pair.severity = "Medium"
        return

    if pair.error_category == "FX":
        # FX rate mismatch — USD matches but rates don't
        pair.severity = "Medium"
        return

    # Dollar-based severity
    if diff <= config.SEVERITY_THRESHOLDS["low_max"]:
        pair.severity = "Low"
    elif diff <= config.SEVERITY_THRESHOLDS["medium_max"]:
        pair.severity = "Medium"
    elif diff <= config.SEVERITY_THRESHOLDS["high_max"]:
        pair.severity = "High"
    else:
        pair.severity = "Critical"
