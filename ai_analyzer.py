"""
AI Analyzer — Claude API integration.
Generates root cause analysis, impact assessment, and resolution recommendations
for IC reconciliation findings.
"""

import os
import sys
import re
from decimal import Decimal
from typing import List, Tuple
from models import ICPair, Finding
import config

# Attempt to import anthropic; if not installed, AI features degrade gracefully
try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False


SYSTEM_PROMPT = """You are a senior consolidation accountant at a publicly traded technology company 
reviewing intercompany elimination results for the quarterly close. You have 15+ years of experience 
with multi-entity consolidations under US GAAP.

Your role is to analyze intercompany reconciliation findings and provide:
1. ROOT CAUSE ANALYSIS — Why did this mismatch occur? What is the most likely operational explanation?
2. IMPACT ASSESSMENT — What happens to the consolidated financial statements if this is not corrected? 
   Be specific about which line items are affected and by how much.
3. RECOMMENDED RESOLUTION — Specific, actionable steps with entity attribution. Who needs to do what?
4. ASC REFERENCE — The most relevant ASC guidance, with section number.

Be direct, specific, and quantitative. Avoid generic statements. Every finding should reference 
actual dollar amounts, entity names, and account codes from the data provided.

Format your response exactly as:
ROOT CAUSE: [analysis]
IMPACT: [assessment]
RESOLUTION: [recommendations]
ASC_REF: [reference]"""


def analyze_all(ic_pairs: List[ICPair]) -> List[Finding]:
    """Analyze all non-matched IC pairs and return Finding objects."""
    findings = []
    error_pairs = [p for p in ic_pairs if p.match_status != "Matched"]
    client = _get_client()

    for i, pair in enumerate(error_pairs):
        finding_id = f"FIND-{i + 1:03d}"
        finding = _build_finding(finding_id, pair)

        if client:
            ai_result = _call_claude(client, pair, finding)
            finding.root_cause = ai_result.get("root_cause", "AI analysis unavailable.")
            finding.impact_assessment = ai_result.get("impact", "AI analysis unavailable.")
            finding.recommended_resolution = ai_result.get("resolution", "AI analysis unavailable.")
            finding.asc_reference = ai_result.get("asc_ref", "")

            # Validate AI output — zero token cost, catches hallucinated amounts/entities
            warnings = _validate_ai_response(finding, pair)
            finding.ai_validation_warnings = warnings
            finding.ai_validated = True
            if warnings:
                print(f"    ⚠ {finding_id} AI VALIDATION: {'; '.join(warnings)}")
        else:
            _generate_fallback_analysis(finding, pair)
            finding.ai_validated = False

        findings.append(finding)
        status_icon = "✓" if not finding.ai_validation_warnings else "⚠"
        print(f"    {status_icon} Analyzed {finding_id}: {pair.entity_a_id} <-> {pair.entity_b_id} [{pair.error_category}]")

    # Generate executive summary
    if client and findings:
        exec_summary = _generate_executive_summary(client, ic_pairs, findings)
    else:
        exec_summary = _fallback_executive_summary(ic_pairs, findings)

    return findings, exec_summary


def _get_client():
    if not HAS_ANTHROPIC:
        print("    WARNING: anthropic package not installed. Using fallback analysis.")
        return None

    api_key = _load_api_key()
    if not api_key:
        print("    WARNING: No API key found. Using fallback analysis.")
        return None

    return anthropic.Anthropic(api_key=api_key)


def _load_api_key() -> str:
    key_file_names = ["API_Key.txt", "api_key.txt", "API_KEY.txt", "Api_Key.txt"]
    search_dirs = [os.path.dirname(os.path.abspath(__file__)), os.getcwd()]

    for d in search_dirs:
        for name in key_file_names:
            path = os.path.join(d, name)
            if os.path.exists(path):
                with open(path, "r") as f:
                    key = f.read().strip()
                if key:
                    return key

    env_key = os.environ.get("ANTHROPIC_API_KEY", "")
    return env_key


def _build_finding(finding_id: str, pair: ICPair) -> Finding:
    entity_a_data = {
        "entity_id": pair.entity_a_id,
        "entity_name": config.ENTITIES.get(pair.entity_a_id, {}).get("name", ""),
        "total_usd": float(pair.entity_a_total_usd),
        "transaction_count": len(pair.entity_a_transactions),
        "fx_rate": float(pair.fx_rate_a) if pair.fx_rate_a else None,
        "account_code": pair.account_code_a,
        "transactions": [
            {
                "id": t.transaction_id,
                "account": t.account_code,
                "account_name": t.account_name,
                "usd_amount": float(t.usd_amount),
                "local_amount": float(t.local_amount),
                "currency": t.local_currency,
                "fx_rate": float(t.fx_rate),
                "settlement_status": t.settlement_status,
                "explanation": t.explanation,
            }
            for t in pair.entity_a_transactions
        ],
    }

    entity_b_data = {
        "entity_id": pair.entity_b_id,
        "entity_name": config.ENTITIES.get(pair.entity_b_id, {}).get("name", ""),
        "total_usd": float(pair.entity_b_total_usd),
        "transaction_count": len(pair.entity_b_transactions),
        "fx_rate": float(pair.fx_rate_b) if pair.fx_rate_b else None,
        "account_code": pair.account_code_b,
        "transactions": [
            {
                "id": t.transaction_id,
                "account": t.account_code,
                "account_name": t.account_name,
                "usd_amount": float(t.usd_amount),
                "local_amount": float(t.local_amount),
                "currency": t.local_currency,
                "fx_rate": float(t.fx_rate),
                "settlement_status": t.settlement_status,
                "explanation": t.explanation,
            }
            for t in pair.entity_b_transactions
        ],
    }

    priority = 3
    if pair.severity == "Critical":
        priority = 1
    elif pair.severity == "High":
        priority = 1
    elif pair.severity == "Medium":
        priority = 2

    return Finding(
        finding_id=finding_id,
        ic_pair=pair,
        error_category=pair.error_category,
        entity_a_data=entity_a_data,
        entity_b_data=entity_b_data,
        difference_usd=pair.net_difference,
        priority=priority,
    )


def _call_claude(client, pair: ICPair, finding: Finding) -> dict:
    prompt = _build_prompt(pair, finding)

    try:
        response = client.messages.create(
            model=config.AI_MODEL,
            max_tokens=config.AI_MAX_TOKENS,
            temperature=config.AI_TEMPERATURE,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        )
        return _parse_response(response.content[0].text)
    except Exception as e:
        print(f"    API error for {finding.finding_id}: {e}")
        return {}


def _build_prompt(pair: ICPair, finding: Finding) -> str:
    a = finding.entity_a_data
    b = finding.entity_b_data

    a_detail = ""
    for t in a["transactions"]:
        a_detail += (
            f"  - {t['id']}: Acct {t['account']} ({t['account_name']}), "
            f"${t['usd_amount']:,.2f} USD, {t['currency']} {t['local_amount']:,.2f} @ {t['fx_rate']}, "
            f"Status: {t['settlement_status']}\n"
            f"    Explanation: \"{t['explanation']}\"\n"
        )

    b_detail = ""
    if b["transactions"]:
        for t in b["transactions"]:
            b_detail += (
                f"  - {t['id']}: Acct {t['account']} ({t['account_name']}), "
                f"${t['usd_amount']:,.2f} USD, {t['currency']} {t['local_amount']:,.2f} @ {t['fx_rate']}, "
                f"Status: {t['settlement_status']}\n"
                f"    Explanation: \"{t['explanation']}\"\n"
            )
    else:
        b_detail = "  NO TRANSACTIONS BOOKED\n"

    return f"""Analyze this intercompany reconciliation finding for {config.REPORT_PERIOD}:

IC PAIR: {pair.entity_a_id} ({a['entity_name']}) <-> {pair.entity_b_id} ({b['entity_name']})
TRANSACTION TYPE: {pair.transaction_type}
ERROR CATEGORY: {pair.error_category}
MATCH STATUS: {pair.match_status}
SEVERITY: {pair.severity}

ENTITY A ({pair.entity_a_id}) DETAIL:
  Total USD: ${a['total_usd']:,.2f}
  FX Rate Used: {a['fx_rate']}
  Transactions:
{a_detail}
ENTITY B ({pair.entity_b_id}) DETAIL:
  Total USD: ${b['total_usd']:,.2f}
  FX Rate Used: {b['fx_rate']}
  Transactions:
{b_detail}
NET DIFFERENCE: ${float(pair.net_difference):,.2f}

Expected account mappings for {pair.transaction_type}: {config.EXPECTED_ACCOUNTS.get(pair.transaction_type, {})}

Provide your analysis."""


def _parse_response(text: str) -> dict:
    result = {}
    sections = {
        "root_cause": "ROOT CAUSE:",
        "impact": "IMPACT:",
        "resolution": "RESOLUTION:",
        "asc_ref": "ASC_REF:",
    }

    for key, marker in sections.items():
        start = text.find(marker)
        if start == -1:
            continue
        start += len(marker)
        # Find the next section marker
        next_starts = []
        for other_marker in sections.values():
            if other_marker != marker:
                idx = text.find(other_marker, start)
                if idx != -1:
                    next_starts.append(idx)
        end = min(next_starts) if next_starts else len(text)
        result[key] = text[start:end].strip()

    return result


def _validate_ai_response(finding: Finding, pair: ICPair) -> list:
    """Post-processing validation of Claude's output. Zero API cost.
    Checks that the AI response references correct entities, amounts,
    and doesn't contain obviously hallucinated data."""
    warnings = []
    full_text = " ".join([
        finding.root_cause,
        finding.impact_assessment,
        finding.recommended_resolution,
    ])

    if not full_text.strip() or full_text.strip() == "AI analysis unavailable.":
        warnings.append("AI returned empty or unavailable response")
        return warnings

    # CHECK 1: Response should reference at least one of the two entities
    a_id = pair.entity_a_id
    b_id = pair.entity_b_id
    a_name = config.ENTITIES.get(a_id, {}).get("name", "")
    b_name = config.ENTITIES.get(b_id, {}).get("name", "")

    a_mentioned = a_id in full_text or a_name in full_text
    b_mentioned = b_id in full_text or b_name in full_text

    if not a_mentioned and not b_mentioned:
        warnings.append(f"Response doesn't reference either entity ({a_id} or {b_id})")

    # CHECK 2: If there's a net difference, response should cite a dollar amount
    # that's close to the actual difference (not a hallucinated number)
    if pair.net_difference > Decimal("0"):
        diff_float = float(pair.net_difference)
        # Extract all dollar amounts from the response
        dollar_amounts = re.findall(r'\$[\d,]+(?:\.\d{2})?', full_text)
        parsed_amounts = []
        for amt_str in dollar_amounts:
            try:
                cleaned = amt_str.replace("$", "").replace(",", "")
                parsed_amounts.append(float(cleaned))
            except ValueError:
                continue

        # Check if the actual difference appears (within 10% tolerance)
        if parsed_amounts:
            closest = min(parsed_amounts, key=lambda x: abs(x - diff_float))
            if abs(closest - diff_float) / max(diff_float, 1) > 0.10:
                # The AI cited dollar amounts but none are close to the actual difference
                # Only warn if the difference is material enough to expect a mention
                if diff_float > 1000:
                    warnings.append(
                        f"Net difference is ${diff_float:,.2f} but closest amount "
                        f"cited by AI is ${closest:,.2f}"
                    )

    # CHECK 3: Response should not reference entity IDs that aren't in this pair
    all_entity_ids = set(config.ENTITIES.keys())
    other_entities = all_entity_ids - {a_id, b_id}
    for other_id in other_entities:
        # Only flag if the other entity is mentioned as a primary subject,
        # not just in passing (e.g., "unlike the APX-SS relationship...")
        # Simple heuristic: flag if the entity ID appears more than twice
        if full_text.count(other_id) > 2:
            warnings.append(f"Response references unrelated entity {other_id} multiple times")

    # CHECK 4: ASC reference should be a real ASC section
    asc_ref = finding.asc_reference
    if asc_ref:
        valid_asc_prefixes = ["ASC 810", "ASC 830", "ASC 850", "ASC 820", "ASC 840",
                              "ASC 805", "ASC 842", "ASC 606", "ASC 815", "ASC 320"]
        has_valid_ref = any(prefix in asc_ref for prefix in valid_asc_prefixes)
        if not has_valid_ref and "ASC" in asc_ref:
            warnings.append(f"ASC reference may be invalid: '{asc_ref}'")

    # CHECK 5: Verify error category alignment
    category_keywords = {
        "OneSided": ["one-sided", "orphan", "unmatched", "no offsetting", "missing",
                      "not booked", "no corresponding", "one side"],
        "FX": ["exchange rate", "fx", "foreign currency", "rate", "translation",
               "revaluation", "currency"],
        "Settlement": ["settlement", "repayment", "partial", "remaining balance",
                       "wire transfer", "payment"],
        "Classification": ["account", "misclassif", "reclassif", "wrong account",
                           "incorrect account", "account code"],
        "Unauthorized": ["unauthorized", "not authorized", "agreement schedule",
                         "not listed", "approval", "not on the"],
    }

    expected_keywords = category_keywords.get(pair.error_category, [])
    if expected_keywords:
        text_lower = full_text.lower()
        keyword_found = any(kw in text_lower for kw in expected_keywords)
        if not keyword_found:
            warnings.append(
                f"Response doesn't contain expected keywords for "
                f"'{pair.error_category}' error category"
            )

    return warnings


def _generate_fallback_analysis(finding: Finding, pair: ICPair):
    """Generate deterministic analysis when Claude API is unavailable."""
    a_name = config.ENTITIES.get(pair.entity_a_id, {}).get("name", pair.entity_a_id)
    b_name = config.ENTITIES.get(pair.entity_b_id, {}).get("name", pair.entity_b_id)

    if pair.error_category == "OneSided":
        posting_entity = pair.entity_a_id if pair.entity_a_transactions else pair.entity_b_id
        missing_entity = pair.entity_b_id if pair.entity_a_transactions else pair.entity_a_id
        amt = pair.entity_a_total_usd if pair.entity_a_transactions else pair.entity_b_total_usd
        finding.root_cause = (
            f"{posting_entity} booked IC {pair.transaction_type} transactions totaling "
            f"${amt:,.2f} with partner {missing_entity}, but {missing_entity} has no "
            f"offsetting entry. Most likely cause: invoice not yet received or posted by "
            f"{missing_entity} due to period-end timing cutoff."
        )
        finding.impact_assessment = (
            f"If uncorrected, the consolidated balance sheet will overstate IC receivables "
            f"by ${amt:,.2f} on one side with no offsetting payable. IC revenue/expense "
            f"will not net to zero, overstating consolidated revenue by ${amt:,.2f}."
        )
        finding.recommended_resolution = (
            f"1. {missing_entity} controller to confirm receipt of invoice/documentation. "
            f"2. If valid, {missing_entity} must book offsetting entry before close. "
            f"3. If invalid, {posting_entity} must reverse the orphan posting."
        )
        finding.asc_reference = "ASC 810-10-45 (Consolidation — Other Presentation Matters)"

    elif pair.error_category == "FX":
        finding.root_cause = (
            f"Both entities booked the same IC transaction but used different FX rates. "
            f"{pair.entity_a_id} used rate {pair.fx_rate_a}, {pair.entity_b_id} used rate "
            f"{pair.fx_rate_b}. USD amounts appear to match, but the underlying local "
            f"currency amounts diverge. One entity likely used the transaction date rate "
            f"while the other used the month-end rate."
        )
        finding.impact_assessment = (
            f"While the USD elimination appears clean, the FX rate inconsistency will "
            f"generate a spurious FX gain/loss at the next balance sheet revaluation. "
            f"The rate difference will create an unreconciled CTA (cumulative translation "
            f"adjustment) variance."
        )
        finding.recommended_resolution = (
            f"1. Treasury to confirm the correct rate for this transaction. "
            f"2. Entity using the incorrect rate must rebook at the agreed rate. "
            f"3. Establish IC FX rate policy: all entities use transaction date rate."
        )
        finding.asc_reference = "ASC 830-10-45-1 (Foreign Currency — Transaction Gains and Losses)"

    elif pair.error_category == "Settlement":
        finding.root_cause = (
            f"Both entities acknowledge the IC {pair.transaction_type} relationship but "
            f"disagree on the remaining balance. {pair.entity_a_id} shows "
            f"${pair.entity_a_total_usd:,.2f}, {pair.entity_b_id} shows "
            f"${pair.entity_b_total_usd:,.2f}. Difference: ${pair.net_difference:,.2f}. "
            f"Most likely cause: one entity recorded a different partial repayment amount."
        )
        finding.impact_assessment = (
            f"Consolidated IC loan balances will not eliminate cleanly, leaving a net "
            f"${pair.net_difference:,.2f} residual on the consolidated balance sheet. "
            f"This overstates either IC receivables or IC payables."
        )
        finding.recommended_resolution = (
            f"1. Both entities to pull bank statements for all repayment wire transfers. "
            f"2. Reconcile actual cash movements against recorded settlements. "
            f"3. Entity with the incorrect balance must post a correcting entry."
        )
        finding.asc_reference = "ASC 810-10-45 (Consolidation — Other Presentation Matters)"

    elif pair.error_category == "Classification":
        finding.root_cause = (
            f"USD amounts match between {pair.entity_a_id} and {pair.entity_b_id}, but "
            f"the account codes used do not align with expected mappings for "
            f"{pair.transaction_type} transactions. One entity posted to an incorrect "
            f"IC account. Expected accounts: {config.EXPECTED_ACCOUNTS.get(pair.transaction_type, {})}."
        )
        finding.impact_assessment = (
            f"Generating elimination entries against misclassified accounts would cause "
            f"the consolidated P&L or balance sheet to misstate specific line items. "
            f"The total IC amount is correct but the financial statement presentation "
            f"would be wrong."
        )
        finding.recommended_resolution = (
            f"1. Identify which entity used the incorrect account code. "
            f"2. Post a reclassification entry at that entity to move the balance "
            f"to the correct IC account. "
            f"3. Re-run elimination after reclassification is posted."
        )
        finding.asc_reference = "ASC 810-10-45 (Consolidation — Other Presentation Matters)"

    elif pair.error_category == "Unauthorized":
        finding.root_cause = (
            f"IC transactions exist between {pair.entity_a_id} and {pair.entity_b_id} "
            f"for type '{pair.transaction_type}', but this entity pair and/or transaction "
            f"type is not listed on the IC Agreement Schedule. The transaction may have "
            f"been initiated without proper corporate authorization."
        )
        finding.impact_assessment = (
            f"Unauthorized IC transactions present a controls risk and may require "
            f"related party disclosure under ASC 850-10. The transaction amount of "
            f"${max(pair.entity_a_total_usd, pair.entity_b_total_usd):,.2f} cannot be "
            f"eliminated until authorized or reversed."
        )
        finding.recommended_resolution = (
            f"1. Investigate who authorized this transaction at both entities. "
            f"2. If legitimate, obtain retroactive approval and update the IC Agreement Schedule. "
            f"3. If not authorized, both entities must reverse the transaction. "
            f"4. Assess related party disclosure requirements under ASC 850-10."
        )
        finding.asc_reference = "ASC 850-10-50 (Related Party Disclosures)"


def _generate_executive_summary(client, ic_pairs: List[ICPair], findings: List[Finding]) -> str:
    matched = [p for p in ic_pairs if p.match_status == "Matched"]
    mismatched = [p for p in ic_pairs if p.match_status != "Matched"]
    total_matched_usd = sum(p.entity_a_total_usd for p in matched)
    critical = [f for f in findings if f.priority == 1]

    summary_data = (
        f"Period: {config.REPORT_PERIOD}\n"
        f"Total IC pairs analyzed: {len(ic_pairs)}\n"
        f"Clean matches: {len(matched)} (${total_matched_usd:,.2f})\n"
        f"Findings requiring review: {len(findings)}\n"
        f"Critical/High priority: {len(critical)}\n\n"
        f"Findings summary:\n"
    )
    for f in findings:
        summary_data += (
            f"- {f.finding_id}: {f.ic_pair.entity_a_id} <-> {f.ic_pair.entity_b_id}, "
            f"{f.error_category}, ${f.difference_usd:,.2f} diff, Priority {f.priority}\n"
        )

    try:
        response = client.messages.create(
            model=config.AI_MODEL,
            max_tokens=config.AI_MAX_TOKENS,
            temperature=config.AI_TEMPERATURE,
            system=(
                "You are a senior consolidation accountant writing an executive summary "
                "of intercompany elimination results for the CFO. Be concise, lead with "
                "the most important findings, and quantify everything. Use professional "
                "but direct language. Do not use bullet points — write in paragraphs."
            ),
            messages=[{
                "role": "user",
                "content": f"Write an executive summary based on these results:\n\n{summary_data}",
            }],
        )
        return response.content[0].text
    except Exception as e:
        print(f"    Executive summary API error: {e}")
        return _fallback_executive_summary(ic_pairs, findings)


def _fallback_executive_summary(ic_pairs: List[ICPair], findings: List[Finding]) -> str:
    matched = [p for p in ic_pairs if p.match_status == "Matched"]
    total_matched_usd = sum(p.entity_a_total_usd for p in matched)
    critical = [f for f in findings if f.priority == 1]
    review = [f for f in findings if f.priority == 2]

    summary = (
        f"# Intercompany Elimination & Reconciliation Report\n"
        f"## Period: {config.REPORT_PERIOD}\n\n"
        f"### Summary Statistics\n"
        f"- Total IC pairs analyzed: {len(ic_pairs)}\n"
        f"- Clean matches: {len(matched)} (${total_matched_usd:,.2f})\n"
        f"- Findings requiring review: {len(findings)}\n"
        f"- Critical/High priority findings: {len(critical)}\n"
        f"- Medium priority findings: {len(review)}\n\n"
    )

    if critical:
        summary += "### Critical Findings\n\n"
        for f in critical:
            summary += (
                f"**{f.finding_id}**: {f.ic_pair.entity_a_id} ↔ {f.ic_pair.entity_b_id} — "
                f"{f.error_category}\n\n"
                f"Root Cause: {f.root_cause}\n\n"
                f"Impact: {f.impact_assessment}\n\n"
                f"Resolution: {f.recommended_resolution}\n\n"
                f"---\n\n"
            )

    if review:
        summary += "### Findings Requiring Review\n\n"
        for f in review:
            summary += (
                f"**{f.finding_id}**: {f.ic_pair.entity_a_id} ↔ {f.ic_pair.entity_b_id} — "
                f"{f.error_category}\n\n"
                f"Root Cause: {f.root_cause}\n\n"
                f"Resolution: {f.recommended_resolution}\n\n"
                f"---\n\n"
            )

    return summary
