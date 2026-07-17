"""
Flag engine — generates contextual warnings, errors, and notes for each phase
and provides a Streamlit render helper.

Flags are data objects; no Streamlit imports here so the module stays testable.
Call render_flags() (imported separately) in the UI layer.
"""

import math
import re
from dataclasses import dataclass
from enum import Enum
from typing import List, Optional


class FlagLevel(Enum):
    ERROR = "error"      # blocks or strongly warns  — shown expanded
    WARNING = "warning"  # needs analyst review       — shown expanded
    INFO = "info"        # informational note         — shown collapsed


@dataclass
class Flag:
    level: FlagLevel
    code: str
    title: str
    detail: Optional[str] = None


def _is_null(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    return False


def _get(d: dict, *keys, default=None):
    """Safe nested dict getter."""
    for k in keys:
        if not isinstance(d, dict):
            return default
        d = d.get(k, default)
        if d is None:
            return default
    return d


# ── Phase 1 ──────────────────────────────────────────────────────────────────

def get_phase1_flags(phase1_data: dict, available_years: list) -> List[Flag]:
    flags: List[Flag] = []
    if not phase1_data or not available_years:
        return flags

    financials = phase1_data.get("financials", {}) or {}

    line_items = [
        ("sales_and_services",      "Sales and Services"),
        ("other_revenues",           "Other Revenues"),
        ("cost_of_goods_sold",       "Cost of Goods Sold"),
        ("sales_expenses",           "Sales Expenses"),
        ("administrative_expenses",  "Administrative Expenses"),
        ("other_expenses",           "Other Expenses"),
        ("depreciation_amortization","Depreciation & Amortization"),
        ("interest_expenses",        "Interest Expenses"),
        ("tax",                      "Tax"),
    ]
    # These are legitimately absent if combined into admin
    often_absent = {"sales_expenses", "other_expenses"}

    for key, label in line_items:
        yearly = financials.get(key, {}) or {}
        null_years = [str(y) for y in available_years if _is_null(yearly.get(str(y)))]
        if not null_years:
            continue

        if key in often_absent:
            flags.append(Flag(
                FlagLevel.INFO, f"NULL_{key.upper()}",
                f"{label}: no data for {', '.join(null_years)}",
                "Likely combined into Administrative Expenses — treated as 0. "
                "Verify against source documents."
            ))
        elif key == "sales_and_services":
            flags.append(Flag(
                FlagLevel.ERROR, "MISSING_REVENUE",
                f"Revenue (Sales & Services) is missing for: {', '.join(null_years)}",
                "Cannot compute any valuation metrics without revenue. "
                "Re-run extraction or enter values manually."
            ))
        else:
            flags.append(Flag(
                FlagLevel.WARNING, f"MISSING_{key.upper()}",
                f"{label} is missing for: {', '.join(null_years)}",
                "This line will appear blank in the exported template. "
                "Check source documents or enter manually."
            ))

    # Revenue decline > 40% YoY
    sales = financials.get("sales_and_services", {}) or {}
    years_sorted = sorted(available_years)
    for i in range(1, len(years_sorted)):
        prev_v = sales.get(str(years_sorted[i - 1]))
        curr_v = sales.get(str(years_sorted[i]))
        if prev_v and curr_v and not _is_null(prev_v) and not _is_null(curr_v) and prev_v > 0:
            decline_pct = (prev_v - curr_v) / prev_v * 100
            if decline_pct > 40:
                flags.append(Flag(
                    FlagLevel.WARNING, f"REVENUE_DECLINE_{years_sorted[i]}",
                    f"Revenue fell {decline_pct:.0f}% from {years_sorted[i-1]} "
                    f"→ {years_sorted[i]}",
                    "Large declines reduce peer comparability. "
                    "Verify accuracy and flag the trend in the report."
                ))

    # Negative computed EBITDA
    for year in available_years:
        y = str(year)
        def gv(k):
            v = financials.get(k, {}).get(y)
            return 0 if _is_null(v) else (v or 0)
        total_rev = gv("sales_and_services") + gv("other_revenues")
        if total_rev > 0:
            # EBIT-first waterfall: opex includes D&A as reported;
            # EBITDA = EBIT + D&A add-back.
            ebitda = (total_rev
                      - gv("cost_of_goods_sold")
                      - gv("sales_expenses")
                      - gv("administrative_expenses")
                      - gv("other_expenses")
                      + gv("depreciation_amortization"))
            if ebitda < 0:
                flags.append(Flag(
                    FlagLevel.ERROR, f"NEGATIVE_EBITDA_{year}",
                    f"Computed EBITDA is negative in {year} "
                    f"({ebitda:,.0f} THB)",
                    "Negative-EBITDA companies are excluded from peer sets per policy. "
                    "Review P&L line items — expenses may be overstated."
                ))

    # D&A unusually high (>30% of revenue)
    da_yearly = financials.get("depreciation_amortization", {}) or {}
    for year in available_years:
        y = str(year)
        da_v = da_yearly.get(y)
        sales_v = sales.get(y)
        if (not _is_null(da_v) and not _is_null(sales_v)
                and sales_v and sales_v > 0 and da_v):
            da_pct = da_v / sales_v * 100
            if da_pct > 30:
                flags.append(Flag(
                    FlagLevel.INFO, f"HIGH_DA_{year}",
                    f"D&A is {da_pct:.1f}% of revenue in {year}",
                    "Higher than typical — may reflect recent capex (new building/machinery). "
                    "Confirm against balance sheet additions."
                ))

    # Cross-validation: extracted expenses vs source net profit
    verification = phase1_data.get("_verification") or {}
    source_np = verification.get("source_net_profit") or {}
    source_gp = verification.get("source_gross_profit") or {}

    for year in available_years:
        y = str(year)
        total_rev = (sales.get(y) or 0) + ((financials.get("other_revenues") or {}).get(y) or 0)
        if total_rev <= 0:
            continue

        # Check against source gross profit → validates COGS
        sgp = source_gp.get(y)
        if sgp is not None and not _is_null(sgp):
            cogs_v = (financials.get("cost_of_goods_sold") or {}).get(y) or 0
            expected_cogs = total_rev - sgp
            if expected_cogs > 0:
                cogs_gap_pct = abs(cogs_v - expected_cogs) / total_rev * 100
                if cogs_gap_pct > 2:
                    flags.append(Flag(
                        FlagLevel.WARNING, f"COGS_MISMATCH_{year}",
                        f"COGS mismatch in {year}: extracted {cogs_v:,.0f} vs "
                        f"implied {expected_cogs:,.0f} (from source Gross Profit {sgp:,.0f})",
                        "COGS may include items that belong in Operating Expenses, "
                        "or Gross Profit was extracted as COGS. Review the P&L mapping."
                    ))

        # Check against source net profit → validates total expenses
        snp = source_np.get(y)
        if snp is not None and not _is_null(snp):
            implied_expenses = total_rev - snp
            extracted_expenses = sum(
                (financials.get(k) or {}).get(y) or 0
                for k in [
                    "cost_of_goods_sold", "sales_expenses", "administrative_expenses",
                    "other_expenses", "depreciation_amortization", "interest_expenses", "tax",
                ]
            )
            gap = extracted_expenses - implied_expenses
            gap_pct = abs(gap) / total_rev * 100

            # Detect EBIT-first P&L format: gap matches D&A exactly
            da_val = (financials.get("depreciation_amortization") or {}).get(y) or 0
            gap_is_da = da_val > 0 and abs(gap - da_val) / total_rev < 0.005

            if gap_is_da:
                flags.append(Flag(
                    FlagLevel.INFO, f"EBIT_FIRST_FORMAT_{year}",
                    f"Source P&L uses EBIT-first format in {year} — D&A ({da_val:,.0f}) is an add-back",
                    f"D&A is not in the Revenue→Net Profit expense chain in this source. "
                    f"The {abs(gap):,.0f} THB gap equals D&A exactly. Numbers are correct — "
                    f"no double-counting."
                ))
            elif gap_pct > 2:
                direction = "exceed" if gap > 0 else "fall short of"
                flags.append(Flag(
                    FlagLevel.ERROR, f"EXPENSE_CROSSCHECK_{year}",
                    f"Expenses {direction} source Net Profit by {abs(gap):,.0f} THB in {year} "
                    f"({gap_pct:.1f}% of revenue)",
                    f"Source states Net Profit = {snp:,.0f}. "
                    f"Implied total expenses = {implied_expenses:,.0f}, but extracted = {extracted_expenses:,.0f}. "
                    "This typically means line items are double-counted or misclassified between COGS and Admin. "
                    "Edit the P&L table to correct before proceeding."
                ))

    # Verification notes from AI or post-processing
    v_notes = verification.get("notes")
    if v_notes and "AUTO-DETECTED" in str(v_notes):
        flags.append(Flag(
            FlagLevel.WARNING, "EXTRACTION_QUALITY_NOTES",
            "Extraction quality issues detected",
            v_notes,
        ))

    # Missing key business model fields
    sec_a = _get(phase1_data, "business_model", "section_a") or {}
    for field, label in [
        ("company_name",      "Company Name"),
        ("registered_capital","Registered Capital"),
        ("shareholders",      "Shareholders"),
    ]:
        if not sec_a.get(field):
            flags.append(Flag(
                FlagLevel.WARNING, f"MISSING_BM_{field.upper()}",
                f"Business Model — Section A: {label} is missing",
                "Will appear blank in the exported Business Model tab."
            ))

    return flags


# ── Phase 2 ──────────────────────────────────────────────────────────────────

def get_phase2_flags(phase1_data: dict, selected_peers: list) -> List[Flag]:
    flags: List[Flag] = []
    if not selected_peers:
        return flags

    key_clients_text = _get(phase1_data, "business_model", "section_b", "key_clients") or ""

    for peer in selected_peers:
        ticker  = peer.get("identifier", "?")
        company = peer.get("company_name", "") or ""
        ring    = peer.get("ring") or 0

        # Weak sub-sector fit
        if ring >= 3:
            flags.append(Flag(
                FlagLevel.WARNING, f"WEAK_FIT_{ticker}",
                f"{ticker} ({company}): Ring {ring} — weak sub-sector match",
                peer.get("ring_justification")
                or "Consider replacing with a closer comp or adding a caveat."
            ))

        # Scale mismatch
        if peer.get("scale_warning"):
            flags.append(Flag(
                FlagLevel.WARNING, f"SCALE_{ticker}",
                f"{ticker} ({company}): {peer['scale_warning']}",
                "Apply a size discount or note the mismatch explicitly in the report."
            ))

        # Peer appears to be a customer of the target
        if company and key_clients_text:
            words = [w for w in re.split(r"\W+", company.lower()) if len(w) > 4]
            if any(w in key_clients_text.lower() for w in words):
                flags.append(Flag(
                    FlagLevel.WARNING, f"CLIENT_AS_COMP_{ticker}",
                    f"{ticker} ({company}) appears to be a direct customer of the target",
                    "Downstream customers trade at different multiples (branded finished goods "
                    "vs unbranded raw materials). Note this caveat prominently."
                ))

    return flags


# ── Phase 3 ──────────────────────────────────────────────────────────────────

def _norm_ticker(ticker: str) -> str:
    m = re.match(r"^([A-Za-z0-9]+?)([a-z])(\.[A-Z]+)$", ticker)
    return (m.group(1) + m.group(3)).upper() if m else ticker.upper()


_CORP_SUFFIX_RE = re.compile(
    r'\b(public\s+company\s+limited|joint\s+stock\s+company|berhad|bhd|pcl'
    r'|co\.?,?\s*ltd\.?|limited|ltd\.?|inc\.?|corp\.?|plc\.?|s\.?a\.?|n\.?v\.?'
    r'|tbk\.?|pte\.?|jsc|pt)\b',
    re.IGNORECASE,
)


def _norm_company_name(name: str) -> str:
    """Strip legal suffixes and punctuation for fuzzy matching."""
    n = _CORP_SUFFIX_RE.sub('', name.lower())
    return ' '.join(re.sub(r'[^a-z0-9\s]', ' ', n).split())


def get_phase3_flags(
    selected_peers: list,
    lseg_parsed_peers: list,
    deep_dive: dict,
) -> List[Flag]:
    flags: List[Flag] = []

    # Build normalized LSEG lookup — match by LSEG identifier, normalized ticker,
    # filename-derived ticker, AND company name (from file content + filename)
    lseg_lookup: dict = {}
    for p in (lseg_parsed_peers or []):
        raw = (p.get("identifier") or "").upper()
        if raw:
            lseg_lookup[raw] = p
        norm = _norm_ticker(raw)
        if norm and norm != raw:
            lseg_lookup[norm] = p
        # Index by filename-derived ticker (e.g., "D.BK" from "D.BK.xlsx")
        fn_ticker = (p.get("filename_ticker") or "").upper()
        if fn_ticker and fn_ticker not in lseg_lookup:
            lseg_lookup[fn_ticker] = p
        # Index by company name from inside the LSEG file
        cname = (p.get("company_name") or "").strip().lower()
        if cname:
            lseg_lookup[f"__name__{cname}"] = p
        norm_cname = _norm_company_name(p.get("company_name") or "")
        if norm_cname:
            lseg_lookup[f"__name__{norm_cname}"] = p
        # Index by filename as company name (e.g., "Wilmar International Ltd.xlsx" → "wilmar international")
        fn_raw = p.get("filename_raw") or p.get("filename_ticker") or ""
        if fn_raw:
            # Only use as company name if it looks like a name (has spaces), not a ticker
            if ' ' in fn_raw or len(fn_raw) > 10:
                fn_name_norm = _norm_company_name(fn_raw)
                if fn_name_norm and f"__name__{fn_name_norm}" not in lseg_lookup:
                    lseg_lookup[f"__name__{fn_name_norm}"] = p

    # Collect all normalized LSEG names for fuzzy fallback
    _lseg_name_entries = []
    for p in (lseg_parsed_peers or []):
        for src in [p.get("company_name", ""), p.get("filename_raw", "")]:
            norm = _norm_company_name(src)
            if norm:
                words = set(norm.split())
                _lseg_name_entries.append((words, norm, p))

    def _fuzzy_name_match(peer_company):
        """Last-resort: match if peer and LSEG share a distinctive word (3+ chars)."""
        peer_norm = _norm_company_name(peer_company)
        if not peer_norm:
            return None
        peer_words = {w for w in peer_norm.split() if len(w) >= 3}
        best, best_score = None, 0
        for lseg_words, _, lseg_p in _lseg_name_entries:
            sig_words = {w for w in lseg_words if len(w) >= 3}
            overlap = peer_words & sig_words
            if overlap and len(overlap) >= 1:
                # Score by overlap ratio relative to the smaller set
                score = len(overlap) / min(len(peer_words), len(sig_words))
                if score > best_score:
                    best_score = score
                    best = lseg_p
        # Require at least one shared significant word
        return best if best_score > 0 else None

    for peer in (selected_peers or []):
        ticker  = (peer.get("identifier") or "").upper()
        company = peer.get("company_name", "") or ""
        # Try matching: exact ticker → normalized ticker → exact company name → normalized company name → fuzzy
        lseg = (
            lseg_lookup.get(ticker)
            or lseg_lookup.get(_norm_ticker(ticker))
            or lseg_lookup.get(f"__name__{company.strip().lower()}")
            or lseg_lookup.get(f"__name__{_norm_company_name(company)}")
            or _fuzzy_name_match(company)
            or {}
        )

        # No LSEG file at all
        if not lseg:
            if lseg_parsed_peers:
                # Files were uploaded but this ticker wasn't among them
                flags.append(Flag(
                    FlagLevel.WARNING, f"NO_LSEG_{ticker}",
                    f"{ticker} ({company}): no matching LSEG file found",
                    f"Uploaded LSEG files: "
                    f"{', '.join(p.get('identifier','?') for p in lseg_parsed_peers)}. "
                    "Upload the LSEG Valuation .xlsx for this peer."
                ))
            continue

        # Ticker matched via normalization — inform the user
        raw_id = lseg.get("identifier", "")
        if raw_id.upper() != ticker:
            flags.append(Flag(
                FlagLevel.INFO, f"TICKER_NORM_{ticker}",
                f'{ticker}: matched LSEG file as "{raw_id}"',
                "LSEG uses sub-board suffixes (e.g. NTSCm.BK for mai, XOm.BK). "
                "Handled automatically — no action needed."
            ))

        # All-null multiples for recent years
        for metric, label in [
            ("ev_ebitda", "EV/EBITDA"),
            ("pe",        "P/E"),
            ("ev_revenue","EV/Revenue"),
        ]:
            yearly = lseg.get(metric, {}) or {}
            null_recent = [y for y in ("2023", "2024", "2025") if _is_null(yearly.get(y))]
            if len(null_recent) == 3:
                flags.append(Flag(
                    FlagLevel.WARNING, f"NULL_{metric.upper()}_{ticker}",
                    f"{ticker}: {label} is null for 2023–2025",
                    "Row label in the LSEG file may not match the parser's expected format, "
                    "or the company had no coverage in those years."
                ))

        # High multiples
        for metric, label, threshold in [
            ("ev_ebitda", "EV/EBITDA",  15.0),
            ("pe",        "P/E",        30.0),
            ("ev_revenue","EV/Revenue",  3.0),
        ]:
            yearly = lseg.get(metric, {}) or {}
            high_years = [
                y for y in ("2021", "2022", "2023", "2024", "2025")
                if (v := yearly.get(y)) is not None and not _is_null(v) and v > threshold
            ]
            if high_years:
                sample_v = yearly[high_years[0]]
                flags.append(Flag(
                    FlagLevel.WARNING, f"HIGH_{metric.upper()}_{ticker}",
                    f"{ticker}: {label} exceeds {threshold:.0f}x in "
                    f"{', '.join(high_years)} "
                    f"(e.g. {sample_v:.1f}x in {high_years[0]})",
                    "Premium likely reflects brand value or growth outlook not present in the target. "
                    "Apply a discount or note explicitly."
                ))

    # AI-estimated deep dive financials (round numbers signal)
    if deep_dive:
        round_count = sum(
            1
            for row in (deep_dive.get("financials_comparison") or [])
            for field in ("revenue_2023", "revenue_2024", "revenue_2025")
            if (v := row.get(field)) and not _is_null(v) and v == round(v, -2)
        )
        if round_count >= 3:
            flags.append(Flag(
                FlagLevel.INFO, "AI_ESTIMATED_FINANCIALS",
                "Peer revenue/EBITDA in the Comparison tab are AI estimates",
                "Not sourced from actual filings — treat as indicative only. "
                "Verify key numbers against LSEG data or public annual reports before sharing."
            ))

    # No transactions selected
    return flags


# ── Streamlit render helper ───────────────────────────────────────────────────

def render_flags(flags: List[Flag], phase_label: str = "") -> None:
    """
    Render a flag summary expander using Streamlit callouts.
    Auto-expanded when errors or warnings exist; collapsed for info-only.
    Import this in UI modules (not at module level to keep flag_engine testable).
    """
    import streamlit as st  # local import — keeps module importable without Streamlit

    if not flags:
        return

    errors   = [f for f in flags if f.level == FlagLevel.ERROR]
    warnings = [f for f in flags if f.level == FlagLevel.WARNING]
    infos    = [f for f in flags if f.level == FlagLevel.INFO]

    parts = []
    if errors:
        parts.append(f"🔴 {len(errors)} error{'s' if len(errors) > 1 else ''}")
    if warnings:
        parts.append(f"🟡 {len(warnings)} warning{'s' if len(warnings) > 1 else ''}")
    if infos:
        parts.append(f"🔵 {len(infos)} note{'s' if len(infos) > 1 else ''}")

    label = ("Flags" + (f" — {phase_label}" if phase_label else "")) + " · " + " · ".join(parts)
    should_expand = bool(errors or warnings)

    with st.expander(label, expanded=should_expand):
        for f in errors:
            body = f"**{f.title}**"
            if f.detail:
                body += f"\n\n{f.detail}"
            st.error(body)
        for f in warnings:
            body = f"**{f.title}**"
            if f.detail:
                body += f"\n\n{f.detail}"
            st.warning(body)
        for f in infos:
            body = f"**{f.title}**"
            if f.detail:
                body += f"\n\n{f.detail}"
            st.info(body)
