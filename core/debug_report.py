"""
Debug report generator for MAX Comparable Valuation App.

Produces a single JSON blob capturing every piece of data that flows through
the pipeline — raw AI outputs, computed metrics, and a cell-by-cell preview
of what would be written to the Excel template.  Intended for download and
sharing with Claude for troubleshooting / optimization.
"""

import json
import math
import re
from datetime import datetime



# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize_ticker(ticker: str) -> str:
    """Strip LSEG sub-board suffixes: NTSCm.BK -> NTSC.BK."""
    if not ticker:
        return ticker
    m = re.match(r'^([A-Za-z0-9]+?)([a-z])(\.[A-Z]+)$', ticker)
    if m:
        return (m.group(1) + m.group(3)).upper()
    return ticker.upper()


_CORP_SUFFIX_RE = re.compile(
    r'\b(public\s+company\s+limited|berhad|bhd|pcl|co\.?,?\s*ltd\.?|limited|ltd\.?|inc\.?'
    r'|corp\.?|plc\.?|s\.?a\.?|n\.?v\.?|tbk\.?|pte\.?)\b',
    re.IGNORECASE,
)


def _normalize_company_name(name: str) -> str:
    n = _CORP_SUFFIX_RE.sub('', name.lower())
    return ' '.join(re.sub(r'[^a-z0-9\s]', ' ', n).split())


def _build_lseg_lookup(lseg_parsed_peers: list) -> dict:
    lookup = {}
    for p in lseg_parsed_peers:
        original = p.get("identifier") or ""
        raw_upper = original.upper()
        if raw_upper:
            lookup[raw_upper] = p
        normalised = _normalize_ticker(original)
        if normalised and normalised != raw_upper:
            lookup[normalised] = p
        cname = (p.get("company_name") or "").strip().lower()
        if cname:
            lookup[f"__name__{cname}"] = p
        norm_cname = _normalize_company_name(p.get("company_name") or "")
        if norm_cname and f"__name__{norm_cname}" not in lookup:
            lookup[f"__name__{norm_cname}"] = p
    return lookup


def _lseg_peer(lseg_by_ticker: dict, ticker: str, company_name: str) -> dict:
    t = ticker or ""
    return (
        lseg_by_ticker.get(t.upper())
        or lseg_by_ticker.get(_normalize_ticker(t))
        or lseg_by_ticker.get(f"__name__{(company_name or '').strip().lower()}")
        or lseg_by_ticker.get(f"__name__{_normalize_company_name(company_name or '')}")
        or {}
    )


def _safe(v):
    """Return v; convert NaN to None and non-serialisable types to strings."""
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, (int, float, str, bool)):
        return v
    return str(v)


def _null_fields(obj, path=""):
    """Recursively collect paths whose value is None."""
    nulls = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            child = f"{path}.{k}" if path else k
            nulls.extend(_null_fields(v, child))
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            nulls.extend(_null_fields(v, f"{path}[{i}]"))
    elif obj is None:
        nulls.append(path)
    return nulls


# ---------------------------------------------------------------------------
# Cell injection previews  (mirror excel_bridge.py without writing to a file)
# ---------------------------------------------------------------------------

def _preview_phase1(data, available_years):
    cells = {}

    # Preliminary P&L — years D-H since the July 2026 FCF revision
    pnl = {}
    pnl["C2"] = "Currency : Thai Baht"
    years_str = "-".join(map(str, available_years))
    pnl["C3"] = f"Source : {data.get('client_name', 'Company')} Financial Statements ({years_str})"

    financials = data.get("financials", {})
    # Actual template row map (EBIT-first layout)
    row_map = {
        "sales_and_services": 6,
        "other_revenues": 7,
        "cost_of_goods_sold": 11,
        "sales_expenses": 18,
        "administrative_expenses": 19,
        "other_expenses": 20,
        "depreciation_amortization": 27,
        "interest_expenses": 32,
        "tax": 36,
    }
    year_col = {2021: 'D', 2022: 'E', 2023: 'F', 2024: 'G', 2025: 'H'}
    for year in available_years:
        col = year_col.get(year)
        if not col:
            continue
        for key, row_num in row_map.items():
            val = financials.get(key, {}).get(str(year))
            pnl[f"{col}{row_num}"] = _safe(val)

    # FCF section working-capital / PPE links to the Balance Sheet
    bs_year_col = {2021: 'C', 2022: 'D', 2023: 'E', 2024: 'F', 2025: 'G'}
    for year in available_years:
        pcol = year_col.get(year)
        bcol = bs_year_col.get(year)
        if pcol and bcol:
            pnl[f"{pcol}51"] = f"='3. Balance Sheet'!{bcol}7"
            pnl[f"{pcol}52"] = f"='3. Balance Sheet'!{bcol}18"
            pnl[f"{pcol}56"] = f"='3. Balance Sheet'!{bcol}12"
        prev_bcol = bs_year_col.get(year - 1)
        if pcol and prev_bcol and (year - 1) in available_years:
            pnl[f"{pcol}57"] = f"='3. Balance Sheet'!{prev_bcol}12"

    cells["Preliminary_PnL"] = pnl

    # Balance Sheet (new tab)
    bs_cells = {}
    balance_sheet = data.get("balance_sheet", {}) or {}
    if balance_sheet:
        bs_cells["B3"] = "Source : Audited Financial Statements"
        bs_row_map = {
            "cash_and_equivalents": 6,
            "accounts_receivable": 7,
            "short_term_loans_receivable": 8,
            "inventories": 9,
            "ppe_net": 12,
            "accounts_payable": 18,
            "short_term_loans": 19,
            "other_current_liabilities": 20,
            "long_term_loans": 23,
            "paid_up_capital": 29,
            "retained_earnings": 30,
        }
        for year in available_years:
            col = bs_year_col.get(year)
            if not col:
                continue
            for key, row_num in bs_row_map.items():
                val = balance_sheet.get(key, {}).get(str(year))
                bs_cells[f"{col}{row_num}"] = _safe(val)
    cells["Balance_Sheet"] = bs_cells

    # Business Model
    bm_cells = {}
    bm = data.get("business_model", {})
    bm_cells["A1"] = f"{data.get('client_name', 'Company').upper()} – BUSINESS MODEL OVERVIEW"

    sec_a = bm.get("section_a", {}) or {}
    rows_a = ["company_name", "founded", "type", "registered_capital",
              "shareholders", "location", "focus", "certifications",
              "operating_history", "website"]
    for i, field in enumerate(rows_a):
        bm_cells[f"B{4+i}"] = _safe(sec_a.get(field))

    sec_b = bm.get("section_b", {}) or {}
    rows_b = ["primary_revenue", "key_clients", "pricing_strategy",
              "customer_mix", "client_segmentation", "credit_terms",
              "secondary_revenue"]
    for i, field in enumerate(rows_b):
        bm_cells[f"B{16+i}"] = _safe(sec_b.get(field))

    sec_c = bm.get("section_c", []) or []
    for i, moat in enumerate(sec_c[:6]):
        bm_cells[f"A{25+i}"] = _safe(moat.get("title"))
        bm_cells[f"B{25+i}"] = _safe(moat.get("description"))

    sec_d = bm.get("section_d", {}) or {}
    for cell, field in [
        ("B33", "revenue_fy2022"), ("B34", "revenue_fy2023"),
        ("B35", "revenue_fy2024"), ("B36", "revenue_fy2025"),
        ("B37", "gross_margin"), ("B38", "ebitda_normalization"),
        ("B39", "liabilities"), ("B40", "cash_flow_note"),
    ]:
        bm_cells[cell] = _safe(sec_d.get(field))

    sec_e = bm.get("section_e", {}) or {}
    rows_e = ["seller", "exit_motivation", "asking_price",
              "property", "transaction_type", "next_steps"]
    for i, field in enumerate(rows_e):
        bm_cells[f"B{43+i}"] = _safe(sec_e.get(field))

    sec_f = bm.get("section_f", {}) or {}
    rows_f = ["total_staff", "shift_structure", "inbound_process",
              "facility", "machinery", "logistics_fleet", "it_systems"]
    for i, field in enumerate(rows_f):
        bm_cells[f"B{51+i}"] = _safe(sec_f.get(field))

    cells["Business_Model"] = bm_cells
    return cells


def _preview_phase2(selected_peers, not_selected_peers, rejection_rationales):
    furniture = {}

    for i, peer in enumerate(selected_peers[:7]):
        row = 3 + i
        furniture[f"A{row}"] = _safe(peer.get("identifier"))
        furniture[f"B{row}"] = _safe(peer.get("company_name"))
        furniture[f"C{row}"] = _safe(peer.get("trbc_activity"))
        furniture[f"D{row}"] = _safe(peer.get("country"))
        furniture[f"E{row}"] = _safe(peer.get("business_description"))
        furniture[f"F{row}"] = _safe(peer.get("market_cap_thb_m"))
        furniture[f"G{row}"] = _safe(peer.get("ring_justification") or "Selected as comparable peer")

    for i, peer in enumerate(not_selected_peers):
        row = 11 + i
        ticker = peer.get("identifier")
        furniture[f"A{row}"] = _safe(ticker)
        furniture[f"B{row}"] = _safe(peer.get("company_name"))
        furniture[f"C{row}"] = _safe(peer.get("trbc_activity"))
        furniture[f"D{row}"] = _safe(peer.get("country"))
        furniture[f"E{row}"] = _safe(peer.get("business_description"))
        furniture[f"F{row}"] = _safe(peer.get("market_cap_thb_m"))
        furniture[f"G{row}"] = _safe(rejection_rationales.get(ticker, "Not a strong fit"))

    return {"Comps_Profile": furniture}


def _preview_phase3(deep_dive, lseg_parsed_peers, selected_peers, deal_code, client_name, latest_year):
    cells = {}
    qualitative = {q.get("identifier"): q for q in deep_dive.get("qualitative", [])}
    financials = {f.get("identifier"): f for f in deep_dive.get("financials_comparison", [])}
    lseg_by_ticker = _build_lseg_lookup(lseg_parsed_peers)

    # Comparison tab — qualitative rows 5-10 (cols C-J), financial rows 18-23
    comp = {}
    for i, peer in enumerate(selected_peers[:6]):
        row = 5 + i
        ticker = peer.get("identifier")
        q = qualitative.get(ticker, {})
        comp[f"C{row}"] = _safe(ticker)
        comp[f"D{row}"] = _safe(peer.get("company_name"))
        comp[f"E{row}"] = _safe(q.get("core_business_model"))
        comp[f"F{row}"] = _safe(q.get("product_focus"))
        comp[f"G{row}"] = _safe(q.get("similarity"))
        comp[f"H{row}"] = _safe(q.get("comparison_points"))
        comp[f"I{row}"] = _safe(q.get("differentiation_points"))
        comp[f"J{row}"] = _safe(peer.get("country"))

    # Financial rows 18-23 — FIVE year blocks (Revenue/EBITDA/Margin formula):
    #   2021: D/E/F, 2022: G/H/I, 2023: J/K/L, 2024: M/N/O, 2025: P/Q/R
    year_blocks = [
        ("2021", "D", "E", "F"), ("2022", "G", "H", "I"),
        ("2023", "J", "K", "L"), ("2024", "M", "N", "O"),
        ("2025", "P", "Q", "R"),
    ]
    for i, peer in enumerate(selected_peers[:6]):
        row = 18 + i
        ticker = peer.get("identifier")
        f = financials.get(ticker, {})
        # C has formula linking to qualitative ticker — skip
        for year, rev_c, ebitda_c, margin_c in year_blocks:
            rev = f.get(f"revenue_{year}")
            ebitda = f.get(f"ebitda_{year}")
            comp[f"{rev_c}{row}"] = _safe(rev)
            comp[f"{ebitda_c}{row}"] = _safe(ebitda)
            if rev is not None and ebitda is not None:
                comp[f"{margin_c}{row}"] = f'=IFERROR({ebitda_c}{row}/{rev_c}{row},"")'

    cells["Comparison"] = comp

    # Summary peer multiples table — rows 33-38 (July 2026 template)
    # C=ticker, E-I=EV/EBITDA 2021-2025, J-N=EV/Revenue, O-S=P/E
    summary_peers = {}
    summary_years = ["2021", "2022", "2023", "2024", "2025"]
    metric_start_cols = [("ev_ebitda", 5), ("ev_revenue", 10), ("pe", 15)]
    from openpyxl.utils import get_column_letter
    for i, peer in enumerate(selected_peers[:6]):
        row = 33 + i
        ticker = peer.get("identifier")
        summary_peers[f"C{row}"] = _safe(ticker)
        lseg = _lseg_peer(lseg_by_ticker, ticker, peer.get("company_name"))
        for metric, start_col in metric_start_cols:
            values = lseg.get(metric, {}) or {}
            for y_idx, year in enumerate(summary_years):
                v = values.get(year)
                if v is not None:
                    summary_peers[f"{get_column_letter(start_col + y_idx)}{row}"] = _safe(v)
    cells["Summary_Peer_Multiples"] = summary_peers

    # Appendix Hist Trading Performance tab
    hist = {}
    hist["B2"] = f"Project {deal_code} - {client_name} Comps Tables"

    # Actual template: annual data columns 2021→D, 2022→E, 2023→F, 2024→G, 2025→H
    years_annual_cols = [
        ("2021", "D"), ("2022", "E"), ("2023", "F"), ("2024", "G"), ("2025", "H"),
    ]

    for i, peer in enumerate(selected_peers[:6]):
        ticker = peer.get("identifier")
        lseg = _lseg_peer(lseg_by_ticker, ticker, peer.get("company_name"))

        # EV/EBITDA rows 9-14
        r1 = 9 + i
        hist[f"B{r1}"] = _safe(peer.get("company_name"))
        hist[f"C{r1}"] = _safe(ticker)
        ev_ebitda = lseg.get("ev_ebitda", {}) or {}
        for year, col in years_annual_cols:
            v = ev_ebitda.get(year)
            hist[f"{col}{r1}"] = _safe(v)

        r2 = 20 + i  # P/E rows 20-25
        pe = lseg.get("pe", {}) or {}
        for year, col in years_annual_cols:
            hist[f"{col}{r2}"] = _safe(pe.get(year))

        r3 = 31 + i  # EV/Revenue rows 31-36
        ev_rev = lseg.get("ev_revenue", {}) or {}
        for year, col in years_annual_cols:
            hist[f"{col}{r3}"] = _safe(ev_rev.get(year))

    cells["Appendix_Hist_Trading_Performan"] = hist
    return cells


def _preview_phase35(transactions):
    sheet = {}
    for i, tx in enumerate(transactions[:7]):
        row = 9 + i  # header at row 8, data starts row 9
        sheet[f"B{row}"] = _safe(tx.get("target") or tx.get("Target Full Name"))
        sheet[f"C{row}"] = _safe(tx.get("acquirer") or tx.get("Acquiror Full Name"))
        sheet[f"D{row}"] = _safe(tx.get("date") or tx.get("Date Announced"))
        sheet[f"E{row}"] = _safe(tx.get("region") or tx.get("Target Nation"))
        sheet[f"F{row}"] = _safe(tx.get("relevance") or "")
        sheet[f"G{row}"] = _safe(tx.get("deal_value_usd_m") or tx.get("Deal Value (USD, Millions)"))
        sheet[f"H{row}"] = _safe(tx.get("ev_ebitda") or tx.get("Ratio of Enterprise Value to EBITDA"))
        sheet[f"I{row}"] = _safe(tx.get("notes_and_caveats") or "")
    return {"Precedent_Transactions": sheet}


def _preview_phase4(data, latest_year):
    deal_code = data.get("deal_code", "DF-XXX")
    client_name = data.get("client_name", "Company")
    return {
        "Summary": {
            "B2": f"{deal_code} {client_name} - Comps Tables",
            # II. Sensitivity Analysis — @EV(M) ladder (G15-G18) + rationale (I15-I18)
            "G15": '=IFERROR(ROUND($E$14,-2)-100,"")',
            "G16": '=IFERROR(G15+100,"")',
            "G17": '=IFERROR(G16+100,"")',
            "G18": '=IFERROR(G17+100,"")',
            "I15": "Conservative — below the median implied EV from the discounted EBITDA multiple",
            "I16": "Base case — approximately the median implied EV (EBITDA multiple approach)",
            "I17": "Upside — modest premium for growth prospects or buyer synergies",
            "I18": "Ceiling — full strategic premium scenario",
        }
    }


# ---------------------------------------------------------------------------
# Computed metrics
# ---------------------------------------------------------------------------

def _compute_metrics(financials, available_years):
    metrics = {}
    for year in available_years:
        y = str(year)
        sales = (financials.get("sales_and_services") or {}).get(y) or 0
        other = (financials.get("other_revenues") or {}).get(y) or 0
        cogs  = (financials.get("cost_of_goods_sold") or {}).get(y) or 0
        s_exp = (financials.get("sales_expenses") or {}).get(y) or 0
        admin = (financials.get("administrative_expenses") or {}).get(y) or 0
        o_exp = (financials.get("other_expenses") or {}).get(y) or 0
        da    = (financials.get("depreciation_amortization") or {}).get(y) or 0
        interest = (financials.get("interest_expenses") or {}).get(y) or 0
        tax   = (financials.get("tax") or {}).get(y) or 0

        total_rev = sales + other
        gp        = total_rev - cogs
        opex      = s_exp + admin + o_exp
        # EBIT-first waterfall (matches the template): opex includes D&A as
        # reported; D&A is an add-back to get EBITDA, not an extra expense.
        ebit      = gp - opex
        ebitda    = ebit + da
        pbt       = ebit - interest
        net_profit = pbt - tax

        metrics[y] = {
            "total_revenue": round(total_rev, 2),
            "gross_profit": round(gp, 2),
            "gp_margin_pct": round(gp / total_rev * 100, 2) if total_rev else None,
            "opex": round(opex, 2),
            "ebitda": round(ebitda, 2),
            "ebitda_margin_pct": round(ebitda / total_rev * 100, 2) if total_rev else None,
            "ebit": round(ebit, 2),
            "pbt": round(pbt, 2),
            "net_profit": round(net_profit, 2),
        }
    return metrics


# ---------------------------------------------------------------------------
# Null / missing value audit
# ---------------------------------------------------------------------------

def _audit_nulls(phase1, lseg_parsed_peers, deep_dive, available_years):
    report = {}

    # Phase 1 financials
    fin_nulls = []
    for key in ["sales_and_services", "other_revenues", "cost_of_goods_sold",
                "sales_expenses", "administrative_expenses", "other_expenses",
                "depreciation_amortization", "interest_expenses", "tax"]:
        for y in available_years:
            val = (phase1.get("financials", {}) or {}).get(key, {}).get(str(y))
            if val is None:
                fin_nulls.append(f"financials.{key}.{y}")
    report["phase1_financial_nulls"] = fin_nulls

    # Phase 1 balance sheet
    bs_nulls = []
    for key in ["cash_and_equivalents", "accounts_receivable", "short_term_loans_receivable",
                "inventories", "ppe_net", "accounts_payable", "short_term_loans",
                "other_current_liabilities", "long_term_loans", "paid_up_capital",
                "retained_earnings"]:
        for y in available_years:
            val = (phase1.get("balance_sheet", {}) or {}).get(key, {}).get(str(y))
            if val is None:
                bs_nulls.append(f"balance_sheet.{key}.{y}")
    report["phase1_balance_sheet_nulls"] = bs_nulls

    # Phase 1 business model
    bm_nulls = _null_fields(phase1.get("business_model", {}), "business_model")
    report["phase1_business_model_nulls"] = bm_nulls

    # LSEG multiples
    lseg_nulls = []
    for peer in lseg_parsed_peers:
        ticker = peer.get("identifier", "?")
        if peer.get("error"):
            lseg_nulls.append(f"{ticker}: parse error — {peer['error']}")
            continue
        for metric in ["ev_ebitda", "pe", "ev_revenue"]:
            for year, val in (peer.get(metric) or {}).items():
                if val is None:
                    lseg_nulls.append(f"{ticker}.{metric}.{year}")
    report["lseg_multiples_nulls"] = lseg_nulls

    # Deep dive financials
    dd_nulls = []
    for row in (deep_dive.get("financials_comparison") or []):
        ticker = row.get("identifier", "?")
        for field in ["revenue_2021", "revenue_2022", "revenue_2023",
                      "revenue_2024", "revenue_2025",
                      "ebitda_2021", "ebitda_2022", "ebitda_2023",
                      "ebitda_2024", "ebitda_2025"]:
            if row.get(field) is None:
                dd_nulls.append(f"{ticker}.{field}")
    report["deep_dive_financials_nulls"] = dd_nulls

    return report


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def generate_debug_report(session_state: dict) -> bytes:
    """
    Compile all session data into a debug JSON report and return as UTF-8 bytes.
    `session_state` should be a plain dict snapshot of st.session_state.
    """
    phase1           = session_state.get("phase1_data") or {}
    available_years  = session_state.get("selected_years") or []
    selected_peers   = session_state.get("final_selected_peers") or []
    not_selected     = session_state.get("final_not_selected_peers") or []
    rejections       = session_state.get("rejection_rationales") or {}
    lseg_peers       = session_state.get("lseg_parsed_peers") or []
    deep_dive        = session_state.get("deep_dive") or {}
    transactions     = session_state.get("selected_transactions") or []
    sanity           = session_state.get("sanity_result")
    broad_list       = session_state.get("peer_list") or []

    latest_year = max(available_years) if available_years else 2024
    deal_code   = phase1.get("deal_code", "DF-XXX")
    client_name = phase1.get("client_name", "Company")

    # Cell injection previews
    inj_p1  = _preview_phase1(phase1, available_years)
    inj_p2  = _preview_phase2(selected_peers, not_selected, rejections)
    inj_p3  = _preview_phase3(deep_dive, lseg_peers, selected_peers,
                               deal_code, client_name, latest_year)
    inj_p35 = _preview_phase35(transactions)
    inj_p4  = _preview_phase4(phase1, latest_year)

    all_injections = {}
    for d in [inj_p1, inj_p2, inj_p3, inj_p35, inj_p4]:
        all_injections.update(d)

    computed = _compute_metrics(phase1.get("financials", {}), available_years)
    null_audit = _audit_nulls(phase1, lseg_peers, deep_dive, available_years)

    report = {
        "report_meta": {
            "generated_at": datetime.utcnow().isoformat() + "Z",
            "app": "MAX Comparable Valuation App",
            "deal_code": deal_code,
            "client_name": client_name,
            "selected_years": available_years,
            "latest_year": latest_year,
        },

        # ── Phase 1 ────────────────────────────────────────────────────────
        "phase1_raw_extraction": phase1,
        "phase1_computed_metrics": computed,

        # ── Phase 2 ────────────────────────────────────────────────────────
        "phase2_broad_list": broad_list,
        "phase2_selected_peers": selected_peers,
        "phase2_not_selected_peers": not_selected,
        "phase2_rejection_rationales": rejections,

        # ── Phase 3 ────────────────────────────────────────────────────────
        "phase3_lseg_parsed_peers": lseg_peers,
        "phase3_deep_dive_raw": deep_dive,

        # ── Phase 3.5 ──────────────────────────────────────────────────────
        "phase35_selected_transactions": transactions,

        # ── Sanity check ───────────────────────────────────────────────────
        "phase4_sanity_check": sanity,

        # ── Excel injection preview ────────────────────────────────────────
        # Every sheet -> cell -> value that would be written to the template
        "excel_injection_preview": all_injections,

        # ── Null / missing value audit ─────────────────────────────────────
        "null_value_audit": null_audit,
    }

    def _json_default(obj):
        if isinstance(obj, float) and math.isnan(obj):
            return None
        return str(obj)

    return json.dumps(report, indent=2, ensure_ascii=False, default=_json_default).encode("utf-8")
