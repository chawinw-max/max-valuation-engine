import math
import re
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter


def _normalize_ticker(ticker: str) -> str:
    """
    Strip LSEG sub-board suffixes so NTSCm.BK -> NTSC.BK, XOm.BK -> XO.BK.
    Also normalises KWNF.KL-style LSEG codes to uppercase.
    """
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
    """Strip legal suffixes and punctuation for fuzzy matching. 'Kawan Food Berhad' -> 'kawan food'"""
    n = _CORP_SUFFIX_RE.sub('', name.lower())
    return ' '.join(re.sub(r'[^a-z0-9\s]', ' ', n).split())


def _build_lseg_lookup(lseg_parsed_peers: list) -> dict:
    """
    Build a ticker -> peer dict with three match strategies:
    1. Exact uppercase ticker (e.g. SAUCE.BK, KWNF.KL)
    2. Normalised ticker stripping board suffix (NTSCm.BK -> NTSC.BK)
    3. Fuzzy company name after stripping legal suffixes (Kawan Food Berhad ~ Kawan Food Bhd)
    """
    lookup = {}
    for p in lseg_parsed_peers:
        original = p.get("identifier") or ""
        raw_upper = original.upper()
        if raw_upper:
            lookup[raw_upper] = p
        normalised = _normalize_ticker(original)  # normalize before uppercasing
        if normalised and normalised != raw_upper:
            lookup[normalised] = p
        cname = (p.get("company_name") or "").strip().lower()
        if cname:
            lookup[f"__name__{cname}"] = p
        norm_cname = _normalize_company_name(p.get("company_name") or "")
        if norm_cname and f"__name__{norm_cname}" not in lookup:
            lookup[f"__name__{norm_cname}"] = p
        # Also index by filename-derived ticker (e.g., "D.BK" from "D.BK.xlsx")
        fn_ticker = (p.get("filename_ticker") or "").upper()
        if fn_ticker and fn_ticker not in lookup:
            lookup[fn_ticker] = p
        # Index by filename as company name (e.g., "Wilmar International Ltd.xlsx")
        fn_raw = p.get("filename_raw") or ""
        if fn_raw and (' ' in fn_raw or len(fn_raw) > 10):
            fn_name_norm = _normalize_company_name(fn_raw)
            if fn_name_norm and f"__name__{fn_name_norm}" not in lookup:
                lookup[f"__name__{fn_name_norm}"] = p
    return lookup


def _fuzzy_name_match(lseg_by_ticker: dict, peer_company: str):
    """Last-resort: match if peer and LSEG share a distinctive word (3+ chars)."""
    peer_norm = _normalize_company_name(peer_company or "")
    if not peer_norm:
        return None
    peer_words = {w for w in peer_norm.split() if len(w) >= 3}
    best, best_score = None, 0
    for key, p in lseg_by_ticker.items():
        if not key.startswith("__name__"):
            continue
        lseg_words = {w for w in key[8:].split() if len(w) >= 3}
        overlap = peer_words & lseg_words
        if overlap:
            score = len(overlap) / min(len(peer_words), len(lseg_words))
            if score > best_score:
                best_score = score
                best = p
    return best if best_score > 0 else None


def _lseg_peer(lseg_by_ticker: dict, ticker: str, company_name: str) -> dict:
    """Look up LSEG data for a peer using ticker then company name fallbacks."""
    t = ticker or ""
    return (
        lseg_by_ticker.get(t.upper())
        or lseg_by_ticker.get(_normalize_ticker(t))
        or lseg_by_ticker.get(f"__name__{(company_name or '').strip().lower()}")
        or lseg_by_ticker.get(f"__name__{_normalize_company_name(company_name or '')}")
        or _fuzzy_name_match(lseg_by_ticker, company_name)
        or {}
    )


def _safe_write(sheet, cell_coord, value):
    """
    Writes a value to a cell ONLY if the cell does not already contain a formula.
    Handles merged cells by writing to the merge's anchor (top-left) cell.
    """
    cell = sheet[cell_coord]

    # If the target is a non-anchor merged cell, find and write to the anchor instead
    if isinstance(cell, MergedCell):
        for merged_range in sheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                break
        else:
            print(f"WARNING: {cell_coord} is a MergedCell with no resolvable anchor on {sheet.title}. Skipping.")
            return False

    if isinstance(cell.value, str) and cell.value.startswith('='):
        print(f"WARNING: Attempted to overwrite formula in {cell_coord} on sheet {sheet.title}. Skipping.")
        return False

    # Treat float NaN as empty — writing NaN to Excel causes #NUM! errors
    if isinstance(value, float) and math.isnan(value):
        return False

    cell.value = value
    return True

def inject_phase1_data(workbook, data, available_years):
    """
    Injects Phase 1 data (Preliminary P&L and Business Model) into the workbook.
    `data` is the JSON parsed dict from Gemini.
    """
    # 1. Preliminary P&L
    if 'Preliminary P&L' in workbook.sheetnames:
        sheet = workbook['Preliminary P&L']
        
        # B2: Currency label
        _safe_write(sheet, 'B2', "Currency : Thai Baht")
        
        # B3: Source description
        years_str = "-".join(map(str, available_years))
        _safe_write(sheet, 'B3', f"{data.get('client_name', 'Company')} – Audited Financial Statements FY{years_str}")
        
        financials = data.get('financials', {})

        # FINAL template row map (V2 / Dental2 layout):
        #   Row 6: Sales, Row 7: Other Rev, Row 8: Total Rev [FORMULA]
        #   Row 11: COGS input, Row 12: Total COGS =SUM(C11) [FORMULA]
        #   Row 14: GP =C8-SUM(C11) [FORMULA], Row 15: GP Margin [FORMULA]
        #   Row 18: Sales Exp, Row 19: Admin Exp, Row 20: CEO Salary (MANUAL)
        #   Row 21: Total OpEx =SUM(C18:C20) [FORMULA]
        #   Row 23: EBITDA =C14-C21 [FORMULA]
        #   Row 26: D&A, Row 27: EBIT =C23-C26 [FORMULA]
        #   Row 30: Interest, Row 31: EBT =C27-C30 [FORMULA]
        #   Row 34: Tax, Row 35: NP =C31-C34 [FORMULA]
        #   Rows 37-41: EBITDA Add-Back section (structure only — CEO values manual)
        row_map = {
            'sales_and_services': 6,
            'other_revenues': 7,
            'cost_of_goods_sold': 11,
            'sales_expenses': 18,
            'administrative_expenses': 19,
            # Row 20 = CEO Salary (Normalised) — left blank by app, filled manually
            # Row 21 = Total OpEx [FORMULA] — _safe_write will skip
            'depreciation_amortization': 26,
            'interest_expenses': 30,
            'tax': 34,
        }

        # Template has FIXED year-to-column mapping (Summary formulas depend on it):
        #   C=2021, D=2022, E=2023, F=2024, G=2025
        # Year labels are already pre-filled — don't overwrite them.
        year_col = {2021: 'C', 2022: 'D', 2023: 'E', 2024: 'F', 2025: 'G'}

        for year in available_years:
            col_letter = year_col.get(year)
            if not col_letter:
                continue
            for key, row_num in row_map.items():
                val = financials.get(key, {}).get(str(year))
                if val is not None:
                    _safe_write(sheet, f'{col_letter}{row_num}', val)

        # --- EBITDA Add-Back section (rows 37-41) — structure only ---
        # Labels
        _safe_write(sheet, 'A37', 'EBITDA Add-Back (M&A Normalisation)')
        _safe_write(sheet, 'H37', 'Normalisation adjustments for M&A valuation purposes')
        _safe_write(sheet, 'B38', '+ CEO Salary Normalisation Add-Back')
        _safe_write(sheet, 'H38', 'Owner-manager salary added back; market-rate management cost assumed separately by buyer')
        # E38-G38: LEFT BLANK — CEO salary is manual entry
        _safe_write(sheet, 'B39', '+ Other Non-recurring / One-time Items')
        _safe_write(sheet, 'H39', 'None identified')
        # Default "other non-recurring" to 0 for years with data
        for year in available_years:
            col_letter = year_col.get(year)
            if col_letter and col_letter in ('E', 'F', 'G'):
                _safe_write(sheet, f'{col_letter}39', 0)
        _safe_write(sheet, 'A40', 'Total EBITDA Add-Back')
        _safe_write(sheet, 'A41', 'Adjusted EBITDA (for Valuation)')
        _safe_write(sheet, 'H41', 'D5 in Summary = J45 (2025 Adj EBITDA)')
        # Formulas for Total Add-Back and Adjusted EBITDA (cols C-G)
        for col_letter in ('C', 'D', 'E', 'F', 'G'):
            _safe_write(sheet, f'{col_letter}40', f'=SUM({col_letter}38:{col_letter}39)')
            _safe_write(sheet, f'{col_letter}41', f'={col_letter}23+{col_letter}40')

    # 2. Business Model (V2 template — 7-section layout)
    if 'Business Model' in workbook.sheetnames:
        sheet = workbook['Business Model']
        bm = data.get('business_model', {})
        deal_code = data.get('deal_code', 'DF-XXX')
        client_name_bm = data.get('client_name', 'Company')

        # Title — V2 format: 'BUSINESS MODEL  |  [Company Name]  |  [Deal Code]'
        _safe_write(sheet, 'A1', f"BUSINESS MODEL  |  {client_name_bm}  |  {deal_code}")

        # Section 1: BUSINESS OVERVIEW (A3, rows 4-7)
        sec_1 = bm.get('section_1', bm.get('section_a', {})) or {}
        _safe_write(sheet, 'A3', '1. BUSINESS OVERVIEW')
        overview_fields = ['company_name', 'founder', 'branches', 'branch_history',
                           'key_strengths']
        for i, field in enumerate(overview_fields[:4]):
            _safe_write(sheet, f'B{4+i}', sec_1.get(field))
            # Also write corresponding right-side entries (cols D-E)
        # Backward compat: also try old section_a fields
        if not sec_1.get('company_name') and bm.get('section_a', {}).get('company_name'):
            sec_a = bm['section_a']
            _safe_write(sheet, 'B4', sec_a.get('company_name'))
            _safe_write(sheet, 'B5', sec_a.get('founded'))
            _safe_write(sheet, 'B6', sec_a.get('location'))
            _safe_write(sheet, 'B7', sec_a.get('focus'))

        # Section 2: FINANCIAL SNAPSHOT (A9, rows 10-13)
        sec_2 = bm.get('section_2', bm.get('section_d', {})) or {}
        _safe_write(sheet, 'A9', '2. FINANCIAL SNAPSHOT')
        fin_fields = ['current_revenue', 'peak_revenue', 'revenue_target',
                       'declared_revenue', 'net_profit_margin', 'payment_mix',
                       'bookkeeping']
        for i, field in enumerate(fin_fields[:4]):
            _safe_write(sheet, f'B{10+i}', sec_2.get(field))
        # Backward compat: old section_d
        if not sec_2.get('current_revenue') and bm.get('section_d'):
            sec_d = bm['section_d']
            _safe_write(sheet, 'B10', sec_d.get('revenue_fy2024') or sec_d.get('revenue_fy2025'))
            _safe_write(sheet, 'B11', sec_d.get('gross_margin'))
            _safe_write(sheet, 'B12', sec_d.get('ebitda_normalization'))
            _safe_write(sheet, 'B13', sec_d.get('liabilities'))

        # Section 3: COST STRUCTURE (A15, rows 16-22)
        sec_3 = bm.get('section_3', {}) or {}
        _safe_write(sheet, 'A15', '3. COST STRUCTURE (Monthly Estimates)')
        cost_items = sec_3.get('items', [])
        for i, item in enumerate(cost_items[:7]):
            _safe_write(sheet, f'B{16+i}', item.get('item'))
            _safe_write(sheet, f'C{16+i}', item.get('amount'))
            _safe_write(sheet, f'D{16+i}', item.get('notes'))

        # Section 4: SERVICES & MARKETING (A24, rows 25-28)
        sec_4 = bm.get('section_4', bm.get('section_b', {})) or {}
        _safe_write(sheet, 'A24', '4. SERVICES & MARKETING')
        svc_fields = ['core_service', 'other_services', 'marketing_channel',
                       'competitive_position', 'risk']
        for i, field in enumerate(svc_fields[:4]):
            _safe_write(sheet, f'B{25+i}', sec_4.get(field))
        # Backward compat
        if not sec_4.get('core_service') and bm.get('section_b'):
            sec_b = bm['section_b']
            _safe_write(sheet, 'B25', sec_b.get('primary_revenue'))
            _safe_write(sheet, 'B26', sec_b.get('pricing_strategy'))
            _safe_write(sheet, 'B27', sec_b.get('customer_mix'))
            _safe_write(sheet, 'B28', sec_b.get('key_clients'))

        # Section 5: OPERATIONAL & COMPLIANCE NOTES (A30, rows 31-33)
        sec_5 = bm.get('section_5', {}) or {}
        _safe_write(sheet, 'A30', '5. OPERATIONAL & COMPLIANCE NOTES')
        ops_fields = ['tax_filing_gap', 'management_issues', 'systems_records']
        for i, field in enumerate(ops_fields):
            _safe_write(sheet, f'B{31+i}', sec_5.get(field))
        # Backward compat with old section_f
        if not sec_5.get('tax_filing_gap') and bm.get('section_f'):
            sec_f = bm['section_f']
            _safe_write(sheet, 'B31', sec_f.get('facility'))
            _safe_write(sheet, 'B32', sec_f.get('it_systems'))
            _safe_write(sheet, 'B33', sec_f.get('total_staff'))

        # Section 6: M&A CONSIDERATIONS (A35, rows 36-40)
        sec_6 = bm.get('section_6', bm.get('section_e', {})) or {}
        _safe_write(sheet, 'A35', '6. M&A CONSIDERATIONS')
        ma_fields = ['reason_for_sale', 'transaction_type', 'asking_price',
                      'implied_ev_ebitda', 'earn_out', 'growth_angle',
                      'key_risks', 'key_attractions', 'recommended_next_steps',
                      'valuation_view']
        for i, field in enumerate(ma_fields[:5]):
            _safe_write(sheet, f'B{36+i}', sec_6.get(field))
        # Backward compat
        if not sec_6.get('reason_for_sale') and bm.get('section_e'):
            sec_e = bm['section_e']
            _safe_write(sheet, 'B36', sec_e.get('exit_motivation'))
            _safe_write(sheet, 'B37', sec_e.get('transaction_type'))
            _safe_write(sheet, 'B38', sec_e.get('asking_price'))
            _safe_write(sheet, 'B39', sec_e.get('next_steps'))

        # Section 7: BRANCH ASSET INVENTORY (A42, rows 43-54)
        sec_7 = bm.get('section_7', {}) or {}
        _safe_write(sheet, 'A42', '7. BRANCH ASSET INVENTORY')
        assets = sec_7.get('assets', [])
        for i, asset in enumerate(assets[:12]):
            _safe_write(sheet, f'B{43+i}', asset.get('name'))
            _safe_write(sheet, f'C{43+i}', asset.get('value'))
            _safe_write(sheet, f'D{43+i}', asset.get('notes'))

        # Footer
        _safe_write(sheet, 'A56',
                     f"* Source: Deal notes prepared by Max Solutions. "
                     f"All figures approximate and subject to verification.")

def inject_phase2_data(workbook, selected_peers, not_selected_peers, rejection_rationales):
    """
    Furniture tab: selected peers in rows 3–8 (and row 9 if 7 selected),
    not-selected peers in rows 11+. Columns A–G.
    """
    if 'Furniture' not in workbook.sheetnames:
        return
    sheet = workbook['Furniture']

    # Selected peers — first 7 fit in rows 3–9
    for i, peer in enumerate(selected_peers[:7]):
        row = 3 + i
        _safe_write(sheet, f'A{row}', peer.get('identifier'))
        _safe_write(sheet, f'B{row}', peer.get('company_name'))
        _safe_write(sheet, f'C{row}', peer.get('trbc_activity'))
        _safe_write(sheet, f'D{row}', peer.get('country'))
        _safe_write(sheet, f'E{row}', peer.get('business_description'))
        _safe_write(sheet, f'F{row}', peer.get('market_cap_thb_m'))
        # Reasoning for selection: prefer ring_justification, otherwise blank
        _safe_write(sheet, f'G{row}', peer.get('ring_justification') or "Selected as comparable peer")

    # Not selected peers — start at row 11
    for i, peer in enumerate(not_selected_peers):
        row = 11 + i
        ticker = peer.get('identifier')
        _safe_write(sheet, f'A{row}', ticker)
        _safe_write(sheet, f'B{row}', peer.get('company_name'))
        _safe_write(sheet, f'C{row}', peer.get('trbc_activity'))
        _safe_write(sheet, f'D{row}', peer.get('country'))
        _safe_write(sheet, f'E{row}', peer.get('business_description'))
        _safe_write(sheet, f'F{row}', peer.get('market_cap_thb_m'))
        _safe_write(sheet, f'G{row}', rejection_rationales.get(ticker, "Not a strong fit"))


def inject_phase3_data(workbook, deep_dive, lseg_parsed_peers, selected_peers, deal_code, client_name, latest_year):
    """
    Comparison tab + Appendix Hist Trading Performan tab.
    Peer order is taken from `selected_peers` (Phase 2 ordering).
    """
    qualitative = {q.get('identifier'): q for q in deep_dive.get('qualitative', [])}
    financials = {f.get('identifier'): f for f in deep_dive.get('financials_comparison', [])}
    lseg_by_ticker = _build_lseg_lookup(lseg_parsed_peers)

    # ----- Comparison tab -----
    if 'Comparison' in workbook.sheetnames:
        sheet = workbook['Comparison']

        # Qualitative — rows 3–8 (max 6 peers)
        for i, peer in enumerate(selected_peers[:6]):
            row = 3 + i
            ticker = peer.get('identifier')
            q = qualitative.get(ticker, {})
            _safe_write(sheet, f'A{row}', ticker)
            _safe_write(sheet, f'B{row}', peer.get('company_name'))
            _safe_write(sheet, f'C{row}', q.get('core_business_model'))
            _safe_write(sheet, f'D{row}', q.get('product_focus'))
            _safe_write(sheet, f'E{row}', q.get('similarity'))
            _safe_write(sheet, f'F{row}', q.get('comparison_points'))
            _safe_write(sheet, f'G{row}', q.get('differentiation_points'))

        # Quantitative — rows 12–17, columns A–G ONLY (H–L are formulas)
        for i, peer in enumerate(selected_peers[:6]):
            row = 12 + i
            ticker = peer.get('identifier')
            f = financials.get(ticker, {})
            _safe_write(sheet, f'A{row}', ticker)
            _safe_write(sheet, f'B{row}', f.get('revenue_2023'))
            _safe_write(sheet, f'C{row}', f.get('revenue_2024'))
            _safe_write(sheet, f'D{row}', f.get('revenue_2025'))
            _safe_write(sheet, f'E{row}', f.get('ebitda_2023'))
            _safe_write(sheet, f'F{row}', f.get('ebitda_2024'))
            _safe_write(sheet, f'G{row}', f.get('ebitda_2025'))

    # ----- Appendix Hist Trading Performan tab -----
    if 'Appendix Hist Trading Performan' in workbook.sheetnames:
        sheet = workbook['Appendix Hist Trading Performan']

        # Header inputs
        _safe_write(sheet, 'B2', f"Project {deal_code} - {client_name} Comps Tables")
        _safe_write(sheet, 'C4', latest_year)
        _safe_write(sheet, 'C5', "THB (actual / millions as noted)")

        # V2 Template structure (Dental2 — quarterly layout):
        #   Columns E-P = 12 quarterly multiples (Q1 2023 through Q4 2025)
        #     E-H = 2023 Q1-Q4, I-L = 2024 Q1-Q4, M-P = 2025 Q1-Q4
        #   Q = 12-Quarter Average, R = 12-Quarter Median (formulas)
        #   S-V = LTM quarters, W = LTM Median (formulas)
        #
        #   EV/EBITDA data rows: 10-16  (7 peers)
        #   P/E data rows:       23-29  (7 peers)
        #   EV/Revenue data rows: 34-40  (7 peers)
        #
        # LSEG parser returns annual data — we write to the Q4 column for each
        # year (H, L, P) as the best annual proxy.  Quarterly slots
        # (Q1-Q3) are left empty for manual entry from LSEG quarterly exports.

        # Clear INPUT cells for all 7 peer slots before writing.
        # EV/EBITDA: rows 10-16, clear B and E-P (cols 2, 5-16)
        # P/E: rows 22-28, clear E-P
        # EV/Revenue: rows 33-39, clear E-P
        for _r_off in range(7):
            for _col in [2] + list(range(5, 17)):  # B=2, E=5 through P=16
                _cell = sheet.cell(row=10 + _r_off, column=_col)
                if not isinstance(_cell, MergedCell) and not (
                    isinstance(_cell.value, str) and _cell.value.startswith('=')
                ):
                    _cell.value = None
            for _col in range(5, 17):  # E-P for P/E
                _cell = sheet.cell(row=23 + _r_off, column=_col)
                if not isinstance(_cell, MergedCell) and not (
                    isinstance(_cell.value, str) and _cell.value.startswith('=')
                ):
                    _cell.value = None
            for _col in range(5, 17):  # E-P for EV/Revenue
                _cell = sheet.cell(row=34 + _r_off, column=_col)
                if not isinstance(_cell, MergedCell) and not (
                    isinstance(_cell.value, str) and _cell.value.startswith('=')
                ):
                    _cell.value = None

        # EV/EBITDA: rows 10-16 ; P/E: rows 22-28 ; EV/Revenue: rows 33-39
        # Annual data mapped to Q4 column for each year:
        #   2023 Q4 = col H, 2024 Q4 = col L, 2025 Q4 = col P
        years_q4_cols = [
            ('2023', 'H'), ('2024', 'L'), ('2025', 'P'),
        ]

        for i, peer in enumerate(selected_peers[:7]):
            ticker = peer.get('identifier')
            lseg = _lseg_peer(lseg_by_ticker, ticker, peer.get('company_name'))

            # Section 1: EV/EBITDA — rows 10-16
            r1 = 10 + i
            _safe_write(sheet, f'C{r1}', ticker)
            ev_ebitda = lseg.get('ev_ebitda', {})
            for year, col in years_q4_cols:
                v = ev_ebitda.get(year)
                if v is not None:
                    _safe_write(sheet, f'{col}{r1}', v)

            # Section 2: P/E — rows 23-29 (C column is FORMULA, skip)
            r2 = 23 + i
            pe = lseg.get('pe', {})
            for year, col in years_q4_cols:
                v = pe.get(year)
                if v is not None:
                    _safe_write(sheet, f'{col}{r2}', v)

            # Section 3: EV/Revenue — rows 34-40 (C column is FORMULA, skip)
            r3 = 34 + i
            ev_rev = lseg.get('ev_revenue', {})
            for year, col in years_q4_cols:
                v = ev_rev.get(year)
                if v is not None:
                    _safe_write(sheet, f'{col}{r3}', v)


def inject_phase35_data(workbook, transactions, deal_code):
    """
    Precedent Transactions tab (V2 column structure).
    Columns: A=SDC Deal No, B=Date, C=Target, D=Nation, E=Industry,
             F=Acquiror, G=Deal Value USD M, H=Deal Value THB M (~35x),
             I=EV/EBITDA, J=EV/Revenue, K=Notes/Relevance
    Data starts at row 5 (header at row 4).
    """
    sheet_name = None
    for candidate in ('Precedent Transactions', 'Precedent Tx', 'Precedent_Transactions'):
        if candidate in workbook.sheetnames:
            sheet_name = candidate
            break
    if sheet_name is None:
        return
    sheet = workbook[sheet_name]

    # Locate header row by scanning rows 1-15 for "Target" + "Acquir" markers
    header_row = None
    for r in range(1, 16):
        row_vals = [str(sheet.cell(row=r, column=c).value or '').strip().lower()
                    for c in range(1, 12)]
        joined = ' | '.join(row_vals)
        if 'target' in joined and 'acquir' in joined:
            header_row = r
            break
    if header_row is None:
        header_row = 4  # V2 template default

    # Clear demo data rows (header_row+1 through header_row+10, columns A-K)
    for r in range(header_row + 1, header_row + 11):
        for c_letter in ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'):
            cell = sheet[f'{c_letter}{r}']
            v = cell.value
            if v is not None and not (isinstance(v, str) and v.startswith('=')):
                cell.value = None

    # Write up to 10 transactions (V2 column layout: A-K)
    for i, tx in enumerate(transactions[:10]):
        row = header_row + 1 + i

        sdc_no = tx.get('sdc_deal_no') or tx.get('SDC Deal No') or ''
        date = tx.get('date') or tx.get('Date Announced') or ''
        target = tx.get('target') or tx.get('Target Full Name') or ''
        nation = tx.get('region') or tx.get('Target Nation') or ''
        industry = tx.get('industry') or tx.get('Target Mid Industry') or ''
        acquirer = tx.get('acquirer') or tx.get('Acquiror Full Name') or ''
        deal_value_usd = tx.get('deal_value_usd_m') or tx.get('Deal Value (USD, Millions)')
        ev_ebitda = tx.get('ev_ebitda') or tx.get('Ratio of Enterprise Value to EBITDA')
        ev_revenue = tx.get('ev_revenue') or tx.get('ev_rev')
        notes = tx.get('relevance') or tx.get('notes_and_caveats') or ''

        # Compute THB equivalent at ~35x rate
        deal_value_thb = None
        if deal_value_usd is not None:
            try:
                deal_value_thb = round(float(deal_value_usd) * 35, 0)
            except (ValueError, TypeError):
                pass

        _safe_write(sheet, f'A{row}', sdc_no)
        _safe_write(sheet, f'B{row}', date)
        _safe_write(sheet, f'C{row}', target)
        _safe_write(sheet, f'D{row}', nation)
        _safe_write(sheet, f'E{row}', industry)
        _safe_write(sheet, f'F{row}', acquirer)
        _safe_write(sheet, f'G{row}', deal_value_usd)
        _safe_write(sheet, f'H{row}', deal_value_thb)
        _safe_write(sheet, f'I{row}', ev_ebitda)
        _safe_write(sheet, f'J{row}', ev_revenue)
        _safe_write(sheet, f'K{row}', notes)


def inject_phase4_data(workbook, data, latest_year):
    """
    Summary tab (FINAL V2 template).
    Writes: header, key metrics formulas, discount table, valuation table with
    implied EV, fair value rows, peer table with direct Appendix refs,
    EBITDA build table, and asset table structure.
    Formula guard (_safe_write) prevents overwriting any existing formulas.
    """
    if 'Summary' not in workbook.sheetnames:
        return
    sheet = workbook['Summary']

    deal_code = data.get('deal_code', 'DF-XXX')
    client_name = data.get('client_name', 'Company')
    bm = data.get('business_model', {})
    legal_name = (
        (bm.get('section_1', {}) or {}).get('company_name')
        or (bm.get('section_a', {}) or {}).get('company_name')
        or client_name
    )

    # --- 1.1 Title & header ---
    _safe_write(sheet, 'B2', f"Project {deal_code} – {client_name} ({legal_name}) Comps Tables")
    _safe_write(sheet, 'D3', f"From {latest_year}")

    # --- 1.2 Key Metrics (rows 4-5) ---
    _safe_write(sheet, 'C4', 'Revenue')
    _safe_write(sheet, 'D4', 'EBITDA')
    _safe_write(sheet, 'E4', 'Net Profit')
    _safe_write(sheet, 'B5', f'{latest_year} Adjusted')
    _safe_write(sheet, 'C5', "='Preliminary P&L'!G8")       # Total Revenue (absolute THB)
    _safe_write(sheet, 'D5', '=J45')                         # 2025 Adjusted EBITDA from build table
    _safe_write(sheet, 'E5', "='Preliminary P&L'!G35")       # Net Profit row 35
    _safe_write(sheet, 'F5', 'mm THB')

    # --- 1.3 Discount table (rows 6-9) ---
    _safe_write(sheet, 'B6', 'Discount Table')
    _safe_write(sheet, 'C7', '% of Discount')
    _safe_write(sheet, 'D7', '=Comparison!F39')              # Dynamic scale discount
    _safe_write(sheet, 'E7', 'Scale Discount')
    _safe_write(sheet, 'D8', '=Comparison!L18')              # Country discount (new column)
    _safe_write(sheet, 'E8', 'Country Discount')
    _safe_write(sheet, 'D9', '=1-(1-D7)*(1-D8)')            # Combined discount
    _safe_write(sheet, 'E9', 'Summary Discount')

    # --- 1.4 Valuation table — headers (rows 11-12) ---
    _safe_write(sheet, 'B11', 'EBITDA Multiple')
    _safe_write(sheet, 'E11', 'Discounted Multiple')
    _safe_write(sheet, 'F11', 'Implied Enterprise Value')
    _safe_write(sheet, 'I11', 'Last 12 Quarters')
    _safe_write(sheet, 'J11', 'Last 12 Quarters')
    _safe_write(sheet, 'K11', f'LTM 4Q, {latest_year}')
    _safe_write(sheet, 'L11', f'LTM 4Q, {latest_year}')
    _safe_write(sheet, 'H12', 'Metrics')
    _safe_write(sheet, 'I12', 'Industry Median ')
    _safe_write(sheet, 'J12', 'Industry Average')
    _safe_write(sheet, 'K12', 'Industry Median')
    _safe_write(sheet, 'L12', 'Industry Average')

    # EV/EBITDA (rows 12-14): industry data → I13/J13/K13/L13
    _safe_write(sheet, 'I13', '=F38')   # EV/EBITDA 12Q Median from peer table
    _safe_write(sheet, 'J13', '=E38')   # EV/EBITDA 12Q Average
    _safe_write(sheet, 'K13', '=H38')   # EV/EBITDA LTM Median
    _safe_write(sheet, 'L13', '=G38')   # EV/EBITDA LTM Average
    _safe_write(sheet, 'D12', f'LTM 4Q, {latest_year}')
    _safe_write(sheet, 'E12', '=(MIN(I13:L13)*(1-$D$7)*(1-$D$8))')
    _safe_write(sheet, 'F12', '=$D$5*E12')
    _safe_write(sheet, 'D13', 'Last 12 Quarters')
    _safe_write(sheet, 'E13', '=(MAX(I13:L13)*(1-$D$7)*(1-$D$8))')
    _safe_write(sheet, 'F13', '=$D$5*E13')
    _safe_write(sheet, 'E14', '=AVERAGE(E12:E13)')
    _safe_write(sheet, 'F14', '=$D$5*E14')

    # P/E (rows 15-18): industry data → I14/J14/K14/L14
    _safe_write(sheet, 'B15', 'P/E Multiple')
    _safe_write(sheet, 'E15', 'Discounted Multiple')
    _safe_write(sheet, 'F15', 'Implied Equity Value')
    _safe_write(sheet, 'I14', '=J38')   # P/E 12Q Median
    _safe_write(sheet, 'J14', '=I38')   # P/E 12Q Average
    _safe_write(sheet, 'K14', '=L38')   # P/E LTM Median
    _safe_write(sheet, 'L14', '=K38')   # P/E LTM Average
    _safe_write(sheet, 'D16', f'LTM 4Q, {latest_year}')
    _safe_write(sheet, 'E16', '=(MIN(I14:L14)*(1-$D$7)*(1-$D$8))')
    _safe_write(sheet, 'F16', '=$E$5*E16')
    _safe_write(sheet, 'D17', 'Last 12 Quarters')
    _safe_write(sheet, 'E17', '=(MAX(I14:L14)*(1-$D$7)*(1-$D$8))')
    _safe_write(sheet, 'F17', '=$E$5*E17')
    _safe_write(sheet, 'E18', '=AVERAGE(E16:E17)')
    _safe_write(sheet, 'F18', '=$E$5*E18')

    # EV/Revenue (rows 19-22): industry data → I15/J15/K15/L15
    _safe_write(sheet, 'B19', 'Revenue Multiple')
    _safe_write(sheet, 'E19', 'Discounted Multiple')
    _safe_write(sheet, 'F19', 'Implied Enterprise Value')
    _safe_write(sheet, 'I15', '=M38')   # EV/Rev 12Q Median
    _safe_write(sheet, 'J15', '=N38')   # EV/Rev 12Q Average (corrected from report)
    _safe_write(sheet, 'K15', '=O38')   # EV/Rev LTM Median
    _safe_write(sheet, 'L15', '=P38')   # EV/Rev LTM Average
    _safe_write(sheet, 'D20', f'LTM 4Q, {latest_year}')
    _safe_write(sheet, 'E20', '=(MIN(I15:L15)*(1-$D$7)*(1-$D$8))')
    _safe_write(sheet, 'F20', '=$C$5*E20')
    _safe_write(sheet, 'D21', 'Last 12 Quarters')
    _safe_write(sheet, 'E21', '=(MAX(I15:L15)*(1-$D$7)*(1-$D$8))')
    _safe_write(sheet, 'F21', '=$C$5*E21')
    _safe_write(sheet, 'E22', '=AVERAGE(E20:E21)')
    _safe_write(sheet, 'F22', '=$C$5*E22')

    # --- 1.5 Fair Value Summary (rows 24-26) ---
    _safe_write(sheet, 'E24', 'Floor')
    _safe_write(sheet, 'F24', '=AVERAGE(F12,F16,F20)')
    _safe_write(sheet, 'E25', 'Average Fair Value')
    _safe_write(sheet, 'F25', '=AVERAGE(F18,F14,F22)')
    _safe_write(sheet, 'E26', 'Ceiling')
    _safe_write(sheet, 'F26', '=AVERAGE(F13,F17,F21)')

    # --- 1.6 Comparable Company table — headers (rows 28-30) ---
    _safe_write(sheet, 'E28', 'EV/EBITDA')
    _safe_write(sheet, 'I28', 'P/E')
    _safe_write(sheet, 'M28', 'EV/Revenue')
    _safe_write(sheet, 'B29', 'Note')
    _safe_write(sheet, 'C29', 'Ticker')
    _safe_write(sheet, 'D29', 'Country')
    for col, label in [('E', 'Last 12Q'), ('F', ' Last 12Q'), ('G', f'LTM {latest_year}'), ('H', f'LTM {latest_year}')]:
        _safe_write(sheet, f'{col}29', label)
    for col, label in [('I', 'Last 12Q'), ('J', ' Last 12Q'), ('K', f'LTM {latest_year}'), ('L', f'LTM {latest_year}')]:
        _safe_write(sheet, f'{col}29', label)
    for col, label in [('M', 'Last 12Q'), ('N', ' Last 12Q'), ('O', f'LTM {latest_year}'), ('P', f'LTM {latest_year}')]:
        _safe_write(sheet, f'{col}29', label)
    for col in ('E', 'G', 'I', 'K', 'M', 'O'):
        _safe_write(sheet, f'{col}30', 'Avg')
    for col in ('F', 'H', 'J', 'L', 'N', 'P'):
        _safe_write(sheet, f'{col}30', 'Median')

    # Peer rows (31-37) — direct Appendix references for 7 peers
    for idx in range(7):
        row = 31 + idx
        app_ev = 10 + idx     # Appendix EV/EBITDA peer row
        app_pe = 23 + idx     # Appendix P/E peer row
        app_rv = 34 + idx     # Appendix EV/Revenue peer row
        _safe_write(sheet, f'A{row}', idx + 1)
        _safe_write(sheet, f'C{row}', f"='Appendix Hist Trading Performan'!C{app_ev}")
        # EV/EBITDA: 12Q Avg (P), 12Q Median (Q), LTM Avg (V), LTM Median (W)
        _safe_write(sheet, f'E{row}', f"='Appendix Hist Trading Performan'!P{app_ev}")
        _safe_write(sheet, f'F{row}', f"='Appendix Hist Trading Performan'!Q{app_ev}")
        _safe_write(sheet, f'G{row}', f"='Appendix Hist Trading Performan'!V{app_ev}")
        _safe_write(sheet, f'H{row}', f"='Appendix Hist Trading Performan'!W{app_ev}")
        # P/E
        _safe_write(sheet, f'I{row}', f"='Appendix Hist Trading Performan'!P{app_pe}")
        _safe_write(sheet, f'J{row}', f"='Appendix Hist Trading Performan'!Q{app_pe}")
        _safe_write(sheet, f'K{row}', f"='Appendix Hist Trading Performan'!V{app_pe}")
        _safe_write(sheet, f'L{row}', f"='Appendix Hist Trading Performan'!W{app_pe}")
        # EV/Revenue
        _safe_write(sheet, f'M{row}', f"='Appendix Hist Trading Performan'!P{app_rv}")
        _safe_write(sheet, f'N{row}', f"='Appendix Hist Trading Performan'!Q{app_rv}")
        _safe_write(sheet, f'O{row}', f"='Appendix Hist Trading Performan'!V{app_rv}")
        _safe_write(sheet, f'P{row}', f"='Appendix Hist Trading Performan'!W{app_rv}")

    # Average and Median rows (38-39)
    _safe_write(sheet, 'D38', 'Average')
    _safe_write(sheet, 'D39', 'Median')
    for col in ('E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'):
        _safe_write(sheet, f'{col}38', f'=AVERAGE({col}31:{col}37)')
        _safe_write(sheet, f'{col}39', f'=MEDIAN({col}31:{col}37)')

    # --- 1.7 EBITDA Build Table (rows 41-45) ---
    # Headers
    for col, label in [('B', 'Year'), ('C', 'Revenue'), ('D', 'Net Profit'),
                        ('E', 'Interest'), ('F', 'Tax'), ('G', 'Depreciation'),
                        ('H', 'EBITDA'), ('I', 'Add Back'), ('J', 'Adjusted EBITDA')]:
        _safe_write(sheet, f'{col}41', label)

    # Data rows: 42=2022, 43=2023, 44=2024, 45=2025
    pnl_year_cols = {2022: 'D', 2023: 'E', 2024: 'F', 2025: 'G'}
    for build_row, year in [(42, 2022), (43, 2023), (44, 2024), (45, 2025)]:
        pnl_col = pnl_year_cols[year]
        _safe_write(sheet, f'B{build_row}', str(year))
        _safe_write(sheet, f'C{build_row}', f"='Preliminary P&L'!{pnl_col}8")    # Revenue
        _safe_write(sheet, f'D{build_row}', f"='Preliminary P&L'!{pnl_col}35")   # Net Profit
        _safe_write(sheet, f'E{build_row}', f"='Preliminary P&L'!{pnl_col}30")   # Interest
        _safe_write(sheet, f'F{build_row}', f"='Preliminary P&L'!{pnl_col}34")   # Tax
        _safe_write(sheet, f'G{build_row}', f"='Preliminary P&L'!{pnl_col}26")   # D&A
        _safe_write(sheet, f'H{build_row}', f"='Preliminary P&L'!{pnl_col}23")   # EBITDA
        if year == 2022:
            _safe_write(sheet, f'I{build_row}', 0)  # No add-back for oldest year
        else:
            _safe_write(sheet, f'I{build_row}', f"='Preliminary P&L'!{pnl_col}40")  # Total Add-Back
        _safe_write(sheet, f'J{build_row}', f'=H{build_row}+I{build_row}')  # Adj EBITDA
        if build_row >= 43:  # EBITDA margin for 2023+
            _safe_write(sheet, f'K{build_row}', f'=J{build_row}/C{build_row}')

    # --- 1.8 Asset Table (rows 47-54) ---
    _safe_write(sheet, 'B47', 'Assets to add value(M)')
    _safe_write(sheet, 'C47', 'Value')
    _safe_write(sheet, 'D47', 'Notes')
    asset_labels = ['Machines', 'Buildings', 'Land ', 'Vehicles', 'Inventory']
    for i, label in enumerate(asset_labels):
        _safe_write(sheet, f'B{48+i}', label)
    _safe_write(sheet, 'B53', 'Total Asset')
    _safe_write(sheet, 'C53', '=SUM(C48:C52)')
    _safe_write(sheet, 'D53', 'Sum of all transferable physical assets.')
    _safe_write(sheet, 'B54', 'Total EV')
    _safe_write(sheet, 'C54', '=F25+C53')
    _safe_write(sheet, 'D54', 'Average Fair Value (income approach, multiple-based) + Total tangible asset value.')
