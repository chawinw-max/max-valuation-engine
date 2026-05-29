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

        row_map = {
            'sales_and_services': 6,
            'other_revenues': 7,
            'cost_of_goods_sold': 11,
            'sales_expenses': 19,
            'administrative_expenses': 20,
            'other_expenses': 21,
            'depreciation_amortization': 28,
            'interest_expenses': 33,
            'tax': 37,
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
                    
    # 2. Business Model
    if 'Business Model' in workbook.sheetnames:
        sheet = workbook['Business Model']
        bm = data.get('business_model', {})
        
        # Title
        _safe_write(sheet, 'A1', f"{data.get('client_name', 'Company').upper()} – BUSINESS MODEL OVERVIEW")
        
        # Section A (Rows 4-13)
        sec_a = bm.get('section_a', {})
        rows_a = ['company_name', 'founded', 'type', 'registered_capital', 'shareholders', 'location', 'focus', 'certifications', 'operating_history', 'website']
        for i, field in enumerate(rows_a):
            _safe_write(sheet, f'B{4+i}', sec_a.get(field))
            
        # Section B (Rows 16-22)
        sec_b = bm.get('section_b', {})
        rows_b = ['primary_revenue', 'key_clients', 'pricing_strategy', 'customer_mix', 'client_segmentation', 'credit_terms', 'secondary_revenue']
        for i, field in enumerate(rows_b):
            _safe_write(sheet, f'B{16+i}', sec_b.get(field))
            
        # Section C (Rows 25-30)
        sec_c = bm.get('section_c', [])
        for i, moat in enumerate(sec_c[:6]): # Max 6 moats
            _safe_write(sheet, f'A{25+i}', moat.get('title'))
            _safe_write(sheet, f'B{25+i}', moat.get('description'))
            
        # Section D (Rows 33-40)
        sec_d = bm.get('section_d', {})
        _safe_write(sheet, 'B33', sec_d.get('revenue_fy2022'))
        _safe_write(sheet, 'B34', sec_d.get('revenue_fy2023'))
        _safe_write(sheet, 'B35', sec_d.get('revenue_fy2024'))
        _safe_write(sheet, 'B36', sec_d.get('revenue_fy2025'))
        _safe_write(sheet, 'B37', sec_d.get('gross_margin'))
        _safe_write(sheet, 'B38', sec_d.get('ebitda_normalization'))
        _safe_write(sheet, 'B39', sec_d.get('liabilities'))
        _safe_write(sheet, 'B40', sec_d.get('cash_flow_note'))
        
        # Section E (Rows 43-48)
        sec_e = bm.get('section_e', {})
        rows_e = ['seller', 'exit_motivation', 'asking_price', 'property', 'transaction_type', 'next_steps']
        for i, field in enumerate(rows_e):
            _safe_write(sheet, f'B{43+i}', sec_e.get(field))
            
        # Section F (Rows 51-57)
        sec_f = bm.get('section_f', {})
        rows_f = ['total_staff', 'shift_structure', 'inbound_process', 'facility', 'machinery', 'logistics_fleet', 'it_systems']
        for i, field in enumerate(rows_f):
            _safe_write(sheet, f'B{51+i}', sec_f.get(field))

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

        # Template structure (updated — quarterly layout):
        #   Columns D-W = quarterly multiples (Q1 2021 through Q4 2025)
        #     D-G = 2021 Q1-Q4, H-K = 2022 Q1-Q4, L-O = 2023 Q1-Q4,
        #     P-S = 2024 Q1-Q4, T-W = 2025 Q1-Q4
        #   X = Average (last 12 quarters), Y = Median (last 12 quarters)
        #   Z-AC = LTM quarters, AD-AE = LTM Avg/Median
        #
        #   EV/EBITDA data rows: 11-17  (A=position, B=ArrayFormula, C=ticker INPUT, D-W=quarterly INPUT)
        #   P/E data rows:       23-29  (C=formula =C11..C17, D-W=quarterly INPUT)
        #   EV/Revenue data rows:35-41  (C=formula =C23..C29, D-W=quarterly INPUT)
        #
        # LSEG parser returns annual data — we write to the Q4 column for each
        # year (G, K, O, S, W) as the best annual proxy.  Quarterly slots
        # (Q1-Q3) are left empty for manual entry from LSEG quarterly exports.

        # Clear INPUT cells for all 7 peer slots before writing.
        # EV/EBITDA: skip A (pre-filled position numbers), clear B and D-W.
        # P/E / EV/Revenue: skip C (formula), clear D-W.
        for _r_off in range(7):
            for _col in [2] + list(range(4, 24)):  # B=2, D=4 through W=23
                _cell = sheet.cell(row=11 + _r_off, column=_col)
                if not isinstance(_cell, MergedCell) and not (
                    isinstance(_cell.value, str) and _cell.value.startswith('=')
                ):
                    _cell.value = None
            for _col in range(4, 24):  # D-W for P/E
                _cell = sheet.cell(row=23 + _r_off, column=_col)
                if not isinstance(_cell, MergedCell) and not (
                    isinstance(_cell.value, str) and _cell.value.startswith('=')
                ):
                    _cell.value = None
            for _col in range(4, 24):  # D-W for EV/Revenue
                _cell = sheet.cell(row=35 + _r_off, column=_col)
                if not isinstance(_cell, MergedCell) and not (
                    isinstance(_cell.value, str) and _cell.value.startswith('=')
                ):
                    _cell.value = None

        # EV/EBITDA: rows 11-17 ; P/E: rows 23-29 ; EV/Revenue: rows 35-41
        # Annual data mapped to Q4 column for each year:
        #   2021 Q4 = col G, 2022 Q4 = col K, 2023 Q4 = col O,
        #   2024 Q4 = col S, 2025 Q4 = col W
        years_q4_cols = [
            ('2021', 'G'), ('2022', 'K'), ('2023', 'O'),
            ('2024', 'S'), ('2025', 'W'),
        ]

        for i, peer in enumerate(selected_peers[:7]):
            ticker = peer.get('identifier')
            lseg = _lseg_peer(lseg_by_ticker, ticker, peer.get('company_name'))

            # Section 1: EV/EBITDA — rows 11-17
            r1 = 11 + i
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

            # Section 3: EV/Revenue — rows 35-41 (C column is FORMULA, skip)
            r3 = 35 + i
            ev_rev = lseg.get('ev_revenue', {})
            for year, col in years_q4_cols:
                v = ev_rev.get(year)
                if v is not None:
                    _safe_write(sheet, f'{col}{r3}', v)


def inject_phase35_data(workbook, transactions, deal_code):
    """
    Precedent Transactions tab. Header location varies; we find the row that
    looks like the header by searching for "Target" + "Acquirer" markers, then
    write data rows below. Formula guard prevents overwriting Min/Median/etc.
    """
    sheet_name = None
    for candidate in ('Precedent Transactions', 'Precedent Tx', 'Precedent_Transactions'):
        if candidate in workbook.sheetnames:
            sheet_name = candidate
            break
    if sheet_name is None:
        return
    sheet = workbook[sheet_name]

    # Locate header row by scanning rows 1-15 for "Target" in column B or nearby
    header_row = None
    for r in range(1, 16):
        row_vals = [str(sheet.cell(row=r, column=c).value or '').strip().lower()
                    for c in range(1, 10)]
        joined = ' | '.join(row_vals)
        if 'target' in joined and 'acquir' in joined:
            header_row = r
            break
    if header_row is None:
        header_row = 8  # template default

    # Update "Relevance to ..." header with actual deal code
    rel_col = 6  # column F
    for c in range(1, 10):
        val = str(sheet.cell(row=header_row, column=c).value or '')
        if 'relevance' in val.lower():
            rel_col = c
            _safe_write(sheet, f'{get_column_letter(c)}{header_row}',
                        f'Relevance to {deal_code}')
            break

    # Clear demo data rows (header_row+1 through header_row+7, columns B-I)
    for r in range(header_row + 1, header_row + 8):
        for c_letter in ('B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'):
            cell = sheet[f'{c_letter}{r}']
            v = cell.value
            if v is not None and not (isinstance(v, str) and v.startswith('=')):
                cell.value = None

    # Write up to 10 transactions starting from header_row + 1 (columns B-I)
    for i, tx in enumerate(transactions[:10]):
        row = header_row + 1 + i
        target = tx.get('target') or tx.get('Target Full Name')
        acquirer = tx.get('acquirer') or tx.get('Acquiror Full Name')
        date = tx.get('date') or tx.get('Date Announced')
        region = tx.get('region') or tx.get('Target Nation')
        relevance = tx.get('relevance') or ''
        deal_value = tx.get('deal_value_usd_m') or tx.get('Deal Value (USD, Millions)')
        ev_ebitda = tx.get('ev_ebitda') or tx.get('Ratio of Enterprise Value to EBITDA')
        notes = tx.get('notes_and_caveats') or ''

        _safe_write(sheet, f'B{row}', target)
        _safe_write(sheet, f'C{row}', acquirer)
        _safe_write(sheet, f'D{row}', date)
        _safe_write(sheet, f'E{row}', region)
        _safe_write(sheet, f'F{row}', relevance)
        _safe_write(sheet, f'G{row}', deal_value)
        _safe_write(sheet, f'H{row}', ev_ebitda)
        _safe_write(sheet, f'I{row}', notes)


def inject_phase4_data(workbook, data, latest_year):
    """
    Summary tab. Almost everything is a formula — we only fill the small set
    of input cells. Formula guard takes care of accidents.
    """
    if 'Summary' not in workbook.sheetnames:
        return
    sheet = workbook['Summary']

    deal_code = data.get('deal_code', 'DF-XXX')
    client_name = data.get('client_name', 'Company')
    bm = data.get('business_model', {})
    legal_name = (bm.get('section_a', {}) or {}).get('company_name') or client_name

    _safe_write(sheet, 'B2', f"Project {deal_code} – {client_name} ({legal_name}) Comps Tables")
    _safe_write(sheet, 'D3', f"From {latest_year}")
