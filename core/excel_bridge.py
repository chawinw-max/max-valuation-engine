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


def _force_write(sheet, cell_coord, value):
    """
    Force-writes a value to a cell, even if it contains a formula.
    Used when the template structure changed and old formulas must be replaced.
    Handles merged cells like _safe_write.
    """
    cell = sheet[cell_coord]
    if isinstance(cell, MergedCell):
        for merged_range in sheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                break
        else:
            return False
    if isinstance(value, float) and math.isnan(value):
        return False
    cell.value = value
    return True


def _clear_cell(sheet, cell_coord):
    """Clear a cell's value (including formulas). For template migration."""
    cell = sheet[cell_coord]
    if isinstance(cell, MergedCell):
        for merged_range in sheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                break
        else:
            return
    cell.value = None


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
        _safe_write(sheet, 'B3', f"{data.get('client_name', 'Company')} — Management Accounts ({years_str})")

        financials = data.get('financials', {})

        # Actual template row map (EBIT-first layout):
        #   Row 6: Sales, Row 7: Other Rev, Row 8: Total Rev [FORMULA]
        #   Row 11: COGS, Row 14: GP [FORMULA]
        #   Row 18: Sales Exp, Row 19: Admin Exp, Row 20: Other Exp
        #   Row 24: EBIT = GP - OpEx [FORMULA]
        #   Row 27: D&A (B27='Depreciation and Amortization')
        #   Row 28: EBITDA = EBIT + D&A [FORMULA]
        #   Row 32: Interest (B32='Interest Expeneses')
        #   Row 33: EBT = EBIT - Interest [FORMULA]
        #   Row 36: Tax (B36='Tax')
        #   Row 39: NP = EBT - Tax [FORMULA]
        row_map = {
            'sales_and_services': 6,
            'other_revenues': 7,
            'cost_of_goods_sold': 11,
            'sales_expenses': 18,
            'administrative_expenses': 19,
            'other_expenses': 20,
            'depreciation_amortization': 27,
            'interest_expenses': 32,
            'tax': 36,
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

    # 2. Business Model (OLD template — 6-section layout: A through F)
    if 'Business Model' in workbook.sheetnames:
        sheet = workbook['Business Model']
        bm = data.get('business_model', {})
        client_name_bm = data.get('client_name', 'Company')

        # A1: Title
        _safe_write(sheet, 'A1', f"{client_name_bm.upper()} – BUSINESS MODEL OVERVIEW")

        # Section A: rows 4-13 (10 fields in B column)
        sec_a = bm.get('section_a', {}) or {}
        rows_a = ['company_name', 'founded', 'type', 'registered_capital',
                  'shareholders', 'location', 'focus', 'certifications',
                  'operating_history', 'website']
        for i, field in enumerate(rows_a):
            _safe_write(sheet, f'B{4+i}', sec_a.get(field))

        # Section B: rows 16-22 (7 fields in B column)
        sec_b = bm.get('section_b', {}) or {}
        rows_b = ['primary_revenue', 'key_clients', 'pricing_strategy',
                  'customer_mix', 'client_segmentation', 'credit_terms',
                  'secondary_revenue']
        for i, field in enumerate(rows_b):
            _safe_write(sheet, f'B{16+i}', sec_b.get(field))

        # Section C: rows 25-30 (moats, A=title, B=description)
        sec_c = bm.get('section_c', []) or []
        for i, moat in enumerate(sec_c[:6]):
            _safe_write(sheet, f'A{25+i}', moat.get('title'))
            _safe_write(sheet, f'B{25+i}', moat.get('description'))

        # Section D: rows 33-40 (8 fields in B column)
        sec_d = bm.get('section_d', {}) or {}
        rows_d = ['revenue_fy2022', 'revenue_fy2023', 'revenue_fy2024',
                  'revenue_fy2025', 'gross_margin', 'ebitda_normalization',
                  'liabilities', 'cash_flow_note']
        for i, field in enumerate(rows_d):
            _safe_write(sheet, f'B{33+i}', sec_d.get(field))

        # Section E: rows 43-48 (6 fields in B column)
        sec_e = bm.get('section_e', {}) or {}
        rows_e = ['seller', 'exit_motivation', 'asking_price',
                  'property', 'transaction_type', 'next_steps']
        for i, field in enumerate(rows_e):
            _safe_write(sheet, f'B{43+i}', sec_e.get(field))

        # Section F: rows 51-57 (7 fields in B column)
        sec_f = bm.get('section_f', {}) or {}
        rows_f = ['total_staff', 'shift_structure', 'inbound_process',
                  'facility', 'machinery', 'logistics_fleet', 'it_systems']
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

        # Qualitative — rows 5–10, columns C–J (max 6 peers)
        for i, peer in enumerate(selected_peers[:6]):
            row = 5 + i
            ticker = peer.get('identifier')
            q = qualitative.get(ticker, {})
            _safe_write(sheet, f'C{row}', ticker)
            _safe_write(sheet, f'D{row}', peer.get('company_name'))
            _safe_write(sheet, f'E{row}', q.get('core_business_model'))
            _safe_write(sheet, f'F{row}', q.get('product_focus'))
            _safe_write(sheet, f'G{row}', q.get('similarity'))
            _safe_write(sheet, f'H{row}', q.get('comparison_points'))
            _safe_write(sheet, f'I{row}', q.get('differentiation_points'))
            _safe_write(sheet, f'J{row}', peer.get('country'))

        # Financial — rows 18–23, columns C–N (max 6 peers)
        # Clear old IF(MATCH) formulas before writing
        for i in range(6):
            row = 18 + i
            for col in ('D', 'E', 'G', 'H', 'J', 'K', 'M', 'N'):
                _force_write(sheet, f'{col}{row}', None)

        for i, peer in enumerate(selected_peers[:6]):
            row = 18 + i
            ticker = peer.get('identifier')
            f = financials.get(ticker, {})
            # C has =C5..=C10 formulas linking to qualitative ticker — skip C
            # 4 years of Revenue + EBITDA: D/E=2021, G/H=2022, J/K=2023, M/N=2024
            _force_write(sheet, f'D{row}', f.get('revenue_2021'))
            _force_write(sheet, f'E{row}', f.get('ebitda_2021'))
            _force_write(sheet, f'G{row}', f.get('revenue_2022'))
            _force_write(sheet, f'H{row}', f.get('ebitda_2022'))
            _force_write(sheet, f'J{row}', f.get('revenue_2023'))
            _force_write(sheet, f'K{row}', f.get('ebitda_2023'))
            _force_write(sheet, f'M{row}', f.get('revenue_2024'))
            _force_write(sheet, f'N{row}', f.get('ebitda_2024'))

    # ----- Appendix Hist Trading Performan tab -----
    if 'Appendix Hist Trading Performan' in workbook.sheetnames:
        sheet = workbook['Appendix Hist Trading Performan']

        # Header inputs
        _safe_write(sheet, 'B2', f"Project {deal_code} - {client_name} Comps Tables")
        _safe_write(sheet, 'C4', latest_year)
        _safe_write(sheet, 'C5', "THB (actual / millions as noted)")

        # Actual template structure (annual layout):
        #   EV/EBITDA: 6 peer rows at 9-14
        #     B=company name, C=ticker, D=2021, E=2022, F=2023, G=2024, H=2025
        #     I=Average formula, J=Median formula (don't touch)
        #     K-N=LTM quarterly (leave empty), O-P=LTM formulas (don't touch)
        #   P/E: 6 peer rows at 20-25
        #     C=formula linking to EV/EBITDA ticker (don't write to C)
        #     D-H=annual P/E data
        #   EV/Revenue: 6 peer rows at 31-36
        #     C=formula linking to EV/EBITDA ticker (don't write to C)
        #     D-H=annual EV/Revenue data

        # Annual data column mapping: 2021→D, 2022→E, 2023→F, 2024→G, 2025→H
        years_annual_cols = [
            ('2021', 'D'), ('2022', 'E'), ('2023', 'F'), ('2024', 'G'), ('2025', 'H'),
        ]

        # Clear INPUT cells for all 6 peer slots before writing
        for _r_off in range(6):
            # EV/EBITDA rows 9-14: clear B (col 2) and D-H (cols 4-8)
            for _col in [2, 4, 5, 6, 7, 8]:
                _cell = sheet.cell(row=9 + _r_off, column=_col)
                if not isinstance(_cell, MergedCell) and not (
                    isinstance(_cell.value, str) and _cell.value.startswith('=')
                ):
                    _cell.value = None
            # P/E rows 20-25: clear D-H (cols 4-8)
            for _col in [4, 5, 6, 7, 8]:
                _cell = sheet.cell(row=20 + _r_off, column=_col)
                if not isinstance(_cell, MergedCell) and not (
                    isinstance(_cell.value, str) and _cell.value.startswith('=')
                ):
                    _cell.value = None
            # EV/Revenue rows 31-36: clear D-H (cols 4-8)
            for _col in [4, 5, 6, 7, 8]:
                _cell = sheet.cell(row=31 + _r_off, column=_col)
                if not isinstance(_cell, MergedCell) and not (
                    isinstance(_cell.value, str) and _cell.value.startswith('=')
                ):
                    _cell.value = None

        for i, peer in enumerate(selected_peers[:6]):
            ticker = peer.get('identifier')
            lseg = _lseg_peer(lseg_by_ticker, ticker, peer.get('company_name'))

            # Section 1: EV/EBITDA — rows 9-14
            r1 = 9 + i
            _safe_write(sheet, f'B{r1}', peer.get('company_name'))  # Company name
            _safe_write(sheet, f'C{r1}', ticker)
            ev_ebitda = lseg.get('ev_ebitda', {})
            for year, col in years_annual_cols:
                v = ev_ebitda.get(year)
                if v is not None:
                    _safe_write(sheet, f'{col}{r1}', v)

            # Section 2: P/E — rows 20-25 (C column is FORMULA, skip)
            r2 = 20 + i
            pe = lseg.get('pe', {})
            for year, col in years_annual_cols:
                v = pe.get(year)
                if v is not None:
                    _safe_write(sheet, f'{col}{r2}', v)

            # Section 3: EV/Revenue — rows 31-36 (C column is FORMULA, skip)
            r3 = 31 + i
            ev_rev = lseg.get('ev_revenue', {})
            for year, col in years_annual_cols:
                v = ev_rev.get(year)
                if v is not None:
                    _safe_write(sheet, f'{col}{r3}', v)


def inject_phase35_data(workbook, transactions, deal_code):
    """
    Precedent Transactions tab (OLD column structure).
    Columns B-I: B=target, C=acquirer, D=date, E=region,
                 F=relevance, G=deal_value_usd_m, H=ev_ebitda, I=notes
    Default header at row 8, data starts at row 9.
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
        header_row = 8  # OLD template default

    # Clear demo data rows (header_row+1 through header_row+7, columns B-I)
    for r in range(header_row + 1, header_row + 8):
        for c_letter in ('B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'):
            cell = sheet[f'{c_letter}{r}']
            v = cell.value
            if v is not None and not (isinstance(v, str) and v.startswith('=')):
                cell.value = None

    # Write up to 7 transactions (OLD column layout: B-I)
    for i, tx in enumerate(transactions[:7]):
        row = header_row + 1 + i

        target = tx.get('target') or tx.get('Target Full Name') or ''
        acquirer = tx.get('acquirer') or tx.get('Acquiror Full Name') or ''
        date = tx.get('date') or tx.get('Date Announced') or ''
        region = tx.get('region') or tx.get('Target Nation') or ''
        relevance = tx.get('relevance') or ''
        deal_value_usd = tx.get('deal_value_usd_m') or tx.get('Deal Value (USD, Millions)')
        ev_ebitda = tx.get('ev_ebitda') or tx.get('Ratio of Enterprise Value to EBITDA')
        notes = tx.get('notes_and_caveats') or ''

        _safe_write(sheet, f'B{row}', target)
        _safe_write(sheet, f'C{row}', acquirer)
        _safe_write(sheet, f'D{row}', date)
        _safe_write(sheet, f'E{row}', region)
        _safe_write(sheet, f'F{row}', relevance)
        _safe_write(sheet, f'G{row}', deal_value_usd)
        _safe_write(sheet, f'H{row}', ev_ebitda)
        _safe_write(sheet, f'I{row}', notes)


def inject_phase4_data(workbook, data, latest_year):
    """
    Summary tab — MINIMAL writes. The template handles almost everything
    via formulas that auto-calculate from the P&L and Appendix data.
    Only write the project title and latest year reference.
    """
    if 'Summary' not in workbook.sheetnames:
        return
    sheet = workbook['Summary']

    deal_code = data.get('deal_code', 'DF-XXX')
    client_name = data.get('client_name', 'Company')
    bm = data.get('business_model', {})
    legal_name = (
        (bm.get('section_a', {}) or {}).get('company_name')
        or client_name
    )

    # B2: Project title
    _safe_write(sheet, 'B2', f"Project {deal_code} – {client_name} ({legal_name}) Comps Tables")

    # D3: Latest year reference
    _safe_write(sheet, 'D3', f"From {latest_year}")
