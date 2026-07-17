import io
import fitz  # PyMuPDF
import docx
import pandas as pd

def extract_text_from_file(uploaded_file):
    """
    Extracts text from an uploaded Streamlit file based on its extension.
    """
    filename = uploaded_file.name.lower()
    file_bytes = uploaded_file.read()
    
    # Try parsing based on extension
    if filename.endswith(".pdf"):
        return _extract_from_pdf(file_bytes)
    elif filename.endswith(".docx"):
        return _extract_from_docx(file_bytes)
    elif filename.endswith(".xlsx") or filename.endswith(".xls"):
        return _extract_from_excel(file_bytes)
    elif filename.endswith(".csv"):
        return _extract_from_csv(file_bytes)
    elif filename.endswith((".png", ".jpg", ".jpeg")):
        # We will pass images directly to Gemini instead of extracting text
        return f"[IMAGE FILE: {uploaded_file.name}] - To be processed visually."
    else:
        # Fallback to standard text decoding
        try:
            return file_bytes.decode('utf-8')
        except UnicodeDecodeError:
            return f"[Unsupported file format: {uploaded_file.name}]"

def _extract_from_pdf(file_bytes):
    try:
        pdf_document = fitz.open(stream=file_bytes, filetype="pdf")
        text = ""
        for page_num in range(pdf_document.page_count):
            page = pdf_document.load_page(page_num)
            text += page.get_text("text") + "\n"
        return text
    except Exception as e:
        return f"[Error extracting PDF: {str(e)}]"

def _extract_from_docx(file_bytes):
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    except Exception as e:
        return f"[Error extracting DOCX: {str(e)}]"

def _extract_from_excel(file_bytes):
    try:
        # Read all sheets
        dfs = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
        text = ""
        for sheet_name, df in dfs.items():
            text += f"--- Sheet: {sheet_name} ---\n"
            text += df.to_string(index=False) + "\n\n"
        return text
    except Exception as e:
        return f"[Error extracting EXCEL: {str(e)}]"

def _extract_from_csv(file_bytes):
    try:
        df = pd.read_csv(io.BytesIO(file_bytes))
        return df.to_string(index=False)
    except Exception as e:
        return f"[Error extracting CSV: {str(e)}]"

TARGET_YEARS = ["2021", "2022", "2023", "2024", "2025"]


def _lseg_find_row(df, start_str, skip_str=None):
    """Index of the first row whose col-A value starts with start_str."""
    for i, row in df.iterrows():
        val = str(row.iloc[0]).strip()
        if val.startswith(start_str):
            if skip_str and skip_str in val:
                continue
            return i
    return None


def _lseg_safe_val(v):
    if pd.isna(v):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _lseg_col_year_map(df):
    """Column index -> year map from the 'Statement Data' header row."""
    year_header_idx = _lseg_find_row(df, "Statement Data")
    col_year_map = {}
    if year_header_idx is not None:
        header_row = df.iloc[year_header_idx]
        for col_i in range(1, len(header_row)):
            cell = header_row.iloc[col_i]
            if pd.notna(cell):
                try:
                    year = int(float(str(cell).strip()))
                    if 2000 <= year <= 2100:
                        col_year_map[col_i] = str(year)
                except (ValueError, TypeError):
                    continue
    return col_year_map


def _lseg_yearly_data(df, idx, col_year_map):
    out = {y: None for y in TARGET_YEARS}
    if idx is None or df is None:
        return out
    row = df.iloc[idx]
    for col_i, year in col_year_map.items():
        if year in out and col_i < len(row):
            out[year] = _lseg_safe_val(row.iloc[col_i])
    return out


def _lseg_label_value(df, label):
    """Value of a 'Label | value' metadata row (e.g., 'Standardized Currency')."""
    idx = _lseg_find_row(df, label)
    if idx is None:
        return None
    v = df.iloc[idx, 1]
    return str(v).strip() if pd.notna(v) else None


def parse_lseg_peer_data(file_bytes, filename: str = ""):
    """Parses an LSEG export file for one peer.

    Handles exports containing any combination of sheets:
    - 'Valuation'          → EV/EBITDA, P/E, EV/Revenue multiples by year
    - 'Financial Summary'  → Revenue / EBITDA / Net Income fundamentals by year
    - 'Income Statement'   → fundamentals fallback

    Sheet names are matched case-insensitively. A file without a Valuation
    sheet still parses (multiples empty, warning set) so the pipeline can
    continue with fundamentals-only data.
    """
    try:
        all_sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None)

        def find_sheet(*name_parts):
            for name, d in all_sheets.items():
                if any(p in name.lower() for p in name_parts):
                    return d
            return None

        val_df = find_sheet('valuation')
        fund_df = find_sheet('financial summary')
        if fund_df is None:
            fund_df = find_sheet('income statement')
        base_df = val_df if val_df is not None else fund_df
        if base_df is None:
            return {"error": f"No 'Valuation' or fundamentals sheet found. "
                             f"Sheets in file: {list(all_sheets.keys())}"}

        # Company name / ticker from the metadata block (same layout on all sheets)
        title_val = str(base_df.iloc[1, 1])
        if '(' in title_val:
            company_name = title_val.split('(')[0].strip()
            ticker = title_val.split('(')[-1].replace(')', '').strip()
        else:
            company_name = ""
            ticker = title_val

        # --- Multiples (Valuation sheet) ---
        if val_df is not None:
            val_years = _lseg_col_year_map(val_df)
            ev_ebitda_idx = _lseg_find_row(
                val_df, "Enterprise Value to Earnings before Interest, Taxes, Depreci",
                skip_str="5 Year Average")
            pe_idx = _lseg_find_row(
                val_df, "Price to EPS - Diluted - excluding Extraordinary Items Applicable")
            if pe_idx is None:
                pe_idx = _lseg_find_row(
                    val_df, "Price to EPS - Diluted - excluding Extraordinary Items - Nor")
            ev_rev_idx = _lseg_find_row(
                val_df, "Enterprise Value to Revenue from Business Activities - Total",
                skip_str="5 Year Average")
            ev_ebitda = _lseg_yearly_data(val_df, ev_ebitda_idx, val_years)
            pe = _lseg_yearly_data(val_df, pe_idx, val_years)
            ev_revenue = _lseg_yearly_data(val_df, ev_rev_idx, val_years)
        else:
            ev_ebitda = {y: None for y in TARGET_YEARS}
            pe = {y: None for y in TARGET_YEARS}
            ev_revenue = {y: None for y in TARGET_YEARS}

        # --- Fundamentals (Financial Summary / Income Statement sheet) ---
        revenue = {y: None for y in TARGET_YEARS}
        ebitda = {y: None for y in TARGET_YEARS}
        net_income = {y: None for y in TARGET_YEARS}
        fundamentals_currency = None
        fundamentals_scaling = None
        if fund_df is not None:
            fund_years = _lseg_col_year_map(fund_df)
            rev_idx = _lseg_find_row(fund_df, "Revenue from Business Activities - Total")
            ebitda_idx = _lseg_find_row(
                fund_df, "Earnings before Interest, Taxes, Depreci")
            ni_idx = _lseg_find_row(
                fund_df, "Income before Discontinued Operations")
            revenue = _lseg_yearly_data(fund_df, rev_idx, fund_years)
            ebitda = _lseg_yearly_data(fund_df, ebitda_idx, fund_years)
            net_income = _lseg_yearly_data(fund_df, ni_idx, fund_years)
            fundamentals_currency = _lseg_label_value(fund_df, "Standardized Currency")
            if fundamentals_currency is None:
                cur_idx = _lseg_find_row(fund_df, "Standardized Currency")
                if cur_idx is not None:
                    fundamentals_currency = str(fund_df.iloc[cur_idx, 1]).strip()
            fundamentals_scaling = _lseg_label_value(fund_df, "Scaling")

        # Derive a ticker from the filename (e.g., "D.BK.xlsx" → "D.BK")
        filename_ticker = ""
        filename_raw = ""
        if filename:
            import os
            base = os.path.splitext(filename)[0]  # "D.BK.xlsx" → "D.BK"
            if base:
                filename_ticker = base.strip().upper()
                filename_raw = base.strip()

        result = {
            "identifier": ticker,
            "filename_ticker": filename_ticker,
            "filename_raw": filename_raw,
            "company_name": company_name,
            "ev_ebitda": ev_ebitda,
            "pe": pe,
            "ev_revenue": ev_revenue,
            "revenue": revenue,
            "ebitda": ebitda,
            "net_income": net_income,
            "fundamentals_currency": fundamentals_currency,
            "fundamentals_scaling": fundamentals_scaling,
        }
        if val_df is None:
            result["warning"] = (
                "No 'Valuation' sheet in this export — trading multiples "
                "(EV/EBITDA, P/E, EV/Revenue) will be blank for this peer. "
                "Re-export from LSEG with the Valuation template included."
            )
        return result
    except Exception as e:
        return {"error": str(e)}

def parse_lseg_transactions(file_bytes):
    """Parses the LSEG M&A Precedent Transactions export file.
    Tolerates non-standard sheet names (e.g. PDF-converted files) and varying
    column headers — picks the sheet with the most rows.
    """
    try:
        all_sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
        if not all_sheets:
            return {"error": "No sheets found in file."}

        # Prefer LSEG's native sheet name; otherwise pick the largest sheet
        if "Current Screen Template" in all_sheets:
            df = all_sheets["Current Screen Template"]
        else:
            df = max(all_sheets.values(), key=lambda d: len(d))

        # Drop rows where ALL of the key value columns are missing — case-insensitive
        # match on column names so PDF-converted variants still work.
        cols_lower = {c.lower(): c for c in df.columns}
        deal_value_col = next(
            (cols_lower[c] for c in cols_lower
             if "deal value" in c or "rank value" in c),
            None,
        )
        ev_ebitda_col = next(
            (cols_lower[c] for c in cols_lower
             if "enterprise value to ebitda" in c or "ev/ebitda" in c),
            None,
        )
        date_col = next(
            (cols_lower[c] for c in cols_lower if "date announced" in c),
            None,
        )

        subset = [c for c in (deal_value_col, ev_ebitda_col) if c]
        if subset:
            df = df.dropna(subset=subset, how='all')
        if date_col:
            df[date_col] = df[date_col].astype(str)

        df = df.where(pd.notnull(df), None)
        return df.to_dict(orient='records')
    except Exception as e:
        return {"error": str(e)}
