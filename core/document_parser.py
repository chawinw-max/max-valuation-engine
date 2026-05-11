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

def parse_lseg_peer_data(file_bytes, filename: str = ""):
    """Parses an LSEG Valuation export file to extract yearly multiples.
    Dynamically reads the year header row so files with different year coverage
    (e.g. missing 2022) parse correctly.
    """
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Valuation", header=None)

        def find_row(start_str, skip_str=None):
            for i, row in df.iterrows():
                val = str(row[0]).strip()
                if val.startswith(start_str):
                    if skip_str and skip_str in val:
                        continue
                    return i
            return None

        title_val = str(df.iloc[1, 1])
        if '(' in title_val:
            company_name = title_val.split('(')[0].strip()
            ticker = title_val.split('(')[-1].replace(')', '').strip()
        else:
            company_name = ""
            ticker = title_val

        # Build a column index -> year map from the "Statement Data" header row.
        year_header_idx = find_row("Statement Data")
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

        target_years = ["2021", "2022", "2023", "2024", "2025"]

        ev_ebitda_idx = find_row("Enterprise Value to Earnings before Interest, Taxes, Depreci", skip_str="5 Year Average")
        pe_idx = find_row("Price to EPS - Diluted - excluding Extraordinary Items Applicable")
        if pe_idx is None:
            pe_idx = find_row("Price to EPS - Diluted - excluding Extraordinary Items - Nor")
        ev_rev_idx = find_row("Enterprise Value to Revenue from Business Activities - Total", skip_str="5 Year Average")

        def safe_val(v):
            if pd.isna(v):
                return None
            try:
                return float(v)
            except (ValueError, TypeError):
                return None

        def get_yearly_data(idx):
            out = {y: None for y in target_years}
            if idx is None:
                return out
            row = df.iloc[idx]
            for col_i, year in col_year_map.items():
                if year in out and col_i < len(row):
                    out[year] = safe_val(row.iloc[col_i])
            return out

        # Derive a ticker from the filename (e.g., "D.BK.xlsx" → "D.BK")
        filename_ticker = ""
        if filename:
            import os
            base = os.path.splitext(filename)[0]  # "D.BK.xlsx" → "D.BK"
            if base:
                filename_ticker = base.strip().upper()

        return {
            "identifier": ticker,
            "filename_ticker": filename_ticker,
            "company_name": company_name,
            "ev_ebitda": get_yearly_data(ev_ebitda_idx),
            "pe": get_yearly_data(pe_idx),
            "ev_revenue": get_yearly_data(ev_rev_idx)
        }
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
