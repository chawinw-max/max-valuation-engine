import json
from google import genai
from google.genai import types
import yfinance as yf
import config


_GEOGRAPHY_SCORES = {
    "Thailand": 100,
    "Singapore": 70, "Malaysia": 70, "Indonesia": 70,
    "Philippines": 70, "Vietnam": 70,
    "Japan": 40, "South Korea": 40, "Hong Kong": 40,
    "Taiwan": 40, "Australia": 40, "New Zealand": 40,
    "India": 40, "China": 40,
}


def _score_country(country: str) -> int:
    if not country:
        return 10
    return _GEOGRAPHY_SCORES.get(country, 10)


def _peer_record(identifier, company_name, country, business_description,
                 market_cap_thb_m, trbc_activity, source):
    geo = _score_country(country)
    if geo == 100:
        ring = 2
    elif geo == 70:
        ring = 2
    else:
        ring = 3
    return {
        "identifier": identifier.upper(),
        "company_name": company_name or "",
        "trbc_activity": trbc_activity or "Manually Added",
        "country": country or "",
        "business_description": business_description or "",
        "market_cap_thb_m": market_cap_thb_m,
        "ebitda_positive": True,
        "fit_rank": 2,
        "geography_score": geo,
        "ring": ring,
        "ring_justification": f"Manually added via {source} lookup.",
        "scale_warning": None,
    }

# Lazy-initialized GenAI Client (st.secrets isn't available at import time on Streamlit Cloud)
_client = None

def _get_client():
    global _client
    if _client is None:
        _client = genai.Client(api_key=config.GEMINI_API_KEY)
    return _client

def test_api():
    """Simple test to verify the API key is working with the correct model."""
    response = _get_client().models.generate_content(
        model='gemini-2.5-flash',
        contents='Say "API is working"'
    )
    return response.text

import mimetypes
from core.document_parser import extract_text_from_file

def _file_to_part(f):
    filename = f.name.lower()
    file_bytes = f.read()
    f.seek(0) # Reset pointer
    
    if filename.endswith('.pdf'):
        return types.Part.from_bytes(data=file_bytes, mime_type='application/pdf')
    elif filename.endswith(('.png', '.jpg', '.jpeg')):
        mime_type, _ = mimetypes.guess_type(filename)
        return types.Part.from_bytes(data=file_bytes, mime_type=mime_type or 'image/jpeg')
    else:
        text = extract_text_from_file(f)
        return f"--- {f.name} ---\n{text}"

def extract_financials_and_business_model(company_files, financial_files, available_years, notes: str = "", business_type: str = "Service"):
    """
    Phase 1 AI logic: Extracts financials and business model from provided files.
    """
    
    system_prompt = """
    You are a Senior M&A Analyst at MAX Solutions, a Thai M&A advisory firm. You specialize in Comparable Company Analysis (Comps).

    STRICT RULES:
    1. NO HALLUCINATIONS: Only use data from the documents provided. Never invent financial figures. IF YOU CANNOT FIND THE EXACT NUMBER IN THE DOCUMENT, RETURN null.
    2. CURRENCY: All financial figures in Thai Baht (THB). If source uses another currency, convert and state the rate used.
    3. RESPOND ONLY IN VALID JSON matching the schema provided. No markdown, no commentary outside the JSON.
    4. COPY ALL NUMBERS EXACTLY as they appear in source documents. Do NOT round, reformat, or convert. Preserve ALL decimal places down to the exact cent.

    THAI FINANCIAL STATEMENT SUPPORT:
    You can read Thai-language financial documents. Key term mappings:
    - รายได้จากการขายและบริการ = Sales and Services Revenue
    - รายได้อื่น = Other Revenues
    - ต้นทุนขาย / ต้นทุนการให้บริการ = Cost of Goods Sold
    - ค่าใช้จ่ายในการขาย = Sales Expenses
    - ค่าใช้จ่ายในการบริหาร = Administrative Expenses
    - ค่าใช้จ่ายอื่น = Other Expenses
    - ค่าเสื่อมราคาและค่าตัดจำหน่าย = Depreciation and Amortization
    - ต้นทุนทางการเงิน / ดอกเบี้ยจ่าย = Finance Costs / Interest Expenses
    - ภาษีเงินได้ = Income Tax
    - กำไรสุทธิ = Net Profit
    - กำไรขั้นต้น = Gross Profit
    - กำไรจากการดำเนินงาน = Operating Profit

    IMPORTANT FOR THAI DOCUMENTS:
    - Thai fiscal years often use Buddhist Era (พ.ศ.). Convert: พ.ศ. 2565 = 2022 CE, พ.ศ. 2566 = 2023, พ.ศ. 2567 = 2024.
    - Numbers may use Thai digits (๐-๙) — convert to Arabic numerals.
    - Some P&L statements combine "Sales and Administrative Expenses" (ค่าใช้จ่ายในการขายและบริหาร) — split proportionally or assign to Administrative and note it.
    """

    prompt = f"""
    You are given two groups of documents (images, PDFs, and text):

    AVAILABLE YEARS: {available_years}

    TASK 1 — From the COMPANY INFORMATION DOCUMENTS, extract:
    1. deal_code: The deal code (e.g., "DF-117"). If not found, use "DF-XXX".
    2. client_name: The client company name.
    3. client_overview: A 1-2 sentence summary of what they do, who they serve, how big they are.
    4. business_attributes:
       - specific_subsector: Be SPECIFIC (e.g., "Fresh produce B2B supply to HORECA" NOT "food company")
       - channel: (e.g., "B2B wholesale direct to 4-5 star hotels" NOT "wholesale")
       - model_type: (e.g., "Trader/distributor with GMP-certified light processing")
       - operating_geography: (e.g., "Thailand (Bangkok/Nonthaburi)")
    5. business_model: Extract ALL fields for Sections A through F of the Business Model page. See schema below.

    TASK 2 — From the FINANCIAL STATEMENT DOCUMENTS, extract the 9 financial line items.

    These 9 line items feed a P&L waterfall template. The template uses formulas to compute
    derived rows from your 9 inputs. You MUST extract values that fit this waterfall correctly:

    ┌─ REVENUE ──────────────────────────────────────────────────────────┐
    │  Row 6: sales_and_services      ← Top-line operating revenue      │
    │  Row 7: other_revenues          ← Non-operating / misc income     │
    │  Row 8: TOTAL REVENUE = Row 6 + Row 7          [FORMULA]          │
    ├─ COST OF SALES ────────────────────────────────────────────────────┤
    │  Row 11: cost_of_goods_sold     ← Direct costs of production      │
    │  Row 12: GROSS PROFIT = Row 8 − Row 11         [FORMULA]          │
    ├─ OPERATING EXPENSES ───────────────────────────────────────────────┤
    │  Row 18: sales_expenses         ← Selling & distribution costs    │
    │  Row 19: administrative_expenses← G&A, office, salaries, etc.     │
    │  Row 20: other_expenses         ← Non-recurring or misc OpEx      │
    │  Row 21: EBITDA = Row 12 − (Row 18+19+20)      [FORMULA]          │
    ├─ BELOW EBITDA ─────────────────────────────────────────────────────┤
    │  Row 26: depreciation_amortization ← D&A (below EBITDA)           │
    │  Row 27: EBIT = Row 21 − Row 26                [FORMULA]          │
    │  Row 30: interest_expenses      ← Finance costs / debt service    │
    │  Row 31: PBT = Row 27 − Row 30                 [FORMULA]          │
    │  Row 34: tax                    ← Income tax expense              │
    │  Row 35: NET PROFIT = Row 31 − Row 34           [FORMULA]          │
    └────────────────────────────────────────────────────────────────────┘

    SIGN CONVENTION: ALL 9 values must be POSITIVE numbers (absolute values).
    The template SUBTRACTS costs automatically via formulas. Do NOT negate expenses.
    Example: If COGS is shown as (50,000) or -50,000 in the source, extract as 50000.

    CRITICAL EXTRACTION RULES:
    1. SANITY CHECK — After extraction, mentally verify for each year:
       - Gross Profit = Revenue − COGS → should be positive for a viable business
       - EBITDA = Gross Profit − OpEx → check this is reasonable (typically 5–30% of revenue)
       - If your extracted numbers produce negative EBITDA or >90% gross margin, re-read the source
    2. COGS vs GROSS PROFIT — Some Thai P&L statements show กำไรขั้นต้น (Gross Profit) directly
       instead of ต้นทุนขาย (COGS). If only Gross Profit is shown, BACK-CALCULATE:
       cost_of_goods_sold = Total Revenue − Gross Profit
       IMPORTANT: Mark in _verification that COGS was "derived" (not directly extracted).
    3. D&A PLACEMENT — D&A goes in Row 26 (below EBITDA), NOT inside admin expenses.
       If the source buries D&A inside admin or operating expenses, you must SEPARATE it:
       - Extract the D&A amount into depreciation_amortization
       - Subtract that amount from administrative_expenses to avoid double-counting
       If D&A cannot be separated (not disclosed), set depreciation_amortization = null

       ⚠️  CRITICAL D&A + DERIVED COGS INTERACTION:
       If COGS was derived using Rule #2 (back-calculated from Gross Profit), the derived
       COGS already embeds any D&A that the source included in its Cost of Sales (e.g.,
       ค่าเสื่อมราคา-เครื่องมือทันตกรรม from Note 10). You CANNOT separate that D&A out of
       the derived COGS because the derivation was: COGS = Revenue − Gross Profit.
       In this situation:
       - Still separate D&A from administrative_expenses (subtract from admin, add to D&A)
       - Do NOT add COGS-embedded D&A to the depreciation_amortization line
       - The depreciation_amortization value should ONLY contain D&A extracted from admin/opex
       - If no D&A can be separated from admin, set depreciation_amortization = null
       - Note in _verification which D&A components are embedded vs separated
    4. COMBINED SG&A — If the source combines Selling + Admin into one line:
       - Put the full amount in administrative_expenses
       - Set sales_expenses = 0
    5. OTHER REVENUES vs OTHER EXPENSES — Do not confuse them:
       - other_revenues (Row 7): income items like interest income, FX gains, dividend income
       - other_expenses (Row 20): non-recurring charges, write-downs, loss on disposal
       If no separate "Other Expenses" line exists, set other_expenses = 0
    6. INTEREST EXPENSES — Extract ONLY interest/finance costs on debt. Do NOT include:
       - Lease liability interest (IFRS 16) unless that's the only interest figure shown
       - Interest income (that goes in other_revenues)
    7. TAX — Extract the income tax expense line. If the source shows "tax benefit" (negative),
       extract as a negative number (this is the one exception to the positive-sign rule).
    8. UNITS — Pay attention to whether figures are in units, thousands, or millions.
       Thai audited statements are often in Thai Baht (units). Internal summaries may use thousands.
       The header/footnote usually states the unit. Convert everything to UNITS (actual THB).
    9. MULTI-BRANCH / SUB-ROW AGGREGATION — Many Thai SME spreadsheets break expenses by branch
       (e.g., ค่าเช่า สาขารามคำแหง + ค่าเช่า สาขาพลัส). Rules:
       - Sum all branch sub-rows into ONE total per expense category per month/year
       - Do NOT count the parent-level total AND the child-level breakdown — pick one
       - If the spreadsheet shows both a per-branch breakdown AND a pre-summed total row,
         use the pre-summed total. If only sub-rows exist, sum them yourself.
    10. ONE-TIME / ANNUAL ITEMS — Some expenses appear only in one month (e.g., วัสดุสิ้นเปลือง
        / รายปี in December only). When the source is a monthly spreadsheet:
        - Sum the 12 monthly columns to get the annual total — do NOT multiply a single month by 12
        - If a line item has a value in only 1 month, that IS the annual total for that item
    11. COGS CLASSIFICATION — The analyst has specified this is a **{business_type}** business.
        Apply COGS rules strictly based on this business type:

        IF business_type == "Service" (clinic, salon, consulting, restaurant, professional services):
          COGS includes ALL direct costs of delivering the service:
          - Facility rent at service locations (ค่าเช่า at each branch)
          - Staff salaries for service delivery personnel (เงินเดือนพนักงาน)
          - Practitioner/professional fees (e.g., dentist fees / DF, consultant fees)
          - Utilities at service locations (ค่าไฟฟ้า, ค่าน้ำ)
          - Materials and lab costs consumed in service delivery (ค่า Lab, วัสดุ)
          - Placement/licensing fees for practitioners (วางใบ)
          COGS does NOT include: credit card processing fees, social security, parking, phone/comms,
          annual consumables unrelated to service delivery — those go to administrative_expenses.

        IF business_type == "Manufacturing":
          COGS includes: raw materials, direct production labor, factory overhead, factory rent,
          factory utilities, production equipment maintenance.
          COGS does NOT include: office rent, admin salaries, sales team costs — those are admin/sales.

        IF business_type == "Trading / Distribution":
          COGS includes: purchase cost of goods resold, warehouse rent, logistics/freight costs,
          warehouse labor.
          COGS does NOT include: office rent, admin salaries, sales commissions — those are admin/sales.

        CRITICAL: Follow the business type classification EXACTLY. Do NOT override with your own
        judgment about what "should" be COGS. The analyst has made this decision.
    12. MANDATORY CROSS-VALIDATION — After extracting all 9 line items, perform this check:
        a) Find the source document's stated Net Profit (กำไรสุทธิ / กำไรหลังหักค่าใช้จ่าย)
           or Gross Profit (กำไรขั้นต้น) if available.
        b) Compute: Implied Expenses = Revenue − Source Net Profit
        c) Compare against: Extracted Expenses = COGS + Sales + Admin + Other + D&A + Interest + Tax
        d) If the gap exceeds 2%% of revenue, you have a double-counting or misclassification error.
           Re-examine which line items overlap and fix before returning.
        Report the source-stated figures in the _verification field (see schema).

    CRITICAL CONSTRAINTS:
    - Extract data ONLY for years in available_years: {available_years}
    - If a value is not found for an available year, use null (not 0, not omit the key)
    - Every line item MUST contain a key for EVERY year in available_years
    - Copy all financial figures EXACTLY as they appear. Preserve ALL decimal places down to the exact cent.
    - Do NOT estimate, interpolate, or generate fake data for any year NOT in available_years
    - If documents are in Thai, convert Buddhist Era years: พ.ศ. 2565=2022, 2566=2023, 2567=2024

    Return ONLY valid JSON matching this schema:
    {{
      "deal_code": "DF-117",
      "client_name": "Company Name",
      "client_overview": "1-2 sentence overview...",
      "business_attributes": {{
        "specific_subsector": "...",
        "channel": "...",
        "model_type": "...",
        "operating_geography": "..."
      }},
      "financials": {{
        "sales_and_services": {{"2022": 1234567.89, "2023": 1234567.89}},
        "other_revenues": {{"2022": 1234567.89, "2023": 1234567.89}},
        "cost_of_goods_sold": {{"2022": 1234567.89, "2023": 1234567.89}},
        "sales_expenses": {{"2022": 1234567.89, "2023": 1234567.89}},
        "administrative_expenses": {{"2022": 1234567.89, "2023": 1234567.89}},
        "other_expenses": {{"2022": 1234567.89, "2023": 1234567.89}},
        "depreciation_amortization": {{"2022": 1234567.89, "2023": 1234567.89}},
        "interest_expenses": {{"2022": 1234567.89, "2023": 1234567.89}},
        "tax": {{"2022": 1234567.89, "2023": 1234567.89}}
      }},
      "_verification": {{
        "source_net_profit": {{"2022": null, "2023": null}},
        "source_gross_profit": {{"2022": null, "2023": null}},
        "notes": "Any observations about data quality, ambiguous classifications, or items that need owner confirmation"
      }},
      "business_model": {{
        "section_a": {{
          "company_name": "...", "founded": "...", "type": "...", "registered_capital": "...",
          "shareholders": "...", "location": "...", "focus": "...", "certifications": "...",
          "operating_history": "...", "website": "..."
        }},
        "section_b": {{
          "primary_revenue": "...", "key_clients": "...", "pricing_strategy": "...",
          "customer_mix": "...", "client_segmentation": "...", "credit_terms": "...",
          "secondary_revenue": "..."
        }},
        "section_c": [
          {{"title": "1. Moat Title", "description": "Explanation..."}}
        ],
        "section_d": {{
          "revenue_fy2022": "...", "revenue_fy2023": "...", "revenue_fy2024": "...",
          "revenue_fy2025": "...", "gross_margin": "...", "ebitda_normalization": "...",
          "liabilities": "...", "cash_flow_note": "..."
        }},
        "section_e": {{
          "seller": "...", "exit_motivation": "...", "asking_price": "...",
          "property": "...", "transaction_type": "...", "next_steps": "..."
        }},
        "section_f": {{
          "total_staff": "...", "shift_structure": "...", "inbound_process": "...",
          "facility": "...", "machinery": "...", "logistics_fleet": "...", "it_systems": "..."
        }}
      }}
    }}
    """
    
    notes_block = f"\n\nANALYST NOTES (treat as clarification context):\n{notes.strip()}" if notes and notes.strip() else ""
    contents = [
        system_prompt,
        prompt + notes_block,
        "=== COMPANY INFORMATION DOCUMENTS ==="
    ]
    for f in company_files:
        contents.append(_file_to_part(f))
        
    contents.append("=== FINANCIAL STATEMENT DOCUMENTS ===")
    for f in financial_files:
        contents.append(_file_to_part(f))

    response = _get_client().models.generate_content(
        model='gemini-2.5-flash',
        contents=contents,
        config=types.GenerateContentConfig(
            temperature=0.0,
            response_mime_type="application/json",
        )
    )

    result = json.loads(response.text)
    result['_business_type'] = business_type
    _post_process_financials(result, available_years)
    return result


def _post_process_financials(result: dict, available_years: list):
    """
    Fix common extraction errors in-place:
    - Flip negative expenses to positive (sign convention)
    - Back-calculate COGS from source gross profit if available
    - Cross-validate total expenses against source net profit
    """
    fin = result.get('financials')
    if not fin or not isinstance(fin, dict):
        return

    verification = result.get('_verification') or {}
    source_gp = verification.get('source_gross_profit') or {}
    source_np = verification.get('source_net_profit') or {}

    expense_keys = [
        'cost_of_goods_sold', 'sales_expenses', 'administrative_expenses',
        'other_expenses', 'depreciation_amortization', 'interest_expenses',
    ]

    warnings = []

    for year in available_years:
        y = str(year)

        # 1. Flip negative expenses to positive (sign convention fix)
        for key in expense_keys:
            val = (fin.get(key) or {}).get(y)
            if val is not None and val < 0:
                fin[key][y] = abs(val)

        sales = (fin.get('sales_and_services') or {}).get(y) or 0
        other = (fin.get('other_revenues') or {}).get(y) or 0
        total_rev = sales + other
        cogs = (fin.get('cost_of_goods_sold') or {}).get(y)

        # 2. If source states Gross Profit, use it to fix COGS
        sgp = source_gp.get(y)
        if sgp is not None and total_rev > 0:
            expected_cogs = total_rev - sgp
            if expected_cogs > 0 and cogs is not None:
                gap_pct = abs(cogs - expected_cogs) / total_rev * 100
                if gap_pct > 2:
                    fin['cost_of_goods_sold'][y] = expected_cogs
                    warnings.append(
                        f"{y}: COGS corrected from {cogs:,.0f} to {expected_cogs:,.0f} "
                        f"(derived from source Gross Profit {sgp:,.0f})"
                    )
        elif total_rev > 0 and cogs is not None:
            # Fallback: if "COGS" produces >95% gross margin, it's likely Gross Profit not COGS
            gross_margin = (total_rev - cogs) / total_rev
            if gross_margin > 0.95 and cogs < total_rev * 0.5:
                fin['cost_of_goods_sold'][y] = total_rev - cogs

        # 3. Cross-validate total expenses against source net profit
        snp = source_np.get(y)
        if snp is not None and total_rev > 0:
            implied_total_expenses = total_rev - snp
            extracted_total = sum(
                (fin.get(k) or {}).get(y) or 0 for k in expense_keys
            )
            # Also add tax
            extracted_total += (fin.get('tax') or {}).get(y) or 0

            gap = extracted_total - implied_total_expenses
            gap_pct = abs(gap) / total_rev * 100
            if gap_pct > 2:
                warnings.append(
                    f"{y}: Extracted expenses ({extracted_total:,.0f}) exceed implied expenses "
                    f"({implied_total_expenses:,.0f}) by {gap:,.0f} THB ({gap_pct:.1f}% of revenue). "
                    f"Likely double-counting or misclassification."
                )

    if warnings:
        existing = verification.get('notes') or ''
        auto_notes = "AUTO-DETECTED ISSUES:\n" + "\n".join(f"- {w}" for w in warnings)
        verification['notes'] = (existing + "\n\n" + auto_notes).strip() if existing else auto_notes
        result['_verification'] = verification


import re

def _clean_json(text):
    text = text.strip()
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    # Remove trailing commas (e.g. `},]`)
    text = re.sub(r',\s*([\]}])', r'\1', text)
    return text.strip()

def _compute_target_metrics(financials: dict, available_years: list) -> dict:
    """Compute key financial metrics from Phase 1 data for peer screening context."""
    if not financials or not available_years:
        return {}

    sorted_years = sorted(available_years)
    latest = str(sorted_years[-1])

    def _get_val(key, year):
        return (financials.get(key) or {}).get(year) or 0

    sales = _get_val('sales_and_services', latest)
    other_rev = _get_val('other_revenues', latest)
    cogs = _get_val('cost_of_goods_sold', latest)
    s_exp = _get_val('sales_expenses', latest)
    admin = _get_val('administrative_expenses', latest)
    o_exp = _get_val('other_expenses', latest)

    total_rev = sales + other_rev
    gross_profit = total_rev - cogs
    ebitda = gross_profit - s_exp - admin - o_exp

    metrics = {
        'latest_year': int(sorted_years[-1]),
        'total_revenue': total_rev,
        'gross_profit': gross_profit,
        'gross_margin_pct': round(gross_profit / total_rev * 100, 1) if total_rev else None,
        'ebitda': ebitda,
        'ebitda_margin_pct': round(ebitda / total_rev * 100, 1) if total_rev else None,
    }

    if len(sorted_years) >= 2:
        prev = str(sorted_years[-2])
        prev_total = _get_val('sales_and_services', prev) + _get_val('other_revenues', prev)
        if prev_total:
            metrics['revenue_growth_pct'] = round((total_rev - prev_total) / prev_total * 100, 1)

    return metrics


def generate_peer_list(client_overview, business_attributes, latest_year_revenue,
                       notes: str = "", financials: dict = None, available_years: list = None):
    """
    Phase 2 AI logic: Generates a list of up to 30 comparable companies based on the Ring framework.
    """
    system_prompt = """
    You are screening publicly listed comparable companies for a Comparable Company Valuation.

    STRICT RULES:
    1. POSITIVE EBITDA RULE: Exclude any company with negative EBITDA from peer selection.
    2. RESPOND ONLY IN VALID JSON matching the schema provided.
    3. REAL COMPANIES ONLY: You must exhaust your knowledge base to find REAL public companies (e.g. on SET, mai, SGX). Generate as many as you can confidently verify (up to 30).
    4. NO HALLUCINATIONS: Every single company and ticker you list MUST be a real, verifiable publicly traded company. DO NOT invent, mock, or fake any tickers. If you only know 12 real companies, output 12.
    5. ACCURATE NAMES: Cross-check your memory. Ensure the `company_name` perfectly matches the ticker `identifier`. Do not mix them up (e.g., SUN.BK is Sunsweet PCL, not Siam Agro Food).
    6. MARKET CAP IN MILLIONS: Output the market cap in millions of THB (e.g., 1800 for 1.8 billion THB).
    """

    target_metrics = _compute_target_metrics(financials or {}, available_years or [])
    financial_profile = ""
    if target_metrics:
        fy = target_metrics.get('latest_year', '?')
        lines = [f"\n    ## TARGET COMPANY FINANCIAL PROFILE (FY{fy})"]
        tr = target_metrics.get('total_revenue')
        if tr:
            lines.append(f"    - Total Revenue: {tr:,.0f} THB")
        rg = target_metrics.get('revenue_growth_pct')
        if rg is not None:
            lines.append(f"    - Revenue Growth (YoY): {rg:+.1f}%")
        gm = target_metrics.get('gross_margin_pct')
        if gm is not None:
            lines.append(f"    - Gross Margin: {gm:.1f}%")
        eb = target_metrics.get('ebitda')
        if eb:
            lines.append(f"    - EBITDA: {eb:,.0f} THB")
        em = target_metrics.get('ebitda_margin_pct')
        if em is not None:
            lines.append(f"    - EBITDA Margin: {em:.1f}%")
        lines.append("")
        lines.append("    Use these metrics to assess SCALE PROXIMITY and PROFITABILITY FIT.")
        lines.append("    Prefer peers with EBITDA margins within ±15pp of target.")
        financial_profile = "\n".join(lines)

    prompt = f"""
    ## TARGET COMPANY PROFILE
    {client_overview}

    Key attributes:
    - Sub-sector: {business_attributes.get('specific_subsector', '')}
    - Channel: {business_attributes.get('channel', '')}
    - Revenue scale: {latest_year_revenue} THB/year
    - Geography: {business_attributes.get('operating_geography', '')}
    - Business model: {business_attributes.get('model_type', '')}
    {financial_profile}

    ## MATCHING RULES — STRICT HIERARCHY

    ### Rule 1: Sub-Sector Specificity (MOST IMPORTANT)
    BEST MATCH (Ring 1): Exact same product + same channel
    ACCEPTABLE (Ring 2): Same product category + adjacent channel
    WEAK (Ring 3): Adjacent product + same channel
    REJECT (Ring 4+): Same broad industry but fundamentally different business — DO NOT INCLUDE

    CRITICAL SUB-SECTOR ALIGNMENT: Match the target's `specific_subsector` at the NARROWEST level, not the broad industry.
    - STRICT REJECTION RULE: If a company operates in the same broad industry (e.g., "Food & Beverage") but serves a fundamentally different product category (e.g., target sells fresh vegetables but candidate processes meat/poultry/seafood), it is Ring 4 — REJECT, regardless of geography or scale.

      Test: "Would a buyer acquiring the target consider this peer's revenue mix, supply chain, and
      customer base to be a meaningful pricing reference?" If no → REJECT.

    - WHEN THAI MATCHES ARE EXHAUSTED: Do NOT relax sub-sector criteria to fill the list.
      A Malaysian fresh produce distributor is a better comp than a Thai meat processor.
      Expand GEOGRAPHY before expanding PRODUCT SCOPE. Target 20–30 companies across regions.


    ### Rule 2: Scale Proximity
    Revenue range: 0.2x to 5x of target. Market cap > 50x target's implied value = REJECT.

    ### Rule 3: Geography Hierarchy (HARD SCORING)
    | Priority | Geography | Score |
    |----------|-----------|-------|
    | 1st | Thailand (SET/mai) | 100 |
    | 2nd | ASEAN (SGX, IDX, BEX, PSE, HOSE) | 70 |
    | 3rd | APAC (ASX, NZX, TSE, KRX, HKEX, BSE/NSE) | 40 |
    | 4th | Global (LSE, NYSE, NASDAQ, etc.) | 10 |

    CRITICAL INSTRUCTION FOR 30 COMPANIES:
    If you cannot find 30 highly relevant companies in Thailand (1st Priority), you MUST expand your search to ASEAN (2nd Priority) and APAC (3rd Priority) until you reach 30 companies. Do NOT stop at 10 companies just because you ran out of Thai companies.

    A Thai Ring-2 peer ALWAYS ranks above a US Ring-1 peer.

    ## REAL COMPANY FORMAT EXAMPLES (Do not use unless relevant to target's sub-sector):
    - Tech/Telecom: ADVANC.BK, TRUE.BK, GULF.BK
    - Retail/Commerce: CPALL.BK, CRC.BK, COM7.BK
    - Healthcare: BDMS.BK, BH.BK, BCH.BK
    - Food/Agri: CPF.BK, TU.BK, CBG.BK, SUN.BK

    CRITICAL: You must find REAL companies that match the target's specific sub-sector. Do not invent fake tickers (e.g., do not output "FRESH.SET" or "TECH.BK"). If you only know 10 real relevant companies, return 10.

    ## OUTPUT: JSON array of up to 30 REAL companies, sorted by geography_score DESC, ring ASC, fit_rank ASC.
    {{
      "broad_list": [
        {{
          "identifier": "TICKER.EXCHANGE",
          "company_name": "Full legal name",
          "trbc_activity": "TRBC Activity Name",
          "country": "Country",
          "business_description": "2-3 specific sentences",
          "market_cap_thb_m": 1800,
          "ebitda_positive": true,
          "fit_rank": 1,
          "geography_score": 100,
          "ring": 1,
          "ring_justification": "1 sentence",
          "scale_warning": null
        }}
      ]
    }}
    """

    notes_block = f"\n\nANALYST NOTES (treat as clarification context):\n{notes.strip()}" if notes and notes.strip() else ""
    response = _get_client().models.generate_content(
        model='gemini-2.5-flash',
        contents=[system_prompt, prompt + notes_block],
        config=types.GenerateContentConfig(
            temperature=0.2,
            response_mime_type="application/json",
            max_output_tokens=65536,
            thinking_config=types.ThinkingConfig(thinking_budget=8192),
        )
    )

    raw_text = response.text or ""

    # Check if output was truncated (finish_reason != STOP)
    finish_reason = None
    try:
        finish_reason = response.candidates[0].finish_reason
    except Exception:
        pass

    try:
        result = json.loads(_clean_json(raw_text))
    except json.JSONDecodeError as e:
        raise RuntimeError(
            f"Peer list JSON parse failed (finish_reason={finish_reason}, "
            f"output_len={len(raw_text)} chars): {e}. "
            f"Last 200 chars of output: ...{raw_text[-200:]!r}"
        ) from e

    # Handle empty or missing broad_list
    peers_found = len(result.get('broad_list', []))
    if peers_found == 0:
        # Gemini returned valid JSON but no peers — check if the list is nested differently
        # Some responses wrap it as {"peers": [...]} or just a raw list
        if isinstance(result, list):
            result = {"broad_list": result}
            peers_found = len(result["broad_list"])
        else:
            # Try common alternative keys
            for alt_key in ["peers", "companies", "comparable_companies", "results"]:
                if alt_key in result and isinstance(result[alt_key], list) and result[alt_key]:
                    result["broad_list"] = result[alt_key]
                    peers_found = len(result["broad_list"])
                    break

    if peers_found == 0:
        raise RuntimeError(
            f"Gemini returned 0 peers. This may happen with very niche businesses. "
            f"Try broadening the sub-sector in analyst notes (e.g., 'expand to healthcare services, "
            f"outpatient clinics, aesthetics chains in ASEAN')."
        )

    return result

def verify_peer_list(peers: list, progress_callback=None) -> list:
    """
    Batch-validate AI-generated peer tickers via yfinance.
    For verified .BK tickers, overwrites market_cap_thb_m with real data.
    Updates company_name and trbc_activity with authoritative yfinance data.
    Returns the list with a 'verified' bool column added to each record.
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    def _verify_one(peer: dict) -> dict:
        record = {**peer}
        ticker = peer.get('identifier', '')

        for variant in _yfinance_ticker_variants(ticker):
            try:
                info = yf.Ticker(variant).info or {}
            except Exception:
                continue
            name = info.get('longName') or info.get('shortName')
            if name:
                record['verified'] = True
                record['company_name'] = name
                record['identifier'] = variant.upper()
                mcap = info.get('marketCap')
                if mcap and variant.upper().endswith('.BK'):
                    record['market_cap_thb_m'] = round(mcap / 1_000_000, 0)
                industry = info.get('industry')
                if industry:
                    record['trbc_activity'] = industry
                return record

        # Could not verify — mark as unverified but DON'T discard
        record['verified'] = False
        return record

    total = len(peers)
    with ThreadPoolExecutor(max_workers=4) as executor:  # Reduced from 8 to avoid rate limits
        futures = {executor.submit(_verify_one, p): i for i, p in enumerate(peers)}
        results = [None] * total
        done_count = 0
        for future in as_completed(futures):
            idx = futures[future]
            try:
                results[idx] = future.result()
            except Exception:
                results[idx] = {**peers[idx], 'verified': False}
            done_count += 1
            if progress_callback:
                progress_callback(done_count, total)

    return results


def generate_rejection_rationales(client_overview, not_selected_peers):
    """
    Phase 2 AI logic: Generates short rejection rationales for peers that were not selected.
    """
    peers_json = json.dumps(not_selected_peers, indent=2)
    prompt = f"""
    TARGET COMPANY: {client_overview}
    
    The user decided NOT to select the following companies for the Comparable Company Valuation.
    For each company, provide a very short 3-5 word rejection reason (e.g., "Different product focus", "Too large scale", "Different geography").
    
    UNSELECTED PEERS:
    {peers_json}
    
    Return ONLY JSON:
    {{
      "rationales": {{
        "TICKER.EXCHANGE": "Reason here",
        ...
      }}
    }}
    """

    response = _get_client().models.generate_content(
        model='gemini-2.5-flash',
        contents=[prompt],
        config=types.GenerateContentConfig(
            temperature=0.0,
            response_mime_type="application/json",
            max_output_tokens=32768,
            thinking_config=types.ThinkingConfig(thinking_budget=0),
        )
    )
    return json.loads(response.text)

def generate_deep_dive(client_overview, selected_peers, notes: str = ""):
    """
    Phase 3 AI logic: Generates qualitative and quantitative comparisons for selected peers.
    """
    peers_json = json.dumps(selected_peers, indent=2)
    prompt = f"""
    TARGET COMPANY: {client_overview}
    
    For each selected peer, write a qualitative comparison to the target company.
    Then, provide Revenue and EBITDA for 2023, 2024, and 2025.
    All figures in MILLIONS THB. Convert from other currencies if needed, stating the rate.
    If 2025 is unavailable, use LTM or null.
    Do NOT include Country — it's auto-calculated by the template.
    
    SELECTED PEERS:
    {peers_json}
    
    Return JSON:
    {{
      "qualitative": [
        {{
          "identifier": "TICKER",
          "company_name": "Name",
          "core_business_model": "2-3 sentences",
          "product_focus": "main products",
          "similarity": "High/Moderate/Low + explanation",
          "comparison_points": "what makes them a valid comp",
          "differentiation_points": "key differences"
        }}
      ],
      "financials_comparison": [
        {{
          "identifier": "TICKER",
          "revenue_2023": 95.5, "revenue_2024": 102.3, "revenue_2025": 110.0,
          "ebitda_2023": 12.1, "ebitda_2024": 14.5, "ebitda_2025": 16.2
        }}
      ]
    }}
    """

    notes_block = f"\n\nANALYST NOTES (treat as clarification context):\n{notes.strip()}" if notes and notes.strip() else ""
    response = _get_client().models.generate_content(
        model='gemini-2.5-flash',
        contents=[prompt + notes_block],
        config=types.GenerateContentConfig(
            temperature=0.0,
            response_mime_type="application/json",
            max_output_tokens=32768,
            thinking_config=types.ThinkingConfig(thinking_budget=0),
        )
    )
    return json.loads(response.text)

def select_precedent_transactions(client_data, parsed_transactions, notes: str = ""):
    """
    Phase 3.5 AI logic: Selects the top 10 best-fit M&A transactions.
    """
    client_name = client_data.get('client_name', '')
    client_overview = client_data.get('client_overview', '')
    specific_subsector = client_data.get('business_attributes', {}).get('specific_subsector', '')
    operating_geography = client_data.get('business_attributes', {}).get('operating_geography', '')
    deal_code = client_data.get('deal_code', 'Target')
    
    # We slice to first 50 transactions to avoid context length explosion
    tx_json = json.dumps(parsed_transactions[:50], indent=2)
    
    prompt = f"""
    You are selecting the most relevant precedent M&A transactions for a comparable valuation.

    TARGET COMPANY: {client_name}
    {client_overview}
    Sub-sector: {specific_subsector}
    Geography: {operating_geography}

    Below is a list of M&A transactions from LSEG. Select the TOP 10 most relevant transactions.

    SELECTION CRITERIA (in priority order):
    1. REGION: Thailand > ASEAN > APAC > Global
    2. INDUSTRY RELEVANCE: How closely the target company's business matches the subject
    3. DEAL SIZE: Prefer transactions in comparable size range
    4. DATA COMPLETENESS: MUST have both Deal Value and EV/EBITDA available — skip transactions missing either
    5. RECENCY: More recent transactions preferred

    TRANSACTIONS:
    {tx_json}

    For each selected transaction, generate:
    - relevance: 1-2 sentences explaining why this transaction is relevant to {deal_code}
    - notes_and_caveats: Key considerations (size mismatch, strategic premium, distressed, etc.)

    Return JSON:
    {{
      "selected_transactions": [
        {{
          "target": "Target Full Name",
          "acquirer": "Acquirer Full Name",
          "date": "YYYY",
          "region": "Country",
          "relevance": "Why relevant to our target...",
          "deal_value_usd_m": 50.0,
          "ev_ebitda": 8.5,
          "notes_and_caveats": "Key considerations..."
        }}
      ]
    }}
    """

    notes_block = f"\n\nANALYST NOTES (treat as clarification context):\n{notes.strip()}" if notes and notes.strip() else ""
    response = _get_client().models.generate_content(
        model='gemini-2.5-flash',
        contents=[prompt + notes_block],
        config=types.GenerateContentConfig(
            temperature=0.0,
            response_mime_type="application/json",
            max_output_tokens=32768,
            thinking_config=types.ThinkingConfig(thinking_budget=0),
        )
    )
    return json.loads(response.text)


def _yfinance_ticker_variants(raw: str):
    """
    Yield ticker variants to try on Yahoo. Handles LSEG/Refinitiv sub-board suffixes
    (e.g., NTSCm.BK for mai board, BBLf.BK for foreign board) which Yahoo doesn't use.
    """
    raw = raw.strip()
    if not raw:
        return
    # 1. As entered (uppercased)
    yield raw.upper()
    # 2. Strip a single lowercase letter immediately before the exchange suffix
    #    e.g., NTSCm.BK -> NTSC.BK ; BBLf.BK -> BBL.BK
    m = re.match(r'^([A-Za-z0-9]+?)([a-z])(\.[A-Z]+)$', raw)
    if m:
        yield (m.group(1) + m.group(3)).upper()


def _resolve_via_yfinance_search(query: str):
    """
    Use yfinance.Search to resolve a name (or fuzzy ticker) to a real symbol.
    Prefers SET listings (Thailand) since the app is Thai-focused, then ASEAN, then others.
    Skips NVDR/depository receipt variants (e.g. SUN-R.BK).
    Returns the best symbol string or None.
    """
    try:
        results = yf.Search(query, max_results=10).quotes or []
    except Exception:
        return None

    # Exchange priority: SET (Thailand) -> ASEAN -> rest
    exchange_priority = {
        "SET": 0,
        "SES": 1, "KLS": 1, "JKT": 1, "PHS": 1, "HOM": 1,  # Singapore, Malaysia, Indonesia, Philippines, Vietnam
    }

    def rank(quote):
        symbol = quote.get("symbol", "") or ""
        is_nvdr = "-R." in symbol
        exch = quote.get("exchange", "") or ""
        return (is_nvdr, exchange_priority.get(exch, 99))

    sorted_quotes = sorted(results, key=rank)
    for q in sorted_quotes:
        symbol = q.get("symbol")
        if symbol and "-R." not in symbol:
            return symbol
    # Fallback: first non-empty symbol even if NVDR
    for q in sorted_quotes:
        if q.get("symbol"):
            return q["symbol"]
    return None


def run_sanity_check(deep_dive: dict, lseg_parsed_peers: list, client_phase1_data: dict, latest_year: int):
    """
    Programmatic sanity check (no Gemini call). Returns:
      {"warnings": [str, ...], "critical_issues": [str, ...], "passed": bool}
    Thresholds (from MEGAPROMPT 5.6):
      - EV/EBITDA > 15x
      - Diluted P/E > 30x
      - EV/Revenue > 3x
      - Negative peer EBITDA
      - Client vs peer EBITDA margin gap > 15pp
    """
    warnings = []
    critical = []

    # 1-3. Multiple thresholds across all peer/year cells
    multiple_checks = [
        ('ev_ebitda', 'EV/EBITDA', 15.0),
        ('pe', 'P/E', 30.0),
        ('ev_revenue', 'EV/Revenue', 3.0),
    ]
    for peer in lseg_parsed_peers:
        ticker = peer.get('identifier', '?')
        for key, label, threshold in multiple_checks:
            yearly = peer.get(key) or {}
            for year, val in yearly.items():
                if val is not None and val > threshold:
                    warnings.append(f"{ticker}: {label} for {year} is {val:.2f} (>{threshold:.0f}x threshold).")

    # 4. Negative peer EBITDA (any year, from deep_dive financials)
    fc = deep_dive.get('financials_comparison', []) or []
    for row in fc:
        ticker = row.get('identifier', '?')
        for y in ('2023', '2024', '2025'):
            v = row.get(f'ebitda_{y}')
            if v is not None and v < 0:
                critical.append(f"{ticker}: Negative EBITDA in {y} ({v}). Should be excluded from peer set.")

    # 5. Client EBITDA margin vs peer EBITDA margin gap > 15pp (using latest year)
    try:
        fin = client_phase1_data.get('financials', {}) or {}
        sales = (fin.get('sales_and_services', {}) or {}).get(str(latest_year))
        cogs = (fin.get('cost_of_goods_sold', {}) or {}).get(str(latest_year)) or 0
        sga = (fin.get('sales_expenses', {}) or {}).get(str(latest_year)) or 0
        admin = (fin.get('administrative_expenses', {}) or {}).get(str(latest_year)) or 0
        oth = (fin.get('other_expenses', {}) or {}).get(str(latest_year)) or 0
        da = (fin.get('depreciation_amortization', {}) or {}).get(str(latest_year)) or 0
        if sales:
            client_ebitda = sales - cogs - sga - admin - oth + da
            client_margin = client_ebitda / sales
            for row in fc:
                ticker = row.get('identifier', '?')
                rev = row.get(f'revenue_{latest_year}') or row.get('revenue_2024')
                ebitda = row.get(f'ebitda_{latest_year}') or row.get('ebitda_2024')
                if rev and ebitda is not None:
                    peer_margin = ebitda / rev
                    gap_pp = abs(client_margin - peer_margin) * 100
                    if gap_pp > 15:
                        warnings.append(
                            f"{ticker}: EBITDA margin gap vs client is {gap_pp:.1f}pp "
                            f"(client {client_margin*100:.1f}% vs peer {peer_margin*100:.1f}%)."
                        )
    except Exception:
        # Margin check is best-effort; skip silently if data shape unexpected
        pass

    return {
        "warnings": warnings,
        "critical_issues": critical,
        "passed": len(critical) == 0,
    }


def parse_lseg_transactions_pdf(file_bytes: bytes):
    """
    Extract precedent M&A transactions from a multi-page LSEG/Refinitiv PDF export.
    The PDF table is typically split horizontally across two halves (first half: deal
    basics; second half: EV/EBITDA and counterparty details), joined by SDC Deal No.
    Returns a list of transaction dicts in the same schema as the xlsx parser, or
    {"error": "..."} on failure.
    """
    pdf_part = types.Part.from_bytes(data=file_bytes, mime_type='application/pdf')

    prompt = """
    You are extracting M&A precedent transactions from an LSEG/Refinitiv PDF export.

    IMPORTANT STRUCTURE NOTE:
    This PDF table is split HORIZONTALLY across two halves (e.g. first ~half pages, then
    second ~half pages). Each transaction has ONE row in the first half and a CORRESPONDING
    row in the second half, joined by SDC Deal No. You must merge these into a single
    transaction record per SDC Deal No.

    EXTRACT every transaction in the document. For each one, return:
    - "SDC Deal No": string
    - "Date Announced": string (e.g. "12/23/2025")
    - "Deal Value (USD, Millions)": number or null  (this is "Rank Value Including Net Debt of Target")
    - "Target Full Name": string
    - "Target Nation": string
    - "Target Mid Industry": string
    - "Target Business Description": string or null
    - "Acquiror Full Name": string
    - "Acquiror Nation": string or null
    - "Acquiror Mid Industry": string or null
    - "Acquiror Business Description": string or null
    - "M&A TRBC Activity": string or null
    - "Ratio of Enterprise Value to EBITDA": number or null

    STRICT RULES:
    1. Extract the data EXACTLY as it appears. Do not invent values. Use null when missing.
    2. Numbers should be plain numeric values (38.02, not "38.02" and not "38.02M").
    3. Skip the column-header rows and footer/copyright lines.
    4. If you cannot match a row across halves with confidence, still include it with the
       fields you do have and set the missing fields to null.

    Return ONLY this JSON shape:
    {
      "transactions": [
        {"SDC Deal No": "...", "Date Announced": "...", ...},
        ...
      ]
    }
    """

    try:
        response = _get_client().models.generate_content(
            model='gemini-2.5-flash',
            contents=[prompt, pdf_part],
            config=types.GenerateContentConfig(
                temperature=0.0,
                response_mime_type="application/json",
                max_output_tokens=32768,
                thinking_config=types.ThinkingConfig(thinking_budget=0),
            )
        )
    except Exception as e:
        return {"error": f"Gemini call failed: {e}"}

    raw_text = response.text or ""
    try:
        data = json.loads(_clean_json(raw_text))
    except json.JSONDecodeError as e:
        finish_reason = None
        try:
            finish_reason = response.candidates[0].finish_reason
        except Exception:
            pass
        return {
            "error": (
                f"Could not parse Gemini PDF response (finish_reason={finish_reason}, "
                f"output_len={len(raw_text)} chars): {e}"
            )
        }

    transactions = data.get("transactions", [])
    if not isinstance(transactions, list):
        return {"error": "Gemini response did not contain a 'transactions' list."}
    return transactions


def lookup_ticker_yfinance(query: str):
    """
    Look up a company via yfinance by ticker OR name.
    1. Try the input as a ticker (with LSEG sub-board variant fallback)
    2. If that fails, treat it as a name and use yf.Search to resolve to a ticker
    Returns a peer record dict or None if nothing found.
    For .BK tickers, market cap is converted to THB millions.
    """
    if not query or not query.strip():
        return None

    info = {}
    resolved_ticker = None

    # Pass 1: treat input as ticker
    for variant in _yfinance_ticker_variants(query):
        try:
            candidate = yf.Ticker(variant).info or {}
        except Exception:
            continue
        if candidate.get("longName") or candidate.get("shortName"):
            info = candidate
            resolved_ticker = variant
            break

    # Pass 2: treat input as company name and resolve via search
    if not resolved_ticker:
        searched = _resolve_via_yfinance_search(query)
        if searched:
            try:
                candidate = yf.Ticker(searched).info or {}
            except Exception:
                candidate = {}
            if candidate.get("longName") or candidate.get("shortName"):
                info = candidate
                resolved_ticker = searched

    if not resolved_ticker:
        return None
    ticker = resolved_ticker

    name = info.get("longName") or info.get("shortName")
    if not name:
        return None

    country = info.get("country") or ""
    summary = info.get("longBusinessSummary") or ""
    if len(summary) > 400:
        summary = summary[:400].rsplit(" ", 1)[0] + "..."

    market_cap = info.get("marketCap")
    market_cap_thb_m = None
    if market_cap and ticker.endswith(".BK"):
        market_cap_thb_m = round(market_cap / 1_000_000, 0)

    trbc = info.get("industry") or info.get("sector") or ""

    return _peer_record(
        identifier=ticker,
        company_name=name,
        country=country,
        business_description=summary,
        market_cap_thb_m=market_cap_thb_m,
        trbc_activity=trbc,
        source="Yahoo Finance",
    )


def lookup_ticker_via_gemini(query: str):
    """
    Fallback lookup via Gemini. Accepts a ticker OR company name.
    Returns a peer record dict (with canonical ticker) or None if Gemini
    doesn't recognize the company.
    """
    query = query.strip()
    if not query:
        return None

    prompt = f"""
    You are looking up a single publicly listed company. The user's input may be either
    a ticker symbol OR a company name.

    USER INPUT: {query}

    STRICT RULES:
    1. NO HALLUCINATIONS. If you cannot confidently identify this company, return {{"found": false}}.
    2. The company must be a real, verifiable publicly traded company on a real exchange.
    3. Do NOT invent tickers or companies. Better to return found:false than to guess.
    4. Return the canonical ticker in standard exchange format (e.g., SUN.BK, AAPL, 005930.KS).

    If you DO know the company, return ONLY this JSON:
    {{
      "found": true,
      "identifier": "TICKER.EXCHANGE",
      "company_name": "Full legal name",
      "country": "Country of primary listing",
      "business_description": "2-3 specific sentences about what the company does",
      "market_cap_thb_m": 1800,
      "trbc_activity": "Industry / TRBC activity"
    }}

    For market_cap_thb_m: provide the market cap in millions of THB. If the company is listed
    in a different currency, convert at a reasonable recent rate. If you are unsure, use null.

    If you do NOT recognize the input, return ONLY:
    {{"found": false}}
    """

    response = _get_client().models.generate_content(
        model='gemini-2.5-flash',
        contents=[prompt],
        config=types.GenerateContentConfig(
            temperature=0.0,
            response_mime_type="application/json",
        )
    )

    try:
        data = json.loads(_clean_json(response.text or ""))
    except json.JSONDecodeError:
        return None

    if not data.get("found"):
        return None

    canonical_ticker = (data.get("identifier") or query).strip().upper()

    return _peer_record(
        identifier=canonical_ticker,
        company_name=data.get("company_name", ""),
        country=data.get("country", ""),
        business_description=data.get("business_description", ""),
        market_cap_thb_m=data.get("market_cap_thb_m"),
        trbc_activity=data.get("trbc_activity", ""),
        source="Gemini",
    )


def lookup_ticker(ticker: str):
    """
    Main entry point: try yfinance first, fall back to Gemini.
    Returns (record_dict, source_str) or (None, None) if neither found it.
    source_str is "yfinance" or "gemini".
    """
    record = lookup_ticker_yfinance(ticker)
    if record:
        return record, "yfinance"

    record = lookup_ticker_via_gemini(ticker)
    if record:
        return record, "gemini"

    return None, None


def generate_extraction_report(phase1_data: dict, business_type: str = "Service") -> str:
    """
    Generate a detailed breakdown report explaining how each P&L line item
    was derived from the source documents. Returns Markdown text.
    """
    financials = phase1_data.get("financials", {})
    verification = phase1_data.get("_verification", {})
    client_name = phase1_data.get("client_name", "the company")
    available_years = sorted(
        set(
            y for vals in financials.values()
            if isinstance(vals, dict)
            for y in vals.keys()
        )
    )

    # Build a compact financials summary for the prompt
    fin_summary = json.dumps(financials, indent=2, default=str)
    verif_summary = json.dumps(verification, indent=2, default=str)

    prompt = f"""
You are a Senior M&A Analyst at MAX Solutions writing an internal extraction audit report.

The AI just extracted a Preliminary P&L from uploaded financial documents for: {client_name}
Business type selected by the analyst: **{business_type}**
Available years: {available_years}

EXTRACTED FINANCIALS (THB):
{fin_summary}

VERIFICATION DATA:
{verif_summary}

Write a detailed extraction breakdown report in Markdown. The report should help
the analyst verify every number. Structure it as follows:

## 1. Revenue
For each year, explain:
- What source line items were used for `sales_and_services` (cite the Thai label if applicable)
- What went into `other_revenues` and why
- If multiple branches/segments were summed, list them

## 2. Cost of Goods Sold (COGS)
This is the MOST IMPORTANT section. For each year:
- List every source line item that was included in COGS
- Explain WHY each item is classified as COGS based on the "{business_type}" business type rule
- If items like rent, staff salary, or utilities were included/excluded from COGS, explicitly state why
- Show the subtotals that sum to the final COGS number

## 3. Operating Expenses (Sales, Admin, Other)
For each year:
- What went into sales_expenses vs administrative_expenses vs other_expenses
- List the specific source line items in each category
- If SG&A items (credit card fees, social security, parking, phone, consumables) were classified here, list them

## 4. Below-EBITDA Items
- D&A: where it came from, or why it's null. IMPORTANT: If COGS was derived (back-calculated
  from Revenue − Gross Profit), note that the derived COGS already embeds any D&A that was
  part of the source's Cost of Sales. The D&A line should ONLY contain D&A separated from
  admin/opex — NOT D&A embedded in derived COGS. If both admin-D&A and COGS-D&A were
  combined into the D&A line, flag this as a potential double-counting issue.
- Interest: where it came from, or why it's null
- Tax: where it came from, or why it's null

## 5. Cross-Validation Summary
For each year, show:
- Source-stated Net Profit (if available) vs computed Net Profit from extracted numbers
- Total Revenue − Total Expenses = Implied Net Profit
- Gap amount and percentage
- Whether the extraction passes the 2% threshold

## 6. Data Quality Notes
- Any ambiguities, assumptions, or items that need analyst confirmation
- Items where classification was uncertain

RULES:
- Be SPECIFIC — cite actual Thai labels from the source documents (e.g., "ค่าเช่า สาขารามคำแหง")
- Show arithmetic: e.g., "Rent total = 58,154 × 12 months × 2 branches = 1,395,696"
- If you don't know the exact source breakdown (because it wasn't preserved), say so honestly
- Use tables where they help readability
- Keep it factual, not promotional
"""

    response = _get_client().models.generate_content(
        model='gemini-2.5-flash',
        contents=[prompt],
        config=types.GenerateContentConfig(
            temperature=0.1,
            max_output_tokens=8192,
        )
    )
    return response.text or ""


def generate_owner_questions(phase1_data: dict, flags: list) -> dict:
    """
    Phase 1 helper: generates a structured list of clarification / due-diligence
    questions to send to the business owner, informed by the extracted data and
    any flags raised during review.

    Returns:
      {"questions": [{"category": str, "question": str, "context": str}, ...]}
    """
    client_name     = phase1_data.get("client_name", "the company")
    client_overview = phase1_data.get("client_overview", "")
    financials      = phase1_data.get("financials", {})
    bm              = phase1_data.get("business_model", {})

    # Summarise flags so the AI can generate targeted questions
    flag_lines = "\n".join(
        f"- [{f.level.value.upper()}] {f.title}" + (f": {f.detail}" if f.detail else "")
        for f in flags
    ) if flags else "No flags raised."

    prompt = f"""
You are a senior M&A analyst at MAX Solutions preparing a focused question list
for a client meeting with the business owner. The goal of this meeting is to
collect the specific information needed to complete a Comparable Company Valuation
(Trading Performance analysis). Do NOT generate generic due diligence questions —
every question must serve one of these purposes:

  A) Fill a gap or resolve an anomaly in the financial data so the P&L can be
     correctly populated (9 line items × available years).
  B) Produce or validate the EBITDA normalization figure (the adjusted EBITDA
     a buyer should rely on — this is a required input in the valuation template).
  C) Establish the Net Debt position (needed for Enterprise Value bridge).
  D) Anchor the deal valuation context (asking price rationale, asset vs. business
     value split) so the team understands the seller's implied EV/EBITDA expectation.
  E) Clarify any flag the system raised that directly affects the above.

CLIENT: {client_name}
OVERVIEW: {client_overview}

EXTRACTED FINANCIALS (THB, actual figures):
{json.dumps(financials, indent=2, default=str)}

EXTRACTED BUSINESS MODEL HIGHLIGHTS:
Section A (Identity): {json.dumps(bm.get('section_a', {}), default=str)}
Section B (Revenue):  {json.dumps(bm.get('section_b', {}), default=str)}
Section D (Financial narrative): {json.dumps(bm.get('section_d', {}), default=str)}
Section E (Deal):     {json.dumps(bm.get('section_e', {}), default=str)}
Section F (Ops):      {json.dumps(bm.get('section_f', {}), default=str)}

DATA FLAGS RAISED BY THE SYSTEM (prioritise questions that address these):
{flag_lines}

RULES:
- 8–12 questions maximum. Quality over quantity.
- Reference specific numbers from the data (e.g. "the 48% revenue decline",
  "D&A of 1.59M THB in 2024", "the ~5–6M THB bank debt").
- Write as rough analyst notes — the IB team will polish the wording.
- DO NOT ask about things already clearly captured in the data above.
- Always include questions to request these financial documents if not already
  provided: (1) actual FY2025 revenue to date, (2) accounts receivable aging,
  (3) full fixed-asset register with book values.

CATEGORIES to use (only use a category if you have a real question for it):
  - P&L Gaps & Anomalies
  - EBITDA Normalization
  - Net Debt & Liabilities
  - Deal Valuation & Asset Breakdown
  - Supporting Financial Documents

Return ONLY valid JSON:
{{
  "questions": [
    {{
      "category": "P&L Gaps & Anomalies",
      "question": "...",
      "context": "..."
    }}
  ]
}}
"""

    response = _get_client().models.generate_content(
        model="gemini-2.5-flash",
        contents=[prompt],
        config=types.GenerateContentConfig(
            temperature=0.3,
            response_mime_type="application/json",
            max_output_tokens=8192,
            thinking_config=types.ThinkingConfig(thinking_budget=0),
        ),
    )
    return json.loads(response.text)
