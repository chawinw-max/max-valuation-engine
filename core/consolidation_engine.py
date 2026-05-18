"""
Consolidation AI engine — Gemini prompts for financial statement
extraction, adjustment detection, intercompany elimination, and
report narrative generation.

Reuses _get_client() and _file_to_part() from ai_engine.py.
"""

import json
import re
from google.genai import types
from core.ai_engine import _get_client, _file_to_part, _clean_json


def _safe_parse_json(text: str) -> dict:
    """Parse JSON with multiple fallback strategies for malformed Gemini output."""
    cleaned = _clean_json(text)
    # Attempt 1: direct parse
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass
    # Attempt 2: fix common issues — unescaped newlines in strings, trailing commas
    fixed = re.sub(r'(?<!\\)\n', '\\n', cleaned)
    fixed = re.sub(r',\s*([\]}])', r'\1', fixed)
    try:
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass
    # Attempt 3: extract the first JSON object/array
    match = re.search(r'(\{[\s\S]*\}|\[[\s\S]*\])', cleaned)
    if match:
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            pass
    # Attempt 4: ask Gemini to fix it
    try:
        repair_response = _get_client().models.generate_content(
            model="gemini-2.5-flash",
            contents=[f"Fix this malformed JSON and return ONLY valid JSON, nothing else:\n\n{cleaned}"],
            config=types.GenerateContentConfig(
                temperature=0.0,
                max_output_tokens=8192,
                response_mime_type="application/json",
            ),
        )
        return json.loads(_clean_json(repair_response.text))
    except Exception:
        raise ValueError(f"Could not parse Gemini response as JSON. Raw output:\n{text[:500]}")


# ── Prompt 0: Detect entities from documents ─────────────────────────────────

def detect_entities(financial_files, meeting_notes_files=None):
    """Analyze uploaded documents to identify legal entities and their business types."""

    system_prompt = """
You are a Senior M&A Analyst at MAX Solutions, a Thai M&A advisory firm.
Analyze the uploaded documents to identify all legal entities (companies) mentioned.

For Thai companies, look for:
- Company names in Thai (บริษัท ... จำกัด) and English
- Entity names on financial statement headers, audit reports, tax filings
- Related entities mentioned in meeting notes or footnotes
- Parent/subsidiary relationships

For each entity, determine the most appropriate business type:
- "Service" — consulting, healthcare, education, hospitality, logistics, professional services
- "Manufacturing" — factories, production, food processing, assembly
- "Trading / Distribution" — wholesale, retail, import/export, distribution

Return VALID JSON only. No markdown.
"""

    user_prompt = """
Analyze these documents and identify all distinct legal entities that have financial statements or are discussed as part of the deal.

Return:
{
    "entities": [
        {
            "name": "Full legal entity name",
            "business_type": "Service" or "Manufacturing" or "Trading / Distribution"
        }
    ],
    "ai_notes": "Brief notes about the entities detected — relationships between them, which documents belong to which entity, and any ambiguities."
}

Only include entities that have financial data in the documents or are explicitly identified as part of the deal scope in meeting notes. Do NOT include entities that are merely mentioned as customers, suppliers, or unrelated parties.
"""

    parts = []
    for f in (financial_files or []):
        parts.append(_file_to_part(f))
        f.seek(0)
    for f in (meeting_notes_files or []):
        parts.append(f"--- MEETING NOTES: {f.name} ---")
        parts.append(_file_to_part(f))
        f.seek(0)
    parts.append(user_prompt)

    response = _get_client().models.generate_content(
        model="gemini-2.5-flash",
        contents=parts,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=0.0,
            max_output_tokens=4096,
            response_mime_type="application/json",
        ),
    )

    return _safe_parse_json(response.text)


# ── Prompt 1: Extract entity financials ──────────────────────────────────────

def extract_entity_financials(
    financial_files,
    meeting_notes_files,
    entity_name: str,
    fiscal_years: list,
    business_type: str = "Service",
):
    """Extract P&L and balance sheet for a single entity from uploaded docs."""

    year_list = ", ".join(str(y) for y in sorted(fiscal_years))

    system_prompt = f"""
You are a Senior M&A Analyst at MAX Solutions, a Thai M&A advisory firm.
You are consolidating messy client financial documents for: {entity_name}

STRICT RULES:
1. NO HALLUCINATIONS: Only use data from the documents provided. Never invent financial figures. If you cannot find the exact number, return null.
2. CURRENCY: All financial figures in Thai Baht (THB). If source uses another currency, convert and note the rate.
3. RESPOND ONLY IN VALID JSON matching the schema. No markdown, no commentary outside the JSON.
4. COPY ALL NUMBERS EXACTLY as they appear in the source documents. Do NOT round.
5. Buddhist Era conversion: Thai years (พ.ศ.) = CE + 543. So พ.ศ. 2565 = 2022, พ.ศ. 2566 = 2023, พ.ศ. 2567 = 2024. Also handle abbreviated forms (e.g. "66" = 2566 = 2023).

THAI FINANCIAL STATEMENT TERMS:
- รายได้จากการขาย / รายได้จากการให้บริการ = Sales & Service Revenue
- รายได้อื่น = Other Revenue
- ต้นทุนขาย / ต้นทุนบริการ = Cost of Goods Sold / Cost of Services
- ค่าใช้จ่ายในการขาย = Selling Expenses
- ค่าใช้จ่ายในการบริหาร = Administrative Expenses
- ค่าใช้จ่ายอื่น = Other Expenses
- ค่าเสื่อมราคา = Depreciation
- ค่าตัดจำหน่าย = Amortization
- ดอกเบี้ยจ่าย = Interest Expense
- ภาษีเงินได้ = Income Tax
- สินทรัพย์รวม = Total Assets
- หนี้สินรวม = Total Liabilities
- ส่วนของผู้ถือหุ้น = Total Equity
- เงินสดและรายการเทียบเท่าเงินสด = Cash & Cash Equivalents
- เงินกู้ยืม = Borrowings / Debt
- ลูกหนี้การค้า = Trade Receivables
- สินค้าคงเหลือ = Inventory
- เจ้าหนี้การค้า = Trade Payables
- เงินให้กู้ยืมแก่กรรมการ / เงินให้กู้ยืมแก่กิจการที่เกี่ยวข้องกัน = Related-Party Loans

CRITICAL — DEPRECIATION & AMORTIZATION (D&A) SEPARATION:
The #1 error to avoid is leaving D&A bundled inside Administrative Expenses.
D&A MUST be extracted as a SEPARATE line item (depreciation_amortization).
- Thai financial statements often embed ค่าเสื่อมราคา (depreciation) inside ค่าใช้จ่ายในการบริหาร (admin expenses).
- You MUST find the D&A amount — check these sources in order:
  1. หมายเหตุประกอบงบการเงิน (Notes to Financial Statements) — look for Note about ที่ดิน อาคาร และอุปกรณ์ (Property, Plant & Equipment) which shows ค่าเสื่อมราคา for the year
  2. งบกระแสเงินสด (Cash Flow Statement) — D&A is always shown as an addback in operating activities
  3. Breakdown of admin expenses in the notes
  4. If D&A is shown as a sub-line under admin expenses, extract it
- Once you find D&A, SUBTRACT it from administrative_expenses and put it in depreciation_amortization
- Example: if admin total is 3,300,000 and D&A within admin is 820,000, then:
    administrative_expenses = 2,480,000 (admin MINUS D&A)
    depreciation_amortization = 820,000
- If D&A appears in BOTH admin and COGS (e.g., factory depreciation in COGS + office depreciation in admin), extract TOTAL D&A and subtract from the relevant line items
- NEVER leave D&A as zero/null when admin expenses are non-zero — there is almost always some depreciation

INCOME TAX:
- Always extract ภาษีเงินได้ (income tax expense) separately, even for small amounts
- Check both the P&L face and the tax note (หมายเหตุภาษีเงินได้)
- For entities with low profit, tax may be small but must still be captured

BUSINESS TYPE: {business_type}
FISCAL YEARS TO EXTRACT: {year_list}
ENTITY NAME: {entity_name}

If there are documents for multiple entities, only extract data for {entity_name}.
If you find meeting notes referencing actual/real numbers that differ from the formal statements, include the formal statement numbers in the main extraction and note the discrepancy in ai_notes.
"""

    user_prompt = f"""
Extract the P&L and balance sheet for entity "{entity_name}" for fiscal years: {year_list}.

CRITICAL INSTRUCTION FOR D&A:
- depreciation_amortization MUST be a separate line item, NOT embedded in administrative_expenses
- Find D&A from notes to financial statements, cash flow statement, or admin expense breakdown
- administrative_expenses should be the PURE admin figure AFTER removing D&A
- This is essential because EBITDA = Revenue - COGS - SGA - Other (before D&A). If D&A is inside admin, EBITDA will be wrong.

Return JSON in this exact schema:
{{
    "entity_name": "{entity_name}",
    "pnl": {{
        "sales_and_services": {{"YEAR": number_or_null, ...}},
        "other_revenues": {{"YEAR": number_or_null, ...}},
        "cost_of_goods_sold": {{"YEAR": number_or_null, ...}},
        "selling_expenses": {{"YEAR": number_or_null, ...}},
        "administrative_expenses": {{"YEAR": number_or_null, ...}},
        "other_expenses": {{"YEAR": number_or_null, ...}},
        "depreciation_amortization": {{"YEAR": number_or_null, ...}},
        "interest_expense": {{"YEAR": number_or_null, ...}},
        "income_tax": {{"YEAR": number_or_null, ...}}
    }},
    "balance_sheet": {{
        "total_assets": {{"YEAR": number_or_null, ...}},
        "total_liabilities": {{"YEAR": number_or_null, ...}},
        "total_equity": {{"YEAR": number_or_null, ...}},
        "cash": {{"YEAR": number_or_null, ...}},
        "total_debt": {{"YEAR": number_or_null, ...}},
        "accounts_receivable": {{"YEAR": number_or_null, ...}},
        "inventory": {{"YEAR": number_or_null, ...}},
        "accounts_payable": {{"YEAR": number_or_null, ...}},
        "related_party_loans": {{"YEAR": number_or_null, ...}}
    }},
    "ai_notes": "Free-text notes about data quality issues, discrepancies between documents, missing data, or notable observations from meeting notes.",
    "source_documents": ["list of filenames that contained data for this entity"]
}}

Replace YEAR placeholders with the actual years: {year_list}.
If balance sheet data is not available, set all balance_sheet values to null.
"""

    parts = []
    for f in (financial_files or []):
        parts.append(_file_to_part(f))
        f.seek(0)
    for f in (meeting_notes_files or []):
        parts.append(f"--- MEETING NOTES: {f.name} ---")
        parts.append(_file_to_part(f))
        f.seek(0)
    parts.append(user_prompt)

    response = _get_client().models.generate_content(
        model="gemini-2.5-flash",
        contents=parts,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=0.0,
            max_output_tokens=8192,
            response_mime_type="application/json",
        ),
    )

    return _safe_parse_json(response.text)


# ── Prompt 2: Detect normalization adjustments ───────────────────────────────

def detect_adjustments(extractions: list, meeting_notes_text: str, entity_configs: list):
    """Cross-reference extracted financials with meeting notes to find adjustments."""

    entities_summary = json.dumps(extractions, indent=2, ensure_ascii=False, default=str)

    system_prompt = """
You are a Senior M&A Analyst at MAX Solutions, a Thai advisory firm.
You are reviewing extracted financial statements and meeting notes to identify normalization adjustments.

Your job is to find items that distort the company's true economic performance and propose adjustments.
You must ONLY propose adjustments you have evidence for — either from meeting notes or clear anomalies in the financials.

ADJUSTMENT CATEGORIES:
- "owner_expense": Personal expenses run through the business (owner's car, travel, family school fees, personal insurance, home renovation costs disguised as admin or selling expenses)
- "related_party": Transactions with related parties at non-arm's-length terms (loans to directors, above-market rent to owner's property company, management fees to holding entity, transfer pricing between related companies)
- "tax_adjustment": Items motivated by tax minimization (accelerated depreciation beyond economic life, understated revenue, overstated COGS, excessive provisions)
- "non_recurring": One-time items that don't reflect ongoing operations (legal settlements, asset write-offs, restructuring costs, COVID-related grants, insurance payouts)
- "intercompany": Transactions between entities being consolidated (only if multiple entities)

THAI SME PATTERNS TO LOOK FOR:
- ค่าใช้จ่ายในการบริหาร (SG&A) that is unusually high relative to revenue — often contains owner personal expenses
- เงินให้กู้ยืมแก่กรรมการ (loans to directors) on the balance sheet — cash drain not reflected in P&L
- Significant consulting fees or management fees to related parties
- Rent expense to a company owned by the same family
- Vehicle leases and fuel costs that seem personal
- Sharp year-over-year changes without business explanation

RULES:
1. Only propose adjustments with clear evidence. State the evidence source.
2. For each adjustment, explain the impact on EBITDA (positive = increases EBITDA).
3. Be conservative — when uncertain, flag it but set confidence to "low".
4. Return VALID JSON only. No markdown.
"""

    user_prompt = f"""
Here are the extracted financials for each entity:
{entities_summary}

Here are the meeting notes and context from the client:
---
{meeting_notes_text or "No meeting notes provided."}
---

Entity configurations:
{json.dumps(entity_configs, indent=2, ensure_ascii=False)}

Analyze these financials and meeting notes. Return a JSON array of proposed adjustments:

{{
    "adjustments": [
        {{
            "entity": "entity name",
            "line_item": "pnl field name (e.g. administrative_expenses, selling_expenses, cost_of_goods_sold)",
            "year": "2023",
            "original_amount": 5000000,
            "adjustment_amount": -800000,
            "adjusted_amount": 4200000,
            "category": "owner_expense",
            "description": "Short title (e.g. Owner vehicle lease)",
            "explanation": "Detailed explanation of why this adjustment is proposed, citing evidence from meeting notes or financial anomalies",
            "evidence_source": "meeting_notes" or "financial_analysis",
            "confidence": "high" or "medium" or "low",
            "ebitda_impact": 800000
        }}
    ],
    "flags": [
        {{
            "entity": "entity name",
            "title": "Short flag title",
            "detail": "Detailed explanation of the issue that needs IB team attention",
            "severity": "warning" or "info"
        }}
    ],
    "meeting_notes_summary": "Summary of key disclosures from the meeting notes and how they were used in the analysis"
}}

Important: adjustment_amount should be NEGATIVE to reduce an expense (which increases EBITDA), POSITIVE to increase an expense (which decreases EBITDA).
For revenue adjustments: POSITIVE to increase revenue, NEGATIVE to decrease.
"""

    response = _get_client().models.generate_content(
        model="gemini-2.5-flash",
        contents=[user_prompt],
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=0.0,
            max_output_tokens=8192,
            response_mime_type="application/json",
        ),
    )

    return _safe_parse_json(response.text)


# ── Prompt 3: Detect intercompany transactions ──────────────────────────────

def detect_intercompany(extractions: list, entity_configs: list):
    """Identify intercompany transactions to eliminate (multi-entity only)."""

    entities_summary = json.dumps(extractions, indent=2, ensure_ascii=False, default=str)

    system_prompt = """
You are a Senior M&A Analyst consolidating multiple Thai entities.
Identify intercompany transactions that must be eliminated to avoid double-counting.

Look for:
- Revenue in one entity matching COGS in another (intercompany sales)
- Related-party loans between entities
- Management fees between entities
- Shared services charges
- Intercompany receivables/payables

Only flag transactions you have evidence for. Return VALID JSON only.
"""

    user_prompt = f"""
Entities and their financials:
{entities_summary}

Entity configurations:
{json.dumps(entity_configs, indent=2, ensure_ascii=False)}

Return:
{{
    "intercompany_eliminations": [
        {{
            "from_entity": "entity selling/providing",
            "to_entity": "entity buying/receiving",
            "description": "Short description",
            "revenue_line": "sales_and_services",
            "cost_line": "cost_of_goods_sold",
            "year": "2023",
            "amount": 2000000,
            "explanation": "Why this is intercompany"
        }}
    ]
}}

Return empty array if no intercompany transactions are detected.
"""

    response = _get_client().models.generate_content(
        model="gemini-2.5-flash",
        contents=[user_prompt],
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=0.0,
            max_output_tokens=4096,
            response_mime_type="application/json",
        ),
    )

    return _safe_parse_json(response.text)


# ── Prompt 4: Generate consolidation narrative for the report ────────────────

def generate_consolidation_narrative(
    normalized_data: dict,
    adjustments: list,
    flags: list,
    meeting_notes_summary: str,
    entity_configs: list,
):
    """Generate structured narrative text for the Word consolidation report."""

    system_prompt = """
You are a Senior M&A Analyst at MAX Solutions writing a consolidation report for the IB team.
Write clear, professional prose. Use Thai Baht figures with proper formatting.
This report will be reviewed by the investment banking team — be precise and cite specific numbers.
Return VALID JSON only.
"""

    user_prompt = f"""
Based on the following consolidation results, generate the narrative sections for the report.

Normalized financial data:
{json.dumps(normalized_data, indent=2, ensure_ascii=False, default=str)}

Adjustments made:
{json.dumps(adjustments, indent=2, ensure_ascii=False, default=str)}

Flags for IB attention:
{json.dumps(flags, indent=2, ensure_ascii=False, default=str)}

Meeting notes summary:
{meeting_notes_summary or "No meeting notes provided."}

Entity configurations:
{json.dumps(entity_configs, indent=2, ensure_ascii=False)}

Return JSON:
{{
    "executive_summary": "1-2 paragraph executive summary covering entities consolidated, fiscal years, key finding (normalized vs reported EBITDA with % change)",
    "data_quality_notes": "Paragraph about data quality, completeness, and any caveats about the source documents",
    "adjustment_narratives": {{
        "owner_expense": "Paragraph summarizing all owner expense adjustments",
        "related_party": "Paragraph summarizing related-party adjustments",
        "tax_adjustment": "Paragraph summarizing tax-motivated adjustments",
        "non_recurring": "Paragraph summarizing non-recurring items",
        "intercompany": "Paragraph summarizing intercompany eliminations"
    }},
    "warnings_section": "Bulleted list (as string) of items requiring IB team attention",
    "meeting_notes_section": "Summary of what the owner disclosed and how each disclosure was reflected in adjustments",
    "conclusion": "Brief concluding paragraph with final normalized EBITDA and recommended next steps"
}}

Omit categories with no adjustments (set to null).
"""

    response = _get_client().models.generate_content(
        model="gemini-2.5-flash",
        contents=[user_prompt],
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=0.1,
            max_output_tokens=8192,
            response_mime_type="application/json",
        ),
    )

    return _safe_parse_json(response.text)
