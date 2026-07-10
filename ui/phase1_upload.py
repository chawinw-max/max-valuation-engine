import streamlit as st
import pandas as pd
from core.document_parser import extract_text_from_file
from core.ai_engine import extract_financials_and_business_model, generate_owner_questions, generate_extraction_report
from core.flag_engine import get_phase1_flags, render_flags

def render_phase1():
    st.header("Phase 1: Upload & Financial Baseline")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Company Information")
        st.write("Meeting notes, company profiles, business descriptions, deal memos")
        company_files = st.file_uploader(
            "Upload Company Docs", 
            type=["pdf", "docx", "xlsx", "xls", "csv", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key="company_docs"
        )
        if company_files:
            for f in company_files:
                st.caption(f.name)
                
    with col2:
        st.subheader("Financial Statements")
        st.write("Income statements, P&L reports, audited financials (Thai or English)")
        financial_files = st.file_uploader(
            "Upload Financial Docs", 
            type=["pdf", "docx", "xlsx", "xls", "csv", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key="financial_docs"
        )
        if financial_files:
            for f in financial_files:
                st.caption(f.name)
                
    st.divider()
    
    st.subheader("Select Available Years")
    st.write("Check the years covered by the uploaded financial documents (minimum 3 years).")
    
    years = [2020, 2021, 2022, 2023, 2024, 2025]
    cols = st.columns(len(years))
    
    selected_years = []
    for i, year in enumerate(years):
        with cols[i]:
            if st.checkbox(str(year), value=(year in [2022, 2023, 2024]), key=f"yr_{year}"):
                selected_years.append(year)
                
    st.divider()

    st.subheader("Business Type")
    st.write("Select the business type to ensure correct COGS classification.")
    business_type = st.selectbox(
        "Business Type",
        options=["Service", "Manufacturing", "Trading / Distribution"],
        index=0,
        key="business_type_select",
        help=(
            "**Service** (clinic, salon, consulting, restaurant): "
            "Facility rent, practitioner fees, and staff delivering services → COGS.\n\n"
            "**Manufacturing**: Raw materials, direct labor, factory overhead → COGS.\n\n"
            "**Trading / Distribution**: Purchase cost of goods, warehouse, logistics → COGS."
        )
    )
    st.session_state.business_type = business_type

    can_extract = bool(company_files and financial_files and len(selected_years) >= 3)

    with st.expander("Notes for AI (optional)", expanded=False):
        st.caption("Add any clarifications to help the AI — e.g. currency, which year columns map to which fiscal year, known data issues.")
        phase1_notes = st.text_area(
            "Notes", value=st.session_state.get("phase1_notes", ""),
            height=100, label_visibility="collapsed", key="phase1_notes_input"
        )
        st.session_state.phase1_notes = phase1_notes

    if st.button("Extract Data", disabled=not can_extract, type="primary"):
        with st.spinner("Analyzing with Gemini 2.5 Flash..."):
            try:
                result = extract_financials_and_business_model(
                    company_files, financial_files, selected_years,
                    notes=st.session_state.get("phase1_notes", ""),
                    business_type=st.session_state.get("business_type", "Service")
                )
                st.session_state.phase1_data = result
                st.session_state.selected_years = selected_years
                st.success("Extraction Complete!")
            except Exception as e:
                st.error(f"Error during extraction: {e}")
                return
    
    # Display Results if extraction is done
    if "phase1_data" in st.session_state:
        data = st.session_state.phase1_data
        
        st.divider()
        st.subheader("Extracted Results")
        
        # 1. Company Overview Card
        st.info(f"**{data.get('deal_code', 'DF-XXX')} - {data.get('client_name', 'Company')}**\n\n{data.get('client_overview', '')}")
        
        # Show verification notes if present
        v_notes = (data.get('_verification') or {}).get('notes')
        if v_notes:
            st.warning(f"**Extraction Notes:**\n\n{v_notes}")

        # 2. Editable P&L Table
        st.markdown("#### Preliminary P&L")
        
        fin_data = data.get('financials', {})
        
        # Build DataFrame for data_editor
        rows = [
            "Sales and Services", "Other Revenues", "Cost of good sold", 
            "Sales Expenses", "Administrative Expenses", "Other Expenses", 
            "Depreciation and Amortization", "Interest Expenses", "Tax"
        ]
        
        keys = [
            "sales_and_services", "other_revenues", "cost_of_goods_sold",
            "sales_expenses", "administrative_expenses", "other_expenses",
            "depreciation_amortization", "interest_expenses", "tax"
        ]
        
        df_dict = {"Line Item": rows}
        for year in st.session_state.selected_years:
            df_dict[str(year)] = [fin_data.get(k, {}).get(str(year), 0.0) for k in keys]
            
        df = pd.DataFrame(df_dict)
        
        edited_df = st.data_editor(
            df,
            hide_index=True,
            use_container_width=True,
            column_config={
                str(y): st.column_config.NumberColumn(str(y), format="%.2f") for y in st.session_state.selected_years
            }
        )
        
        # Update session state with edited P&L
        for y in st.session_state.selected_years:
            for i, k in enumerate(keys):
                if k not in st.session_state.phase1_data['financials']:
                    st.session_state.phase1_data['financials'][k] = {}
                st.session_state.phase1_data['financials'][k][str(y)] = edited_df.loc[i, str(y)]
        
        # Calculated metrics display
        st.markdown("**Calculated Metrics (Auto-updated)**")
        calc_cols = st.columns(len(st.session_state.selected_years))
        for i, y in enumerate(st.session_state.selected_years):
            with calc_cols[i]:
                st.markdown(f"**{y}**")
                sales = edited_df.loc[0, str(y)] or 0
                other = edited_df.loc[1, str(y)] or 0
                cogs = edited_df.loc[2, str(y)] or 0
                
                total_rev = sales + other
                gp = total_rev - cogs
                gp_margin = (gp / total_rev * 100) if total_rev else 0
                
                sales_exp = edited_df.loc[3, str(y)] or 0
                admin_exp = edited_df.loc[4, str(y)] or 0
                other_exp = edited_df.loc[5, str(y)] or 0
                opex = sales_exp + admin_exp + other_exp
                
                ebitda = gp - opex
                
                st.write(f"Total Revenue: {total_rev:,.2f}")
                st.write(f"Gross Profit: {gp:,.2f} ({gp_margin:.1f}%)")
                st.write(f"EBITDA: {ebitda:,.2f}")
                
        # 2a-2. Editable Balance Sheet Table (from AUDITED financial statements)
        st.markdown("#### Balance Sheet (Audited FS)")
        st.caption("Derived from audited financial statements only — not internal accounts. "
                   "Feeds the FCF working-capital/CapEx lines, the SME tax-rate test, and the Equity Value bridge.")

        bs_data = data.get('balance_sheet', {}) or {}
        bs_rows = [
            "Cash and cash equivalents", "Accounts receivable",
            "Short-term loans receivable", "Inventories - net",
            "Property, plant and equipment - net",
            "Accounts payable", "Short-term loans",
            "Other current liabilities", "Long-term loans",
            "Issued and paid-up capital", "Retained earnings",
        ]
        bs_keys = [
            "cash_and_equivalents", "accounts_receivable",
            "short_term_loans_receivable", "inventories",
            "ppe_net",
            "accounts_payable", "short_term_loans",
            "other_current_liabilities", "long_term_loans",
            "paid_up_capital", "retained_earnings",
        ]
        bs_df_dict = {"Line Item": bs_rows}
        for year in st.session_state.selected_years:
            bs_df_dict[str(year)] = [bs_data.get(k, {}).get(str(year), 0.0) for k in bs_keys]

        edited_bs_df = st.data_editor(
            pd.DataFrame(bs_df_dict),
            hide_index=True,
            use_container_width=True,
            key="bs_editor",
            column_config={
                str(y): st.column_config.NumberColumn(str(y), format="%.2f")
                for y in st.session_state.selected_years
            }
        )
        if 'balance_sheet' not in st.session_state.phase1_data:
            st.session_state.phase1_data['balance_sheet'] = {}
        for y in st.session_state.selected_years:
            for i, k in enumerate(bs_keys):
                if k not in st.session_state.phase1_data['balance_sheet']:
                    st.session_state.phase1_data['balance_sheet'][k] = {}
                st.session_state.phase1_data['balance_sheet'][k][str(y)] = edited_bs_df.loc[i, str(y)]

        # 2b. Extraction Breakdown Report
        st.markdown("---")
        report_col1, report_col2 = st.columns([3, 1])
        with report_col1:
            st.markdown("#### Extraction Breakdown")
            st.caption("Generate a detailed report explaining how each P&L number was derived from the source documents.")
        with report_col2:
            gen_report = st.button("📋 Generate Report", key="gen_extraction_report", use_container_width=True)

        if gen_report:
            with st.spinner("Generating extraction breakdown..."):
                try:
                    report_md = generate_extraction_report(
                        st.session_state.phase1_data,
                        business_type=st.session_state.get("business_type", "Service")
                    )
                    st.session_state.extraction_report = report_md
                except Exception as e:
                    st.error(f"Failed to generate report: {e}")

        if st.session_state.get("extraction_report"):
            with st.expander("📋 Extraction Breakdown Report", expanded=True):
                st.markdown(st.session_state.extraction_report)
                deal_code = data.get("deal_code", "Draft")
                st.download_button(
                    "⬇ Download Report (.md)",
                    data=st.session_state.extraction_report,
                    file_name=f"extraction_breakdown_{deal_code}.md",
                    mime="text/markdown",
                    key="dl_extraction_report"
                )
                if st.button("🔄 Re-generate Report", key="regen_extraction_report"):
                    del st.session_state.extraction_report
                    st.rerun()

        # 3. Editable Business Model
        st.markdown("#### Business Model Details")
        st.caption("Edit any fields below before confirming.")
        
        bm = data.get("business_model", {})
        
        with st.expander("Section A - Business Identity", expanded=False):
            sec_a = bm.get("section_a", {})
            st.session_state.phase1_data['business_model']['section_a']['company_name'] = st.text_input("Company Name", sec_a.get("company_name", ""))
            st.session_state.phase1_data['business_model']['section_a']['founded'] = st.text_input("Founded", sec_a.get("founded", ""))
            st.session_state.phase1_data['business_model']['section_a']['type'] = st.text_input("Type", sec_a.get("type", ""))
            st.session_state.phase1_data['business_model']['section_a']['registered_capital'] = st.text_input("Registered Capital", sec_a.get("registered_capital", ""))
            
        with st.expander("Section B - Revenue Model", expanded=False):
            sec_b = bm.get("section_b", {})
            st.session_state.phase1_data['business_model']['section_b']['primary_revenue'] = st.text_area("Primary Revenue", sec_b.get("primary_revenue", ""))
            st.session_state.phase1_data['business_model']['section_b']['key_clients'] = st.text_area("Key Clients", sec_b.get("key_clients", ""))
            
        with st.expander("Section E - Deal Structure", expanded=False):
            sec_e = bm.get("section_e", {})
            st.session_state.phase1_data['business_model']['section_e']['seller'] = st.text_input("Seller", sec_e.get("seller", ""))
            st.session_state.phase1_data['business_model']['section_e']['exit_motivation'] = st.text_input("Exit Motivation", sec_e.get("exit_motivation", ""))
            
        st.divider()
        flags = get_phase1_flags(
            st.session_state.phase1_data,
            st.session_state.get("selected_years", [])
        )
        render_flags(flags, "Phase 1")

        # ── Owner Questions ──────────────────────────────────────────────
        st.subheader("Owner Questions")
        st.caption(
            "Generate a draft question list for the IB team to refine before "
            "the next client meeting."
        )
        if st.button("Generate Owner Questions", key="gen_owner_q"):
            with st.spinner("Drafting questions..."):
                try:
                    result = generate_owner_questions(
                        st.session_state.phase1_data, flags
                    )
                    st.session_state.owner_questions = result.get("questions", [])
                except Exception as e:
                    st.error(f"Failed to generate questions: {e}")

        if st.session_state.get("owner_questions"):
            questions = st.session_state.owner_questions

            # Group by category
            from collections import defaultdict
            by_cat = defaultdict(list)
            for q in questions:
                by_cat[q.get("category", "General")].append(q)

            for cat, items in by_cat.items():
                st.markdown(f"**{cat}**")
                for i, q in enumerate(items, 1):
                    with st.expander(f"{i}. {q.get('question', '')}", expanded=False):
                        st.caption(f"Why: {q.get('context', '')}")

            # Plain-text download
            lines = []
            for cat, items in by_cat.items():
                lines.append(f"\n{cat.upper()}")
                lines.append("-" * len(cat))
                for i, q in enumerate(items, 1):
                    lines.append(f"{i}. {q.get('question', '')}")
                    lines.append(f"   [Context: {q.get('context', '')}]")
            txt = "\n".join(lines).strip()
            deal_code = st.session_state.phase1_data.get("deal_code", "Draft")
            st.download_button(
                "⬇ Download as .txt",
                data=txt,
                file_name=f"owner_questions_{deal_code}.txt",
                mime="text/plain",
            )

            if st.button("Re-generate Questions", key="regen_owner_q"):
                del st.session_state.owner_questions
                st.rerun()

        st.divider()
        if st.button("Confirm & Proceed to Phase 2", type="primary"):
            st.session_state.current_phase = 2
            st.rerun()
