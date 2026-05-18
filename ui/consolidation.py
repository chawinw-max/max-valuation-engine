"""
Consolidation UI — 5-step flow for financial statement consolidation.
Step 1: Upload & Configure
Step 2: Review Extracted Data
Step 3: Review Adjustments
Step 4: Preview Consolidated Financials
Step 5: Download
"""

import streamlit as st
import pandas as pd
import json
from core.document_parser import extract_text_from_file


def render_consolidation():
    """Main entry point for the consolidation tool."""

    # Initialize step tracking
    if "consol_step" not in st.session_state:
        st.session_state.consol_step = 1

    # Step indicator
    steps = [
        "1. Upload & Configure",
        "2. Review Extraction",
        "3. Review Adjustments",
        "4. Preview Consolidated",
        "5. Download",
    ]
    cols = st.columns(len(steps))
    for i, step in enumerate(steps):
        step_num = i + 1
        with cols[i]:
            if st.session_state.consol_step == step_num:
                st.markdown(f"**:green[{step}]**")
            elif st.session_state.consol_step > step_num:
                st.markdown(f":white_check_mark: {step}")
            else:
                st.markdown(f":white_circle: {step}")

    st.divider()

    # Route to current step
    if st.session_state.consol_step == 1:
        _render_step1()
    elif st.session_state.consol_step == 2:
        _render_step2()
    elif st.session_state.consol_step == 3:
        _render_step3()
    elif st.session_state.consol_step == 4:
        _render_step4()
    elif st.session_state.consol_step == 5:
        _render_step5()


# ── Step 1: Upload & Configure ───────────────────────────────────────────────

def _render_step1():
    st.subheader("Step 1: Upload Documents & Configure")

    # Deal code
    deal_code = st.text_input(
        "Deal Code (optional)",
        value=st.session_state.get("consol_deal_code", ""),
        placeholder="e.g., DF-125",
    )
    st.session_state.consol_deal_code = deal_code

    # Entity configuration
    st.markdown("#### Entities")

    ai_define = st.checkbox(
        "Let AI define entities from uploaded documents",
        value=st.session_state.get("consol_ai_define_entities", False),
        help="The AI will analyze your uploaded documents and meeting notes to identify all legal entities and their business types automatically.",
        key="consol_ai_define_cb",
    )
    st.session_state.consol_ai_define_entities = ai_define

    if ai_define:
        st.caption("The AI will detect entities from your uploaded financial documents and meeting notes during extraction.")
    else:
        st.caption("Add each legal entity to consolidate. For a single company, just enter one.")

        if "consol_entities" not in st.session_state:
            st.session_state.consol_entities = [{"name": "", "business_type": "Service"}]

        entities = st.session_state.consol_entities
        business_types = ["Service", "Manufacturing", "Trading / Distribution"]

        for i, entity in enumerate(entities):
            col1, col2, col3 = st.columns([3, 2, 1])
            with col1:
                entities[i]["name"] = st.text_input(
                    f"Entity {i + 1} Name",
                    value=entity["name"],
                    placeholder="e.g., Company A Co., Ltd.",
                    key=f"entity_name_{i}",
                )
            with col2:
                idx = business_types.index(entity["business_type"]) if entity["business_type"] in business_types else 0
                entities[i]["business_type"] = st.selectbox(
                    f"Business Type",
                    business_types,
                    index=idx,
                    key=f"entity_btype_{i}",
                )
            with col3:
                if i > 0:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("Remove", key=f"remove_entity_{i}"):
                        entities.pop(i)
                        st.rerun()

        if st.button("+ Add Entity"):
            entities.append({"name": "", "business_type": "Service"})
            st.rerun()

    # Fiscal years
    st.markdown("#### Fiscal Years")
    year_options = list(range(2019, 2027))
    selected_years = st.multiselect(
        "Select fiscal years to extract",
        year_options,
        default=st.session_state.get("consol_years", [2022, 2023, 2024]),
    )
    st.session_state.consol_years = selected_years

    # File uploads
    st.markdown("#### Financial Documents")
    st.caption("Upload financial statements, tax filings (ภ.ง.ด.50), bank statements — PDF, DOCX, or XLSX.")
    financial_files = st.file_uploader(
        "Financial Documents",
        type=["pdf", "docx", "xlsx", "xls"],
        accept_multiple_files=True,
        key="consol_fin_files",
        label_visibility="collapsed",
    )

    st.markdown("#### Meeting Notes & Context")
    st.caption("Upload meeting notes, owner interview transcripts, or deal memos. These help the AI identify owner expenses, real revenue, and related-party items.")
    notes_files = st.file_uploader(
        "Meeting Notes",
        type=["pdf", "docx", "xlsx", "txt"],
        accept_multiple_files=True,
        key="consol_notes_files",
        label_visibility="collapsed",
    )

    # File-to-entity tagging (if multiple entities, manual mode only)
    entities = st.session_state.get("consol_entities", [])
    if not ai_define and len(entities) > 1 and financial_files:
        st.markdown("#### Tag Files to Entities")
        st.caption("Assign each financial document to its entity. Meeting notes apply to all entities.")

        if "consol_file_tags" not in st.session_state:
            st.session_state.consol_file_tags = {}

        entity_names = [e["name"] or f"Entity {i+1}" for i, e in enumerate(entities)]
        entity_names.append("All Entities")

        for f in financial_files:
            current = st.session_state.consol_file_tags.get(f.name, "All Entities")
            st.session_state.consol_file_tags[f.name] = st.selectbox(
                f"{f.name}",
                entity_names,
                index=entity_names.index(current) if current in entity_names else len(entity_names) - 1,
                key=f"file_tag_{f.name}",
            )

    # Validation & proceed
    st.divider()
    ai_define = st.session_state.get("consol_ai_define_entities", False)
    valid = True
    if not financial_files:
        st.warning("Upload at least one financial document.")
        valid = False
    if not selected_years or len(selected_years) < 1:
        st.warning("Select at least one fiscal year.")
        valid = False
    if not ai_define and not any(e["name"].strip() for e in st.session_state.get("consol_entities", [])):
        st.warning("Enter at least one entity name.")
        valid = False

    if st.button("Extract & Analyze", type="primary", disabled=not valid):
        # Store file references (use different keys than widget keys to avoid Streamlit error)
        st.session_state.consol_financial_files_stored = financial_files
        st.session_state.consol_notes_files_stored = notes_files

        # Extract meeting notes text for later use
        notes_text_parts = []
        for f in (notes_files or []):
            text = extract_text_from_file(f)
            f.seek(0)
            notes_text_parts.append(f"--- {f.name} ---\n{text}")
        st.session_state.consol_notes_text = "\n\n".join(notes_text_parts)

        # If AI-define mode, detect entities first
        if ai_define:
            from core.consolidation_engine import detect_entities

            with st.spinner("AI is analyzing documents to identify entities..."):
                # Reset file pointers
                for f in financial_files:
                    f.seek(0)
                for f in (notes_files or []):
                    f.seek(0)

                detected = detect_entities(
                    financial_files=financial_files,
                    meeting_notes_files=notes_files,
                )

            detected_entities = detected.get("entities", [])
            if not detected_entities:
                st.error("AI could not detect any entities from the documents. Please define entities manually.")
                return

            st.session_state.consol_entities = detected_entities
            st.success(
                f"AI detected **{len(detected_entities)}** entit{'ies' if len(detected_entities) > 1 else 'y'}: "
                + ", ".join(f"**{e['name']}** ({e['business_type']})" for e in detected_entities)
            )
            if detected.get("ai_notes"):
                st.info(f"**AI Notes:** {detected['ai_notes']}")
        else:
            # Filter out empty entities
            valid_entities = [e for e in st.session_state.consol_entities if e["name"].strip()]
            st.session_state.consol_entities = valid_entities

        valid_entities = st.session_state.consol_entities

        # Run extraction for each entity
        from core.consolidation_engine import extract_entity_financials

        extractions = []
        multi = len(valid_entities) > 1
        file_tags = st.session_state.get("consol_file_tags", {})

        progress = st.progress(0, text="Extracting financials...")
        for ei, entity in enumerate(valid_entities):
            ename = entity["name"]
            progress.progress(
                (ei) / len(valid_entities),
                text=f"Extracting {ename}..."
            )

            # Filter files for this entity
            if multi and not ai_define:
                entity_files = [
                    f for f in financial_files
                    if file_tags.get(f.name, "All Entities") in (ename, "All Entities")
                ]
            else:
                entity_files = list(financial_files)

            # Reset file pointers
            for f in entity_files:
                f.seek(0)
            for f in (notes_files or []):
                f.seek(0)

            result = extract_entity_financials(
                financial_files=entity_files,
                meeting_notes_files=notes_files,
                entity_name=ename,
                fiscal_years=selected_years,
                business_type=entity["business_type"],
            )
            extractions.append(result)

        progress.progress(1.0, text="Extraction complete!")
        st.session_state.consol_extractions = extractions
        st.session_state.consol_step = 2
        st.rerun()


# ── Step 2: Review Extracted Data ────────────────────────────────────────────

def _render_step2():
    st.subheader("Step 2: Review Extracted Data")

    if st.button("Back to Upload"):
        st.session_state.consol_step = 1
        for key in ["consol_extractions"]:
            st.session_state.pop(key, None)
        st.rerun()

    extractions = st.session_state.get("consol_extractions", [])
    years = [str(y) for y in sorted(st.session_state.get("consol_years", []))]

    for ext in extractions:
        ename = ext.get("entity_name", "Unknown")
        pnl = ext.get("pnl", {})
        bs = ext.get("balance_sheet", {})

        st.markdown(f"### {ename}")

        # P&L table
        pnl_rows = []
        pnl_labels = {
            "sales_and_services": "Sales & Service Revenue",
            "other_revenues": "Other Revenues",
            "cost_of_goods_sold": "Cost of Goods Sold",
            "selling_expenses": "Selling Expenses",
            "administrative_expenses": "Administrative Expenses",
            "other_expenses": "Other Expenses",
            "depreciation_amortization": "Depreciation & Amortization",
            "interest_expense": "Interest Expense",
            "income_tax": "Income Tax",
        }

        for key, label in pnl_labels.items():
            row = {"Line Item": label}
            for yr in years:
                row[yr] = (pnl.get(key) or {}).get(yr)
            pnl_rows.append(row)

        df = pd.DataFrame(pnl_rows)
        st.dataframe(
            df,
            hide_index=True,
            use_container_width=True,
            column_config={
                yr: st.column_config.NumberColumn(yr, format="%,.2f") for yr in years
            },
        )

        # Balance sheet summary (if available)
        has_bs = any((bs.get(k) or {}).get(yr) is not None for k in bs for yr in years)
        if has_bs:
            with st.expander("Balance Sheet"):
                bs_rows = []
                bs_labels = {
                    "total_assets": "Total Assets",
                    "total_liabilities": "Total Liabilities",
                    "total_equity": "Total Equity",
                    "cash": "Cash & Cash Equivalents",
                    "total_debt": "Total Debt",
                    "related_party_loans": "Related-Party Loans",
                }
                for key, label in bs_labels.items():
                    row = {"Line Item": label}
                    for yr in years:
                        row[yr] = (bs.get(key) or {}).get(yr)
                    bs_rows.append(row)
                st.dataframe(
                    pd.DataFrame(bs_rows),
                    hide_index=True,
                    use_container_width=True,
                    column_config={
                        yr: st.column_config.NumberColumn(yr, format="%,.2f") for yr in years
                    },
                )

        # AI notes
        if ext.get("ai_notes"):
            st.info(f"**AI Notes:** {ext['ai_notes']}")

    # Source documents
    with st.expander("Source Documents Used"):
        for ext in extractions:
            sources = ext.get("source_documents", [])
            if sources:
                st.write(f"**{ext.get('entity_name', '')}:** {', '.join(sources)}")

    st.divider()
    if st.button("Confirm Extraction & Detect Adjustments", type="primary"):
        from core.consolidation_engine import detect_adjustments, detect_intercompany

        entities = st.session_state.consol_entities
        notes_text = st.session_state.get("consol_notes_text", "")

        with st.spinner("AI is analyzing financials and meeting notes for adjustments..."):
            adj_result = detect_adjustments(extractions, notes_text, entities)

        st.session_state.consol_raw_adjustments = adj_result.get("adjustments", [])
        st.session_state.consol_flags = adj_result.get("flags", [])
        st.session_state.consol_meeting_summary = adj_result.get("meeting_notes_summary", "")

        # Add default action to each adjustment
        for adj in st.session_state.consol_raw_adjustments:
            adj["action"] = "accept"

        # Detect intercompany if multi-entity
        if len(entities) > 1:
            with st.spinner("Detecting intercompany transactions..."):
                ic_result = detect_intercompany(extractions, entities)
            ic_elims = ic_result.get("intercompany_eliminations", [])
            for ic in ic_elims:
                ic["action"] = "accept"
            st.session_state.consol_intercompany = ic_elims
        else:
            st.session_state.consol_intercompany = []

        st.session_state.consol_step = 3
        st.rerun()


# ── Step 3: Review Adjustments ───────────────────────────────────────────────

def _render_step3():
    st.subheader("Step 3: Review Proposed Adjustments")

    if st.button("Back to Extraction Review"):
        st.session_state.consol_step = 2
        for key in ["consol_raw_adjustments", "consol_flags", "consol_intercompany", "consol_meeting_summary"]:
            st.session_state.pop(key, None)
        st.rerun()

    adjustments = st.session_state.get("consol_raw_adjustments", [])
    flags = st.session_state.get("consol_flags", [])
    intercompany = st.session_state.get("consol_intercompany", [])
    meeting_summary = st.session_state.get("consol_meeting_summary", "")

    # Meeting notes summary
    if meeting_summary:
        with st.expander("Meeting Notes Summary", expanded=True):
            st.write(meeting_summary)

    # Flags
    if flags:
        st.markdown("#### Flags for Attention")
        for flag in flags:
            severity = flag.get("severity", "info")
            if severity == "warning":
                st.warning(f"**{flag.get('title', '')}:** {flag.get('detail', '')}")
            else:
                st.info(f"**{flag.get('title', '')}:** {flag.get('detail', '')}")

    # Adjustments table
    st.markdown("#### Proposed Adjustments")
    st.caption("Review each adjustment. Set Action to Accept, Reject, or Edit. You can also add manual adjustments below.")

    if adjustments:
        adj_df = pd.DataFrame(adjustments)

        display_cols = [
            "entity", "year", "line_item", "original_amount",
            "adjustment_amount", "ebitda_impact", "category",
            "description", "confidence", "action",
        ]
        # Only show columns that exist
        display_cols = [c for c in display_cols if c in adj_df.columns]

        edited = st.data_editor(
            adj_df[display_cols],
            hide_index=True,
            use_container_width=True,
            key="adj_editor",
            column_config={
                "action": st.column_config.SelectboxColumn(
                    "Action", options=["accept", "reject"], default="accept",
                ),
                "category": st.column_config.SelectboxColumn(
                    "Category",
                    options=["owner_expense", "related_party", "tax_adjustment", "non_recurring", "intercompany"],
                ),
                "original_amount": st.column_config.NumberColumn("Original", format="%,.0f"),
                "adjustment_amount": st.column_config.NumberColumn("Adjustment", format="%,.0f"),
                "ebitda_impact": st.column_config.NumberColumn("EBITDA Impact", format="%,.0f"),
                "confidence": st.column_config.SelectboxColumn(
                    "Confidence", options=["high", "medium", "low"],
                ),
            },
        )

        # Update adjustments with edited values
        for i, row in edited.iterrows():
            if i < len(adjustments):
                for col in display_cols:
                    adjustments[i][col] = row[col]

        accepted = sum(1 for a in adjustments if a.get("action") == "accept")
        rejected = sum(1 for a in adjustments if a.get("action") == "reject")
        st.write(f"**{accepted}** accepted, **{rejected}** rejected out of **{len(adjustments)}** proposed adjustments")
    else:
        st.info("No adjustments were detected. You can add manual adjustments below.")

    # Manual adjustment entry
    with st.expander("Add Manual Adjustment"):
        entities = st.session_state.consol_entities
        years = [str(y) for y in sorted(st.session_state.get("consol_years", []))]

        with st.form("manual_adj_form"):
            m_col1, m_col2 = st.columns(2)
            with m_col1:
                m_entity = st.selectbox("Entity", [e["name"] for e in entities])
                m_year = st.selectbox("Year", years)
                m_line = st.selectbox("Line Item", [
                    "sales_and_services", "other_revenues", "cost_of_goods_sold",
                    "selling_expenses", "administrative_expenses", "other_expenses",
                    "depreciation_amortization", "interest_expense", "income_tax",
                ])
            with m_col2:
                m_original = st.number_input("Original Amount", value=0.0, format="%.2f")
                m_adjustment = st.number_input(
                    "Adjustment Amount",
                    value=0.0,
                    format="%.2f",
                    help="Negative = reduce expense (increases EBITDA). Positive = increase expense.",
                )
                m_category = st.selectbox("Category", [
                    "owner_expense", "related_party", "tax_adjustment", "non_recurring",
                ])
            m_desc = st.text_input("Description", placeholder="e.g., Owner's personal vehicle lease")
            m_explanation = st.text_area("Explanation", placeholder="Why this adjustment is needed", height=80)
            m_submitted = st.form_submit_button("Add Adjustment")

        if m_submitted and m_desc.strip():
            new_adj = {
                "entity": m_entity,
                "line_item": m_line,
                "year": m_year,
                "original_amount": m_original,
                "adjustment_amount": m_adjustment,
                "adjusted_amount": m_original + m_adjustment,
                "category": m_category,
                "description": m_desc.strip(),
                "explanation": m_explanation.strip(),
                "evidence_source": "manual",
                "confidence": "high",
                "ebitda_impact": abs(m_adjustment) if m_line != "sales_and_services" else m_adjustment,
                "action": "accept",
            }
            adjustments.append(new_adj)
            st.session_state.consol_raw_adjustments = adjustments
            st.rerun()

    # Intercompany eliminations
    if intercompany:
        st.markdown("#### Intercompany Eliminations")
        ic_df = pd.DataFrame(intercompany)
        ic_display = ["from_entity", "to_entity", "year", "amount", "description", "action"]
        ic_display = [c for c in ic_display if c in ic_df.columns]

        edited_ic = st.data_editor(
            ic_df[ic_display],
            hide_index=True,
            use_container_width=True,
            key="ic_editor",
            column_config={
                "action": st.column_config.SelectboxColumn(
                    "Action", options=["accept", "reject"], default="accept",
                ),
                "amount": st.column_config.NumberColumn("Amount", format="%,.0f"),
            },
        )

        for i, row in edited_ic.iterrows():
            if i < len(intercompany):
                for col in ic_display:
                    intercompany[i][col] = row[col]

    st.divider()
    if st.button("Confirm Adjustments & Preview", type="primary"):
        st.session_state.consol_step = 4
        st.rerun()


# ── Step 4: Preview Consolidated Financials ──────────────────────────────────

def _render_step4():
    st.subheader("Step 4: Preview Consolidated Financials")

    if st.button("Back to Adjustments"):
        st.session_state.consol_step = 3
        st.rerun()

    extractions = st.session_state.get("consol_extractions", [])
    adjustments = st.session_state.get("consol_raw_adjustments", [])
    intercompany = st.session_state.get("consol_intercompany", [])
    years = [str(y) for y in sorted(st.session_state.get("consol_years", []))]
    entities = st.session_state.consol_entities

    accepted = [a for a in adjustments if a.get("action") != "reject"]

    from core.consolidation_excel import _apply_adjustments, _calc_derived

    # Build consolidated view
    def _build_pnl(kind):
        combined = {}
        for ext in extractions:
            ename = ext.get("entity_name", "")
            pnl = ext.get("pnl", {})
            if kind == "normalized":
                pnl = _apply_adjustments(pnl, accepted, ename, years)
            for k, v in pnl.items():
                if k not in combined:
                    combined[k] = {yr: 0 for yr in years}
                for yr in years:
                    combined[k][yr] = (combined[k].get(yr) or 0) + ((v or {}).get(yr) or 0)
        return combined

    reported = _build_pnl("reported")
    normalized = _build_pnl("normalized")

    # Normalized P&L table
    st.markdown("#### Normalized P&L")
    pnl_lines = [
        ("sales_and_services", "Revenue", False),
        ("cost_of_goods_sold", "COGS", False),
        (None, "Gross Profit", True),
        ("selling_expenses", "Selling Expenses", False),
        ("administrative_expenses", "Admin Expenses", False),
        ("other_expenses", "Other Expenses", False),
        (None, "EBITDA", True),
        ("depreciation_amortization", "D&A", False),
        ("interest_expense", "Interest", False),
        ("income_tax", "Tax", False),
        (None, "Net Profit", True),
    ]

    header_row = {"Line Item": ""}
    for yr in years:
        header_row[f"{yr} Reported"] = ""
        header_row[f"{yr} Adj"] = ""
        header_row[f"{yr} Normalized"] = ""

    rows = []
    for key, label, is_total in pnl_lines:
        row = {"Line Item": label}
        for yr in years:
            if is_total:
                # Map display label to calc key
                calc_map = {"Gross Profit": "Gross Profit", "EBITDA": "EBITDA", "Net Profit": "Net Profit"}
                calc_key = calc_map.get(label, label)
                rep = _calc_derived(calc_key, reported, yr)
                norm = _calc_derived(calc_key, normalized, yr)
            else:
                rep = (reported.get(key) or {}).get(yr)
                norm = (normalized.get(key) or {}).get(yr)

            row[f"{yr} Reported"] = rep
            row[f"{yr} Adj"] = ((norm or 0) - (rep or 0)) if rep is not None else None
            row[f"{yr} Normalized"] = norm
        rows.append(row)

    df = pd.DataFrame(rows)
    num_cols = [c for c in df.columns if c != "Line Item"]
    st.dataframe(
        df,
        hide_index=True,
        use_container_width=True,
        column_config={
            c: st.column_config.NumberColumn(c, format="%,.0f") for c in num_cols
        },
    )

    # EBITDA Bridge
    st.markdown("#### EBITDA Bridge")
    cat_labels = {
        "owner_expense": "Owner Expenses",
        "related_party": "Related-Party",
        "tax_adjustment": "Tax Adjustments",
        "non_recurring": "Non-Recurring",
        "intercompany": "Intercompany",
    }

    bridge_rows = []
    bridge_rows.append({
        "Item": "Reported EBITDA",
        **{yr: _calc_derived("EBITDA", reported, yr) for yr in years},
    })

    for cat, label in cat_labels.items():
        cat_adjs = [a for a in accepted if a.get("category") == cat]
        if not cat_adjs:
            continue
        row = {"Item": f"  + {label}"}
        for yr in years:
            total = sum(a.get("ebitda_impact", 0) or 0 for a in cat_adjs if a.get("year") == yr)
            row[yr] = total if total else None
        bridge_rows.append(row)

    bridge_rows.append({
        "Item": "Normalized EBITDA",
        **{yr: _calc_derived("EBITDA", normalized, yr) for yr in years},
    })

    bridge_df = pd.DataFrame(bridge_rows)
    st.dataframe(
        bridge_df,
        hide_index=True,
        use_container_width=True,
        column_config={
            yr: st.column_config.NumberColumn(yr, format="%,.0f") for yr in years
        },
    )

    # Key metrics
    st.markdown("#### Key Metrics")
    metrics_cols = st.columns(len(years))
    for ci, yr in enumerate(years):
        with metrics_cols[ci]:
            st.metric(f"{yr} Revenue", f"THB {_calc_derived('Total Revenue', normalized, yr):,.0f}")
            ebitda = _calc_derived("EBITDA", normalized, yr)
            rev = _calc_derived("Total Revenue", normalized, yr)
            margin = f"{ebitda / rev * 100:.1f}%" if rev else "N/A"
            rep_ebitda = _calc_derived("EBITDA", reported, yr)
            delta = f"{ebitda - rep_ebitda:+,.0f}" if rep_ebitda else None
            st.metric(f"{yr} EBITDA", f"THB {ebitda:,.0f}", delta=delta)
            st.metric(f"{yr} EBITDA Margin", margin)

    st.divider()
    if st.button("Generate Outputs", type="primary"):
        _generate_outputs()


def _generate_outputs():
    """Generate Excel and Word report, store in session state."""
    from core.consolidation_engine import generate_consolidation_narrative
    from core.consolidation_excel import generate_consolidation_excel
    from core.consolidation_report import generate_consolidation_report

    extractions = st.session_state.consol_extractions
    adjustments = st.session_state.consol_raw_adjustments
    intercompany = st.session_state.get("consol_intercompany", [])
    entities = st.session_state.consol_entities
    years = st.session_state.consol_years
    deal_code = st.session_state.get("consol_deal_code", "")
    flags = st.session_state.get("consol_flags", [])
    meeting_summary = st.session_state.get("consol_meeting_summary", "")

    accepted = [a for a in adjustments if a.get("action") != "reject"]

    # Generate narrative
    with st.spinner("Generating consolidation narrative..."):
        # Build normalized data summary for the narrative prompt
        from core.consolidation_excel import _apply_adjustments, _calc_derived
        str_years = [str(y) for y in sorted(years)]

        def _consol(kind):
            combined = {}
            for ext in extractions:
                pnl = ext.get("pnl", {})
                if kind == "normalized":
                    pnl = _apply_adjustments(pnl, accepted, ext["entity_name"], str_years)
                for k, v in pnl.items():
                    if k not in combined:
                        combined[k] = {yr: 0 for yr in str_years}
                    for yr in str_years:
                        combined[k][yr] = (combined[k].get(yr) or 0) + ((v or {}).get(yr) or 0)
            return combined

        reported = _consol("reported")
        normalized = _consol("normalized")

        normalized_data = {
            "reported_ebitda": {yr: _calc_derived("EBITDA", reported, yr) for yr in str_years},
            "normalized_ebitda": {yr: _calc_derived("EBITDA", normalized, yr) for yr in str_years},
            "reported_revenue": {yr: _calc_derived("Total Revenue", reported, yr) for yr in str_years},
            "normalized_revenue": {yr: _calc_derived("Total Revenue", normalized, yr) for yr in str_years},
        }

        narrative = generate_consolidation_narrative(
            normalized_data=normalized_data,
            adjustments=accepted,
            flags=flags,
            meeting_notes_summary=meeting_summary,
            entity_configs=entities,
        )

    st.session_state.consol_narrative = narrative

    # Generate Excel
    with st.spinner("Generating Excel workbook..."):
        excel_bytes = generate_consolidation_excel(
            extractions=extractions,
            adjustments=adjustments,
            intercompany=intercompany,
            entity_configs=entities,
            fiscal_years=years,
            deal_code=deal_code,
        )
    st.session_state.consol_excel_bytes = excel_bytes

    # Generate Word report
    with st.spinner("Generating consolidation report..."):
        # Collect source filenames
        source_files = []
        for f in st.session_state.get("consol_financial_files_stored", []):
            source_files.append(f.name)
        for f in st.session_state.get("consol_notes_files_stored", []):
            source_files.append(f.name)

        report_bytes = generate_consolidation_report(
            extractions=extractions,
            adjustments=adjustments,
            intercompany=intercompany,
            narrative=narrative,
            entity_configs=entities,
            fiscal_years=years,
            deal_code=deal_code,
            source_files=source_files,
        )
    st.session_state.consol_report_bytes = report_bytes

    st.session_state.consol_step = 5
    st.rerun()


# ── Step 5: Download ─────────────────────────────────────────────────────────

def _render_step5():
    st.subheader("Step 5: Download Outputs")

    if st.button("Back to Preview"):
        st.session_state.consol_step = 4
        st.rerun()

    deal_code = st.session_state.get("consol_deal_code", "")
    prefix = f"{deal_code}_" if deal_code else ""

    st.success("Consolidation complete! Download your files below.")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### Consolidated Excel")
        st.caption("Normalized P&L, EBITDA Bridge, Balance Sheet, and Adjustments Detail.")
        excel_bytes = st.session_state.get("consol_excel_bytes")
        if excel_bytes:
            st.download_button(
                label="Download Consolidated Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"{prefix}Consolidated_Financials.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

    with col2:
        st.markdown("### Consolidation Report")
        st.caption("Audit trail for the IB team — every adjustment explained and sourced.")
        report_bytes = st.session_state.get("consol_report_bytes")
        if report_bytes:
            st.download_button(
                label="Download Consolidation Report (.docx)",
                data=report_bytes,
                file_name=f"{prefix}Consolidation_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )

    # Show narrative preview
    narrative = st.session_state.get("consol_narrative", {})
    if narrative:
        with st.expander("Preview: Report Executive Summary"):
            st.write(narrative.get("executive_summary", ""))

        with st.expander("Preview: Flags & Warnings"):
            st.write(narrative.get("warnings_section", "No warnings."))

    st.divider()
    if st.button("Start New Consolidation"):
        keys_to_clear = [k for k in st.session_state if k.startswith("consol_")]
        for k in keys_to_clear:
            del st.session_state[k]
        st.rerun()
