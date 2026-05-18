import streamlit as st
import pandas as pd
from core.ai_engine import generate_peer_list, verify_peer_list, generate_rejection_rationales, lookup_ticker
from core.flag_engine import get_phase2_flags, render_flags

def render_phase2():
    st.header("Phase 2: Peer Screening Funnel")
    
    data = st.session_state.phase1_data
    client_overview = data.get('client_overview', '')
    business_attributes = data.get('business_attributes', {})
    
    # Get latest year revenue for scale comparison
    financials = data.get('financials', {})
    sales = financials.get('sales_and_services', {})
    # Get the max year available
    latest_year = max(st.session_state.selected_years) if st.session_state.selected_years else 2024
    latest_year_revenue = sales.get(str(latest_year), 0)
    
    # 1. AI Generation
    if "peer_list" not in st.session_state:
        st.write("Generate the initial broad list of up to 30 comparable companies based on the Ring framework.")

        with st.expander("Notes for AI (optional)", expanded=False):
            st.caption("Suggest specific tickers, sub-sectors, or geographies to focus on or exclude.")
            phase2_notes = st.text_area(
                "Notes", value=st.session_state.get("phase2_notes", ""),
                height=100, label_visibility="collapsed", key="phase2_notes_input"
            )
            st.session_state.phase2_notes = phase2_notes

        if st.button("Generate Peer List", type="primary"):
            try:
                with st.spinner("AI is analyzing and sourcing comparable peers (thinking enabled)..."):
                    result = generate_peer_list(
                        client_overview, business_attributes, latest_year_revenue,
                        notes=st.session_state.get("phase2_notes", ""),
                        financials=financials,
                        available_years=st.session_state.get("selected_years", []),
                    )
                    raw_peers = result.get('broad_list', [])

                progress_bar = st.progress(0, text="Verifying tickers via Yahoo Finance...")
                def _update_progress(done, total):
                    progress_bar.progress(done / total, text=f"Verifying tickers... {done}/{total}")

                verified_peers = verify_peer_list(raw_peers, progress_callback=_update_progress)
                progress_bar.empty()

                verified_count = sum(1 for p in verified_peers if p.get('verified'))
                unverified = [p for p in verified_peers if not p.get('verified')]

                # Keep ALL peers (verified + unverified) — let the analyst decide
                st.session_state.peer_list = verified_peers
                st.session_state._peer_verification_summary = {
                    'verified': verified_count,
                    'total': len(verified_peers),
                    'unverified': [p.get('identifier', '?') for p in unverified],
                }
                st.rerun()
            except Exception as e:
                st.error(f"Error generating peers: {e}")
                return
        return # Stop rendering until generated
    
    peers = st.session_state.peer_list
    
    # We maintain a column "Selected" in the dataframe for checkboxes
    if "peer_df" not in st.session_state:
        df = pd.DataFrame(peers)
        # Add selection column if it doesn't exist
        if 'Selected' not in df.columns:
            df.insert(0, 'Selected', False)
        st.session_state.peer_df = df

    # Step 2b: Initial Selection (30 -> 12)
    if not st.session_state.get('top_12_confirmed', False):
        st.subheader("Initial Selection: Top 12 Peers")
        st.write("Select 5 to 12 companies to move to the final shortlist.")

        vsummary = st.session_state.pop('_peer_verification_summary', None)
        if vsummary:
            v = vsummary['verified']
            t = vsummary.get('total', v)
            if v == t:
                st.success(f"Verified all **{v}** peers via Yahoo Finance.")
            elif v > 0:
                st.success(f"Verified **{v}/{t}** peers via Yahoo Finance.")
                st.info(
                    f"**{t - v}** ticker(s) could not be verified (may be delisted, "
                    f"misspelled, or Yahoo Finance unavailable): "
                    + ", ".join(vsummary.get('unverified', []))
                    + "\n\nThey are kept in the list — review and remove any that are incorrect."
                )
            else:
                st.warning(
                    f"Could not verify any tickers via Yahoo Finance (0/{t}). "
                    "This may be due to rate limiting on the cloud server. "
                    "All AI-generated peers are kept — please verify them manually."
                )

        if st.button("← Re-generate Peer List"):
            for key in ['peer_list', 'peer_df', 'peer_df_base']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
        
        # Only init the editor state once
        if "peer_df_base" not in st.session_state:
            st.session_state.peer_df_base = st.session_state.peer_df.copy()

        with st.expander("Add companies manually"):
            tab_lookup, tab_manual = st.tabs(["🔍 Look up by ticker/name", "✏️ Enter manually"])

            with tab_lookup:
                with st.form("add_peer_form"):
                    st.write("Enter a ticker symbol or company name — the app will look it up and auto-fill all fields.")
                    new_query = st.text_input("Ticker or Company Name (e.g., SUN.BK, Sunsweet, Chiangmai Frozen)")
                    submitted = st.form_submit_button("Look up & Add")

                if submitted:
                    query_input = (new_query or "").strip()
                    if not query_input:
                        st.error("Please enter a ticker or company name.")
                    else:
                        with st.spinner(f"Looking up '{query_input}'..."):
                            record, source = lookup_ticker(query_input)

                        if record is None:
                            st.warning(
                                f"Could not find '{query_input}' in Yahoo Finance or Gemini. "
                                "The company may be delisted or not covered. "
                                "Use the **Enter manually** tab to add it with your own details."
                            )
                        else:
                            existing = st.session_state.peer_df_base['identifier'].astype(str).str.upper().tolist()
                            if record['identifier'].upper() in existing:
                                st.warning(
                                    f"{record['identifier']} ({record['company_name']}) is already in the peer list."
                                )
                            else:
                                record['Selected'] = True
                                base = st.session_state.peer_df_base
                                new_row_df = pd.DataFrame([record])
                                st.session_state.peer_df_base = pd.concat(
                                    [new_row_df, base],
                                    ignore_index=True,
                                )[base.columns.tolist()]
                                if source == "yfinance":
                                    st.success(
                                        f"Added {record['identifier']} — {record['company_name']} (Yahoo Finance)"
                                    )
                                else:
                                    st.warning(
                                        f"Added {record['identifier']} — {record['company_name']} "
                                        "(AI-sourced from Gemini — please verify accuracy)"
                                    )
                                st.rerun()

            with tab_manual:
                st.caption("For delisted companies or those not on Yahoo Finance. Fill in the details yourself.")
                with st.form("manual_peer_form"):
                    m_col1, m_col2 = st.columns(2)
                    with m_col1:
                        m_ticker = st.text_input("Ticker *", placeholder="e.g., PSQ.AX")
                        m_name = st.text_input("Company Name *", placeholder="e.g., Pacific Smiles Group Limited")
                        m_country = st.text_input("Country *", placeholder="e.g., Australia")
                    with m_col2:
                        m_industry = st.text_input("Industry / TRBC Activity", placeholder="e.g., Dental Services")
                        m_mcap = st.number_input("Market Cap (THB M)", value=0, min_value=0, help="Leave 0 if unknown")
                        m_ring = st.selectbox("Ring", options=[1, 2, 3], index=1)
                    m_desc = st.text_area("Business Description", placeholder="2-3 sentences about what the company does", height=80)
                    m_submitted = st.form_submit_button("Add to Peer List")

                if m_submitted:
                    if not m_ticker.strip() or not m_name.strip() or not m_country.strip():
                        st.error("Ticker, Company Name, and Country are required.")
                    else:
                        existing = st.session_state.peer_df_base['identifier'].astype(str).str.upper().tolist()
                        if m_ticker.strip().upper() in existing:
                            st.warning(f"{m_ticker.strip().upper()} is already in the peer list.")
                        else:
                            from core.ai_engine import _score_country
                            geo = _score_country(m_country.strip())
                            manual_record = {
                                'identifier': m_ticker.strip().upper(),
                                'company_name': m_name.strip(),
                                'trbc_activity': m_industry.strip() or "Manually Added",
                                'country': m_country.strip(),
                                'business_description': m_desc.strip(),
                                'market_cap_thb_m': m_mcap if m_mcap > 0 else None,
                                'ebitda_positive': True,
                                'fit_rank': 2,
                                'geography_score': geo,
                                'ring': m_ring,
                                'ring_justification': "Manually added by analyst.",
                                'scale_warning': None,
                                'verified': True,
                                'Selected': True,
                            }
                            base = st.session_state.peer_df_base
                            new_row_df = pd.DataFrame([manual_record])
                            st.session_state.peer_df_base = pd.concat(
                                [new_row_df, base],
                                ignore_index=True,
                            )[base.columns.tolist()]
                            st.success(f"Added {manual_record['identifier']} — {manual_record['company_name']} (manual entry)")
                            st.rerun()

        edited_df = st.data_editor(
            st.session_state.peer_df_base,
            hide_index=True,
            use_container_width=True,
            key="peer_editor_step1",
            column_config={
                "Selected": st.column_config.CheckboxColumn("Select", default=False),
                "verified": st.column_config.CheckboxColumn("✓ Verified", default=False),
                "market_cap_thb_m": st.column_config.NumberColumn("Market Cap (THB M)", format="%,.0f"),
            },
            disabled=["identifier", "company_name", "trbc_activity", "country", "business_description", "market_cap_thb_m", "fit_rank", "geography_score", "ring", "ring_justification", "scale_warning", "verified"]
        )
        
        # When Smart Select is clicked, we update the base df
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("Smart Select Top 12"):
                st.session_state.peer_df_base['Selected'] = False
                st.session_state.peer_df_base.loc[:11, 'Selected'] = True
                st.rerun()

        selected_count = edited_df['Selected'].sum()
        st.write(f"**Selected: {selected_count}** (Aim for 5 to 12)")
        
        if st.button("Confirm Top 12", disabled=not (5 <= selected_count <= 12), type="primary"):
            st.session_state.top_12_confirmed = True
            st.session_state.top_12_df = edited_df[edited_df['Selected'] == True].copy()
            st.rerun()
            
    # Step 2c: Final Selection (12 -> 5-7)
    elif not st.session_state.get('final_peers_confirmed', False):
        st.subheader("Final Selection: 5-7 Peers")
        st.write("Narrow down your shortlist to the final 5-7 peers.")

        if st.button("← Back to Top 12 Selection"):
            for key in ['top_12_confirmed', 'top_12_df', 'final_df_base']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()

        if "final_df_base" not in st.session_state:
            st.session_state.final_df_base = st.session_state.top_12_df.copy()
            st.session_state.final_df_base['Selected'] = False # Reset selection for the final cut

        edited_final = st.data_editor(
            st.session_state.final_df_base,
            hide_index=True,
            use_container_width=True,
            key="peer_editor_step2",
            column_config={
                "Selected": st.column_config.CheckboxColumn("Final Selection", default=False),
                "market_cap_thb_m": st.column_config.NumberColumn("Market Cap (THB M)", format="%,.0f"),
            },
            disabled=["identifier", "company_name", "trbc_activity", "country", "business_description", "market_cap_thb_m"]
        )
        
        col_sel1, col_sel2 = st.columns([1, 5])
        with col_sel1:
            if st.button("Select All"):
                st.session_state.final_df_base['Selected'] = True
                st.rerun()

        final_count = edited_final['Selected'].sum()
        st.write(f"**Final Selected: {final_count}** (Need 5 to 7)")
        
        if st.button("Finalize Peers", disabled=not (5 <= final_count <= 7), type="primary"):
            selected_tickers = edited_final[edited_final['Selected'] == True]['identifier'].tolist()
            
            # Unselected are the ones in the Top 12 that aren't selected
            not_selected = st.session_state.top_12_df[~st.session_state.top_12_df['identifier'].isin(selected_tickers)]
            
            not_selected_list = not_selected[['identifier', 'company_name']].to_dict(orient='records')
            
            with st.spinner("Generating rejection rationales..."):
                try:
                    rationales = generate_rejection_rationales(client_overview, not_selected_list)
                    st.session_state.rejection_rationales = rationales.get('rationales', {})
                    st.session_state.final_selected_peers = edited_final[edited_final['Selected'] == True].to_dict(orient='records')
                    st.session_state.final_not_selected_peers = not_selected.to_dict(orient='records')
                    st.session_state.final_peers_confirmed = True
                    st.rerun()
                except Exception as e:
                    st.error(f"Error generating rationales: {e}")
                    
    # Step 2.5: Checkpoint Preview
    else:
        st.subheader("Checkpoint: Peer Selection Confirmed")

        mcap_col_config = {
            "market_cap_thb_m": st.column_config.NumberColumn("Market Cap (THB M)", format="%,.0f"),
        }

        st.markdown("### Selected Peers")
        sel_df = pd.DataFrame(st.session_state.final_selected_peers)[
            ['identifier', 'company_name', 'country', 'market_cap_thb_m', 'business_description']
        ]
        st.dataframe(sel_df, hide_index=True, column_config=mcap_col_config, use_container_width=True)

        st.markdown("### Not Selected Peers")
        not_selected_raw = st.session_state.final_not_selected_peers
        if not_selected_raw:
            not_sel_df = pd.DataFrame(not_selected_raw)
            display_cols = [c for c in ['identifier', 'company_name', 'market_cap_thb_m'] if c in not_sel_df.columns]
            not_sel_df = not_sel_df[display_cols] if display_cols else not_sel_df
            not_sel_df['Reason'] = not_sel_df['identifier'].map(
                lambda x: st.session_state.rejection_rationales.get(x, "Not a strong fit")
            ) if 'identifier' in not_sel_df.columns else "Not a strong fit"
            st.dataframe(not_sel_df, hide_index=True, column_config=mcap_col_config, use_container_width=True)
        else:
            st.caption("All peers were selected — none rejected.")

        st.divider()
        flags = get_phase2_flags(
            st.session_state.get("phase1_data", {}),
            st.session_state.get("final_selected_peers", [])
        )
        render_flags(flags, "Phase 2")

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("← Back to Final Selection"):
                if 'final_peers_confirmed' in st.session_state:
                    del st.session_state['final_peers_confirmed']
                st.rerun()
        with col2:
            if st.button("Reset & Fix Peers"):
                for key in ['peer_list', 'peer_df', 'peer_df_base', 'top_12_confirmed', 'top_12_df', 'final_df_base', 'final_peers_confirmed']:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()
        with col3:
            if st.button("Confirm & Continue to Phase 3", type="primary"):
                st.session_state.current_phase = 3
                st.rerun()
