import streamlit as st
import pandas as pd
from core.document_parser import parse_lseg_peer_data, parse_lseg_transactions
from core.ai_engine import generate_deep_dive, select_precedent_transactions, parse_lseg_transactions_pdf
from core.flag_engine import get_phase3_flags, render_flags

def render_phase3():
    st.header("Phase 3 & 3.5: Deep Dive & Precedent Transactions")
    
    selected_peers = st.session_state.get('final_selected_peers', [])
    if not selected_peers:
        st.warning("No peers were selected in Phase 2.")
        return
        
    st.markdown("### Step 3a: Upload LSEG Peer Data")
    st.write("Upload the individual `.xlsx` LSEG files for the selected peers.")
    
    # Show expected tickers
    expected_tickers = [p['identifier'] for p in selected_peers]
    st.write(f"**Expected Tickers:** {', '.join(expected_tickers)}")
    
    lseg_files = st.file_uploader(
        "Upload LSEG Peer Files",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="lseg_peer_files"
    )
    
    if lseg_files:
        if st.button("Parse LSEG Data"):
            with st.spinner("Parsing LSEG files..."):
                parsed_peers = []
                failed_files = []
                for f in lseg_files:
                    f_bytes = f.read()
                    data = parse_lseg_peer_data(f_bytes, filename=f.name)
                    if "error" in data:
                        failed_files.append((f.name, data["error"]))
                    else:
                        parsed_peers.append(data)

                st.session_state.lseg_parsed_peers = parsed_peers
                st.session_state.lseg_failed_files = failed_files

                if failed_files:
                    st.error(
                        f"Parsed {len(parsed_peers)} of {len(lseg_files)} files. "
                        f"{len(failed_files)} failed:"
                    )
                    for name, err in failed_files:
                        st.write(f"- **{name}**: {err}")
                else:
                    st.success(f"Parsed all {len(parsed_peers)} peer files.")
                st.rerun()

    if "lseg_parsed_peers" in st.session_state:
        st.write("Parsed Data Check:")
        df = pd.DataFrame(st.session_state.lseg_parsed_peers)
        st.dataframe(df, hide_index=True)

        # LSEG coverage flags — immediately after parse
        lseg_flags = get_phase3_flags(
            selected_peers,
            st.session_state.lseg_parsed_peers,
            st.session_state.get("deep_dive"),
        )
        render_flags(lseg_flags, "Phase 3 — LSEG Coverage")

        st.markdown("### Step 3b: AI Deep Dive")
        if "deep_dive" not in st.session_state:
            with st.expander("Notes for AI (optional)", expanded=False):
                st.caption("Provide context about the peers (e.g. currency to use, known revenue figures, or caveats for specific companies).")
                phase3_notes = st.text_area(
                    "Notes", value=st.session_state.get("phase3_notes", ""),
                    height=100, label_visibility="collapsed", key="phase3_notes_input"
                )
                st.session_state.phase3_notes = phase3_notes

            if st.button("Generate AI Deep Dive Comparisons", type="primary"):
                with st.spinner("Generating qualitative and quantitative comparisons..."):
                    client_overview = st.session_state.phase1_data.get('client_overview', '')
                    result = generate_deep_dive(
                        client_overview, selected_peers,
                        notes=st.session_state.get("phase3_notes", "")
                    )
                    st.session_state.deep_dive = result
                    st.rerun()
        else:
            st.success("Deep Dive generated!")
            dd = st.session_state.deep_dive
            st.markdown("#### Qualitative Comparison")
            st.dataframe(pd.DataFrame(dd.get('qualitative', [])), hide_index=True)
            st.markdown("#### Quantitative Forecasts (Millions THB)")
            st.dataframe(pd.DataFrame(dd.get('financials_comparison', [])), hide_index=True)

            # Full Phase 3 flags (includes AI-estimated financials note + high multiples)
            deep_flags = get_phase3_flags(
                selected_peers,
                st.session_state.lseg_parsed_peers,
                dd,
            )
            render_flags(deep_flags, "Phase 3 — Deep Dive")

            st.divider()
            
            # Phase 3.5
            st.markdown("### Phase 3.5: Precedent Transactions")
            st.write("Upload the single LSEG Precedent M&A Excel file to extract top 10 relevant deals.")
            
            tx_file = st.file_uploader(
                "Upload LSEG Precedent Transactions (.xlsx, .xls, .xlsm, or .pdf)",
                type=["xlsx", "xls", "xlsm", "pdf"],
                key="lseg_tx_file"
            )

            if tx_file:
                is_pdf = tx_file.name.lower().endswith(".pdf")
                if is_pdf:
                    st.caption(
                        "PDF detected — extraction takes 30–60 seconds and costs ~THB 3–5. "
                        "Use xlsx when available for faster, more reliable parsing."
                    )

                with st.expander("Notes for AI (optional)", expanded=False):
                    st.caption("E.g. preferred date range, geography focus, or deal types to exclude.")
                    phase35_notes = st.text_area(
                        "Notes", value=st.session_state.get("phase35_notes", ""),
                        height=80, label_visibility="collapsed", key="phase35_notes_input"
                    )
                    st.session_state.phase35_notes = phase35_notes

                if st.button("Extract Best Transactions"):
                    spinner_msg = (
                        "Extracting transactions from PDF via Gemini..."
                        if is_pdf else
                        "Parsing and selecting top 10 transactions..."
                    )
                    with st.spinner(spinner_msg):
                        f_bytes = tx_file.read()
                        if is_pdf:
                            parsed_txs = parse_lseg_transactions_pdf(f_bytes)
                        else:
                            parsed_txs = parse_lseg_transactions(f_bytes)

                        if isinstance(parsed_txs, dict) and "error" in parsed_txs:
                            st.error(f"Failed to parse transactions: {parsed_txs['error']}")
                        else:
                            st.session_state.all_transactions = parsed_txs
                            if is_pdf:
                                st.info(f"Extracted {len(parsed_txs)} transactions from PDF.")

                            client_data = st.session_state.phase1_data
                            tx_result = select_precedent_transactions(
                                client_data, parsed_txs,
                                notes=st.session_state.get("phase35_notes", "")
                            )
                            st.session_state.selected_transactions = tx_result.get('selected_transactions', [])
                            st.rerun()
            
            if "selected_transactions" in st.session_state:
                st.success("Top transactions selected!")
                st.dataframe(pd.DataFrame(st.session_state.selected_transactions), hide_index=True)
                
                st.divider()
                if st.button("Confirm & Proceed to Export (Phase 4)", type="primary"):
                    st.session_state.current_phase = 4
                    st.rerun()
