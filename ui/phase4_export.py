import io
import os
import re
import tempfile
import openpyxl
import streamlit as st

from core.ai_engine import run_sanity_check
from core.drive_api import download_template, upload_output
from core.debug_report import generate_debug_report
from core.excel_bridge import (
    inject_phase1_data,
    inject_phase2_data,
    inject_phase3_data,
    inject_phase35_data,
    inject_phase4_data,
)


def _sanitize_filename(name: str) -> str:
    return re.sub(r'[^A-Za-z0-9 _.\-]', '', name).strip() or "Output"


def render_phase4():
    st.header("Phase 4: Sanity Check & Export")

    # ----- Required state -----
    phase1 = st.session_state.get('phase1_data')
    deep_dive = st.session_state.get('deep_dive')
    lseg_peers = st.session_state.get('lseg_parsed_peers')
    selected_peers = st.session_state.get('final_selected_peers')
    not_selected_peers = st.session_state.get('final_not_selected_peers', [])
    rejection_rationales = st.session_state.get('rejection_rationales', {})
    selected_transactions = st.session_state.get('selected_transactions', [])
    available_years = st.session_state.get('selected_years', [])

    missing = []
    if not phase1: missing.append("Phase 1 (extraction)")
    if not selected_peers: missing.append("Phase 2 (peer selection)")
    if not deep_dive: missing.append("Phase 3 (deep dive)")
    if not lseg_peers: missing.append("Phase 3 (LSEG parsed data)")
    if missing:
        st.error(f"Cannot run Phase 4 — missing data from: {', '.join(missing)}")
        return

    latest_year = max(available_years) if available_years else 2024

    # ----- Step 4a: Sanity Check -----
    st.subheader("Step 4a: Sanity Check")

    if 'sanity_result' not in st.session_state:
        if st.button("Run Sanity Check", type="primary"):
            st.session_state.sanity_result = run_sanity_check(
                deep_dive, lseg_peers, phase1, latest_year
            )
            st.rerun()
        return

    result = st.session_state.sanity_result
    if result['critical_issues']:
        st.error("**Critical issues found:**")
        for issue in result['critical_issues']:
            st.write(f"- {issue}")
    if result['warnings']:
        st.warning(f"**{len(result['warnings'])} warning(s):**")
        for w in result['warnings']:
            st.write(f"- {w}")
    if not result['warnings'] and not result['critical_issues']:
        st.success("All checks passed — no flags.")

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("← Re-run Sanity Check"):
            del st.session_state.sanity_result
            st.rerun()
    with col_b:
        proceed = st.button(
            "Acknowledge & Continue to Export",
            type="primary",
            disabled=len(result['critical_issues']) > 0 and not st.session_state.get('override_critical'),
        )
    if result['critical_issues']:
        st.checkbox(
            "I understand the critical issues and want to export anyway",
            key='override_critical',
        )

    if not proceed and not st.session_state.get('sanity_acknowledged'):
        return

    if proceed:
        st.session_state.sanity_acknowledged = True

    # ----- Step 4b: Export -----
    st.divider()
    st.subheader("Step 4b: Export")

    # Debug report — always visible inline, generates on-demand
    def _make_debug_report():
        state_snapshot = {
            "phase1_data": st.session_state.get("phase1_data"),
            "selected_years": st.session_state.get("selected_years"),
            "peer_list": st.session_state.get("peer_list"),
            "final_selected_peers": st.session_state.get("final_selected_peers"),
            "final_not_selected_peers": st.session_state.get("final_not_selected_peers"),
            "rejection_rationales": st.session_state.get("rejection_rationales"),
            "lseg_parsed_peers": st.session_state.get("lseg_parsed_peers"),
            "deep_dive": st.session_state.get("deep_dive"),
            "selected_transactions": st.session_state.get("selected_transactions"),
            "sanity_result": st.session_state.get("sanity_result"),
        }
        return generate_debug_report(state_snapshot)

    _deal_code_slug = re.sub(r'\D', '', phase1.get('deal_code') or 'XXX') or 'XXX'
    _client_slug = re.sub(r'[^A-Za-z0-9]', '_', phase1.get('client_name', 'Company'))
    st.download_button(
        label="🛠 Download Debug Report (.json)",
        data=_make_debug_report(),
        file_name=f"debug_report_DF{_deal_code_slug}_{_client_slug}.json",
        mime="application/json",
    )
    st.caption("Share this file with Claude to diagnose extraction errors, missing values, or template injection issues.")

    st.divider()

    if 'export_result' in st.session_state:
        result = st.session_state.export_result
        st.success(f"Export complete: **{result['filename']}**")
        st.download_button(
            "⬇ Download .xlsx",
            data=result['bytes'],
            file_name=result['filename'],
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type="primary",
        )
        if result.get('drive_id'):
            drive_url = f"https://drive.google.com/file/d/{result['drive_id']}/view"
            st.markdown(f"📁 Also uploaded to Drive: [{result['filename']}]({drive_url})")
        if st.button("Generate Again"):
            del st.session_state.export_result
            st.rerun()
        return

    if st.button("Generate Excel & Upload to Drive", type="primary"):
        try:
            with st.spinner("Downloading template from Google Drive..."):
                template_bytes = download_template()

            with st.spinner("Injecting data into all tabs..."):
                workbook = openpyxl.load_workbook(template_bytes, data_only=False)

                inject_phase1_data(workbook, phase1, available_years)
                inject_phase2_data(workbook, selected_peers, not_selected_peers, rejection_rationales)
                inject_phase3_data(
                    workbook,
                    deep_dive,
                    lseg_peers,
                    selected_peers,
                    phase1.get('deal_code', 'DF-XXX'),
                    phase1.get('client_name', 'Company'),
                    latest_year,
                )
                if selected_transactions:
                    inject_phase35_data(workbook, selected_transactions, phase1.get('deal_code', 'DF-XXX'))
                inject_phase4_data(workbook, phase1, latest_year)

                # Write to bytes for download + temp file for Drive upload
                buf = io.BytesIO()
                workbook.save(buf)
                buf.seek(0)
                xlsx_bytes = buf.read()

            deal_code = phase1.get('deal_code', 'DFXXX')
            deal_num = re.sub(r'\D', '', deal_code) or 'XXX'
            client_name = phase1.get('client_name', 'Company')
            filename = _sanitize_filename(f"Trading Performance DF{deal_num} - {client_name}") + ".xlsx"

            drive_id = None
            with st.spinner("Uploading to Google Drive..."):
                with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                    tmp.write(xlsx_bytes)
                    tmp_path = tmp.name
                try:
                    drive_id = upload_output(tmp_path, filename)
                except Exception as e:
                    st.warning(f"Drive upload failed (file is still available for download): {e}")
                finally:
                    try:
                        os.unlink(tmp_path)
                    except OSError:
                        pass

            st.session_state.export_result = {
                'filename': filename,
                'bytes': xlsx_bytes,
                'drive_id': drive_id,
            }
            st.rerun()
        except Exception as e:
            st.error(f"Export failed: {e}")
            raise
