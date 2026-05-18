import streamlit as st

st.set_page_config(
    page_title="MAX Solutions Valuation Engine",
    page_icon="📊",
    layout="wide"
)

# Initialize Session State
if "current_phase" not in st.session_state:
    st.session_state.current_phase = 1

# Keys that belong to each phase — used when navigating back to clear downstream state
_PHASE_KEYS = {
    1: ["phase1_data", "selected_years", "owner_questions"],
    2: [
        "peer_list", "peer_df", "peer_df_base",
        "top_12_confirmed", "top_12_df",
        "final_df_base", "final_peers_confirmed",
        "final_selected_peers", "final_not_selected_peers",
        "rejection_rationales",
    ],
    3: [
        "lseg_parsed_peers", "lseg_failed_files",
        "deep_dive", "all_transactions", "selected_transactions",
    ],
    4: [
        "sanity_result", "sanity_acknowledged",
        "override_critical", "export_result",
    ],
}


def _clear_from_phase(from_phase: int):
    """Clear session state for all phases >= from_phase."""
    for phase_num in range(from_phase, 5):
        for key in _PHASE_KEYS.get(phase_num, []):
            st.session_state.pop(key, None)


def _has_downstream_data(target_phase: int) -> bool:
    """Return True if navigating back to target_phase would erase meaningful work."""
    for phase_num in range(target_phase + 1, 5):
        for key in _PHASE_KEYS.get(phase_num, []):
            if key in st.session_state:
                return True
    return False


def main():
    # ── Tool selector ────────────────────────────────────────────────────────
    tool = st.sidebar.radio(
        "Select Tool",
        ["Comparable Valuation", "Financial Consolidation"],
        key="active_tool",
    )

    if tool == "Financial Consolidation":
        st.title("MAX Solutions — Financial Consolidation")
        from ui.consolidation import render_consolidation
        render_consolidation()
        return

    st.title("MAX Solutions Valuation Engine")

    phases = [
        "Phase 1: Upload & Extract",
        "Phase 2: Peer Screening",
        "Phase 3: Deep Dive",
        "Phase 4: Export",
    ]

    # ── Navigation bar ────────────────────────────────────────────────────────
    cols = st.columns(len(phases))
    for i, phase in enumerate(phases):
        phase_num = i + 1
        with cols[i]:
            if st.session_state.current_phase == phase_num:
                # Current phase — not clickable, just highlighted
                st.markdown(f"**🟢 {phase}**")

            elif st.session_state.current_phase > phase_num:
                # Completed phase — clickable to go back
                if st.button(f"✅ {phase}", key=f"nav_back_{phase_num}", use_container_width=True):
                    # Ask for confirmation only when going back would erase downstream data
                    if _has_downstream_data(phase_num):
                        st.session_state["_nav_back_confirm_target"] = phase_num
                    else:
                        _clear_from_phase(phase_num + 1)
                        st.session_state.current_phase = phase_num
                        st.rerun()

            else:
                # Future phase — greyed out
                st.markdown(f"⚪ {phase}")

    # ── Back-navigation confirmation dialog ───────────────────────────────────
    target = st.session_state.get("_nav_back_confirm_target")
    if target is not None:
        target_label = phases[target - 1]
        st.warning(
            f"Going back to **{target_label}** will clear all work done in the later phases. "
            "This cannot be undone."
        )
        c1, c2, _ = st.columns([1, 1, 4])
        with c1:
            if st.button("Yes, go back", type="primary", key="confirm_go_back"):
                _clear_from_phase(target + 1)
                st.session_state.current_phase = target
                del st.session_state["_nav_back_confirm_target"]
                st.rerun()
        with c2:
            if st.button("Cancel", key="cancel_go_back"):
                del st.session_state["_nav_back_confirm_target"]
                st.rerun()
        return  # Don't render the phase content while the dialog is open

    st.divider()

    # ── Route to the correct UI module ────────────────────────────────────────
    if st.session_state.current_phase == 1:
        from ui.phase1_upload import render_phase1
        render_phase1()
    elif st.session_state.current_phase == 2:
        from ui.phase2_screening import render_phase2
        render_phase2()
    elif st.session_state.current_phase == 3:
        from ui.phase3_deepdive import render_phase3
        render_phase3()
    elif st.session_state.current_phase == 4:
        from ui.phase4_export import render_phase4
        render_phase4()


if __name__ == "__main__":
    main()
