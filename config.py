# config.py
# Reads secrets from Streamlit Cloud secrets (st.secrets) or falls back to
# a local config_local.py file for development.

import json

def _get_secret(key: str, default=None):
    """Try st.secrets first, then config_local.py, then default."""
    # 1. Streamlit secrets (available on Streamlit Cloud and local .streamlit/secrets.toml)
    try:
        import streamlit as st
        val = st.secrets.get(key)
        if val is not None:
            return val
    except Exception:
        pass

    # 2. Local config file (for development — git-ignored)
    try:
        import config_local
        val = getattr(config_local, key, None)
        if val is not None:
            return val
    except ImportError:
        pass

    return default


# ── API Keys ─────────────────────────────────────────────────────────────────
GEMINI_API_KEY = _get_secret("GEMINI_API_KEY", "")

# ── Google Drive ─────────────────────────────────────────────────────────────
DRIVE_TEMPLATE_FILE_ID = _get_secret("DRIVE_TEMPLATE_FILE_ID", "")
DRIVE_OUTPUT_FOLDER_ID = _get_secret("DRIVE_OUTPUT_FOLDER_ID", "")

# ── Google Service Account ───────────────────────────────────────────────────
# On Streamlit Cloud: stored as a JSON string in st.secrets["SERVICE_ACCOUNT_JSON"]
# Locally: stored as a dict in config_local.py SERVICE_ACCOUNT_INFO
def _load_service_account():
    # Try Streamlit secrets (JSON string)
    try:
        import streamlit as st
        sa_json = st.secrets.get("SERVICE_ACCOUNT_JSON")
        if sa_json:
            if isinstance(sa_json, str):
                return json.loads(sa_json)
            # st.secrets may already parse it as a dict-like object
            return dict(sa_json)
    except Exception:
        pass

    # Try local config
    try:
        import config_local
        sa = getattr(config_local, "SERVICE_ACCOUNT_INFO", None)
        if sa:
            return sa
    except ImportError:
        pass

    return {}

SERVICE_ACCOUNT_INFO = _load_service_account()
