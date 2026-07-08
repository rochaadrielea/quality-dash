"""
auth.py -- password gate for Streamlit Community Cloud.
Password is read from st.secrets["app_password"], set in the app's
Settings -> Secrets (never stored in the repo).
"""

import hashlib
import streamlit as st


def _hash(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def check_password() -> bool:
    # already signed in this session
    if st.session_state.get("authenticated"):
        return True

    stored = st.secrets.get("app_password", None)
    if stored is None:
        st.error("No password configured. Set 'app_password' in the app's Secrets.")
        return False

    st.title("Quality BRM Dashboard")
    st.caption("Please sign in to continue")
    entered = st.text_input("Password", type="password", key="pw_input")

    if st.button("Sign in"):
        if entered == stored:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Wrong password.")
    return False