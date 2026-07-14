import io
import os
import tempfile
from datetime import datetime
from pathlib import Path

import fitz
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


st.set_page_config(
    page_title="StudyBook Builder",
    page_icon="📚",
    layout="wide",
)


def get_secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return os.getenv(name, default)


APP_PASSWORD = ***"APP_PASSWORD", "")


def check_password():
    if not APP_PASSWORD:
        *** True

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.title("📚 StudyBook Builder")
    password = st.text_input("Password", type="password")

    if st.button("Entra"):
        if password == APP_PASSWORD:
            st.ses…ated = True
            st.rerun()
        else:
            st.error("Password errata.")

    return False


if not check_password():
    st.stop()


st.title("📚 StudyBook Builder")
st.caption("Genera un unico documento Word da appunti, slide e testi di riferimento.")
