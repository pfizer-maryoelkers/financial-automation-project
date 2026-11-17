"""Documentation page: the user, technical and governance guides, read from docs/."""

from pathlib import Path

import streamlit as st

from webapp.pfizer_theme import apply_pfizer_theme, render_footer, render_header

DOCS_DIR = Path(__file__).resolve().parent.parent / "docs"

# (tab label, file, who it is for)
GUIDES = (
    ("User Guide", "user_guide.md", "For anyone running the tool: files, setup, the six steps, exceptions and fixes."),
    ("Technical Guide", "technical_guide.md", "For maintainers: rules, month placement, anchors, code map and safe changes."),
    ("Governance", "governance.md", "Ownership, controlled items, data protection and change management (draft)."),
)

st.set_page_config(page_title="Documentation - Financial Automation", layout="wide",
                   initial_sidebar_state="collapsed")
apply_pfizer_theme()
render_header("Documentation")

st.page_link("app.py", label="Back to the report generator", icon=":material/arrow_back:")
st.divider()

tabs = st.tabs([label for label, _, _ in GUIDES])
for tab, (label, filename, blurb) in zip(tabs, GUIDES):
    with tab:
        st.caption(blurb)
        path = DOCS_DIR / filename
        if path.exists():
            text = path.read_text(encoding="utf-8")
            st.download_button(f"Download {label} (.md)", data=text, file_name=filename,
                               mime="text/markdown", key=f"dl_{filename}")
            st.markdown(text)
        else:
            st.warning(f"{filename} was not found in the docs folder.")

render_footer("Financial Automation Report Generator v1.0")
