"""Documentation view: the user, technical and governance guides, read from docs/."""

from pathlib import Path

import streamlit as st

DOCS_DIR = Path(__file__).resolve().parent.parent / "docs"

# (tab label, file, who it is for)
GUIDES = (
    ("User Guide", "user_guide.md", "For anyone running the tool: files, setup, the six steps, exceptions and fixes."),
    ("Technical Guide", "technical_guide.md", "For maintainers: rules, month placement, anchors, code map and safe changes."),
    ("Governance", "governance.md", "Ownership, controlled items, data protection and change management (draft)."),
)


def render_docs() -> None:
    """Draw the three guides as tabs, each with a download button."""
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
