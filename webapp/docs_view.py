"""Documentation: the user, technical and governance guides, read from docs/."""

from pathlib import Path

import streamlit as st

DOCS_DIR = Path(__file__).resolve().parent.parent / "docs"

# key -> (button / page title, file stem, who it is for)
GUIDES = {
    "user": ("User Guide", "user_guide",
             "For anyone running the tool: files, setup, the six steps, exceptions and fixes."),
    "technical": ("Technical Guide", "technical_guide",
                  "For maintainers: rules, month placement, anchors, code map and safe changes."),
    "governance": ("Governance", "governance",
                   "Ownership, controlled items, data protection and change management (draft)."),
}


def render_doc_buttons() -> None:
    """Main page: three separate, centered buttons; each opens its own guide."""
    st.markdown("<h3 style='text-align:center;'>Documentation</h3>", unsafe_allow_html=True)
    _, c1, c2, c3, _ = st.columns([1, 1.4, 1.4, 1.4, 1], gap="medium")
    for col, (key, (label, _stem, _blurb)) in zip((c1, c2, c3), GUIDES.items()):
        with col:
            if st.button(label, key=f"open_doc_{key}", use_container_width=True):
                st.session_state.view = "docs"
                st.session_state.doc = key
                st.rerun()


def _on_switch() -> None:
    """Switcher callback: open the picked guide; never leave the control empty."""
    picked = st.session_state.get("doc_switch")
    if picked:
        st.session_state.doc = picked
    else:
        st.session_state.doc_switch = st.session_state.get("doc", "user")


def render_doc_switcher(active: str) -> None:
    """Guide pages: one clean segmented control to move between the guides."""
    _, mid, _ = st.columns([1, 2.4, 1])
    with mid:
        st.segmented_control(
            "Guides", list(GUIDES), selection_mode="single", default=[active],
            format_func=lambda k: GUIDES[k][0], key="doc_switch",
            label_visibility="collapsed", width="stretch", on_change=_on_switch,
        )


def render_guide(key: str) -> None:
    """Show one guide with a PDF download button."""
    label, stem, blurb = GUIDES[key]
    st.caption(blurb)
    pdf_path, md_path = DOCS_DIR / f"{stem}.pdf", DOCS_DIR / f"{stem}.md"
    if pdf_path.exists():
        st.download_button(f"Download {label} (PDF)", data=pdf_path.read_bytes(),
                           file_name=f"Financial Automation - {label}.pdf", mime="application/pdf",
                           key=f"dl_{stem}")
    if md_path.exists():
        text = md_path.read_text(encoding="utf-8")
        # The file opens with a label line ("User Guide"); the page title already says it.
        first, _, rest = text.partition("\n")
        st.markdown(rest.lstrip() if first.strip() == label else text)
    else:
        st.warning(f"{stem}.md was not found in the docs folder.")
