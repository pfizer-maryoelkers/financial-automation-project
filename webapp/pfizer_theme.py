"""Pfizer Brand Identity 2.0 styling for the Streamlit app.

Colours, typography, buttons, cards and the logo follow the Pfizer brand guide.
Fonts and logos live in ./static and are served by Streamlit at /app/static/
(enabled in .streamlit/config.toml).
"""

import base64
import functools
import re
from pathlib import Path

import streamlit as st

# Logos and fonts are embedded straight into the page as data URIs, so they show
# up no matter how the app is launched or hosted (no dependence on Streamlit's
# static file serving, the working directory or a URL prefix).
_STATIC_DIR = Path(__file__).resolve().parent.parent / "static"


@functools.lru_cache(maxsize=None)
def _data_uri(relative_path: str, mime: str) -> str:
    """Base64 data URI for a file under ./static ('' if the file is missing)."""
    try:
        raw = (_STATIC_DIR / relative_path).read_bytes()
    except OSError:
        return ""
    return f"data:{mime};base64,{base64.b64encode(raw).decode('ascii')}"

_CSS = """
<style>
@font-face { font-family: 'Pfizer Tomorrow'; font-weight: 400; font-display: swap;
    src: url('__S__/fonts/PfizerTomorrow-Regular.woff2') format('woff2'); }
@font-face { font-family: 'Pfizer Tomorrow'; font-weight: 700; font-display: swap;
    src: url('__S__/fonts/PfizerTomorrow-Bold.woff2') format('woff2'); }
@font-face { font-family: 'Pfizer Diatype'; font-weight: 300; font-display: swap;
    src: url('__S__/fonts/PfizerDiatype-Light.woff2') format('woff2'); }
@font-face { font-family: 'Pfizer Diatype'; font-weight: 400; font-display: swap;
    src: url('__S__/fonts/PfizerDiatype-Regular.woff2') format('woff2'); }
@font-face { font-family: 'Pfizer Diatype'; font-weight: 500; font-display: swap;
    src: url('__S__/fonts/PfizerDiatype-Medium.woff2') format('woff2'); }
@font-face { font-family: 'Pfizer Diatype'; font-weight: 700; font-display: swap;
    src: url('__S__/fonts/PfizerDiatype-Bold.woff2') format('woff2'); }

:root {
    --pfizer-blue-70: #0000C9;
    --pfizer-blue-90: #000067;
    --pfizer-blue-10: #E8F2FF;
    --pfizer-blue-05: #EEF5FB;
    --pfizer-blue-03: #ECF6FF;
    --pfizer-border:  #d4e4f7;
    --pfizer-caption: #4a4a6a;
}

/* ── Typography ───────────────────────────────────────────────────────── */
html, body, [class*="css"], p, label, input, textarea, button, li, td, th,
[data-testid="stMarkdownContainer"], [data-testid="stCaptionContainer"] {
    font-family: 'Pfizer Diatype', 'Helvetica Neue', Arial, sans-serif !important;
    color: var(--pfizer-blue-90);
}
h1, h2, h3, h4,
[data-testid="stHeading"] {
    font-family: 'Pfizer Tomorrow', 'Franklin Gothic Medium', Arial, sans-serif !important;
    color: var(--pfizer-blue-90) !important;
    letter-spacing: 0;
}
h2 { font-size: 1.35rem !important; line-height: 1.2 !important; }
h3 { font-size: 1.1rem !important;  line-height: 1.2 !important; font-weight: 700 !important; }
[data-testid="stCaptionContainer"], small { color: var(--pfizer-caption) !important; }

/* ── Layout ───────────────────────────────────────────────────────────── */
/* Streamlit's 60px top toolbar used to cover the page top, which is why content
   had to start low.  Make the bar see-through (its Deploy / menu buttons stay
   clickable) so the title can sit close to the top edge. */
[data-testid="stHeader"] { background: transparent !important; pointer-events: none; }
[data-testid="stHeader"] button, [data-testid="stHeader"] a,
[data-testid="stToolbar"] { pointer-events: auto; }
.block-container { padding-top: 1.25rem !important; max-width: 1280px; }
[data-testid="stDivider"] hr, hr { border-color: var(--pfizer-border) !important; }

/* ── Header ───────────────────────────────────────────────────────────── */
.pf-header { display: flex; flex-direction: row; align-items: center;
    justify-content: center; gap: 10px; padding: 8px 0 18px 0; flex-wrap: wrap; }
/* The spiral matches the title: as tall as its capital letters (31px at the
   locked 44px title size) and shifted down so top and bottom line up with them */
.pf-header img { height: 31px; width: auto; position: relative; top: 4px; }
/* Streamlit adds an invisible anchor-link icon after headings; it pushes the
   spiral away from the title, so remove it from our own header */
.pf-title a, .pf-header [data-testid="stHeadingActionElements"] { display: none !important; }
.pf-title { font-family: 'Pfizer Tomorrow', Arial, sans-serif; font-weight: 700;
    color: var(--pfizer-blue-90); margin: 0; }
.pf-header .pf-title { font-size: 2.75rem !important; line-height: 1.15 !important; }
.pf-tagline { color: var(--pfizer-caption); font-size: 0.9rem; margin: 2px 0 0 0; }

/* ── Buttons: Electric Blue pills ─────────────────────────────────────── */
.stButton > button, [data-testid="stDownloadButton"] > button,
[data-testid="stFormSubmitButton"] > button {
    border-radius: 100px !important;
    border: 2px solid var(--pfizer-blue-70) !important;
    background: transparent !important;
    color: var(--pfizer-blue-70) !important;
    font-weight: 600 !important;
    padding: 0.5rem 1.6rem !important;
}
.stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
    background: var(--pfizer-blue-70) !important;
    color: #FFFFFF !important;
}
.stButton > button[kind="primary"], [data-testid="stDownloadButton"] > button,
[data-testid="stFormSubmitButton"] > button[kind="primary"] {
    background: var(--pfizer-blue-70) !important;
    color: #FFFFFF !important;
}
.stButton > button[kind="primary"]:hover, [data-testid="stDownloadButton"] > button:hover {
    background: #0000a3 !important;
    border-color: #0000a3 !important;
}
/* Button labels sit inside Streamlit's markdown container, which the global text
   colour above would turn navy — make every part of the label follow the button. */
.stButton > button [data-testid="stMarkdownContainer"],
.stButton > button [data-testid="stMarkdownContainer"] p,
[data-testid="stDownloadButton"] > button [data-testid="stMarkdownContainer"],
[data-testid="stDownloadButton"] > button [data-testid="stMarkdownContainer"] p,
[data-testid="stFormSubmitButton"] > button [data-testid="stMarkdownContainer"] p {
    color: inherit !important;
}
/* Disabled: pale-blue pill with readable navy-grey text (never white-on-pale) */
.stButton > button:disabled, .stButton > button:disabled:hover,
[data-testid="stDownloadButton"] > button:disabled {
    background: var(--pfizer-blue-10) !important;
    border-color: var(--pfizer-border) !important;
    color: var(--pfizer-caption) !important;
    opacity: 1 !important;
    cursor: not-allowed;
}
.stButton > button:disabled p, [data-testid="stDownloadButton"] > button:disabled p {
    color: var(--pfizer-caption) !important;
}

/* Toggle-style and add buttons keep the same pill look */
div.add-type-toggle .stButton > button, div.add-btn .stButton > button {
    border-radius: 100px !important;
}

/* ── Alerts: pale-blue cards with a top border (no side stripes) ──────── */
[data-testid="stAlert"] {
    border: none !important;
    border-top: 3px solid var(--pfizer-blue-70) !important;
    border-radius: 16px !important;
    background: var(--pfizer-blue-10) !important;
    color: var(--pfizer-blue-90) !important;
}
[data-testid="stAlert"] p, [data-testid="stAlert"] div { color: var(--pfizer-blue-90) !important; }
[data-testid="stAlert"] svg { fill: var(--pfizer-blue-70) !important; }
/* Warnings: navy fill so they stand out without warm colours */
[data-testid="stAlert"][kind="warning"], div[data-baseweb="notification"][kind="warning"] {
    background: var(--pfizer-blue-90) !important;
    border-top-color: var(--pfizer-blue-90) !important;
}
[data-testid="stAlert"][kind="warning"] p, [data-testid="stAlert"][kind="warning"] div,
[data-testid="stAlert"][kind="warning"] span { color: #FFFFFF !important; }
[data-testid="stAlert"][kind="warning"] svg { fill: #FFFFFF !important; }
/* Errors stay clearly distinguishable for accessibility */
[data-testid="stAlert"][kind="error"], div[data-baseweb="notification"][kind="negative"] {
    background: #FFFFFF !important;
    border: 1px solid #B3261E !important;
    border-top: 3px solid #B3261E !important;
}
[data-testid="stAlert"][kind="error"] svg { fill: #B3261E !important; }

/* ── File uploader ────────────────────────────────────────────────────── */
[data-testid="stFileUploaderDropzone"] {
    background: var(--pfizer-blue-05) !important;
    border: 2px dashed var(--pfizer-blue-70) !important;
    border-radius: 16px !important;
}
[data-testid="stFileUploaderDropzone"]:hover { background: var(--pfizer-blue-10) !important; }
[data-testid="stFileUploaderDropzone"] button, [data-testid="baseButton-secondary"] {
    border-radius: 100px !important;
    border: 2px solid var(--pfizer-blue-70) !important;
    background: var(--pfizer-blue-70) !important;
    color: #FFFFFF !important;
}
[data-testid="stFileUploaderDropzone"] button:hover { background: #0000a3 !important; }
[data-testid="stFileUploaderFile"] {
    background: #FFFFFF !important;
    border: 1px solid var(--pfizer-border) !important;
    border-radius: 12px !important;
}
[data-testid="stFileUploaderProgressBar"] > div { background-color: var(--pfizer-blue-70) !important; }

/* ── Inputs, selects, tags ────────────────────────────────────────────── */
[data-baseweb="select"] > div, [data-baseweb="input"] > div, textarea {
    border-radius: 12px !important;
    border-color: var(--pfizer-border) !important;
    background: #FFFFFF !important;
}
[data-testid="stMultiSelect"] span[data-baseweb="tag"] {
    background-color: var(--pfizer-blue-70) !important;
    border-radius: 100px !important;
    color: #FFFFFF !important;
}
[data-testid="stMultiSelect"] span[data-baseweb="tag"] svg { fill: #FFFFFF !important; }
[data-testid="stMultiSelect"] li[aria-selected="true"], [data-testid="stMultiSelect"] li:hover {
    background-color: var(--pfizer-blue-10) !important;
    color: var(--pfizer-blue-90) !important;
}

/* ── Cards: expanders, tabs, tables ───────────────────────────────────── */
[data-testid="stExpander"], [data-testid="stVerticalBlockBorderWrapper"] {
    border: 1px solid var(--pfizer-border) !important;
    border-radius: 16px !important;
    background: #FFFFFF;
}
[data-testid="stExpander"] summary { font-weight: 500; }
[data-testid="stTabs"] button[aria-selected="true"] {
    color: var(--pfizer-blue-70) !important;
    border-bottom-color: var(--pfizer-blue-70) !important;
}
[data-testid="stDataFrame"] { border: 1px solid var(--pfizer-border); border-radius: 12px; }
[data-testid="stProgress"] > div > div > div { background-color: var(--pfizer-blue-70) !important; }
[data-testid="stSpinner"] svg, [data-testid="stSpinner"] i { border-top-color: var(--pfizer-blue-70) !important; }

/* ── Template-type bar (OpEx / Project): full-width, high contrast ────── */
.pf-bar { display: flex; align-items: center; justify-content: space-between;
    gap: 16px; flex-wrap: wrap; width: 100%; box-sizing: border-box;
    padding: 20px 32px; margin: 16px 0 8px 0; border-radius: 16px; color: #FFFFFF; }
.pf-bar-opex { background: var(--pfizer-blue-70); }
.pf-bar-project { background: var(--pfizer-blue-90); }
.pf-bar-label { font-size: 0.8rem; letter-spacing: 0.08em; text-transform: uppercase;
    font-weight: 500; color: #FFFFFF !important; opacity: 0.85; }
.pf-bar-value { font-family: 'Pfizer Tomorrow', Arial, sans-serif; font-weight: 700;
    font-size: 1.5rem; line-height: 1.1; color: #FFFFFF !important; }
.pf-bar-sub { font-size: 0.95rem; color: #FFFFFF !important; }

/* ── Footer ───────────────────────────────────────────────────────────── */
.pf-footer { background: var(--pfizer-blue-90); color: #FFFFFF; border-radius: 16px;
    padding: 28px 32px; margin-top: 32px; display: flex; align-items: center;
    justify-content: space-between; gap: 24px; flex-wrap: wrap; }
.pf-footer img { height: 24px; width: auto; }
.pf-footer p { color: #FFFFFF !important; margin: 0; font-size: 0.85rem; }
</style>
"""


def apply_pfizer_theme() -> None:
    """Inject the Pfizer brand CSS (fonts, colours, pill buttons, cards)."""
    css = re.sub(
        r"__S__/fonts/([\w-]+\.woff2)",
        lambda m: _data_uri(f"fonts/{m.group(1)}", "font/woff2"),
        _CSS,
    )
    st.markdown(css, unsafe_allow_html=True)


def render_header(title: str, tagline: str = "") -> None:
    """Centered page title with the Pfizer spiral right after it."""
    tag = f'<p class="pf-tagline">{tagline}</p>' if tagline else ""
    st.markdown(
        f"""
        <div class="pf-header">
            <div style="text-align:center;">
                <h1 class="pf-title">{title}</h1>{tag}
            </div>
            <img src="{_data_uri('Pfizer_Spiral_Small.png', 'image/png')}" alt="Pfizer" />
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_template_pill(is_project: bool) -> None:
    """Full-width bar showing which template type was detected."""
    if is_project:
        value, sub, cls = "Project / P3 template", "P3 ID and WBS pipeline", "pf-bar-project"
    else:
        value, sub, cls = "OpEx template", "Cost center pipeline", "pf-bar-opex"
    st.markdown(
        f"""
        <div class="pf-bar {cls}">
            <div>
                <div class="pf-bar-label">Template detected</div>
                <div class="pf-bar-value">{value}</div>
            </div>
            <div class="pf-bar-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_footer(text: str) -> None:
    """Navy footer with the white Pfizer logo."""
    st.markdown(
        f"""
        <div class="pf-footer">
            <img src="{_data_uri('Pfizer_Logo_White_Small.png', 'image/png')}" alt="Pfizer" />
            <p>{text}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
