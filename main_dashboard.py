import base64
import os
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Exam Timetable Tools",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed"
)


def _get_logo_data_uri():
    """Load app_logo.png (NMIMS logo) once and return it as a base64 data URI."""
    for candidate in (
        os.path.join(os.path.dirname(__file__), "app_logo.png"),
        os.path.join(os.path.dirname(__file__), "assets", "app_logo.png"),
        "app_logo.png",
        "assets/app_logo.png",
    ):
        if os.path.exists(candidate):
            with open(candidate, "rb") as f:
                encoded = base64.b64encode(f.read()).decode()
            return f"data:image/png;base64,{encoded}"
    return None


def _get_brand_theme():
    """
    Fixed NMIMS brand look: white background with red/dark text, always —
    intentionally not reactive to Streamlit's light/dark toggle.
    """
    return {
        "page_bg": "#ffffff",
        "card_bg": "#f8f9fa",
        "card_border": "rgba(21,21,21,0.10)",
        "card_hover_bg": "#f1f1f3",
        "text_strong": "#1a1a1a",
        "text_soft": "rgba(26,26,26,0.62)",
        "text_faint": "rgba(26,26,26,0.48)",
        "arrow": "rgba(26,26,26,0.30)",
        "arrow_hover": "rgba(151,28,28,0.75)",
        "footer": "rgba(26,26,26,0.40)",
        "scrollbar_track": "rgba(0,0,0,0.04)",
        "scrollbar_thumb": "rgba(0,0,0,0.18)",
        "scrollbar_thumb_hover": "rgba(0,0,0,0.32)",
        "app_gradient": "linear-gradient(135deg, #ffffff 0%, #fdf3f0 100%)",
    }


_theme = _get_brand_theme()
_logo_uri = _get_logo_data_uri()

st.markdown(f"""
<style>
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
header {{visibility: hidden;}}
[data-testid="collapsedControl"] {{display: none;}}
section[data-testid="stSidebar"] {{display: none;}}
.block-container {{padding: 0 !important; margin: 0 !important; max-width: 100% !important;}}
.stApp {{background: {_theme['app_gradient']} !important;}}
</style>
""", unsafe_allow_html=True)

_logo_html = f'<img src="{_logo_uri}" class="header-logo" alt="NMIMS Logo">' if _logo_uri else ""
_divider_html = '<div class="header-divider"></div>' if _logo_uri else ""

components.html(f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  :root {{
    --page-bg: {_theme['page_bg']};
    --card-bg: {_theme['card_bg']};
    --card-border: {_theme['card_border']};
    --card-hover-bg: {_theme['card_hover_bg']};
    --text-strong: {_theme['text_strong']};
    --text-soft: {_theme['text_soft']};
    --text-faint: {_theme['text_faint']};
    --arrow: {_theme['arrow']};
    --arrow-hover: {_theme['arrow_hover']};
    --footer-color: {_theme['footer']};
    --scrollbar-track: {_theme['scrollbar_track']};
    --scrollbar-thumb: {_theme['scrollbar_thumb']};
    --scrollbar-thumb-hover: {_theme['scrollbar_thumb_hover']};
    --accent-1: #C73E1D;
    --accent-2: #951C1C;
  }}

  * {{ margin: 0; padding: 0; box-sizing: border-box; }}

  /* Custom sleek scrollbar for the iframe */
  ::-webkit-scrollbar {{ width: 6px; }}
  ::-webkit-scrollbar-track {{ background: var(--scrollbar-track); }}
  ::-webkit-scrollbar-thumb {{ background: var(--scrollbar-thumb); border-radius: 10px; }}
  ::-webkit-scrollbar-thumb:hover {{ background: var(--scrollbar-thumb-hover); }}

  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: var(--page-bg); /* Let the Streamlit background show through */
    min-height: 100vh;
    padding: 2rem 1.5rem 3rem 1.5rem;
    overflow-x: hidden;
  }}

  .header-bar {{
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 1.5rem;
    flex-wrap: wrap;
    max-width: 1000px;
    margin: 0 auto;
    animation: fadeInDown 0.6s ease;
  }}

  .header-logo {{
    height: 90px;
    width: auto;
    border-radius: 10px;
    background: #ffffff;
    padding: 8px 12px;
    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.2);
    flex-shrink: 0;
  }}

  .header-divider {{
    width: 2px;
    height: 70px;
    background: var(--card-border);
    flex-shrink: 0;
  }}

  .header-text {{
    text-align: left;
    min-width: 0;
    flex: 1 1 auto;
  }}

  .hero {{
    text-align: left;
    padding: 0;
    animation: fadeInDown 0.7s ease;
  }}

  .badge {{
    display: inline-block;
    background: rgba(151,28,28,0.12);
    color: #C73E1D;
    border: 1px solid rgba(151,28,28,0.35);
    border-radius: 50px;
    padding: 0.35rem 1.2rem;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    margin-bottom: 1.2rem;
  }}

  .hero-title {{
    font-size: 3rem;
    font-weight: 800;
    background: linear-gradient(90deg, #951C1C, #C73E1D, #E85D3F);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin-bottom: 0.6rem;
    letter-spacing: -1.5px;
    line-height: 1.15;
    overflow-wrap: break-word;
    word-break: break-word;
  }}

  .hero-sub {{
    font-size: 1.05rem;
    color: var(--text-soft);
    margin-bottom: 0;
  }}

  .hero-line {{
    width: 70px;
    height: 3px;
    background: linear-gradient(90deg, var(--accent-1), var(--accent-2));
    margin: 1.2rem 0 0 0;
    border-radius: 2px;
  }}

  .grid {{
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 1.4rem;
    max-width: 1000px;
    margin: 2.5rem auto 0 auto;
    animation: fadeInUp 0.8s ease;
  }}

  .card {{
    position: relative;
    background: var(--card-bg);
    border: 1px solid var(--card-border);
    border-radius: 20px;
    padding: 2rem 1.8rem 1.7rem 1.8rem;
    text-decoration: none;
    display: block;
    overflow: hidden;
    transition: transform 0.28s cubic-bezier(.4,2,.4,1),
                box-shadow 0.28s ease,
                border-color 0.28s ease,
                background 0.28s ease;
  }}

  .card::before {{
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    border-radius: 20px 20px 0 0;
    opacity: 0;
    transition: opacity 0.28s ease;
  }}

  .card:hover {{
    transform: translateY(-8px) scale(1.015);
    background: var(--card-hover-bg);
    text-decoration: none;
  }}
  .card:hover::before {{ opacity: 1; }}

  /* Card 1 — Deep Red */
  .c-purple::before {{ background: linear-gradient(90deg,#951C1C,#C73E1D); }}
  .c-purple:hover   {{ border-color: rgba(151,28,28,0.4); box-shadow: 0 16px 40px rgba(21,21,21,0.12), 0 0 60px rgba(151,28,28,0.15); }}
  .c-purple .icon   {{ background: rgba(151,28,28,0.15); border: 1px solid rgba(151,28,28,0.25); }}
  .c-purple .lbl    {{ color: #C73E1D; }}
  .c-purple .btn    {{ color: #C73E1D; border-color: rgba(151,28,28,0.35); background: rgba(151,28,28,0.08); }}
  .c-purple:hover .btn {{ background: rgba(151,28,28,0.20); }}

  /* Card 2 — Bright Red / Orange-red */
  .c-blue::before {{ background: linear-gradient(90deg,#C73E1D,#E85D3F); }}
  .c-blue:hover   {{ border-color: rgba(199,62,29,0.4); box-shadow: 0 16px 40px rgba(21,21,21,0.12), 0 0 60px rgba(199,62,29,0.15); }}
  .c-blue .icon   {{ background: rgba(199,62,29,0.15); border: 1px solid rgba(199,62,29,0.25); }}
  .c-blue .lbl    {{ color: #E85D3F; }}
  .c-blue .btn    {{ color: #E85D3F; border-color: rgba(199,62,29,0.35); background: rgba(199,62,29,0.08); }}
  .c-blue:hover .btn {{ background: rgba(199,62,29,0.20); }}

  /* Card 3 — Maroon */
  .c-green::before {{ background: linear-gradient(90deg,#7A1515,#A23217); }}
  .c-green:hover   {{ border-color: rgba(122,21,21,0.4); box-shadow: 0 16px 40px rgba(21,21,21,0.12), 0 0 60px rgba(122,21,21,0.15); }}
  .c-green .icon   {{ background: rgba(122,21,21,0.15); border: 1px solid rgba(122,21,21,0.25); }}
  .c-green .lbl    {{ color: #A23217; }}
  .c-green .btn    {{ color: #A23217; border-color: rgba(122,21,21,0.35); background: rgba(122,21,21,0.08); }}
  .c-green:hover .btn {{ background: rgba(122,21,21,0.20); }}

  /* Card 4 — Warm Ember */
  .c-orange::before {{ background: linear-gradient(90deg,#D9481F,#F2673D); }}
  .c-orange:hover   {{ border-color: rgba(217,72,31,0.4); box-shadow: 0 16px 40px rgba(21,21,21,0.12), 0 0 60px rgba(217,72,31,0.15); }}
  .c-orange .icon   {{ background: rgba(217,72,31,0.15); border: 1px solid rgba(217,72,31,0.25); }}
  .c-orange .lbl    {{ color: #F2673D; }}
  .c-orange .btn    {{ color: #F2673D; border-color: rgba(217,72,31,0.35); background: rgba(217,72,31,0.08); }}
  .c-orange:hover .btn {{ background: rgba(217,72,31,0.20); }}

  .icon {{
    width: 54px;
    height: 54px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.6rem;
    margin-bottom: 1.1rem;
  }}

  .lbl {{
    font-size: 0.67rem;
    font-weight: 700;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    margin-bottom: 0.45rem;
    display: block;
  }}

  .title {{
    font-size: 1.12rem;
    font-weight: 700;
    color: var(--text-strong);
    margin-bottom: 0.55rem;
    line-height: 1.35;
    display: block;
    overflow-wrap: break-word;
    word-break: break-word;
  }}

  .desc {{
    font-size: 0.83rem;
    color: var(--text-faint);
    line-height: 1.65;
    margin-bottom: 1.3rem;
    display: block;
  }}

  .btn {{
    display: inline-flex;
    align-items: center;
    gap: 6px;
    font-size: 0.78rem;
    font-weight: 600;
    padding: 0.45rem 1.1rem;
    border-radius: 50px;
    border: 1px solid;
    transition: background 0.2s;
    letter-spacing: 0.3px;
  }}

  .arrow {{
    position: absolute;
    top: 1.4rem;
    right: 1.5rem;
    color: var(--arrow);
    font-size: 1rem;
    transition: color 0.2s, transform 0.2s;
    font-style: normal;
    font-family: sans-serif;
  }}
  .card:hover .arrow {{
    color: var(--arrow-hover);
    transform: translate(3px,-3px);
  }}

  .footer {{
    text-align: center;
    margin-top: 2.5rem;
    color: var(--footer-color);
    font-size: 0.77rem;
    letter-spacing: 0.3px;
  }}

  @keyframes fadeInDown {{
    from {{ opacity:0; transform:translateY(-20px); }}
    to   {{ opacity:1; transform:translateY(0); }}
  }}
  @keyframes fadeInUp {{
    from {{ opacity:0; transform:translateY(20px); }}
    to   {{ opacity:1; transform:translateY(0); }}
  }}

  /* --- RESPONSIVE MEDIA QUERIES --- */

  /* Tablets and smaller screens */
  @media (max-width: 900px) {{
    .grid {{
        gap: 1rem;
        padding: 0 0.5rem;
    }}
  }}

  /* Mobile screens */
  @media (max-width: 650px) {{
    body {{
        padding: 1rem 0.5rem 2rem 0.5rem;
    }}
    .header-logo {{
        height: 56px;
    }}
    .header-bar {{
        justify-content: center;
        text-align: center;
    }}
    .header-text {{
        text-align: center;
    }}
    .hero-line {{
        margin: 1.2rem auto 0 auto;
    }}
    .hero {{
        text-align: center;
        padding: 1rem 0.5rem 0 0.5rem;
    }}
    .hero-title {{
        font-size: 2.2rem;
    }}
    .hero-sub {{
        font-size: 0.95rem;
    }}
    .grid {{
        grid-template-columns: 1fr;
        margin-top: 1.5rem;
    }}
    .card {{
        padding: 1.5rem;
    }}
    .icon {{
        width: 48px;
        height: 48px;
        font-size: 1.3rem;
        margin-bottom: 0.9rem;
    }}
    .title {{
        font-size: 1.05rem;
    }}
  }}
</style>
</head>
<body>

<div class="header-bar">
  {_logo_html}
  {_divider_html}
  <div class="header-text">
    <div class="badge">📚 &nbsp; Exam Management Suite</div>
    <div class="hero-title">Timetable Tools</div>
    <div class="hero-sub">All your exam scheduling and conversion tools — in one place.</div>
    <div class="hero-line"></div>
  </div>
</div>

<div class="grid">

  <a class="card c-purple" href="https://examtimetablescheduler-dqttcyf5vzakkjfpkt6xhp.streamlit.app/" target="_blank" rel="noopener noreferrer">
    <span class="arrow">↗</span>
    <div class="icon">🗓️</div>
    <span class="lbl">Scheduler</span>
    <span class="title">Final Exam Timetable Scheduler</span>
    <span class="desc">Automatically generate conflict-free final exam timetables with intelligent scheduling algorithms.</span>
    <span class="btn">Open Tool &nbsp;→</span>
  </a>

  <a class="card c-blue" href="https://reexamschedulerlatest.streamlit.app/" target="_blank" rel="noopener noreferrer">
    <span class="arrow">↗</span>
    <div class="icon">🔁</div>
    <span class="lbl">Re-Exam</span>
    <span class="title">Re-Exam Timetable Scheduler</span>
    <span class="desc">Schedule re-examination timetables seamlessly, handling student and room constraints with ease.</span>
    <span class="btn">Open Tool &nbsp;→</span>
  </a>

  <a class="card c-green" href="https://verification-file-change-to-pdf-converter.streamlit.app/" target="_blank" rel="noopener noreferrer">
    <span class="arrow">↗</span>
    <div class="icon">📄</div>
    <span class="lbl">PDF Converter</span>
    <span class="title">Final Exam Verification → PDF</span>
    <span class="desc">Convert final exam verification files into clean, formatted PDF documents ready for distribution.</span>
    <span class="btn">Open Tool &nbsp;→</span>
  </a>

  <a class="card c-orange" href="https://re-examtimetablescheduler-gndknuqn7whtdxe6cvubaw.streamlit.app/" target="_blank" rel="noopener noreferrer">
    <span class="arrow">↗</span>
    <div class="icon">🖨️</div>
    <span class="lbl">PDF Converter</span>
    <span class="title">Re-Exam File → PDF Converter</span>
    <span class="desc">Transform re-examination timetable files into professional PDFs with a single click.</span>
    <span class="btn">Open Tool &nbsp;→</span>
  </a>

</div>

<div class="footer">
  Exam Timetable Project &nbsp;·&nbsp; Click any card to open in a new tab
</div>

</body>
</html>
""", height=1480, scrolling=False)
