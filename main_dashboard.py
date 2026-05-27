import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Exam Timetable Tools",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
[data-testid="collapsedControl"] {display: none;}
section[data-testid="stSidebar"] {display: none;}
.block-container {padding: 0 !important; margin: 0 !important; max-width: 100% !important;}
.stApp {background: linear-gradient(135deg, #0f0c29, #302b63, #24243e) !important;}
</style>
""", unsafe_allow_html=True)

components.html("""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }

  /* Custom sleek scrollbar for the iframe */
  ::-webkit-scrollbar { width: 6px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.2); border-radius: 10px; }
  ::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.4); }

  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: transparent; /* Let the Streamlit background show through */
    min-height: 100vh;
    padding: 2rem 1.5rem 3rem 1.5rem;
    overflow-x: hidden;
  }

  .hero {
    text-align: center;
    padding: 2.5rem 1rem 1.5rem 1rem;
    animation: fadeInDown 0.7s ease;
  }

  .badge {
    display: inline-block;
    background: rgba(167,139,250,0.15);
    color: #a78bfa;
    border: 1px solid rgba(167,139,250,0.35);
    border-radius: 50px;
    padding: 0.35rem 1.2rem;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    margin-bottom: 1.2rem;
  }

  .hero-title {
    font-size: 3rem;
    font-weight: 800;
    background: linear-gradient(90deg, #a78bfa, #60a5fa, #34d399);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin-bottom: 0.6rem;
    letter-spacing: -1.5px;
    line-height: 1.15;
  }

  .hero-sub {
    font-size: 1.05rem;
    color: rgba(255,255,255,0.5);
    margin-bottom: 0;
  }

  .hero-line {
    width: 70px;
    height: 3px;
    background: linear-gradient(90deg, #a78bfa, #60a5fa);
    margin: 1.2rem auto 0 auto;
    border-radius: 2px;
  }

  .grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 1.4rem;
    max-width: 1000px;
    margin: 2.5rem auto 0 auto;
    animation: fadeInUp 0.8s ease;
  }

  .card {
    position: relative;
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 20px;
    padding: 2rem 1.8rem 1.7rem 1.8rem;
    text-decoration: none;
    display: block;
    overflow: hidden;
    transition: transform 0.28s cubic-bezier(.4,2,.4,1),
                box-shadow 0.28s ease,
                border-color 0.28s ease,
                background 0.28s ease;
  }

  .card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    border-radius: 20px 20px 0 0;
    opacity: 0;
    transition: opacity 0.28s ease;
  }

  .card:hover {
    transform: translateY(-8px) scale(1.015);
    background: rgba(255,255,255,0.09);
    text-decoration: none;
  }
  .card:hover::before { opacity: 1; }

  /* Purple */
  .c-purple::before { background: linear-gradient(90deg,#a78bfa,#818cf8); }
  .c-purple:hover   { border-color: rgba(167,139,250,0.4); box-shadow: 0 24px 60px rgba(0,0,0,0.5), 0 0 60px rgba(167,139,250,0.15); }
  .c-purple .icon   { background: rgba(167,139,250,0.15); border: 1px solid rgba(167,139,250,0.25); }
  .c-purple .lbl    { color: #a78bfa; }
  .c-purple .btn    { color: #a78bfa; border-color: rgba(167,139,250,0.35); background: rgba(167,139,250,0.08); }
  .c-purple:hover .btn { background: rgba(167,139,250,0.20); }

  /* Blue */
  .c-blue::before { background: linear-gradient(90deg,#60a5fa,#38bdf8); }
  .c-blue:hover   { border-color: rgba(96,165,250,0.4); box-shadow: 0 24px 60px rgba(0,0,0,0.5), 0 0 60px rgba(96,165,250,0.15); }
  .c-blue .icon   { background: rgba(96,165,250,0.15); border: 1px solid rgba(96,165,250,0.25); }
  .c-blue .lbl    { color: #60a5fa; }
  .c-blue .btn    { color: #60a5fa; border-color: rgba(96,165,250,0.35); background: rgba(96,165,250,0.08); }
  .c-blue:hover .btn { background: rgba(96,165,250,0.20); }

  /* Green */
  .c-green::before { background: linear-gradient(90deg,#34d399,#10b981); }
  .c-green:hover   { border-color: rgba(52,211,153,0.4); box-shadow: 0 24px 60px rgba(0,0,0,0.5), 0 0 60px rgba(52,211,153,0.15); }
  .c-green .icon   { background: rgba(52,211,153,0.15); border: 1px solid rgba(52,211,153,0.25); }
  .c-green .lbl    { color: #34d399; }
  .c-green .btn    { color: #34d399; border-color: rgba(52,211,153,0.35); background: rgba(52,211,153,0.08); }
  .c-green:hover .btn { background: rgba(52,211,153,0.20); }

  /* Orange */
  .c-orange::before { background: linear-gradient(90deg,#fb923c,#f97316); }
  .c-orange:hover   { border-color: rgba(251,146,60,0.4); box-shadow: 0 24px 60px rgba(0,0,0,0.5), 0 0 60px rgba(251,146,60,0.15); }
  .c-orange .icon   { background: rgba(251,146,60,0.15); border: 1px solid rgba(251,146,60,0.25); }
  .c-orange .lbl    { color: #fb923c; }
  .c-orange .btn    { color: #fb923c; border-color: rgba(251,146,60,0.35); background: rgba(251,146,60,0.08); }
  .c-orange:hover .btn { background: rgba(251,146,60,0.20); }

  .icon {
    width: 54px;
    height: 54px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.6rem;
    margin-bottom: 1.1rem;
  }

  .lbl {
    font-size: 0.67rem;
    font-weight: 700;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    margin-bottom: 0.45rem;
    display: block;
  }

  .title {
    font-size: 1.12rem;
    font-weight: 700;
    color: rgba(255,255,255,0.93);
    margin-bottom: 0.55rem;
    line-height: 1.35;
    display: block;
  }

  .desc {
    font-size: 0.83rem;
    color: rgba(255,255,255,0.42);
    line-height: 1.65;
    margin-bottom: 1.3rem;
    display: block;
  }

  .btn {
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
  }

  .arrow {
    position: absolute;
    top: 1.4rem;
    right: 1.5rem;
    color: rgba(255,255,255,0.18);
    font-size: 1rem;
    transition: color 0.2s, transform 0.2s;
    font-style: normal;
    font-family: sans-serif;
  }
  .card:hover .arrow {
    color: rgba(255,255,255,0.6);
    transform: translate(3px,-3px);
  }

  .footer {
    text-align: center;
    margin-top: 2.5rem;
    color: rgba(255,255,255,0.2);
    font-size: 0.77rem;
    letter-spacing: 0.3px;
  }

  @keyframes fadeInDown {
    from { opacity:0; transform:translateY(-20px); }
    to   { opacity:1; transform:translateY(0); }
  }
  @keyframes fadeInUp {
    from { opacity:0; transform:translateY(20px); }
    to   { opacity:1; transform:translateY(0); }
  }

  /* --- RESPONSIVE MEDIA QUERIES --- */
  
  /* Tablets and smaller screens */
  @media (max-width: 900px) {
    .grid { 
        gap: 1rem; 
        padding: 0 0.5rem;
    }
  }

  /* Mobile screens */
  @media (max-width: 650px) {
    body {
        padding: 1rem 0.5rem 2rem 0.5rem;
    }
    .hero {
        padding: 1.5rem 0.5rem 1rem 0.5rem;
    }
    .hero-title { 
        font-size: 2.2rem; 
    }
    .hero-sub {
        font-size: 0.95rem;
    }
    .grid { 
        grid-template-columns: 1fr; 
        margin-top: 1.5rem;
    }
    .card {
        padding: 1.5rem;
    }
    .icon {
        width: 48px;
        height: 48px;
        font-size: 1.3rem;
        margin-bottom: 0.9rem;
    }
    .title {
        font-size: 1.05rem;
    }
  }
</style>
</head>
<body>

<div class="hero">
  <div class="badge">📚 &nbsp; Exam Management Suite</div>
  <div class="hero-title">Timetable Tools</div>
  <div class="hero-sub">All your exam scheduling and conversion tools — in one place.</div>
  <div class="hero-line"></div>
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
""", height=1200, scrolling=True)
