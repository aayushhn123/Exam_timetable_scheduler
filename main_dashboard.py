import streamlit as st

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

.block-container {
    padding-top: 1rem !important;
    padding-bottom: 2rem !important;
    max-width: 1100px !important;
}

.stApp {
    background: linear-gradient(135deg, #0f0c29, #302b63, #24243e) !important;
    min-height: 100vh;
}

.hero-section {
    text-align: center;
    padding: 3rem 1rem 2rem 1rem;
    animation: fadeInDown 0.8s ease;
}

.hero-badge {
    display: inline-block;
    background: rgba(167,139,250,0.15);
    color: #a78bfa;
    border: 1px solid rgba(167,139,250,0.3);
    border-radius: 50px;
    padding: 0.35rem 1.1rem;
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 2px;
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
    letter-spacing: -1px;
    line-height: 1.2;
}

.hero-sub {
    font-size: 1.1rem;
    color: rgba(255,255,255,0.55);
    margin-bottom: 0;
}

.hero-line {
    width: 70px;
    height: 3px;
    background: linear-gradient(90deg, #a78bfa, #60a5fa);
    margin: 1.2rem auto 0 auto;
    border-radius: 2px;
}

.tools-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 1.5rem;
    padding: 2rem 1rem 1rem 1rem;
    animation: fadeInUp 0.9s ease;
}

.tool-card {
    position: relative;
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 20px;
    padding: 2rem 1.8rem 1.6rem 1.8rem;
    cursor: pointer;
    text-decoration: none !important;
    display: block;
    overflow: hidden;
    transition: transform 0.25s cubic-bezier(.4,2,.6,1),
                box-shadow 0.25s ease,
                border-color 0.25s ease,
                background 0.25s ease;
}

.tool-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    border-radius: 20px 20px 0 0;
    opacity: 0;
    transition: opacity 0.25s ease;
}

.tool-card:hover {
    transform: translateY(-7px) scale(1.02);
    background: rgba(255,255,255,0.09);
    border-color: rgba(255,255,255,0.22);
    text-decoration: none !important;
}

.tool-card:hover::before { opacity: 1; }

.card-purple::before { background: linear-gradient(90deg, #a78bfa, #818cf8); }
.card-blue::before   { background: linear-gradient(90deg, #60a5fa, #38bdf8); }
.card-green::before  { background: linear-gradient(90deg, #34d399, #10b981); }
.card-orange::before { background: linear-gradient(90deg, #fb923c, #f97316); }

.card-purple:hover { box-shadow: 0 20px 60px rgba(0,0,0,0.5), 0 0 50px rgba(167,139,250,0.18); border-color: rgba(167,139,250,0.35); }
.card-blue:hover   { box-shadow: 0 20px 60px rgba(0,0,0,0.5), 0 0 50px rgba(96,165,250,0.18);  border-color: rgba(96,165,250,0.35); }
.card-green:hover  { box-shadow: 0 20px 60px rgba(0,0,0,0.5), 0 0 50px rgba(52,211,153,0.18);  border-color: rgba(52,211,153,0.35); }
.card-orange:hover { box-shadow: 0 20px 60px rgba(0,0,0,0.5), 0 0 50px rgba(251,146,60,0.18);  border-color: rgba(251,146,60,0.35); }

.card-icon {
    width: 54px;
    height: 54px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.6rem;
    margin-bottom: 1.1rem;
}

.icon-purple { background: rgba(167,139,250,0.15); border: 1px solid rgba(167,139,250,0.25); }
.icon-blue   { background: rgba(96,165,250,0.15);  border: 1px solid rgba(96,165,250,0.25); }
.icon-green  { background: rgba(52,211,153,0.15);  border: 1px solid rgba(52,211,153,0.25); }
.icon-orange { background: rgba(251,146,60,0.15);  border: 1px solid rgba(251,146,60,0.25); }

.card-label {
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 0.4rem;
    display: block;
}
.label-purple { color: #a78bfa; }
.label-blue   { color: #60a5fa; }
.label-green  { color: #34d399; }
.label-orange { color: #fb923c; }

.card-title {
    font-size: 1.15rem;
    font-weight: 700;
    color: rgba(255,255,255,0.93);
    margin-bottom: 0.55rem;
    line-height: 1.35;
    display: block;
}

.card-desc {
    font-size: 0.84rem;
    color: rgba(255,255,255,0.45);
    line-height: 1.65;
    margin-bottom: 1.3rem;
    display: block;
}

.card-launch {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    font-size: 0.78rem;
    font-weight: 600;
    padding: 0.42rem 1.1rem;
    border-radius: 50px;
    border: 1px solid;
    transition: background 0.2s;
    letter-spacing: 0.3px;
}

.launch-purple { color: #a78bfa; border-color: rgba(167,139,250,0.35); background: rgba(167,139,250,0.08); }
.launch-blue   { color: #60a5fa; border-color: rgba(96,165,250,0.35);  background: rgba(96,165,250,0.08); }
.launch-green  { color: #34d399; border-color: rgba(52,211,153,0.35);  background: rgba(52,211,153,0.08); }
.launch-orange { color: #fb923c; border-color: rgba(251,146,60,0.35);  background: rgba(251,146,60,0.08); }

.card-arrow {
    position: absolute;
    top: 1.4rem;
    right: 1.5rem;
    color: rgba(255,255,255,0.18);
    font-size: 1rem;
    transition: color 0.2s, transform 0.2s;
    font-style: normal;
}
.tool-card:hover .card-arrow {
    color: rgba(255,255,255,0.55);
    transform: translate(3px, -3px);
}

.dash-footer {
    text-align: center;
    padding: 1.5rem 0 2rem 0;
    color: rgba(255,255,255,0.22);
    font-size: 0.78rem;
    letter-spacing: 0.3px;
}

@keyframes fadeInDown {
    from { opacity: 0; transform: translateY(-22px); }
    to   { opacity: 1; transform: translateY(0); }
}
@keyframes fadeInUp {
    from { opacity: 0; transform: translateY(22px); }
    to   { opacity: 1; transform: translateY(0); }
}
</style>

<!-- HERO -->
<div class="hero-section">
    <div class="hero-badge">📚 &nbsp; Exam Management Suite</div>
    <div class="hero-title">Timetable Tools</div>
    <div class="hero-sub">All your exam scheduling and conversion tools — in one place.</div>
    <div class="hero-line"></div>
</div>

<!-- CARDS GRID -->
<div class="tools-grid">

    <a class="tool-card card-purple" href="https://examtimetablescheduler-dqttcyf5vzakkjfpkt6xhp.streamlit.app/" target="_blank" rel="noopener noreferrer">
        <i class="card-arrow">↗</i>
        <div class="card-icon icon-purple">🗓️</div>
        <span class="card-label label-purple">Scheduler</span>
        <span class="card-title">Final Exam Timetable Scheduler</span>
        <span class="card-desc">Automatically generate conflict-free final exam timetables with intelligent scheduling algorithms.</span>
        <span class="card-launch launch-purple">Open Tool &nbsp;→</span>
    </a>

    <a class="tool-card card-blue" href="https://reexamschedulerlatest.streamlit.app/" target="_blank" rel="noopener noreferrer">
        <i class="card-arrow">↗</i>
        <div class="card-icon icon-blue">🔁</div>
        <span class="card-label label-blue">Re-Exam</span>
        <span class="card-title">Re-Exam Timetable Scheduler</span>
        <span class="card-desc">Schedule re-examination timetables seamlessly, handling student and room constraints with ease.</span>
        <span class="card-launch launch-blue">Open Tool &nbsp;→</span>
    </a>

    <a class="tool-card card-green" href="https://verification-file-change-to-pdf-converter.streamlit.app/" target="_blank" rel="noopener noreferrer">
        <i class="card-arrow">↗</i>
        <div class="card-icon icon-green">📄</div>
        <span class="card-label label-green">PDF Converter</span>
        <span class="card-title">Final Exam Verification → PDF</span>
        <span class="card-desc">Convert final exam verification files into clean, formatted PDF documents ready for distribution.</span>
        <span class="card-launch launch-green">Open Tool &nbsp;→</span>
    </a>

    <a class="tool-card card-orange" href="https://re-examtimetablescheduler-gndknuqn7whtdxe6cvubaw.streamlit.app/" target="_blank" rel="noopener noreferrer">
        <i class="card-arrow">↗</i>
        <div class="card-icon icon-orange">🖨️</div>
        <span class="card-label label-orange">PDF Converter</span>
        <span class="card-title">Re-Exam File → PDF Converter</span>
        <span class="card-desc">Transform re-examination timetable files into professional PDFs with a single click.</span>
        <span class="card-launch launch-orange">Open Tool &nbsp;→</span>
    </a>

</div>

<!-- FOOTER -->
<div class="dash-footer">
    Exam Timetable Project &nbsp;·&nbsp; Click any card to open in a new tab
</div>
""", unsafe_allow_html=True)
