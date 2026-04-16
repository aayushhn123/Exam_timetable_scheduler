import streamlit as st

st.set_page_config(
    page_title="Exam Timetable Tools",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    /* Hide default streamlit elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1100px;
    }

    /* Page background */
    .stApp {
        background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
        min-height: 100vh;
    }

    /* Hero section */
    .hero-section {
        text-align: center;
        padding: 3rem 1rem 2rem 1rem;
        animation: fadeInDown 0.8s ease;
    }

    .hero-title {
        font-size: 3rem;
        font-weight: 800;
        background: linear-gradient(90deg, #a78bfa, #60a5fa, #34d399);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 0.5rem;
        letter-spacing: -1px;
    }

    .hero-subtitle {
        font-size: 1.15rem;
        color: rgba(255,255,255,0.6);
        margin-bottom: 0.5rem;
        font-weight: 400;
    }

    .hero-divider {
        width: 80px;
        height: 3px;
        background: linear-gradient(90deg, #a78bfa, #60a5fa);
        margin: 1.2rem auto 0 auto;
        border-radius: 2px;
    }

    /* Badge */
    .badge {
        display: inline-block;
        background: rgba(167,139,250,0.15);
        color: #a78bfa;
        border: 1px solid rgba(167,139,250,0.3);
        border-radius: 50px;
        padding: 0.3rem 1rem;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        margin-bottom: 1rem;
    }

    /* Tool cards grid */
    .tools-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 1.5rem;
        padding: 1.5rem 0 2rem 0;
        animation: fadeInUp 0.9s ease;
    }

    @media (max-width: 700px) {
        .tools-grid { grid-template-columns: 1fr; }
        .hero-title { font-size: 2rem; }
    }

    /* Individual tool card */
    .tool-card {
        position: relative;
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.10);
        border-radius: 20px;
        padding: 2rem 1.8rem 1.6rem 1.8rem;
        cursor: pointer;
        text-decoration: none !important;
        display: block;
        transition: transform 0.25s cubic-bezier(.4,2,.6,1),
                    box-shadow 0.25s ease,
                    border-color 0.25s ease,
                    background 0.25s ease;
        overflow: hidden;
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
        transform: translateY(-6px) scale(1.02);
        background: rgba(255,255,255,0.08);
        border-color: rgba(255,255,255,0.25);
        box-shadow: 0 20px 60px rgba(0,0,0,0.4), 0 0 40px rgba(167,139,250,0.10);
        text-decoration: none !important;
    }

    .tool-card:hover::before {
        opacity: 1;
    }

    /* Color accents per card */
    .card-purple::before { background: linear-gradient(90deg, #a78bfa, #818cf8); }
    .card-blue::before   { background: linear-gradient(90deg, #60a5fa, #38bdf8); }
    .card-green::before  { background: linear-gradient(90deg, #34d399, #10b981); }
    .card-orange::before { background: linear-gradient(90deg, #fb923c, #f97316); }

    .card-purple:hover { box-shadow: 0 20px 60px rgba(0,0,0,0.4), 0 0 50px rgba(167,139,250,0.15); }
    .card-blue:hover   { box-shadow: 0 20px 60px rgba(0,0,0,0.4), 0 0 50px rgba(96,165,250,0.15); }
    .card-green:hover  { box-shadow: 0 20px 60px rgba(0,0,0,0.4), 0 0 50px rgba(52,211,153,0.15); }
    .card-orange:hover { box-shadow: 0 20px 60px rgba(0,0,0,0.4), 0 0 50px rgba(251,146,60,0.15); }

    /* Icon circle */
    .card-icon {
        width: 52px;
        height: 52px;
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.5rem;
        margin-bottom: 1.1rem;
    }

    .icon-purple { background: rgba(167,139,250,0.15); border: 1px solid rgba(167,139,250,0.25); }
    .icon-blue   { background: rgba(96,165,250,0.15);  border: 1px solid rgba(96,165,250,0.25); }
    .icon-green  { background: rgba(52,211,153,0.15);  border: 1px solid rgba(52,211,153,0.25); }
    .icon-orange { background: rgba(251,146,60,0.15);  border: 1px solid rgba(251,146,60,0.25); }

    /* Card text */
    .card-label {
        font-size: 0.7rem;
        font-weight: 700;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        margin-bottom: 0.4rem;
    }
    .label-purple { color: #a78bfa; }
    .label-blue   { color: #60a5fa; }
    .label-green  { color: #34d399; }
    .label-orange { color: #fb923c; }

    .card-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: rgba(255,255,255,0.92);
        margin-bottom: 0.5rem;
        line-height: 1.35;
    }

    .card-desc {
        font-size: 0.85rem;
        color: rgba(255,255,255,0.45);
        line-height: 1.6;
        margin-bottom: 1.2rem;
    }

    /* Launch button inside card */
    .card-launch {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        font-size: 0.8rem;
        font-weight: 600;
        padding: 0.4rem 1rem;
        border-radius: 50px;
        border: 1px solid;
        transition: background 0.2s, color 0.2s;
    }
    .launch-purple { color: #a78bfa; border-color: rgba(167,139,250,0.35); background: rgba(167,139,250,0.07); }
    .launch-blue   { color: #60a5fa; border-color: rgba(96,165,250,0.35);  background: rgba(96,165,250,0.07); }
    .launch-green  { color: #34d399; border-color: rgba(52,211,153,0.35);  background: rgba(52,211,153,0.07); }
    .launch-orange { color: #fb923c; border-color: rgba(251,146,60,0.35);  background: rgba(251,146,60,0.07); }

    .tool-card:hover .launch-purple { background: rgba(167,139,250,0.18); }
    .tool-card:hover .launch-blue   { background: rgba(96,165,250,0.18); }
    .tool-card:hover .launch-green  { background: rgba(52,211,153,0.18); }
    .tool-card:hover .launch-orange { background: rgba(251,146,60,0.18); }

    /* Arrow icon on card */
    .card-arrow {
        position: absolute;
        top: 1.4rem;
        right: 1.4rem;
        color: rgba(255,255,255,0.15);
        font-size: 1rem;
        transition: color 0.2s, transform 0.2s;
    }
    .tool-card:hover .card-arrow {
        color: rgba(255,255,255,0.45);
        transform: translate(3px, -3px);
    }

    /* Footer */
    .dashboard-footer {
        text-align: center;
        padding: 1rem 0 2rem 0;
        color: rgba(255,255,255,0.25);
        font-size: 0.8rem;
        animation: fadeIn 1.2s ease;
    }

    /* Animations */
    @keyframes fadeInDown {
        from { opacity: 0; transform: translateY(-20px); }
        to   { opacity: 1; transform: translateY(0); }
    }
    @keyframes fadeInUp {
        from { opacity: 0; transform: translateY(20px); }
        to   { opacity: 1; transform: translateY(0); }
    }
    @keyframes fadeIn {
        from { opacity: 0; }
        to   { opacity: 1; }
    }
</style>
""", unsafe_allow_html=True)

tools = [
    {
        "title": "Final Exam Timetable Scheduler",
        "desc": "Automatically generate conflict-free final exam timetables with intelligent scheduling algorithms.",
        "url": "https://examtimetablescheduler-dqttcyf5vzakkjfpkt6xhp.streamlit.app/",
        "icon": "🗓️",
        "label": "Scheduler",
        "color": "purple",
    },
    {
        "title": "Re-Exam Timetable Scheduler",
        "desc": "Schedule re-examination timetables seamlessly, handling student and room constraints with ease.",
        "url": "https://reexamschedulerlatest.streamlit.app/",
        "icon": "🔁",
        "label": "Re-Exam",
        "color": "blue",
    },
    {
        "title": "Final Exam Verification → PDF",
        "desc": "Convert final exam verification files into clean, formatted PDF documents ready for distribution.",
        "url": "https://verification-file-change-to-pdf-converter.streamlit.app/",
        "icon": "📄",
        "label": "PDF Converter",
        "color": "green",
    },
    {
        "title": "Re-Exam File → PDF Converter",
        "desc": "Transform re-examination timetable files into professional PDFs with a single click.",
        "url": "https://re-examtimetablescheduler-gndknuqn7whtdxe6cvubaw.streamlit.app/",
        "icon": "🖨️",
        "label": "PDF Converter",
        "color": "orange",
    },
]

st.markdown("""
<div class="hero-section">
    <div class="badge">📚 &nbsp; Exam Management Suite</div>
    <div class="hero-title">Timetable Tools</div>
    <div class="hero-subtitle">All your exam scheduling and conversion tools — in one place.</div>
    <div class="hero-divider"></div>
</div>
""", unsafe_allow_html=True)

cards_html = '<div class="tools-grid">'
for tool in tools:
    c = tool["color"]
    cards_html += f"""
    <a class="tool-card card-{c}" href="{tool['url']}" target="_blank" rel="noopener noreferrer">
        <span class="card-arrow">↗</span>
        <div class="card-icon icon-{c}">{tool['icon']}</div>
        <div class="card-label label-{c}">{tool['label']}</div>
        <div class="card-title">{tool['title']}</div>
        <div class="card-desc">{tool['desc']}</div>
        <span class="card-launch launch-{c}">Open Tool &nbsp;→</span>
    </a>
    """
cards_html += '</div>'

st.markdown(cards_html, unsafe_allow_html=True)

st.markdown("""
<div class="dashboard-footer">
    Exam Timetable Project &nbsp;·&nbsp; Click any card to open the tool in a new tab
</div>
""", unsafe_allow_html=True)
