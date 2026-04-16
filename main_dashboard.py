import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Exam Management Dashboard",
    page_icon="📅",
    layout="wide"
)

# Custom CSS for the Tiles
st.markdown("""
    <style>
    /* Main container styling */
    .main {
        background-color: #f8f9fa;
    }
    
    /* Title styling */
    .main-title {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        color: #1E1E1E;
        text-align: center;
        padding: 20px;
        font-weight: 700;
    }

    /* Card styling */
    .card-container {
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 25px;
        padding: 20px;
    }

    .card {
        background-color: white;
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        padding: 30px;
        width: 300px;
        text-align: center;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        text-decoration: none;
        color: inherit;
        border: 1px solid #e0e0e0;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        min-height: 200px;
    }

    .card:hover {
        transform: translateY(-10px);
        box-shadow: 0 10px 20px rgba(0, 0, 0, 0.15);
        border-color: #ff4b4b;
        color: #ff4b4b;
        cursor: pointer;
    }

    .card h3 {
        font-size: 1.2rem;
        margin-bottom: 15px;
    }

    .card p {
        font-size: 0.9rem;
        color: #6c757d;
    }

    .icon {
        font-size: 40px;
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# Dashboard Header
st.markdown("<h1 class='main-title'>📚 Exam Timetable Project Hub</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #666;'>Select a tool below to launch it in a new tab.</p>", unsafe_allow_html=True)
st.write("---")

# Data for the tools
tools = [
    {
        "name": "Final Exam Timetable Scheduler",
        "desc": "Generate and manage primary exam schedules efficiently.",
        "url": "https://examtimetablescheduler-dqttcyf5vzakkjfpkt6xhp.streamlit.app/",
        "icon": "🗓️"
    },
    {
        "name": "Re-exam Timetable Scheduler",
        "desc": "Handle scheduling for supplementary and re-examinations.",
        "url": "https://reexamschedulerlatest.streamlit.app/",
        "icon": "🔄"
    },
    {
        "name": "Final Exam PDF Converter",
        "desc": "Convert verification files into professional PDF documents.",
        "url": "https://verification-file-change-to-pdf-converter.streamlit.app/",
        "icon": "📄"
    },
    {
        "name": "Re-exam PDF Converter",
        "desc": "Specialized conversion tool for re-exam documentation.",
        "url": "https://re-examtimetablescheduler-gndknuqn7whtdxe6cvubaw.streamlit.app/",
        "icon": "📤"
    }
]

# Rendering the Tiles
cols = st.columns(2) # Create a 2x2 grid

for i, tool in enumerate(tools):
    with cols[i % 2]:
        card_html = f"""
            <a href="{tool['url']}" target="_blank" style="text-decoration: none; color: inherit;">
                <div class="card">
                    <div class="icon">{tool['icon']}</div>
                    <h3>{tool['name']}</h3>
                    <p>{tool['desc']}</p>
                    <div style="margin-top: 10px; font-weight: bold; color: #ff4b4b;">Launch Tool →</div>
                </div>
            </a>
        """
        st.markdown(card_html, unsafe_allow_html=True)
        st.write("") # Spacer

# Footer
st.write("---")
st.caption("Centralized Exam Management System | Built with Streamlit")
