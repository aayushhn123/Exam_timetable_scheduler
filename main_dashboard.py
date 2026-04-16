import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Exam Hub",
    page_icon="📚",
    layout="wide"
)

# Advanced CSS for Grid and Visibility
st.markdown("""
    <style>
    /* Remove default padding */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }

    /* Title Styling */
    .header-text {
        text-align: center;
        font-family: 'Inter', sans-serif;
        color: #FFFFFF;
        padding-bottom: 30px;
    }

    /* The Flexbox Grid */
    .dashboard-grid {
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 20px;
        width: 100%;
    }

    /* Individual Card Styling - Glassmorphism */
    .card {
        background: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 15px;
        padding: 25px;
        width: 320px;
        height: 250px;
        text-align: center;
        transition: all 0.3s ease;
        text-decoration: none !important;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }

    .card:hover {
        transform: translateY(-10px);
        background: rgba(255, 255, 255, 0.1);
        border-color: #ff4b4b;
        box-shadow: 0 10px 20px rgba(0,0,0,0.4);
    }

    /* Text Inside Cards */
    .card h3 {
        color: #ffffff !important;
        font-size: 1.2rem !important;
        margin-top: 15px !important;
        margin-bottom: 10px !important;
        font-weight: 600 !important;
    }

    .card p {
        color: #bbbbbb !important;
        font-size: 0.9rem !important;
        line-height: 1.4;
    }

    .card .icon {
        font-size: 45px;
    }

    .launch-btn {
        margin-top: 15px;
        color: #ff4b4b;
        font-weight: bold;
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    /* Link behavior */
    a {
        text-decoration: none !important;
    }
    </style>
    """, unsafe_allow_html=True)

# Header Section
st.markdown("<h1 class='header-text'>📚 Exam Timetable Project Hub</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #888; margin-top: -20px;'>Aayush's Centralized Management System</p>", unsafe_allow_html=True)
st.write("<br>", unsafe_allow_html=True)

# Define Tools
tools = [
    {
        "name": "Final Exam Timetable Scheduler",
        "desc": "Generate and manage primary exam schedules efficiently.",
        "url": "https://examtimetablescheduler-dqttcyf5vzakkjfpkt6xhp.streamlit.app/",
        "icon": "📅"
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

# Build the Grid in HTML
grid_html = "<div class='dashboard-grid'>"

for tool in tools:
    grid_html += f"""
        <a href="{tool['url']}" target="_blank">
            <div class="card">
                <div class="icon">{tool['icon']}</div>
                <h3>{tool['name']}</h3>
                <p>{tool['desc']}</p>
                <div class="launch-btn">Launch Tool →</div>
            </div>
        </a>
    """

grid_html += "</div>"

# Render the HTML
st.markdown(grid_html, unsafe_allow_html=True)

# Footer
st.markdown("<br><br><p style='text-align: center; color: #555; font-size: 0.8rem;'>NMIMS Data Science Project</p>", unsafe_allow_html=True)
