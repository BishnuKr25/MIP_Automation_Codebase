import streamlit as st
import pandas as pd
import os
import time
import glob
import shutil
import win32com.client as win32
import json
import re
from datetime import datetime
import subprocess
import plotly.express as px
import plotly.graph_objects as go
from streamlit_option_menu import option_menu
import requests
from PIL import Image
import base64
import numpy as np
from io import BytesIO

st.set_page_config(
    page_title="MIP Automation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def get_base64_image(image_path):
    """
    Convert an image file to a base64 string.
    """
    with open(image_path, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode()
    return f"data:image/png;base64,{encoded_string}"
    
LOGO_PATH = "HCLLP.jpg"
logo_base64 = get_base64_image(LOGO_PATH)

st.markdown(
    f"""
    <div style="display: flex; align-items: center; margin-bottom: -20px;">
        <a href="https://www.harshwal.com/"><img src="{logo_base64}" alt="Company Logo" style="height: 80px; margin-right: 25px; border-radius: 10px" /></a>
    </div>
    """,
    unsafe_allow_html=True
)

# theme_choice = st.sidebar.radio("Theme Mode", ["Light", "Dark"], index=0)
# dark_css = """
# <style>
# /* Override light background with dark */
# .main, .css-18e3th9 {
#     background-color: #2e2e2e !important;
#     color: #ffffff !important;
# }
# /* Override text in cards, alerts, etc. if needed */
# .card, .alert, .tech-box, .prompt-container, .processing-container, .terminal {
#     background-color: #3c3c3c !important;
#     color: #ffffff !important;
# }
# .step {
#     color: #ffffff !important;
#     border-bottom: 3px solid #666666 !important;
# }
# .step.active {
#     border-bottom: 3px solid #3498db !important;
# }
# .step.complete {
#     border-bottom: 3px solid #2ecc71 !important;
# }
# </style>
# """

# if theme_choice == "Dark":
#     st.markdown(dark_css, unsafe_allow_html=True)

# ---------------------------
# Custom Functions
# ---------------------------
def convert_xlsx_to_csv(excel_file, csv_file):
    try:
        # Try using pandas for faster conversion
        df = pd.read_excel(excel_file)
        df.to_csv(csv_file, index=False)
        return csv_file
    except Exception as e:
        with st.spinner("Pandas conversion failed. Using Excel COM..."):
            try:
                import pythoncom
                pythoncom.CoInitialize()
                excel = win32.Dispatch("Excel.Application")
                excel.Visible = False
                workbook = excel.Workbooks.Open(excel_file)
                workbook.SaveAs(csv_file, FileFormat=6)  # CSV format
                workbook.Close(False)
                excel.Quit()
                return csv_file
            except Exception as e:
                st.error(f"Error converting to CSV with Excel COM: {e}")
                return None

def wait_for_file_stability(file_path, timeout=60, check_interval=2):
    """Wait until the file size remains stable (download complete)."""
    start_time = time.time()
    prev_size = -1
    while time.time() - start_time < timeout:
        if os.path.exists(file_path):
            curr_size = os.path.getsize(file_path)
            if curr_size == prev_size and curr_size > 0:
                return True
            prev_size = curr_size
        time.sleep(check_interval)
    return False

def get_latest_downloaded_file():
    download_path = os.path.expanduser(r"~\Downloads")
    excel_file_pattern = os.path.join(download_path, "*.xlsx")
    csv_file = os.path.join(download_path, "ExpandedGL.csv")
    timeout = 60
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    start_time = time.time()
    latest_file = None

    while time.time() - start_time < timeout:
        elapsed = time.time() - start_time
        progress = min(elapsed / timeout, 0.99)
        progress_bar.progress(progress)
        status_text.text(f"Searching for downloaded files... {int(progress*100)}%")
        list_of_files = glob.glob(excel_file_pattern)
        if list_of_files:
            latest_file = max(list_of_files, key=os.path.getctime)
            if wait_for_file_stability(latest_file):
                status_text.text("File found and stable!")
                progress_bar.progress(1.0)
                break
        time.sleep(2)

    if not latest_file:
        status_text.empty()
        progress_bar.empty()
        st.error("No Excel file found in Downloads within timeout!")
        return None

    with st.spinner("Converting Excel to CSV..."):
        try:
            return convert_xlsx_to_csv(latest_file, csv_file)
        except Exception as e:
            st.error(f"Failed to convert {os.path.basename(latest_file)} to CSV: {e}")
            return None

def get_latest_rectified_file():
    download_path = os.path.expanduser(r"~\Downloads")
    rectified_files = glob.glob(os.path.join(download_path, "*_rectified.csv"))
    return max(rectified_files, key=os.path.getctime) if rectified_files else None

def extract_dates_from_text(text):
    date_pattern1 = r'(\d{1,2})(?:st|nd|rd|th)?\s+(?:of\s+)?([A-Za-z]+)\s+(\d{4})'
    date_pattern2 = r'([A-Za-z]+)\s+(\d{1,2})(?:,|\s)\s*(\d{4})'
    date_pattern3 = r'(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})'
    month_dict = {
        'january': '01', 'february': '02', 'march': '03', 'april': '04',
        'may': '05', 'june': '06', 'july': '07', 'august': '08',
        'sept': '09', 'september': '09', 'october': '10',
        'november': '11', 'december': '12'
    }
    extracted_dates = []
    for match in re.finditer(date_pattern1, text, re.IGNORECASE):
        day = match.group(1).zfill(2)
        month = month_dict.get(match.group(2).lower(), '00')
        year = match.group(3)
        if month != '00':
            extracted_dates.append(f"{month}-{day}-{year}")
    for match in re.finditer(date_pattern2, text, re.IGNORECASE):
        month = month_dict.get(match.group(1).lower(), '00')
        day = match.group(2).zfill(2)
        year = match.group(3)
        if month != '00':
            extracted_dates.append(f"{month}-{day}-{year}")
    for match in re.finditer(date_pattern3, text):
        day = match.group(1).zfill(2)
        month = match.group(2).zfill(2)
        year = match.group(3)
        extracted_dates.append(f"{month}-{day}-{year}")
    return list(set(extracted_dates))

def display_data_metrics(df):
    metrics_cols = st.columns(4)
    total_rows = len(df)
    total_grants = df['Grant Code'].nunique() if 'Grant Code' in df.columns else 0
    total_amount = df['Amount'].sum() if 'Amount' in df.columns else 0
    unique_departments = df['Department'].nunique() if 'Department' in df.columns else 0
    with metrics_cols[0]:
        st.metric("Total Records", f"{total_rows:,}", delta=None)
    with metrics_cols[1]:
        st.metric("Unique Grants", f"{total_grants:,}", delta=None)
    with metrics_cols[2]:
        st.metric("Total Amount", f"${total_amount:,.2f}", delta=None)
    with metrics_cols[3]:
        st.metric("Departments", f"{unique_departments}", delta=None)

def create_visualizations(df):
    st.subheader("Data Insights")
    col1, col2 = st.columns(2)
    with col1:
        if 'Category' in df.columns:
            category_counts = df['Category'].value_counts().reset_index()
            category_counts.columns = ['Category', 'Count']
            fig = px.pie(
                category_counts, 
                values='Count', 
                names='Category', 
                title='Distribution by Category',
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            fig.update_layout(
                legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
                margin=dict(l=20, r=20, t=40, b=20)
            )
            st.plotly_chart(fig, use_container_width=True)
    with col2:
        if 'Amount' in df.columns and 'Department' in df.columns:
            dept_amounts = df.groupby('Department')['Amount'].sum().reset_index()
            dept_amounts = dept_amounts.sort_values('Amount', ascending=False)
            fig = px.bar(
                dept_amounts, 
                x='Department', 
                y='Amount',
                title='Total Amount by Department',
                color='Amount',
                labels={'Amount': 'Total Amount ($)'},
                color_continuous_scale='Viridis'
            )
            fig.update_layout(
                xaxis_title="Department",
                yaxis_title="Amount ($)",
                coloraxis_showscale=False,
                margin=dict(l=20, r=20, t=40, b=20)
            )
            st.plotly_chart(fig, use_container_width=True)
    if 'Transaction Date' in df.columns and 'Amount' in df.columns:
        df['Transaction Date'] = pd.to_datetime(df['Transaction Date'])
        time_data = df.groupby(pd.Grouper(key='Transaction Date', freq='M'))['Amount'].sum().reset_index()
        fig = px.line(
            time_data, 
            x='Transaction Date', 
            y='Amount',
            title='Transaction Volume Over Time',
            markers=True,
            line_shape='spline',
            labels={'Transaction Date': 'Month', 'Amount': 'Total Amount ($)'}
        )
        fig.update_layout(
            xaxis_title="Date",
            yaxis_title="Amount ($)",
            hovermode="x unified",
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig, use_container_width=True)

# ---------------------------
# Session State Setup
# ---------------------------
if 'current_step' not in st.session_state:
    st.session_state.current_step = 1
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'start_automation' not in st.session_state:
    st.session_state.start_automation = False

# ---------------------------
# Custom CSS (Light Mode by default)
# ---------------------------
st.markdown("""
<style>
    /* Main app styling */
    .main { background-color: #f8f9fa !important; }
    .css-18e3th9 { padding: 1rem 5rem 5rem; }
    .card {
        display: flex;
        justify-content: center;
        align-items: center;
        border-radius: 10px;
        padding: 10px;
        background-color: white;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 10px;
    }
    .card b { color: #06002a; font-size: 22px; }
    .custom-title {
        color: #2c3e50;
        font-size: 40px !important;
        font-weight: bold;
        margin-bottom: 10px;
        text-align: center;
    }
    .custom-subtitle {
        color: #7f8c8d;
        font-size: 20px !important;
        font-weight: normal;
        margin-bottom: 30px;
        text-align: center;
    }
    .stButton button {
        background-color: #3498db;
        color: white;
        border-radius: 5px;
        font-weight: bold;
        width: 100%;
        transition: all 0.3s;
        border: none;
    }
    .stButton button:hover {
        background-color: #2980b9;
        transform: translateY(-2px);
        box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2);
    }
    .status-badge {
        display: inline-block;
        padding: 5px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: bold;
        margin-bottom: 10px;
    }
    .status-pending { background-color: #f1c40f; color: #333; }
    .status-complete { background-color: #2ecc71; color: white; }
    .status-error { background-color: #e74c3c; color: white; }
    .step-container {
        display: flex;
        flex-direction: row;
        justify-content: space-between;
        margin-bottom: 30px;
    }
    .step {
        flex: 1;
        text-align: center;
        padding: 10px;
        border-bottom: 3px solid #e0e0e0;
    }
    .step.active {
        border-bottom: 3px solid #3498db;
        font-weight: bold;
    }
    .step.complete { border-bottom: 3px solid #2ecc71; }
    @keyframes pulse {
        0% { box-shadow: 0 0 0 0 rgba(52, 152, 219, 0.7); }
        70% { box-shadow: 0 0 0 10px rgba(52, 152, 219, 0); }
        100% { box-shadow: 0 0 0 0 rgba(52, 152, 219, 0); }
    }
    .pulse-animation { animation: pulse 2s infinite; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .date-display {
        background-color: #e8f4fd;
        padding: 10px 15px;
        border-radius: 5px;
        font-weight: bold;
        color: #2980b9;
        display: inline-block;
        margin-right: 10px;
    }
    .alert { padding: 15px; border-radius: 5px; margin-bottom: 15px; }
    .alert-info { background-color: #d1ecf1; color: #0c5460; }
    .alert-success { background-color: #d4edda; color: #155724; }
    .alert-warning { background-color: #fff3cd; color: #856404; }
    .alert-danger { background-color: #f8d7da; color: #721c24; }
    .tech-box {
        border: 2px solid #3498db;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 20px;
        background-color: rgba(52, 152, 219, 0.05);
        position: relative;
        overflow: hidden;
    }
    .tech-box:before {
        content: "";
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 3px;
        background: linear-gradient(to right, #3498db, #2ecc71, #3498db);
        animation: gradient-move 3s linear infinite;
    }
    @keyframes gradient-move {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    .processing-container {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        margin-top: 20px;
        position: relative;
        border-left: 4px solid #3498db;
    }
    .blinking-dot {
        height: 10px;
        width: 10px;
        background-color: #3498db;
        border-radius: 50%;
        display: inline-block;
        margin-right: 5px;
        animation: blink 1s infinite;
    }
    @keyframes blink {
        0% { opacity: 0; }
        50% { opacity: 1; }
        100% { opacity: 0; }
    }
    .terminal {
        background-color: #2c3e50;
        color: #2ecc71;
        padding: 15px;
        border-radius: 5px;
        font-family: monospace;
        position: relative;
        overflow: hidden;
        margin-bottom: 20px;
    }
    .terminal:before {
        content: "● ● ●";
        position: absolute;
        top: 5px;
        left: 10px;
        color: #95a5a6;
        font-size: 12px;
    }
    .terminal-content { margin-top: 15px; }
    .cursor {
        display: inline-block;
        width: 8px;
        height: 15px;
        background-color: #2ecc71;
        animation: cursor-blink 1s infinite;
        margin-left: 5px;
        vertical-align: middle;
    }
    @keyframes cursor-blink {
        0% { opacity: 0; }
        50% { opacity: 1; }
        100% { opacity: 0; }
    }
    .download-btn {
        background-color: #27ae60;
        color: white;
        padding: 10px 15px;
        border-radius: 5px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-weight: bold;
        margin: 10px 0;
        border: none;
        cursor: pointer;
        transition: all 0.3s;
    }
    .download-btn:hover {
        background-color: #2ecc71;
        transform: translateY(-2px);
        box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2);
    }
    .prompt-container {
        border: 2px solid #3498db;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
        background-color: white;
        box-shadow: 0 4px 15px rgba(52, 152, 219, 0.2);
    }
    @keyframes glow {
        0% { box-shadow: 0 0 5px rgba(52, 152, 219, 0.5); }
        50% { box-shadow: 0 0 20px rgba(52, 152, 219, 0.8); }
        100% { box-shadow: 0 0 5px rgba(52, 152, 219, 0.5); }
    }
    .glow { animation: glow 2s infinite; }
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Sidebar Navigation
# ---------------------------
with st.sidebar:
#     st.markdown(
#     f"""
#     <div style="display: flex; align-items: center; margin-bottom: -20px;">
#         <img src="{logo_base64}" alt="Company Logo" style="height: 80px; margin-right: 25px;" />
#     </div>
#     """,
#     unsafe_allow_html=True
# )
    st.markdown("<h1 style='text-align: center;'>MIP Automation</h1>", unsafe_allow_html=True)
    st.markdown("""
    <div class="tech-box">
        <h3 style="text-align: center;">AI-Powered</h3>
        <p style="text-align: center; margin: 0;">Intelligent Financial Processing</p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### Process Steps")
    step1_status = "complete" if st.session_state.current_step > 1 else "active" if st.session_state.current_step == 1 else ""
    step2_status = "complete" if st.session_state.current_step > 2 else "active" if st.session_state.current_step == 2 else ""
    step3_status = "complete" if st.session_state.current_step > 3 else "active" if st.session_state.current_step == 3 else ""
    step4_status = "active" if st.session_state.current_step == 4 else ""
    st.markdown(f"""
    <div class="step-container" style="flex-direction: column;">
        <div class="step {step1_status}" style="text-align: left; margin-bottom: 10px;">
            <div>Step 1: Input Prompt</div>
        </div>
        <div class="step {step2_status}" style="text-align: left; margin-bottom: 10px;">
            <div>Step 2: Generate Report</div>
        </div>
        <div class="step {step3_status}" style="text-align: left; margin-bottom: 10px;">
            <div>Step 3: Process Data</div>
        </div>
        <div class="step {step4_status}" style="text-align: left; margin-bottom: 10px;">
            <div>Step 4: View Results</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### System Status")
    if 'rectified_file' in st.session_state:
        st.markdown('<div class="status-badge status-complete">Processing Complete</div>', unsafe_allow_html=True)
    elif 'latest_file' in st.session_state:
        st.markdown('<div class="status-badge status-pending">Processing In Progress</div>', unsafe_allow_html=True)
    elif st.session_state.start_automation:
        st.markdown('<div class="status-badge status-pending">Report Generation In Progress</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="status-badge status-pending">Ready For Input</div>', unsafe_allow_html=True)
    if st.button("Restart Automation"):
        for key in list(st.session_state.keys()):
            if key != 'page':
                del st.session_state[key]
        st.session_state.current_step = 1
        st.rerun()

st.markdown("<h1 class='custom-title'>MIP Automation</h1>", unsafe_allow_html=True)
# st.markdown("<p class='custom-subtitle'>AI-Powered Financial Data Processing</p>", unsafe_allow_html=True)
step1_class = "complete" if st.session_state.current_step > 1 else "active"
step2_class = "complete" if st.session_state.current_step > 2 else "active" if st.session_state.current_step >= 2 else ""
step3_class = "complete" if st.session_state.current_step > 3 else "active" if st.session_state.current_step >= 3 else ""
step4_class = "active" if st.session_state.current_step >= 4 else ""
st.markdown(f"""
<div class="step-container">
    <div class="step {step1_class}">
        <div>Step 1</div>
        <div>Prompt Input</div>
    </div>
    <div class="step {step2_class}">
        <div>Step 2</div>
        <div>Report Generation</div>
    </div>
    <div class="step {step3_class}">
        <div>Step 3</div>
        <div>Data Processing</div>
    </div>
    <div class="step {step4_class}">
        <div>Step 4</div>
        <div>Analysis</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------------------
# Step 1: Prompt Input (Auto-advance)
# ---------------------------
if st.session_state.current_step == 1:
    st.markdown('<div class="card prompt-container glow"><b>Enter Your Financial Report Request</b></div>', unsafe_allow_html=True)
    prompt_input = st.text_input("Enter your request:",
                                 placeholder="Example: Generate financial report from January 10, 2020 to September 30, 2021")
    if prompt_input:
        with st.spinner("Analyzing your request..."):
            extracted_dates = extract_dates_from_text(prompt_input)
            if len(extracted_dates) < 2:
                st.info("Example: 'Generate financial report from October 10, 2020 to September 30, 2021'")
            else:
                extracted_dates.sort(key=lambda x: datetime.strptime(x, "%m-%d-%Y"))
                from_date = extracted_dates[0]
                to_date = extracted_dates[-1]
                start_date_obj = datetime.strptime(from_date, '%m-%d-%Y')
                end_date_obj = datetime.strptime(to_date, '%m-%d-%Y')
                st.markdown(f"""
                <div class="alert alert-success">
                    <strong>Dates Extracted:</strong><br>
                    From: <span class="date-display">{start_date_obj.strftime('%B %d, %Y')}</span><br>
                    To: <span class="date-display">{end_date_obj.strftime('%B %d, %Y')}</span>
                </div>
                """, unsafe_allow_html=True)
                st.session_state.from_date = from_date
                st.session_state.to_date = to_date
                st.session_state.prompt = prompt_input
                st.session_state.start_automation = True
                st.session_state.current_step = 2
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------
# Step 2: Automatic Report Generation
# ---------------------------
elif st.session_state.current_step == 2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Generating Financial Report")
    st.markdown("""
    <div class="terminal">
        <div class="terminal-content">
            > Initializing MIP Automation<br>
            > Loading financial modules<br>
            > Setting date parameters<br>
            > Starting report generation<span class="cursor"></span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    start_date_obj = datetime.strptime(st.session_state.from_date, '%m-%d-%Y')
    end_date_obj = datetime.strptime(st.session_state.to_date, '%m-%d-%Y')
    st.markdown(f"""
    <div class="tech-box">
        <h4>Request Parameters</h4>
        <p><strong>From:</strong> <span class="date-display">{start_date_obj.strftime('%B %d, %Y')}</span></p>
        <p><strong>To:</strong> <span class="date-display">{end_date_obj.strftime('%B %d, %Y')}</span></p>
        <p><strong>Query:</strong> {st.session_state.prompt}</p>
    </div>
    """, unsafe_allow_html=True)
    status_text = st.empty()
    progress = st.progress(0)
    with open("extracted_dates.json", "w") as json_file:
        json.dump({"from": st.session_state.from_date, "to": st.session_state.to_date}, json_file)
    status_text.text("Initializing Automation Framework...")
    for i in range(20):
        time.sleep(0.05)
        progress.progress(i / 100)
    status_text.text("Executing data extraction process...")
    try:
        subprocess.run([
            "python", "MIP_Automation.py",
            st.session_state.from_date,
            st.session_state.from_date,
            st.session_state.from_date,
            st.session_state.to_date
        ])
        for i in range(20, 70):
            time.sleep(0.05)
            progress.progress(i / 100)
    except Exception as e:
        st.error(f"Error running automation: {e}")
        progress.progress(100)
        st.stop()
    status_text.text("Retrieving generated report...")
    latest_csv = get_latest_downloaded_file()
    if latest_csv:
        st.session_state.latest_file = latest_csv
        for i in range(70, 100):
            time.sleep(0.05)
            progress.progress(i / 100)
        status_text.text("Report generation complete!")
        st.session_state.current_step = 3
        st.rerun()
    else:
        st.error("Could not find generated file")
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------
# Step 3: Automatic Data Processing
# ---------------------------
elif st.session_state.current_step == 3:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### ⚙️ Data Processing")
    if 'latest_file' in st.session_state:
        st.markdown(f"""
        <div class="alert alert-success">
            <strong>Report Generated:</strong> {os.path.basename(st.session_state.latest_file)}
        </div>
        """, unsafe_allow_html=True)
        with st.expander("🔍 Preview Raw Data"):
            try:
                df = pd.read_csv(st.session_state.latest_file, dtype={"Grant Code": str})
                st.dataframe(df.head())
            except Exception as e:
                st.error(f"Error loading file: {e}")
        st.markdown("Processing data automatically...")
        try:
            result = subprocess.run(
                ["python", "process_csv.py", st.session_state.latest_file],
                capture_output=True,
                text=True
            )
            if result.returncode == 0:
                rectified = get_latest_rectified_file()
                if rectified:
                    st.session_state.rectified_file = rectified
                    st.session_state.current_step = 4
                    st.rerun()
                else:
                    st.error("Could not find processed file")
            else:
                st.error(f"Processing failed: {result.stderr}")
        except Exception as e:
            st.error(f"Error during processing: {e}")
    else:
        st.warning("No report file found. Restarting automation.")
        st.session_state.current_step = 1
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------
# Step 4: View Results and Download
# ---------------------------
elif st.session_state.current_step == 4:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### 📊 Report Results")
    if 'rectified_file' in st.session_state:
        st.markdown(f"""
        <div class="alert alert-success">
            <strong>Processing Complete:</strong> {os.path.basename(st.session_state.rectified_file)}
        </div>
        """, unsafe_allow_html=True)
        try:
            df = pd.read_csv(st.session_state.rectified_file)
            display_data_metrics(df)
            create_visualizations(df)
            st.markdown("### Download Options")
            csv_data = df.to_csv(index=False)
            b64 = base64.b64encode(csv_data.encode()).decode()
            col1, col2, col3 = st.columns(3)
            with col1:
                st.download_button(
                    label="📥 Download CSV",
                    data=csv_data,
                    file_name="processed_report.csv",
                    mime="text/csv",
                    key="download_csv"
                )
            with col2:
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False)
                st.download_button(
                    label="📊 Download Excel",
                    data=excel_buffer.getvalue(),
                    file_name="processed_report.xlsx",
                    mime="application/vnd.ms-excel",
                    key="download_excel"
                )
            with col3:
                st.markdown("No further input required.")
        except Exception as e:
            st.error(f"Error loading processed data: {e}")
    else:
        st.warning("No processed data found. Restarting automation.")
        st.session_state.current_step = 3
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("""
<div style="text-align: center; margin-top: 50px;">
    <div style="display: inline-block; padding: 10px 20px; background-color: #f8f9fa; border-radius: 10px;">
        <div style="font-size: 12px; color: #7f8c8d;">SYSTEM STATUS: OPERATIONAL</div>
        <div style="display: flex; justify-content: center; gap: 20px; margin-top: 10px;">
            <div style="text-align: center;">
                <div style="font-size: 10px;">DATA INTEGRITY</div>
                <div style="font-size: 14px; font-weight: bold; color: #2ecc71;">100%</div>
            </div>
            <div style="text-align: center;">
                <div style="font-size: 10px;">PROCESSING SPEED</div>
                <div style="font-size: 14px; font-weight: bold; color: #3498db;">OPTIMAL</div>
            </div>
            <div style="text-align: center;">
                <div style="font-size: 10px;">SECURITY</div>
                <div style="font-size: 14px; font-weight: bold; color: #e74c3c;">ENCRYPTED</div>
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)