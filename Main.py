import streamlit as st
import pandas as pd
import PyPDF2
import pdfplumber
import re
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from datetime import datetime
from typing import List, Dict
from dotenv import load_dotenv
from extract_Eligibility import extract_Eligibility
from extract_primary_data import extract_primary_data
from create_comprehensive_excel_with_formatting import create_comprehensive_excel_with_formatting
from create_addon import create_addon
from create_AddonCoverages import create_AddonCoverages

# Load environment variables
load_dotenv()

def load_users_from_env():
    """Load users from environment variables"""
    users = {}
    for i in range(1, 4):  # 3 users
        username = os.getenv(f'USER{i}_NAME')
        password = os.getenv(f'USER{i}_PASSWORD')
        if username and password:
            users[username] = password
    return users

def check_authentication(username, password, users):
    """Check if username and password match"""
    return username in users and users[username] == password

def authentication_page():
    """Display authentication page"""
    # Load users from environment
    users = load_users_from_env()
    
    # Simple centered layout using Streamlit columns
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### 🔐 Authentication Required")
        st.markdown("Please login to access the PDF to Excel Extractor")
        st.markdown("---")
        
        username = st.text_input("Username", placeholder="Enter your username", key="login_username")
        password = st.text_input("Password", type="password", placeholder="Enter your password", key="login_password")
        
        if st.button("🚀 Login", use_container_width=True, type="primary"):
            if check_authentication(username, password, users):
                st.session_state.authenticated = True
                st.session_state.username = username
                st.success("✅ Login successful!")
                st.rerun()
            else:
                st.error("❌ Invalid username or password!")
        
        # Display available users for testing
        with st.expander("ℹ️ Test Users", expanded=False):
            st.markdown("**Available Test Users:**")
            for user, pwd in users.items():
                st.code(f"👤 {user} | 🔒 {pwd}", language="text")

# Set page config with custom styling
st.set_page_config(
    page_title="PDF to Excel Extractor",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Complete dark theme CSS
st.markdown("""
<style>
    /* Main app background */
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #0f0f23 0%, #1a1a2e 50%, #16213e 100%);
    }
    
    /* Sidebar background */
    [data-testid="stSidebar"] {
        background: linear-gradient(135deg, #0f0f23 0%, #1a1a2e 50%, #16213e 100%);
    }
    
    /* Main content area */
    .main {
        background: linear-gradient(135deg, #0f0f23 0%, #1a1a2e 50%, #16213e 100%);
    }
    
    /* Block container */
    .block-container {
        background: linear-gradient(135deg, #0f0f23 0%, #1a1a2e 50%, #16213e 100%);
    }
    
    /* All text elements */
    .stMarkdown, .stText, .stDataFrame, .stExpander, .stButton, .stTextInput {
        color: white !important;
        background: transparent !important;
    }
    
    /* Dataframes */
    .stDataFrame > div {
        background: rgba(255, 255, 255, 0.1) !important;
        color: white !important;
    }
    
    /* Input fields */
    .stTextInput > div > div > input {
        background: rgba(255, 255, 255, 0.1) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
    }
    
    .stTextInput > div > div > input::placeholder {
        color: rgba(255, 255, 255, 0.6) !important;
    }
    
    /* Buttons */
    .stButton > button {
        background: rgba(255, 255, 255, 0.1) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
    }
    
    /* Expanders */
    .streamlit-expanderHeader {
        background: rgba(255, 255, 255, 0.1) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
    }
    
    /* Text areas */
    .stTextArea > div > div > textarea {
        background: rgba(255, 255, 255, 0.1) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
    }
    
    /* File uploader */
    .stFileUploader > div {
        background: rgba(255, 255, 255, 0.1) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
    }
    
    /* Tabs */
    .stTabs > div > div > div > div {
        background: rgba(255, 255, 255, 0.1) !important;
        color: white !important;
    }
    
    .stTabs > div > div > div > div[data-baseweb="tab"] {
        background: rgba(255, 255, 255, 0.1) !important;
        color: white !important;
    }
    
    /* Success and error messages */
    .stSuccess, .stError {
        background: rgba(255, 255, 255, 0.1) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
    }
    
    /* Info boxes */
    .stAlert {
        background: rgba(255, 255, 255, 0.1) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
    }
    
    /* Header styling for main app */
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        margin-bottom: 1rem;
        text-align: center;
        color: white !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
    }
    
    .main-header h1 {
        margin: 0;
        font-size: 2.5rem;
        font-weight: 700;
        text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.3);
        color: white !important;
    }
    
    .main-header p {
        margin: 0.5rem 0 0 0;
        font-size: 1.1rem;
        opacity: 0.9;
        color: white !important;
    }
    
    /* Override any white backgrounds */
    * {
        background-color: transparent !important;
    }
    
    /* Ensure all containers have dark background */
    div[data-testid="stVerticalBlock"] {
        background: linear-gradient(135deg, #0f0f23 0%, #1a1a2e 50%, #16213e 100%) !important;
    }
</style>
""", unsafe_allow_html=True)

def extract_text_from_pdf(pdf_file):
    """Extract text from uploaded PDF file using both PyPDF2 and pdfplumber for better accuracy"""
    try:
        # Try pdfplumber first for better text extraction
        text = ""
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        
        # If pdfplumber didn't extract much text, try PyPDF2 as backup
        if len(text.strip()) < 100:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
        
        return text
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None

def display_features():
    """Display feature highlights"""
    pass

def main():
    # Initialize session state for authentication
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    
    # Check if user is authenticated
    if not st.session_state.authenticated:
        authentication_page()
        return
    
    # Main header with user info
    st.markdown(f"""
    <div class="main-header">
        <h1>🚀 PDF to Excel Extractor</h1>
        <p>Welcome, <strong>{st.session_state.username}</strong>! Transform your PDF documents into structured Excel files with intelligent data extraction</p>
    </div>
    """, unsafe_allow_html=True)
    
    # User info and logout in sidebar
    with st.sidebar:
        st.markdown("### 👤 User Session")
        st.info(f"**Logged in as:** {st.session_state.username}")
        
        if st.button("🚪 Logout", type="secondary", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.username = None
            st.rerun()
        
        st.markdown("---")
    
    # Sidebar for file upload
    uploaded_file = st.sidebar.file_uploader(
        "Choose a PDF file",
        type=['pdf'],
        help="Upload a PDF file containing eligibility information"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        
        # Extract text from PDF
        with st.spinner("🔍 Extracting text from PDF..."):
            text = extract_text_from_pdf(uploaded_file)
        
        if text:
            # Text preview
            with st.expander("📖 Preview Extracted Text", expanded=False):
                st.text_area(
                    "Extracted Text Preview", 
                    text[:400000] + "..." if len(text) > 400000 else text, 
                    height=200,
                    label_visibility="collapsed"
                )
            
            # Extract eligibility details
            with st.spinner("🧠 Analyzing and extracting data..."):
                eligibility_data = extract_Eligibility(text)
                primary_data = extract_primary_data(text)
                addon_data = create_addon(text)
                AddonCoverages_data = create_AddonCoverages(text)
                
                # Debug: Print the data structures
                print("DEBUG: Data extraction results:")
                print(f"eligibility_data type: {type(eligibility_data)}, length: {len(eligibility_data) if isinstance(eligibility_data, list) else 'N/A'}")
                print(f"primary_data type: {type(primary_data)}, length: {len(primary_data) if isinstance(primary_data, list) else 'N/A'}")
                print(f"addon_data type: {type(addon_data)}, length: {len(addon_data) if isinstance(addon_data, list) else 'N/A'}")
                print(f"AddonCoverages_data type: {type(AddonCoverages_data)}, length: {len(AddonCoverages_data) if isinstance(AddonCoverages_data, list) else 'N/A'}")
                
                if AddonCoverages_data and isinstance(AddonCoverages_data, list) and len(AddonCoverages_data) > 0:
                    print("DEBUG: AddonCoverages_data first item keys:")
                    for key, value in AddonCoverages_data[0].items():
                        if value:  # Only print non-empty values
                            print(f"  {key}: {value}")
            
            # Display extracted data in tabs
            st.subheader("📊 Extracted Data")
            
            tab1, tab2, tab3, tab4 = st.tabs([
                "🛡️ Eligibility Coverage", 
                "🏥 Primary Coverage", 
                "➕ Addon Coverage", 
                "🔧 Addon Coverages"
            ])
            
            with tab1:
                df_eligibility = pd.DataFrame(eligibility_data)
                st.dataframe(df_eligibility, use_container_width=True, height=300)
            
            with tab2:
                df_primary = pd.DataFrame(primary_data)
                st.dataframe(df_primary, use_container_width=True, height=300)
            
            with tab3:
                df_addon = pd.DataFrame(addon_data)
                st.dataframe(df_addon, use_container_width=True, height=300)
            
            with tab4:
                df_AddonCoverages = pd.DataFrame(AddonCoverages_data)
                st.dataframe(df_AddonCoverages, use_container_width=True, height=300)
            
            # Generate Excel file
            with st.sidebar:
                with st.spinner("Creating Excel file..."):
                    try:
                        # Generate the Excel workbook
                        wb = create_comprehensive_excel_with_formatting(
                            eligibility_data, primary_data, addon_data, AddonCoverages_data
                        )

                        # Save workbook to BytesIO
                        excel_buffer = BytesIO()
                        wb.save(excel_buffer)
                        excel_buffer.seek(0)

                        # Create download button with original PDF filename
                        # Get the original PDF filename without extension
                        pdf_filename = uploaded_file.name
                        base_filename = pdf_filename.rsplit('.', 1)[0] if '.' in pdf_filename else pdf_filename
                        
                        st.download_button(
                            label="📥 Download Excel File",
                            data=excel_buffer.getvalue(),
                            file_name=f"{base_filename}_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                    except Exception as e:
                        st.error(f"❌ Error generating Excel file: {str(e)}")

if __name__ == "__main__":
    main()