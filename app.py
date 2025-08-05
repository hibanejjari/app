import streamlit as st
import importlib

# Set config
st.set_page_config(page_title="PO Workflow Generator", layout="wide")

# Sidebar navigation
st.sidebar.title("Navigation")
page = st.sidebar.radio("Go to", ("🏠 Home", "📊 General Report", "💼 SAP Report", "🌐 Ariba Report"))

# Render selected page
if page == "🏠 Home":
    st.title("👋 Welcome to Your PO Workflow Report Generator")
    st.markdown("""
    Use the sidebar to navigate to:
    - 📊 **General Report**
    - 💼 **SAP Report**
    - 🌐 **Ariba Report**

    This app helps you:
    - ✅ Upload Excel files  
    - 📈 Visualize Purchase Order metrics  
    - 📤 Download automated PowerPoint reports  
    """)
    
elif page == "📊 General Report":
    module = importlib.import_module("pages.General_Report")
    module.main()

elif page == "💼 SAP Report":
    module = importlib.import_module("pages.SAP_Report")
    module.main()

elif page == "🌐 Ariba Report":
    module = importlib.import_module("pages.Ariba_Report")
    module.main()
