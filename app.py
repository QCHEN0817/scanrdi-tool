import streamlit as st
from docxtpl import DocxTemplate
import os
import json
from datetime import datetime, timedelta

# --- PART 1: UI SETUP ---
st.set_page_config(page_title="Sterility OOS Wizard", layout="wide")

# Sidebar for Platform Selection (The Left Side)
st.sidebar.title("OOS Platform")
selection = st.sidebar.radio("Select Test Type:", ["ScanRDI", "Celsis", "USP 71"])

st.title(f"Sterility Investigation Wizard: {selection}")

# --- PART 2: FORM FIELDS ---
# We use st.columns to create the "Boxes"
col1, col2 = st.columns(2)

with col1:
    oos_id = st.text_input("OOS Number", "OOS-250000")
    client_name = st.text_input("Client Name")
    sample_id = st.text_input("Sample ID")

with col2:
    test_date = st.text_input("Test Date (DDMonYY)", datetime.now().strftime("%d%b%y"))
    sample_name = st.text_input("Sample Name")
    lot_number = st.text_input("Lot Number")

# --- PART 3: PLATFORM SPECIFIC LOGIC ---
data = {
    "oos_id": oos_id,
    "client_name": client_name,
    "sample_id": sample_id,
    "test_date": test_date,
    "sample_name": sample_name,
    "lot_number": lot_number,
}

if selection == "ScanRDI":
    st.subheader("ScanRDI Specific Details")
    c3, c4 = st.columns(2)
    with c3:
        data["prepper_name"] = st.text_input("Prepper Name")
        data["prepper_initial"] = st.text_input("Prepper Initials")
        data["analyst_name"] = st.text_input("Processor Name")
        data["analyst_initial"] = st.text_input("Processor Initials")
    with c4:
        data["scan_id"] = st.text_input("ScanRDI ID (E00...)")
        data["cr_suit"] = st.text_input("Suite/Room #")
        data["bsc_id"] = st.text_input("BSC ID (E00...)")
        data["organism_morphology"] = st.text_input("Org Shape", "rod")
        # [cite_start]Ensure these keys exist for the template [cite: 182]
        data["reader_name"] = data["analyst_name"] 
        data["reader_initial"] = data["analyst_initial"]
        data["changeover_name"] = "N/A"
        data["changeover_initial"] = "N/A"

elif selection == "Celsis":
    st.info("Celsis template and fields pending...")
    
elif selection == "USP 71":
    st.info("USP 71 template and fields pending...")

# --- PART 4: GENERATION ---
if st.button("GENERATE & DOWNLOAD"):
    template_file = f"{selection} OOS template.docx"
    
    if not os.path.exists(template_file):
        st.error(f"Template file '{template_file}' not found in GitHub folder!")
    else:
        try:
            doc = DocxTemplate(template_file)
            
            # [cite_start]Add bracketing dates automatically [cite: 184]
            try:
                dt = datetime.strptime(test_date, "%d%b%y")
                data['date_before_test'] = (dt - timedelta(days=1)).strftime("%d%b%y")
                data['date_after_test'] = (dt + timedelta(days=1)).strftime("%d%b%y")
                data['date_of_weekly'] = test_date
            except:
                pass
            
            doc.render(data)
            output_name = f"{oos_id}_{selection}_Report.docx"
            doc.save(output_name)
            
            with open(output_name, "rb") as f:
                st.download_button("Click here to Download", f, file_name=output_name)
            st.success(f"Report Generated: {output_name}")
            
        except Exception as e:
            st.error(f"Error: {e}")
