import streamlit as st
from docxtpl import DocxTemplate
import os
from datetime import datetime, timedelta

# --- PAGE CONFIG ---
st.set_page_config(page_title="Sterility OOS Wizard", layout="wide")

# --- SIDEBAR NAVIGATION ---
st.sidebar.title("Sterility Platform")
selection = st.sidebar.radio("Select Test Type:", ["ScanRDI", "Celsis", "USP 71"])

st.title(f"Investigation Wizard: {selection}")

# --- DATA DICTIONARY ---
# This dictionary will hold every value for the template
data = {}

# --- SECTION 1: GENERAL INFO (COMMON) ---
st.header("1. General Test Details")
col1, col2, col3 = st.columns(3)
with col1:
    data["oos_id"] = st.text_input("OOS Number", "OOS-250000")
    data["client_name"] = st.text_input("Client Name", "Pharmacy Name")
    data["sample_id"] = st.text_input("Sample ID", "SCAN-250000")
with col2:
    data["test_date"] = st.text_input("Test Date (DDMonYY)", datetime.now().strftime("%d%b%y"))
    data["sample_name"] = st.text_input("Sample / Active Name")
    data["lot_number"] = st.text_input("Lot Number")
with col3:
    data["dosage_form"] = st.text_input("Dosage Form", "Injectable")
    data["test_record"] = st.text_input("Record Ref (e.g. 060925-1040)")
    data["monthly_cleaning_date"] = st.text_input("Monthly Cleaning Date", "01Jun25")

# --- SECTION 2: PLATFORM SPECIFIC ---
if selection == "ScanRDI":
    st.header("2. Personnel & Equipment")
    
    # Personnel
    p1, p2, p3, p4 = st.columns(4)
    with p1:
        data["prepper_name"] = st.text_input("Prepper Name")
        data["prepper_initial"] = st.text_input("Prepper Init")
    with p2:
        data["analyst_name"] = st.text_input("Processor Name")
        data["analyst_initial"] = st.text_input("Processor Init")
    with p3:
        data["changeover_name"] = st.text_input("Changeover Name", "N/A")
        data["changeover_initial"] = st.text_input("Changeover Init", "N/A")
    with p4:
        data["reader_name"] = st.text_input("Reader Name")
        data["reader_initial"] = st.text_input("Reader Init")

    # Equipment & Results
    st.header("3. Findings & EM Data")
    e1, e2, e3 = st.columns(3)
    with e1:
        data["scan_id"] = st.text_input("ScanRDI ID (last 4 digits)")
        data["cr_id"] = st.text_input("Room ID (e.g. 117)")
        data["cr_suit"] = st.text_input("Suite # (e.g. 117)")
        data["bsc_id"] = st.text_input("BSC ID (last 4 digits)")
    with e2:
        data["control_positive"] = st.text_input("Pos Control", "B. subtilis")
        data["control_lot"] = st.text_input("Control Lot")
        data["control_data"] = st.text_input("Control Exp Date")
        data["organism_morphology"] = st.text_input("Org Shape (rod/cocci)")
    with e3:
        data["weekly_initial"] = st.text_input("Weekly Monitor Initials")
        data["date_of_weekly"] = st.text_input("Date of Weekly Monitoring")

    # Table 1 EM Observations
    st.subheader("Table 1: EM Observations")
    em1, em2, em3 = st.columns(3)
    with em1:
        data["obs_pers_dur"] = st.text_input("Personnel Obs", "No Growth")
        data["etx_pers_dur"] = st.text_input("Pers ETX #", "N/A")
        data["id_pers_dur"] = st.text_input("Pers ID", "N/A")
    with em2:
        data["obs_surf_dur"] = st.text_input("Surface Obs", "No Growth")
        data["etx_surf_dur"] = st.text_input("Surf ETX #", "N/A")
        data["id_surf_dur"] = st.text_input("Surf ID", "N/A")
    with em3:
        data["obs_sett_dur"] = st.text_input("Settling Obs", "No Growth")
        data["etx_sett_dur"] = st.text_input("Sett ETX #", "N/A")
        data["id_sett_dur"] = st.text_input("Sett ID", "N/A")

    st.subheader("Weekly Results")
    w1, w2 = st.columns(2)
    with w1:
        data["obs_air_wk_of"] = st.text_input("Weekly Air Obs", "No Growth")
        data["etx_air_wk_of"] = st.text_input("Air ETX #", "N/A")
        data["id_air_wk_of"] = st.text_input("Air ID", "N/A")
    with w2:
        data["obs_room_wk_of"] = st.text_input("Weekly Room Obs", "No Growth")
        data["etx_room_wk_of"] = st.text_input("Room ETX #", "N/A")
        data["id_room_wk_of"] = st.text_input("Room ID", "N/A")

    # Narrative Summary Logic
    st.header("4. Smart Summaries")
    if st.button("🪄 Auto-Generate Narrative"):
        if data["obs_pers_dur"] == "No Growth" and data["obs_surf_dur"] == "No Growth":
            data["narrative_summary"] = "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling, surface sampling, or settling plates."
        else:
            data["narrative_summary"] = "Microbial growth was observed during environmental monitoring as detailed below."
        st.success("Narrative Prepared!")

# --- GENERATION ---
if st.button("🚀 GENERATE REPORT"):
    template_path = f"{selection} OOS template.docx"
    if os.path.exists(template_path):
        doc = DocxTemplate(template_path)
        
        # Calculate bracketing dates
        try:
            dt = datetime.strptime(data["test_date"], "%d%b%y")
            data["date_before_test"] = (dt - timedelta(days=1)).strftime("%d%b%y")
            data["date_after_test"] = (dt + timedelta(days=1)).strftime("%d%b%y")
        except:
            pass
            
        doc.render(data)
        out_file = f"{data['oos_id']}_{selection}.docx"
        doc.save(out_file)
        
        with open(out_file, "rb") as f:
            st.download_button("📂 Download Final Report", f, file_name=out_file)
    else:
        st.error(f"Template '{template_path}' missing from GitHub!")
