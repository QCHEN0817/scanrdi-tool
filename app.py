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
data = {}

# --- SECTION 1: GENERAL INFO ---
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
    data["monthly_cleaning_date"] = st.text_input("Monthly Cleaning Date")

# --- SECTION 2: SCANRDI SPECIFIC LOGIC ---
if selection == "ScanRDI":
    st.header("2. Personnel & Equipment")
    
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

    # Equipment Smart Lookup
    st.subheader("Equipment & Facility (Auto-Fill)")
    e1, e2 = st.columns(2)
    with e1:
        # Selection includes your updated list
        bsc_choice = st.selectbox("Select BSC ID", ["1310", "1309", "1311", "1312", "1314", "1313", "1316", "1798", "Other"])
        data["bsc_id"] = bsc_choice if bsc_choice != "Other" else st.text_input("Enter Manual BSC ID")
        
        # Room Mapping Logic
        if data["bsc_id"] in ["1310", "1309"]: 
            data["cr_id"], data["cr_suit"] = "117", "117"
        elif data["bsc_id"] in ["1311", "1312"]: 
            data["cr_id"], data["cr_suit"] = "116", "116"
        elif data["bsc_id"] in ["1314", "1313"]: 
            data["cr_id"], data["cr_suit"] = "115", "115"
        elif data["bsc_id"] in ["1316", "1798"]: 
            data["cr_id"], data["cr_suit"] = "114", "114"
        else:
            data["cr_id"] = st.text_input("Room ID")
            data["cr_suit"] = st.text_input("Suite #")

        # Smart A/B Suffix & Description Logic (Even=B/Innermost, Odd=A/Middle)
        try:
            bsc_num = int(data["bsc_id"])
            suffix = "B" if bsc_num % 2 == 0 else "A"
            location_desc = "innermost ISO 7 room" if suffix == "B" else "middle ISO 7 buffer room"
        except:
            suffix = "B"
            location_desc = "innermost ISO 7 room"

        st.info(f"Automatically Mapped to Room {data['cr_id']} / Suite {data['cr_suit']}{suffix} ({location_desc})")

    with e2:
        data["scan_id"] = st.text_input("ScanRDI ID (last 4 digits)")
        data["organism_morphology"] = st.selectbox("Org Shape", ["rod", "cocci", "yeast/mold"])

    # Table 1 EM Bracketing (Mapping Source 184)
    st.header("3. Table 1: EM Observations")
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

    st.subheader("Weekly Bracketing")
    w1, w2 = st.columns(2)
    with w1:
        data["obs_air_wk_of"] = st.text_input("Weekly Air Obs", "No Growth")
        data["etx_air_wk_of"] = st.text_input("Air ETX #", "N/A")
        data["id_air_wk_of"] = st.text_input("Air ID", "N/A")
        data["weekly_initial"] = st.text_input("Weekly Monitor Initials")
    with w2:
        data["obs_room_wk_of"] = st.text_input("Weekly Room Obs", "No Growth")
        data["etx_room_wk_of"] = st.text_input("Room ETX #", "N/A")
        data["id_room_wk_of"] = st.text_input("Room ID", "N/A")
        data["date_of_weekly"] = st.text_input("Date of Weekly Monitoring")

    # Smart Narrative Calculation
    if st.button("🪄 Auto-Generate Narrative"):
        # Mapping {{ equipment_summary }} (Source 182)
        data["equipment_summary"] = (
            f"The ISO 5 BSC E00{data['bsc_id']}, located in the {location_desc} "
            f"(Suite {data['cr_suit']}{suffix}), was thoroughly cleaned and disinfected "
            f"prior to procedures in accordance with SOP 2.600.018."
        )
        
        # Mapping {{ narrative_summary }} and {{ em_details }} (Source 182)
        if data["obs_pers_dur"] == "No Growth" and data["obs_surf_dur"] == "No Growth":
            data["narrative_summary"] = (
                "Upon analyzing the environmental monitoring results, no microbial growth was observed "
                "in personal sampling (left touch and right touch), surface sampling, or settling plates."
            )
            data["em_details"] = (
                "Weekly active air sampling and weekly surface sampling from both the week of testing "
                "and the week before testing showed no microbial growth."
            )
        else:
            data["narrative_summary"] = "Microbial growth was observed during environmental monitoring as detailed in Table 1."
            data["em_details"] = "Growth details are provided in the attached EM identification reports."
        st.success("Narratives Prepared!")

# --- GENERATION ---
if st.button("🚀 GENERATE REPORT"):
    template_path = f"{selection} OOS template.docx"
    if os.path.exists(template_path):
        doc = DocxTemplate(template_path)
        
        # Calculate bracketing dates (Source 184)
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
        st.error(f"Template '{template_path}' missing from GitHub folder!")
