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
    
    # Restored Changeover Personnel Fields
    p1, p2, p3, p4 = st.columns(4)
    with p1:
        data["prepper_name"] = st.text_input("Prepper Name")
        data["prepper_initial"] = st.text_input("Prepper Init")
    with p2:
        data["analyst_name"] = st.text_input("Processor Name")
        data["analyst_initial"] = st.text_input("Processor Init")
    with p3:
        # CHG Analyst info restored
        data["changeover_name"] = st.text_input("Changeover Name")
        data["changeover_initial"] = st.text_input("Changeover Init")
    with p4:
        data["reader_name"] = st.text_input("Reader Name")
        data["reader_initial"] = st.text_input("Reader Init")

    # Equipment Smart Lookup
    st.subheader("Equipment & Facility (Auto-Fill)")
    e1, e2 = st.columns(2)
    
    # BSC Options based on your facility
    bsc_list = ["1310", "1309", "1311", "1312", "1314", "1313", "1316", "1798", "Other"]

    with e1:
        # Processing BSC
        proc_bsc = st.selectbox("Select Processing BSC ID", bsc_list)
        data["bsc_id"] = proc_bsc if proc_bsc != "Other" else st.text_input("Enter Manual Proc BSC ID")
        
        # Room Mapping for Processing
        if data["bsc_id"] in ["1310", "1309"]: data["cr_id"], data["cr_suit"] = "117", "117"
        elif data["bsc_id"] in ["1311", "1312"]: data["cr_id"], data["cr_suit"] = "116", "116"
        elif data["bsc_id"] in ["1314", "1313"]: data["cr_id"], data["cr_suit"] = "115", "115"
        elif data["bsc_id"] in ["1316", "1798"]: data["cr_id"], data["cr_suit"] = "114", "114"
        else:
            data["cr_id"] = st.text_input("Room ID")
            data["cr_suit"] = st.text_input("Suite #")

        # Even/Odd logic for Proc Room Suffix
        try:
            p_suffix = "B" if int(data["bsc_id"]) % 2 == 0 else "A"
            p_loc = "innermost ISO 7 room" if p_suffix == "B" else "middle ISO 7 buffer room"
        except:
            p_suffix, p_loc = "B", "innermost ISO 7 room"
            
        st.info(f"Processing: Room {data['cr_id']} / Suite {data['cr_suit']}{p_suffix} ({p_loc})")

    with e2:
        # Changeover BSC Logic restored
        chg_bsc = st.selectbox("Select Changeover BSC ID", bsc_list)
        data["chgbsc_id"] = chg_bsc if chg_bsc != "Other" else st.text_input("Enter Manual CHG BSC ID")
        
        # Determine Changeover Suite/Room info
        if data["chgbsc_id"] in ["1310", "1309"]: c_room, c_suite = "117", "117"
        elif data["chgbsc_id"] in ["1311", "1312"]: c_room, c_suite = "116", "116"
        elif data["chgbsc_id"] in ["1314", "1313"]: c_room, c_suite = "115", "115"
        elif data["chgbsc_id"] in ["1316", "1798"]: c_room, c_suite = "114", "114"
        else:
            c_room = st.text_input("CHG Room ID")
            c_suite = st.text_input("CHG Suite #")

        # Even/Odd logic for CHG Room Suffix
        try:
            c_suffix = "B" if int(data["chgbsc_id"]) % 2 == 0 else "A"
            c_loc = "innermost ISO 7 room" if c_suffix == "B" else "middle ISO 7 buffer room"
        except:
            c_suffix, c_loc = "A", "middle ISO 7 buffer room"

        st.info(f"Changeover: Room {c_room} / Suite {c_suite}{c_suffix} ({c_loc})")

    # Findings
    f1, f2 = st.columns(2)
    with f1:
        data["scan_id"] = st.text_input("ScanRDI ID (last 4 digits)")
        data["organism_morphology"] = st.selectbox("Org Shape", ["rod", "cocci", "yeast/mold"])
    with f2:
        data["control_positive"] = st.text_input("Pos Control", "B. subtilis")
        data["control_lot"] = st.text_input("Control Lot")
        data["control_data"] = st.text_input("Control Exp Date")

    # Table 1 EM Data
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

    # Auto-Narrative Generation Logic restored
    if st.button("🪄 Auto-Generate Narrative"):
        # Combined Equipment Summary for Proc and CHG (Source 182)
        data["equipment_summary"] = (
            f"Sample processing was conducted within the ISO 5 BSC in the {p_loc} "
            f"(Suite {data['cr_suit']}{p_suffix}, BSC E00{data['bsc_id']}) by {data['analyst_name']} "
            f"and the changeover step was conducted within the ISO 5 BSC in the {c_loc} "
            f"(Suite {c_suite}{c_suffix}, BSC E00{data['chgbsc_id']}) by {data['changeover_name']}."
        )
        
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
            data["em_details"] = "Identification of growth is provided in the attached reports."
        st.success("Changeover details and narratives prepared!")

# --- GENERATION ---
if st.button("🚀 GENERATE REPORT"):
    template_path = f"{selection} OOS template.docx"
    if os.path.exists(template_path):
        doc = DocxTemplate(template_path)
        
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
        st.error(f"Template '{template_path}' missing from folder!")
