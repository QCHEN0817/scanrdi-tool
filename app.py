import streamlit as st
from docxtpl import DocxTemplate
import os
from datetime import datetime, timedelta

# --- PAGE CONFIG ---
st.set_page_config(page_title="Sterility OOS Wizard", layout="wide")

# --- CUSTOM STYLING ---
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; }
    .main { background-color: #f9f9f9; }
    </style>
    """, unsafe_allow_html=True)

# Helper function with specific exceptions for initials
def get_initials(name):
    if not name:
        return ""
    
    # Standardize name for checking exceptions
    clean_name = name.strip().lower()
    
    # Specific exceptions
    exceptions = {
        "varsha subramanian": "VV",
        "clea s. garza": "CSG",
        "clea garza": "CSG",
        "guanchen li": "GL",
        "guan chen li": "GL",
        "qiyue chen": "QYC"
    }
    
    if clean_name in exceptions:
        return exceptions[clean_name]
    
    # Default logic: First letter of first and last name
    parts = name.strip().split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    elif len(parts) == 1:
        return parts[0][0].upper()
    return ""

# --- SIDEBAR NAVIGATION ---
with st.sidebar:
    st.title("Sterility Platforms")
    st.write("Select a platform:")
    
    if "active_platform" not in st.session_state:
        st.session_state.active_platform = "ScanRDI"

    if st.button("ScanRDI"):
        st.session_state.active_platform = "ScanRDI"
    if st.button("Celsis"):
        st.session_state.active_platform = "Celsis"
    if st.button("USP 71"):
        st.session_state.active_platform = "USP 71"
    
    st.divider()
    st.success(f"Current: {st.session_state.active_platform}")

selection = st.session_state.active_platform
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
    st.header("2. Personnel & Changeover")
    
    p1, p2, p3, p4 = st.columns(4)
    with p1:
        prepper = st.text_input("Prepper Name")
        data["prepper_name"] = prepper
        data["prepper_initial"] = st.text_input("Prepper Init", value=get_initials(prepper))
    with p2:
        analyst = st.text_input("Processor Name")
        data["analyst_name"] = analyst
        data["analyst_initial"] = st.text_input("Processor Init", value=get_initials(analyst))
    with p3:
        changeover = st.text_input("Changeover Name")
        data["changeover_name"] = changeover
        data["changeover_initial"] = st.text_input("Changeover Init", value=get_initials(changeover))
    with p4:
        reader = st.text_input("Reader Name")
        data["reader_name"] = reader
        data["reader_initial"] = st.text_input("Reader Init", value=get_initials(reader))

    st.subheader("Equipment & Facility Smart Lookup")
    e1, e2 = st.columns(2)
    bsc_list = ["1310", "1309", "1311", "1312", "1314", "1313", "1316", "1798", "Other"]

    def get_room_logic(bsc_id):
        try:
            num = int(bsc_id)
            suffix = "B" if num % 2 == 0 else "A"
            location = "innermost ISO 7 room" if suffix == "B" else "middle ISO 7 buffer room"
        except:
            suffix, location = "B", "innermost ISO 7 room"
            
        if bsc_id in ["1310", "1309"]: room, suite = "117", "117"
        elif bsc_id in ["1311", "1312"]: room, suite = "116", "116"
        elif bsc_id in ["1314", "1313"]: room, suite = "115", "115"
        elif bsc_id in ["1316", "1798"]: room, suite = "114", "114"
        else: room, suite = "Unknown", "Unknown"
        return room, suite, suffix, location

    with e1:
        proc_bsc = st.selectbox("Select Processing BSC ID", bsc_list)
        data["bsc_id"] = proc_bsc if proc_bsc != "Other" else st.text_input("Manual Proc BSC")
        p_room, p_suite, p_suffix, p_loc = get_room_logic(data["bsc_id"])
        data["cr_id"], data["cr_suit"] = p_room, p_suite
        st.caption(f"Processor: {p_room} / Suite {p_suite}{p_suffix} ({p_loc})")

    with e2:
        chg_bsc = st.selectbox("Select Changeover BSC ID", bsc_list)
        data["chgbsc_id"] = chg_bsc if chg_bsc != "Other" else st.text_input("Manual CHG BSC")
        c_room, c_suite, c_suffix, c_loc = get_room_logic(data["chgbsc_id"])
        st.caption(f"Changeover: {c_room} / Suite {c_suite}{c_suffix} ({c_loc})")

    st.header("3. Findings & EM Data")
    f1, f2 = st.columns(2)
    with f1:
        scan_id_options = ["1230", "2017", "1040", "1877", "2225", "2132"]
        data["scan_id"] = st.selectbox("ScanRDI ID", scan_id_options)
        data["organism_morphology"] = st.selectbox("Org Shape", ["rod", "cocci", "yeast/mold"])
    with f2:
        control_choices = ["A. brasiliensis", "B. subtilis", "C. albicans", "C. sporogenes", "P. aeruginosa", "S. aureus"]
        data["control_positive"] = st.selectbox("Positive Control", control_choices)
        data["control_lot"] = st.text_input("Control Lot")
        data["control_data"] = st.text_input("Control Exp Date")

    st.subheader("Table 1: EM Observations")
    em_col1, em_col2, em_col3 = st.columns(3)
    with em_col1:
        data["obs_pers_dur"] = st.text_input("Personnel Obs", "No Growth")
        data["etx_pers_dur"] = st.text_input("Pers ETX #", "N/A")
        data["id_pers_dur"] = st.text_input("Pers ID", "N/A")
    with em_col2:
        data["obs_surf_dur"] = st.text_input("Surface Obs", "No Growth")
        data["etx_surf_dur"] = st.text_input("Surf ETX #", "N/A")
        data["id_surf_dur"] = st.text_input("Surf ID", "N/A")
    with em_col3:
        data["obs_sett_dur"] = st.text_input("Settling Obs", "No Growth")
        data["etx_sett_dur"] = st.text_input("Sett ETX #", "N/A")
        data["id_sett_dur"] = st.text_input("Sett ID", "N/A")

    st.subheader("Weekly Bracketing")
    wk1, wk2 = st.columns(2)
    with wk1:
        data["obs_air_wk_of"] = st.text_input("Weekly Air Obs", "No Growth")
        data["etx_air_wk_of"] = st.text_input("Air ETX #", "N/A")
        data["id_air_wk_of"] = st.text_input("Air ID", "N/A")
        data["weekly_initial"] = st.text_input("Weekly Monitor Init")
    with wk2:
        data["obs_room_wk_of"] = st.text_input("Weekly Room Obs", "No Growth")
        data["etx_room_wk_of"] = st.text_input("Room ETX #", "N/A")
        data["id_room_wk_of"] = st.text_input("Room ID", "N/A")
        data["date_of_weekly"] = st.text_input("Date of Weekly Monitoring")

    st.header("4. EM Narratives")
    
    if 'equipment_summary' not in st.session_state: st.session_state.equipment_summary = ""
    if 'narrative_summary' not in st.session_state: st.session_state.narrative_summary = ""
    if 'em_details' not in st.session_state: st.session_state.em_details = ""

    if st.button("🪄 Auto-Generate Narratives"):
        # Equipment Summary (Source 182)
        st.session_state.equipment_summary = (
            f"Sample processing was conducted within the ISO 5 BSC in the {p_loc} "
            f"(Suite {p_suite}{p_suffix}, BSC E00{data['bsc_id']}) by {data['analyst_name']} "
            f"and the changeover step was conducted within the ISO 5 BSC in the {c_loc} "
            f"(Suite {c_suite}{c_suffix}, BSC E00{data['chgbsc_id']}) by {data['changeover_name']}."
        )
        
        # Combined Narrative for "No Growth"
        if data["obs_pers_dur"] == "No Growth" and data["obs_surf_dur"] == "No Growth":
            st.session_state.narrative_summary = (
                "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), "
                "surface sampling, or settling plates. Weekly active air sampling and weekly surface sampling from both the week of testing and the week "
                "before testing showed no microbial growth."
            )
            st.session_state.em_details = ""
        else:
            st.session_state.narrative_summary = "Microbial growth was observed during environmental monitoring as detailed in Table 1."
            st.session_state.em_details = "Growth details are documented in the attached reports."
        st.rerun()

    data["equipment_summary"] = st.session_state.equipment_summary
    data["narrative_summary"] = st.text_area("Narrative Summary (Editable)", value=st.session_state.narrative_summary, height=150)
    data["em_details"] = st.text_area("EM Details (Editable)", value=st.session_state.em_details, height=100)

# --- FINAL GENERATION ---
if st.button("🚀 GENERATE FINAL REPORT"):
    template_name = f"{selection} OOS template.docx"
    if os.path.exists(template_name):
        doc = DocxTemplate(template_name)
        try:
            # Bracketing dates calculation (Source 184)
            dt_obj = datetime.strptime(data["test_date"], "%d%b%y")
            data["date_before_test"] = (dt_obj - timedelta(days=1)).strftime("%d%b%y")
            data["date_after_test"] = (dt_obj + timedelta(days=1)).strftime("%d%b%y")
        except: pass
        doc.render(data)
        out_name = f"{data['oos_id']}_{selection}.docx"
        doc.save(out_name)
        with open(out_name, "rb") as f:
            st.download_button("📂 Download Document", f, file_name=out_name)
    else:
        st.error(f"Template '{template_name}' not found.")
