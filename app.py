import streamlit as st
from docxtpl import DocxTemplate
import os
from datetime import datetime, timedelta

# --- PAGE CONFIG ---
st.set_page_config(page_title="Sterility OOS Wizard", layout="wide")

# --- CUSTOM STYLING ---
st.markdown("""
    <style>
    .main { background-color: #f5f5f5; }
    .stButton>button { width: 100%; height: 3em; background-color: #007bff; color: white; }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR PLATFORM SELECTOR ---
# Restoring the professional sidebar style you liked
with st.sidebar:
    st.title("🛡️ Sterility Tool")
    selection = st.selectbox("Choose Platform:", ["ScanRDI", "Celsis", "USP 71"])
    st.info(f"Currently working on {selection} template.")

st.title(f"Sterility Investigation Wizard: {selection}")

# --- DATA DICTIONARY ---
data = {}

# --- SECTION 1: GENERAL INFO (COMMON) ---
st.header("1. General Test Details")
col1, col2, col3 = st.columns(3)
with col1:
    data["oos_id"] = st.text_input("OOS Number", "OOS-250000") [cite: 182]
    data["client_name"] = st.text_input("Client Name", "Pharmacy Name") [cite: 182]
    data["sample_id"] = st.text_input("Sample ID", "SCAN-250000") [cite: 182]
with col2:
    data["test_date"] = st.text_input("Test Date (DDMonYY)", datetime.now().strftime("%d%b%y")) [cite: 182]
    data["sample_name"] = st.text_input("Sample / Active Name") [cite: 182]
    data["lot_number"] = st.text_input("Lot Number") [cite: 182]
with col3:
    data["dosage_form"] = st.text_input("Dosage Form", "Injectable") [cite: 182]
    data["test_record"] = st.text_input("Record Ref (e.g. 060925-1040)") [cite: 182]
    data["monthly_cleaning_date"] = st.text_input("Monthly Cleaning Date") [cite: 182]

# --- SECTION 2: SCANRDI SPECIFIC LOGIC ---
if selection == "ScanRDI":
    st.header("2. Personnel & Changeover")
    
    # All analysts involved (Source 182)
    p1, p2, p3, p4 = st.columns(4)
    with p1:
        data["prepper_name"] = st.text_input("Prepper Name") [cite: 182]
        data["prepper_initial"] = st.text_input("Prepper Init") [cite: 182]
    with p2:
        data["analyst_name"] = st.text_input("Processor Name") [cite: 182]
        data["analyst_initial"] = st.text_input("Processor Init") [cite: 182]
    with p3:
        # Full Changeover integration restored
        data["changeover_name"] = st.text_input("Changeover Name") [cite: 4]
        data["changeover_initial"] = st.text_input("Changeover Init") [cite: 4]
    with p4:
        data["reader_name"] = st.text_input("Reader Name") [cite: 182]
        data["reader_initial"] = st.text_input("Reader Init") [cite: 182]

    st.subheader("Equipment & Facility Smart Lookup")
    e1, e2 = st.columns(2)
    
    # Master BSC List with mappings
    bsc_list = ["1310", "1309", "1311", "1312", "1314", "1313", "1316", "1798", "Other"]

    def get_room_logic(bsc_id):
        # Even (双数) = B (Innermost), Odd (单数) = A (Middle)
        try:
            num = int(bsc_id)
            suffix = "B" if num % 2 == 0 else "A"
            location = "innermost ISO 7 room" if suffix == "B" else "middle ISO 7 buffer room"
        except:
            suffix, location = "B", "innermost ISO 7 room"
            
        # Room/Suite mapping based on your specific request
        if bsc_id in ["1310", "1309"]: room, suite = "117", "117" [cite: 48, 182]
        elif bsc_id in ["1311", "1312"]: room, suite = "116", "116" [cite: 134, 182]
        elif bsc_id in ["1314", "1313"]: room, suite = "115", "115"
        elif bsc_id in ["1316", "1798"]: room, suite = "114", "114"
        else: room, suite = "Unknown", "Unknown"
        
        return room, suite, suffix, location

    with e1:
        proc_bsc = st.selectbox("Select Processing BSC ID", bsc_list)
        data["bsc_id"] = proc_bsc if proc_bsc != "Other" else st.text_input("Manual Proc BSC") [cite: 182]
        p_room, p_suite, p_suffix, p_loc = get_room_logic(data["bsc_id"])
        data["cr_id"], data["cr_suit"] = p_room, p_suite [cite: 182]
        st.caption(f"📍 Processor Room: {p_room} / Suite {p_suite}{p_suffix} ({p_loc})")

    with e2:
        chg_bsc = st.selectbox("Select Changeover BSC ID", bsc_list)
        data["chgbsc_id"] = chg_bsc if chg_bsc != "Other" else st.text_input("Manual CHG BSC") [cite: 4]
        c_room, c_suite, c_suffix, c_loc = get_room_logic(data["chgbsc_id"])
        st.caption(f"📍 Changeover Room: {c_room} / Suite {c_suite}{c_suffix} ({c_loc})")

    st.header("3. Findings & EM Data (Table 1)")
    f1, f2 = st.columns(2)
    with f1:
        data["scan_id"] = st.text_input("ScanRDI ID (last 4 digits)") [cite: 182]
        data["organism_morphology"] = st.selectbox("Org Shape", ["rod", "cocci", "yeast/mold"]) [cite: 182]
    with f2:
        data["control_positive"] = st.text_input("Pos Control", "B. subtilis") [cite: 182]
        data["control_lot"] = st.text_input("Control Lot") [cite: 182]
        data["control_data"] = st.text_input("Control Exp Date") [cite: 182]

    # Mapping Table 1 Variables (Source 184)
    em_col1, em_col2, em_col3 = st.columns(3)
    with em_col1:
        data["obs_pers_dur"] = st.text_input("Personnel Obs", "No Growth") [cite: 184]
        data["etx_pers_dur"] = st.text_input("Pers ETX #", "N/A") [cite: 184]
        data["id_pers_dur"] = st.text_input("Pers ID", "N/A") [cite: 184]
    with em_col2:
        data["obs_surf_dur"] = st.text_input("Surface Obs", "No Growth") [cite: 184]
        data["etx_surf_dur"] = st.text_input("Surf ETX #", "N/A") [cite: 184]
        data["id_surf_dur"] = st.text_input("Surf ID", "N/A") [cite: 184]
    with em_col3:
        data["obs_sett_dur"] = st.text_input("Settling Obs", "No Growth") [cite: 184]
        data["etx_sett_dur"] = st.text_input("Sett ETX #", "N/A") [cite: 184]
        data["id_sett_dur"] = st.text_input("Sett ID", "N/A") [cite: 184]

    st.subheader("Weekly Bracketing")
    wk1, wk2 = st.columns(2)
    with wk1:
        data["obs_air_wk_of"] = st.text_input("Weekly Air Obs", "No Growth") [cite: 184]
        data["etx_air_wk_of"] = st.text_input("Air ETX #", "N/A") [cite: 184]
        data["id_air_wk_of"] = st.text_input("Air ID", "N/A") [cite: 184]
        data["weekly_initial"] = st.text_input("Weekly Monitor Init") [cite: 184]
    with wk2:
        data["obs_room_wk_of"] = st.text_input("Weekly Room Obs", "No Growth") [cite: 184]
        data["etx_room_wk_of"] = st.text_input("Room ETX #", "N/A") [cite: 184]
        data["id_room_wk_of"] = st.text_input("Room ID", "N/A") [cite: 184]
        data["date_of_weekly"] = st.text_input("Date of Weekly Monitoring") [cite: 184]

    # --- NARRATIVE GENERATOR ---
    st.header("4. Review & Finalize")
    if st.button("🪄 Auto-Generate All Narratives"):
        # Mapping {{ equipment_summary }} based on specific A/B logic
        data["equipment_summary"] = (
            f"Sample processing was conducted within the ISO 5 BSC in the {p_loc} "
            f"(Suite {p_suite}{p_suffix}, BSC E00{data['bsc_id']}) by {data['analyst_name']} "
            f"and the changeover step was conducted within the ISO 5 BSC in the {c_loc} "
            f"(Suite {c_suite}{c_suffix}, BSC E00{data['chgbsc_id']}) by {data['changeover_name']}." [cite: 4, 182]
        )
        
        # Narrative logic for Source 182 Summary
        if data["obs_pers_dur"] == "No Growth" and data["obs_surf_dur"] == "No Growth":
            data["narrative_summary"] = (
                "Upon analyzing the environmental monitoring results, no microbial growth was observed "
                "in personal sampling (left touch and right touch), surface sampling, or settling plates." [cite: 5, 182]
            )
            data["em_details"] = (
                "Weekly active air sampling and weekly surface sampling from both the week of testing "
                "and the week before testing showed no microbial growth." [cite: 50, 182]
            )
        else:
            data["narrative_summary"] = "Microbial growth was observed during environmental monitoring as detailed in Table 1." [cite: 48]
            data["em_details"] = "Growth details are documented in the attached EM identification reports." [cite: 11]
        st.success("All Summary Fields have been populated!")

# --- FINAL GENERATION ---
if st.button("🚀 GENERATE FINAL REPORT"):
    template_name = f"{selection} OOS template.docx"
    if os.path.exists(template_name):
        doc = DocxTemplate(template_name)
        
        # Automatic bracketing dates for Source 184
        try:
            dt_obj = datetime.strptime(data["test_date"], "%d%b%y")
            data["date_before_test"] = (dt_obj - timedelta(days=1)).strftime("%d%b%y") [cite: 48]
            data["date_after_test"] = (dt_obj + timedelta(days=1)).strftime("%d%b%y") [cite: 48]
        except: pass
            
        doc.render(data) [cite: 182]
        out_name = f"{data['oos_id']}_{selection}.docx"
        doc.save(out_name)
        
        with open(out_name, "rb") as f:
            st.download_button("📂 Download Final Document", f, file_name=out_name)
    else:
        st.error(f"Required template '{template_name}' not found on server.")
