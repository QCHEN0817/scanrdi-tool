import streamlit as st
from docxtpl import DocxTemplate
import os
from datetime import datetime, timedelta

# --- PAGE CONFIG ---
st.set_page_config(page_title="Sterility Investigation Tool", layout="wide")

# --- CUSTOM STYLING ---
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #f0f2f6; }
    .main { background-color: #ffffff; }
    </style>
    """, unsafe_allow_html=True)

# Helper function for reverse lookup (Initials -> Full Name)
def get_full_name(initials):
    if not initials: return ""
    lookup = {
        "HS": "Halaina Smith", "DS": "Devanshi Shah", "GS": "Gabbie Surber",
        "MRB": "Muralidhar Bythatagari", "KSM": "Karla Silva", "DT": "Debrework Tassew",
        "PG": "Pagan Gary", "GA": "Gerald Anyangwe", "DH": "Domiasha Harrison",
        "TK": "Tamiru Kotisso", "AO": "Ayomide Odugbesi", "CCD": "Cuong Du",
        "ES": "Alex Saravia", "MJ": "Mukyang Jang", "KA": "Kathleen Aruta",
        "SMO": "Simin Mohammad", "VV": "Varsha Subramanian", "CSG": "Clea S. Garza",
        "GL": "Guanchen Li", "QYC": "Qiyue Chen"
    }
    return lookup.get(initials.upper().strip(), "")

# --- SIDEBAR NAVIGATION ---
with st.sidebar:
    st.title("Sterility Platforms")
    if "active_platform" not in st.session_state:
        st.session_state.active_platform = "ScanRDI"
    if st.button("ScanRDI"): st.session_state.active_platform = "ScanRDI"
    if st.button("Celsis"): st.session_state.active_platform = "Celsis"
    if st.button("USP 71"): st.session_state.active_platform = "USP 71"
    st.divider()
    st.success(f"Active: {st.session_state.active_platform}")

st.title(f"Sterility Investigation & Reporting: {st.session_state.active_platform}")

data = {}

# --- SECTION 1: GENERAL INFO ---
st.header("1. General Test Details")
col1, col2, col3 = st.columns(3)
with col1:
    data["oos_id"] = st.text_input("OOS Number", "OOS-252503")
    data["client_name"] = st.text_input("Client Name", "Northmark Pharmacy")
    data["sample_id"] = st.text_input("Sample ID", "E12955")
with col2:
    data["test_date"] = st.text_input("Test Date (DDMonYY)", datetime.now().strftime("%d%b%y"))
    data["sample_name"] = st.text_input("Sample / Active Name")
    data["lot_number"] = st.text_input("Lot Number")
with col3:
    data["dosage_form"] = st.selectbox("Dosage Form", ["Injectable", "Aqueous Solution", "Liquid", "Solution"])
    data["monthly_cleaning_date"] = st.text_input("Monthly Cleaning Date")

# --- SECTION 2: PERSONNEL & FACILITY ---
if st.session_state.active_platform == "ScanRDI":
    st.header("2. Personnel & Changeover")
    p1, p2, p3, p4 = st.columns(4)
    with p1:
        data["prepper_initial"] = st.text_input("Prepper Initials").upper()
        data["prepper_name"] = st.text_input("Prepper Full Name", value=get_full_name(data["prepper_initial"]))
    with p2:
        data["analyst_initial"] = st.text_input("Processor Initials").upper()
        data["analyst_name"] = st.text_input("Processor Full Name", value=get_full_name(data["analyst_initial"]))
    with p3:
        data["changeover_initial"] = st.text_input("Changeover Initials").upper()
        data["changeover_name"] = st.text_input("Changeover Full Name", value=get_full_name(data["changeover_initial"]))
    with p4:
        data["reader_initial"] = st.text_input("Reader Initials").upper()
        data["reader_name"] = st.text_input("Reader Full Name", value=get_full_name(data["reader_initial"]))

    st.subheader("Equipment & Facility Smart Lookup")
    e1, e2 = st.columns(2)
    bsc_list = ["1310", "1309", "1311", "1312", "1314", "1313", "1316", "1798", "Other"]

    def get_room_logic(bsc_id):
        try:
            num = int(bsc_id)
            suffix = "B" if num % 2 == 0 else "A"
            location = "innermost ISO 7 room" if suffix == "B" else "middle ISO 7 buffer room"
        except: suffix, location = "B", "innermost ISO 7 room"
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
        st.caption(f"Processor: Suite {p_suite}{p_suffix} ({p_loc})")
    with e2:
        chg_bsc = st.selectbox("Select Changeover BSC ID", bsc_list)
        data["chgbsc_id"] = chg_bsc if chg_bsc != "Other" else st.text_input("Manual CHG BSC")
        c_room, c_suite, c_suffix, c_loc = get_room_logic(data["chgbsc_id"])
        st.caption(f"Changeover: Suite {c_suite}{c_suffix} ({c_loc})")

    st.header("3. Findings & EM Data")
    f1, f2 = st.columns(2)
    with f1:
        data["scan_id"] = st.selectbox("ScanRDI ID", ["1230", "2017", "1040", "1877", "2225", "2132"])
        shape_choice = st.selectbox("Org Shape", ["rod", "cocci", "Other"])
        data["organism_morphology"] = shape_choice if shape_choice != "Other" else st.text_input("Enter Manual Org Shape")
        try:
            date_part = datetime.strptime(data["test_date"], "%d%b%y").strftime("%m%d%y")
            default_ref = f"{date_part}-{data['scan_id']}-"
        except: default_ref = ""
        data["test_record"] = st.text_input("Record Ref", value=default_ref)
    with f2:
        data["control_positive"] = st.selectbox("Positive Control", ["A. brasiliensis", "B. subtilis", "C. albicans", "C. sporogenes", "P. aeruginosa", "S. aureus"])
        data["control_lot"] = st.text_input("Control Lot")
        data["control_data"] = st.text_input("Control Exp Date")

    st.subheader("Table 1: EM Observations")
    em1, em2, em3 = st.columns(3)
    with em1:
        data["obs_pers_dur"] = st.text_input("Personnel Obs", "")
        data["etx_pers_dur"] = st.text_input("Pers ETX #", "N/A")
        data["id_pers_dur"] = st.text_input("Pers ID", "N/A")
    with em2:
        data["obs_surf_dur"] = st.text_input("Surface Obs", "")
        data["etx_surf_dur"] = st.text_input("Surf ETX #", "N/A")
        data["id_surf_dur"] = st.text_input("Surf ID", "N/A")
    with em3:
        data["obs_sett_dur"] = st.text_input("Settling Obs", "")
        data["etx_sett_dur"] = st.text_input("Sett ETX #", "N/A")
        data["id_sett_dur"] = st.text_input("Sett ID", "N/A")

    st.subheader("Weekly Bracketing")
    wk1, wk2 = st.columns(2)
    with wk1:
        data["obs_air_wk_of"] = st.text_input("Weekly Air Obs", "")
        data["etx_air_wk_of"] = st.text_input("Air ETX #", "N/A")
        data["id_air_wk_of"] = st.text_input("Air ID", "N/A")
        data["weekly_initial"] = st.text_input("Weekly Monitor Initials")
    with wk2:
        data["obs_room_wk_of"] = st.text_input("Weekly Room Obs", "")
        data["etx_room_wk_of"] = st.text_input("Room ETX #", "N/A")
        data["id_room_wk_of"] = st.text_input("Room ID", "N/A")
        data["date_of_weekly"] = st.text_input("Date of Weekly Monitoring")

    st.header("4. EM Narratives")
    if 'equipment_summary' not in st.session_state: st.session_state.equipment_summary = ""
    if 'narrative_summary' not in st.session_state: st.session_state.narrative_summary = ""
    if 'em_details' not in st.session_state: st.session_state.em_details = ""

    if st.button("🪄 Auto-Generate Narratives"):
        # Equipment Summary
        st.session_state.equipment_summary = f"Sample processing was conducted within the ISO 5 BSC in the {p_loc} (Suite {p_suite}{p_suffix}, BSC E00{data['bsc_id']}) by {data['analyst_name']} and the changeover step was conducted within the ISO 5 BSC in the {c_loc} (Suite {c_suite}{c_suffix}, BSC E00{data['chgbsc_id']}) by {data['changeover_name']}."
        
        # Smart Splicing Logic for Summary
        em_parts = []
        if not data["obs_pers_dur"].strip(): em_parts.append("personal sampling (left touch and right touch)")
        if not data["obs_surf_dur"].strip(): em_parts.append("surface sampling")
        if not data["obs_sett_dur"].strip(): em_parts.append("settling plates")
        
        weekly_parts = []
        if not data["obs_air_wk_of"].strip(): weekly_parts.append("weekly active air sampling")
        if not data["obs_room_wk_of"].strip(): weekly_parts.append("weekly surface sampling")

        growth_sources = []
        if data["obs_pers_dur"].strip(): growth_sources.append(("Personnel Sampling", data["obs_pers_dur"], data["etx_pers_dur"], data["id_pers_dur"]))
        if data["obs_surf_dur"].strip(): growth_sources.append(("Surface Sampling", data["obs_surf_dur"], data["etx_surf_dur"], data["id_surf_dur"]))
        if data["obs_sett_dur"].strip(): growth_sources.append(("Settling Plates", data["obs_sett_dur"], data["etx_sett_dur"], data["id_sett_dur"]))
        if data["obs_air_wk_of"].strip(): growth_sources.append(("Weekly Active Air Sampling", data["obs_air_wk_of"], data["etx_air_wk_of"], data["id_air_wk_of"]))
        # Exact Template logic for Room Sampling
        if data["obs_room_wk_of"].strip(): growth_sources.append(("Surface Sampling of Cleanrooms during weekly room surface sampling", data["obs_room_wk_of"], data["etx_room_wk_of"], data["id_room_wk_of"]))

        if not growth_sources:
            st.session_state.narrative_summary = "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, or settling plates. Weekly active air sampling and weekly surface sampling showed no microbial growth."
            st.session_state.em_details = ""
        else:
            # Build Summary with Clean parts
            summary_text = "Upon analyzing the environmental monitoring results, "
            if em_parts:
                summary_text += "no microbial growth was observed in " + ", ".join(em_parts[:-1]) + (", and " if len(em_parts) > 1 else "") + em_parts[-1] + ". "
            else:
                summary_text += "microbial growth was observed during the testing period. "
            
            if weekly_parts:
                summary_text += "Additionally, " + " and ".join(weekly_parts) + " showed no microbial growth."
            st.session_state.narrative_summary = summary_text

            # Build EM Details using your exact Template
            details_list = []
            for category, obs, etx, org_id in growth_sources:
                para = (f"However, microbial growths were observed in {category} the week of testing. "
                        f"Specifically, {obs}. The plates were submitted for microbial identification under sample IDs "
                        f"{etx}. The organisms identified included {org_id}.")
                details_list.append(para)
            st.session_state.em_details = "\n\n".join(details_list)
        
        st.rerun()

    data["equipment_summary"] = st.session_state.equipment_summary
    data["narrative_summary"] = st.text_area("Narrative Summary (Editable)", value=st.session_state.narrative_summary, height=150)
    data["em_details"] = st.text_area("EM Details (Editable)", value=st.session_state.em_details, height=150)

# --- FINAL GENERATION ---
if st.button("🚀 GENERATE FINAL REPORT"):
    template_name = f"{st.session_state.active_platform} OOS template.docx"
    if os.path.exists(template_name):
        doc = DocxTemplate(template_name)
        try:
            dt_obj = datetime.strptime(data["test_date"], "%d%b%y")
            data["date_before_test"] = (dt_obj - timedelta(days=1)).strftime("%d%b%y")
            data["date_after_test"] = (dt_obj + timedelta(days=1)).strftime("%d%b%y")
            for key in ["obs_pers_dur", "obs_surf_dur", "obs_sett_dur", "obs_air_wk_of", "obs_room_wk_of"]:
                if not data[key].strip(): data[key] = "No Growth"
        except: pass
        doc.render(data)
        out_name = f"{data['oos_id']} {data['client_name']} ({data['sample_id']}) - {st.session_state.active_platform}.docx"
        doc.save(out_name)
        with open(out_name, "rb") as f:
            st.download_button("📂 Download Document", f, file_name=out_name)
    else: st.error(f"Template '{template_name}' not found.")
