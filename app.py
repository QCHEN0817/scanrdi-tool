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

# --- SESSION STATE INITIALIZATION ---
def init_state(key, default_value=""):
    if key not in st.session_state:
        st.session_state[key] = default_value

# Fields to persist
field_keys = [
    "oos_id", "client_name", "sample_id", "test_date", "sample_name", "lot_number", 
    "dosage_form", "monthly_cleaning_date", "prepper_initial", "analyst_initial", 
    "changeover_initial", "reader_initial", "bsc_id", "chgbsc_id", "scan_id", 
    "org_choice", "manual_org", "test_record", "control_pos", "control_lot", 
    "control_exp", "obs_pers", "etx_pers", "id_pers", "obs_surf", "etx_surf", 
    "id_surf", "obs_sett", "etx_sett", "id_sett", "obs_air", "etx_air_weekly", 
    "id_air_weekly", "obs_room", "etx_room_weekly", "id_room_weekly", "weekly_init", 
    "date_weekly", "equipment_summary", "narrative_summary", "em_details", 
    "sample_history_para"
]

for k in field_keys:
    if k == "oos_id": init_state(k, "OOS-252503")
    elif k == "client_name": init_state(k, "Northmark Pharmacy")
    elif k == "sample_id": init_state(k, "E12955")
    elif k == "test_date": init_state(k, datetime.now().strftime("%d%b%y"))
    elif k == "dosage_form": init_state(k, "Injectable")
    elif "etx" in k or "id" in k: init_state(k, "N/A")
    else: init_state(k, "")

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

# --- SECTION 1: GENERAL INFO ---
st.header("1. General Test Details")
col1, col2, col3 = st.columns(3)
with col1:
    st.session_state.oos_id = st.text_input("OOS Number", st.session_state.oos_id)
    st.session_state.client_name = st.text_input("Client Name", st.session_state.client_name)
    st.session_state.sample_id = st.text_input("Sample ID", st.session_state.sample_id)
with col2:
    st.session_state.test_date = st.text_input("Test Date (DDMonYY)", st.session_state.test_date)
    st.session_state.sample_name = st.text_input("Sample / Active Name", st.session_state.sample_name)
    st.session_state.lot_number = st.text_input("Lot Number", st.session_state.lot_number)
with col3:
    dosage_options = ["Injectable", "Aqueous Solution", "Liquid", "Solution"]
    idx = dosage_options.index(st.session_state.dosage_form) if st.session_state.dosage_form in dosage_options else 0
    st.session_state.dosage_form = st.selectbox("Dosage Form", dosage_options, index=idx)
    st.session_state.monthly_cleaning_date = st.text_input("Monthly Cleaning Date", st.session_state.monthly_cleaning_date)

# --- SECTION 2: PERSONNEL & FACILITY ---
if st.session_state.active_platform == "ScanRDI":
    st.header("2. Personnel & Changeover")
    p1, p2, p3, p4 = st.columns(4)
    with p1:
        st.session_state.prepper_initial = st.text_input("Prepper Initials", st.session_state.prepper_initial).upper()
        prepper_full = get_full_name(st.session_state.prepper_initial)
        st.text_input("Prepper Full Name", value=prepper_full, disabled=True)
    with p2:
        st.session_state.analyst_initial = st.text_input("Processor Initials", st.session_state.analyst_initial).upper()
        analyst_full = get_full_name(st.session_state.analyst_initial)
        st.text_input("Processor Full Name", value=analyst_full, disabled=True)
    with p3:
        st.session_state.changeover_initial = st.text_input("Changeover Initials", st.session_state.changeover_initial).upper()
        changeover_full = get_full_name(st.session_state.changeover_initial)
        st.text_input("Changeover Full Name", value=changeover_full, disabled=True)
    with p4:
        st.session_state.reader_initial = st.text_input("Reader Initials", st.session_state.reader_initial).upper()
        reader_full = get_full_name(st.session_state.reader_initial)
        st.text_input("Reader Full Name", value=reader_full, disabled=True)

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
        st.session_state.bsc_id = st.selectbox("Select Processing BSC ID", bsc_list, index=bsc_list.index(st.session_state.bsc_id) if st.session_state.bsc_id in bsc_list else 0)
        p_room, p_suite, p_suffix, p_loc = get_room_logic(st.session_state.bsc_id)
        st.caption(f"Processor: Suite {p_suite}{p_suffix} ({p_loc})")
    with e2:
        st.session_state.chgbsc_id = st.selectbox("Select Changeover BSC ID", bsc_list, index=bsc_list.index(st.session_state.chgbsc_id) if st.session_state.chgbsc_id in bsc_list else 0)
        c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
        st.caption(f"Changeover: Suite {c_suite}{c_suffix} ({c_loc})")

    st.header("3. Findings & EM Data")
    f1, f2 = st.columns(2)
    with f1:
        scan_ids = ["1230", "2017", "1040", "1877", "2225", "2132"]
        st.session_state.scan_id = st.selectbox("ScanRDI ID", scan_ids, index=scan_ids.index(st.session_state.scan_id) if st.session_state.scan_id in scan_ids else 0)
        shape_opts = ["rod", "cocci", "Other"]
        st.session_state.org_choice = st.selectbox("Org Shape", shape_opts, index=shape_opts.index(st.session_state.org_choice) if st.session_state.org_choice in shape_opts else 0)
        if st.session_state.org_choice == "Other":
            st.session_state.manual_org = st.text_input("Enter Manual Org Shape", st.session_state.manual_org)
        
        try:
            date_part = datetime.strptime(st.session_state.test_date, "%d%b%y").strftime("%m%d%y")
            st.session_state.test_record = f"{date_part}-{st.session_state.scan_id}-"
        except: pass
        st.text_input("Record Ref", st.session_state.test_record, disabled=True)

    with f2:
        ctrl_opts = ["A. brasiliensis", "B. subtilis", "C. albicans", "C. sporogenes", "P. aeruginosa", "S. aureus"]
        st.session_state.control_pos = st.selectbox("Positive Control", ctrl_opts, index=ctrl_opts.index(st.session_state.control_pos) if st.session_state.control_pos in ctrl_opts else 0)
        st.session_state.control_lot = st.text_input("Control Lot", st.session_state.control_lot)
        st.session_state.control_exp = st.text_input("Control Exp Date", st.session_state.control_exp)

    st.subheader("Table 1: EM Observations")
    em1, em2, em3 = st.columns(3)
    with em1:
        st.session_state.obs_pers = st.text_input("Personnel Obs", st.session_state.obs_pers)
        st.session_state.etx_pers = st.text_input("Pers ETX #", st.session_state.etx_pers)
        st.session_state.id_pers = st.text_input("Pers ID", st.session_state.id_pers)
    with em2:
        st.session_state.obs_surf = st.text_input("Surface Obs", st.session_state.obs_surf)
        st.session_state.etx_surf = st.text_input("Surf ETX #", st.session_state.etx_surf)
        st.session_state.id_surf = st.text_input("Surf ID", st.session_state.id_surf)
    with em3:
        st.session_state.obs_sett = st.text_input("Settling Obs", st.session_state.obs_sett)
        st.session_state.etx_sett = st.text_input("Sett ETX #", st.session_state.etx_sett)
        st.session_state.id_sett = st.text_input("Sett ID", st.session_state.id_sett)

    st.subheader("Weekly Bracketing")
    wk1, wk2 = st.columns(2)
    with wk1:
        st.session_state.obs_air = st.text_input("Weekly Air Obs", st.session_state.obs_air)
        st.session_state.etx_air_weekly = st.text_input("Weekly Air ETX #", st.session_state.etx_air_weekly)
        st.session_state.id_air_weekly = st.text_input("Weekly Air ID", st.session_state.id_air_weekly)
        st.session_state.weekly_init = st.text_input("Weekly Monitor Initials", st.session_state.weekly_init)
    with wk2:
        st.session_state.obs_room = st.text_input("Weekly Room Obs", st.session_state.obs_room)
        st.session_state.etx_room_weekly = st.text_input("Weekly Surf ETX #", st.session_state.etx_room_weekly)
        st.session_state.id_room_weekly = st.text_input("Weekly Surf ID", st.session_state.id_room_weekly)
        st.session_state.date_weekly = st.text_input("Date of Weekly Monitoring", st.session_state.date_weekly)

    st.header("4. Automated Summaries")
    if st.button("🪄 Auto-Generate All"):
        # Equipment Summary
        st.session_state.equipment_summary = f"Sample processing was conducted within the ISO 5 BSC in the {p_loc} (Suite {p_suite}{p_suffix}, BSC E00{st.session_state.bsc_id}) by {analyst_full} and the changeover step was conducted within the ISO 5 BSC in the {c_loc} (Suite {c_suite}{c_suffix}, BSC E00{st.session_state.chgbsc_id}) by {changeover_full}."
        
        # History Paragraph (Updated Template for failures)
        st.session_state.sample_history_para = f"Analyzing a 6-month sample history for {st.session_state.client_name} ({st.session_state.sample_id}), this specific analyte “{st.session_state.sample_name}” has had no prior failures using the Scan RDI method during this period."

        # Narrative Logic
        em_parts = []
        if not st.session_state.obs_pers.strip(): em_parts.append("personal sampling (left touch and right touch)")
        if not st.session_state.obs_surf.strip(): em_parts.append("surface sampling")
        if not st.session_state.obs_sett.strip(): em_parts.append("settling plates")
        
        weekly_parts = []
        if not st.session_state.obs_air.strip(): weekly_parts.append("weekly active air sampling")
        if not st.session_state.obs_room.strip(): weekly_parts.append("weekly surface sampling")

        growth_sources = []
        if st.session_state.obs_pers.strip(): growth_sources.append(("Personnel Sampling", st.session_state.obs_pers, st.session_state.etx_pers, st.session_state.id_pers))
        if st.session_state.obs_surf.strip(): growth_sources.append(("Surface Sampling", st.session_state.obs_surf, st.session_state.etx_surf, st.session_state.id_surf))
        if st.session_state.obs_sett.strip(): growth_sources.append(("Settling Plates", st.session_state.obs_sett, st.session_state.etx_sett, st.session_state.id_sett))
        if st.session_state.obs_air.strip(): growth_sources.append(("Weekly Active Air Sampling", st.session_state.obs_air, st.session_state.etx_air_weekly, st.session_state.id_air_weekly))
        if st.session_state.obs_room.strip(): growth_sources.append(("Surface Sampling of Cleanrooms during weekly room surface sampling", st.session_state.obs_room, st.session_state.etx_room_weekly, st.session_state.id_room_weekly))

        if not growth_sources:
            st.session_state.narrative_summary = "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, or settling plates. Weekly active air sampling and weekly surface sampling showed no microbial growth."
            st.session_state.em_details = ""
        else:
            summary_text = "Upon analyzing the environmental monitoring results, "
            if em_parts:
                summary_text += "no microbial growth was observed in " + ", ".join(em_parts[:-1]) + (", and " if len(em_parts) > 1 else "") + em_parts[-1] + ". "
            else:
                summary_text += "microbial growth was observed during the testing period. "
            if weekly_parts:
                summary_text += "Additionally, " + " and ".join(weekly_parts) + " showed no microbial growth."
            st.session_state.narrative_summary = summary_text

            details_list = []
            for category, obs, etx, org_id in growth_sources:
                para = (f"However, microbial growths were observed in {category} the week of testing. "
                        f"Specifically, {obs}. The plates were submitted for microbial identification under sample IDs "
                        f"{etx}. The organisms identified included {org_id}.")
                details_list.append(para)
            st.session_state.em_details = "\n\n".join(details_list)
        st.rerun()

    st.session_state.sample_history_para = st.text_area("Sample History (Editable)", value=st.session_state.sample_history_para, height=100)
    st.session_state.narrative_summary = st.text_area("Narrative Summary (Editable)", value=st.session_state.narrative_summary, height=120)
    st.session_state.em_details = st.text_area("EM Growth Details (Editable)", value=st.session_state.em_details, height=120)

# --- FINAL GENERATION ---
if st.button("🚀 GENERATE FINAL REPORT"):
    template_name = f"{st.session_state.active_platform} OOS template.docx"
    if os.path.exists(template_name):
        doc = DocxTemplate(template_name)
        final_data = {k: v for k, v in st.session_state.items()}
        for key in ["obs_pers", "obs_surf", "obs_sett", "obs_air", "obs_room"]:
            if not final_data[key].strip(): final_data[key] = "No Growth"
        try:
            dt_obj = datetime.strptime(st.session_state.test_date, "%d%b%y")
            final_data["date_before_test"] = (dt_obj - timedelta(days=1)).strftime("%d%b%y")
            final_data["date_after_test"] = (dt_obj + timedelta(days=1)).strftime("%d%b%y")
        except: pass
        doc.render(final_data)
        out_name = f"{st.session_state.oos_id} {st.session_state.client_name} ({st.session_state.sample_id}) - {st.session_state.active_platform}.docx"
        doc.save(out_name)
        with open(out_name, "rb") as f:
            st.download_button("📂 Download Document", f, file_name=out_name)
