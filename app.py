import streamlit as st
from docxtpl import DocxTemplate
import os
import re
import json
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

# --- FILE PERSISTENCE (MEMORY) ---
STATE_FILE = "investigation_state.json"

def load_saved_state():
    """Loads saved inputs from the JSON file if it exists."""
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r") as f:
                saved_data = json.load(f)
            for key, value in saved_data.items():
                if key in st.session_state:
                    st.session_state[key] = value
        except Exception as e:
            st.error(f"Could not load saved state: {e}")

def save_current_state():
    """Saves the current session state to a JSON file."""
    data_to_save = {k: v for k, v in st.session_state.items() if k in field_keys}
    try:
        with open(STATE_FILE, "w") as f:
            json.dump(data_to_save, f)
        st.toast("✅ Progress saved!", icon="💾")
    except Exception as e:
        st.error(f"Could not save state: {e}")

# --- HELPER FUNCTIONS ---
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

def get_room_logic(bsc_id):
    # Returns: room_id, suite, suffix (A/B), location_desc
    try:
        num = int(bsc_id)
        if num % 2 == 0:
            suffix = "B"
            location = "innermost ISO 7 room"
        else:
            suffix = "A"
            location = "middle ISO 7 buffer room"
    except: 
        suffix, location = "B", "innermost ISO 7 room"
    
    if bsc_id in ["1310", "1309"]: suite = "117"
    elif bsc_id in ["1311", "1312"]: suite = "116"
    elif bsc_id in ["1314", "1313"]: suite = "115"
    elif bsc_id in ["1316", "1798"]: suite = "114"
    else: suite = "Unknown"
    
    room_map = {"117": "1739", "116": "1738", "115": "1737", "114": "1736"}
    room_id = room_map.get(suite, "Unknown")
    
    return room_id, suite, suffix, location

# --- SESSION STATE INITIALIZATION ---
def init_state(key, default_value=""):
    if key not in st.session_state:
        st.session_state[key] = default_value

field_keys = [
    "oos_id", "client_name", "sample_id", "test_date", "sample_name", "lot_number", 
    "dosage_form", "monthly_cleaning_date", 
    "prepper_initial", "prepper_name", 
    "analyst_initial", "analyst_name",
    "changeover_initial", "changeover_name",
    "reader_initial", "reader_name",
    "bsc_id", "chgbsc_id", "scan_id", 
    "shift_number", "active_platform",
    "org_choice", "manual_org", "test_record", "control_pos", "control_lot", 
    "control_exp", "obs_pers", "etx_pers", "id_pers", "obs_surf", "etx_surf", 
    "id_surf", "obs_sett", "etx_sett", "id_sett", "obs_air", "etx_air_weekly", 
    "id_air_weekly", "obs_room", "etx_room_weekly", "id_room_weekly", "weekly_init", 
    "date_weekly", "equipment_summary", "narrative_summary", "em_details", 
    "sample_history_paragraph", "incidence_count", "oos_refs"
]

for k in field_keys:
    if k == "incidence_count": init_state(k, 0)
    elif k == "shift_number": init_state(k, "1")
    elif "etx" in k or "id" in k: init_state(k, "N/A")
    elif k == "active_platform": init_state(k, "ScanRDI")
    else: init_state(k, "")

# LOAD DATA ON APP START
if "data_loaded" not in st.session_state:
    load_saved_state()
    st.session_state.data_loaded = True

# --- EMAIL PARSER LOGIC ---
def parse_email_text(text):
    oos_match = re.search(r"OOS-(\d+)", text)
    if oos_match: st.session_state.oos_id = oos_match.group(1)

    client_match = re.search(r"([A-Za-z\s]+\(E\d+\))", text)
    if client_match: st.session_state.client_name = client_match.group(1).strip()

    etx_id_match = re.search(r"(ETX-\d{6}-\d{4})", text)
    if etx_id_match: st.session_state.sample_id = etx_id_match.group(1).strip()

    sample_match = re.search(r"Sample\s*Name:\s*(.*)", text, re.IGNORECASE)
    if sample_match: st.session_state.sample_name = sample_match.group(1).strip()

    lot_match = re.search(r"Lot:\s*(\w+)", text, re.IGNORECASE)
    if lot_match: st.session_state.lot_number = lot_match.group(1).strip()

    date_match = re.search(r"testing\s*on\s*(\d{2}\s*\w{3}\s*\d{4})", text, re.IGNORECASE)
    if date_match:
        try:
            d_obj = datetime.strptime(date_match.group(1).strip(), "%d %b %Y")
            st.session_state.test_date = d_obj.strftime("%d%b%y")
        except: pass

    analyst_match = re.search(r"\(\s*([A-Z]{2,3})\s*\d+[a-z]{2}\s*Sample\)", text)
    if analyst_match: 
        st.session_state.analyst_initial = analyst_match.group(1).strip()
        st.session_state.analyst_name = get_full_name(st.session_state.analyst_initial)
    
    save_current_state()

# --- SIDEBAR ---
with st.sidebar:
    st.title("Sterility Platforms")
    if st.button("ScanRDI"): st.session_state.active_platform = "ScanRDI"
    if st.button("Celsis"): st.session_state.active_platform = "Celsis"
    if st.button("USP 71"): st.session_state.active_platform = "USP 71"
    st.divider()
    
    if st.button("💾 Save Current Inputs"):
        save_current_state()
        
    st.success(f"Active: {st.session_state.active_platform}")

st.title(f"Sterility Investigation & Reporting: {st.session_state.active_platform}")

# --- SMART PARSER BOX ---
st.header("📧 Smart Email Import")
email_input = st.text_area("Paste the OOS Notification email here to auto-fill fields:", height=150)
if st.button("🪄 Parse Email & Auto-Fill"):
    if email_input:
        parse_email_text(email_input)
        st.success("Fields updated!")
        st.rerun()

# --- SECTION 1: GENERAL INFO ---
st.header("1. General Test Details")
col1, col2, col3 = st.columns(3)
with col1:
    st.session_state.oos_id = st.text_input("OOS Number (Numbers only)", st.session_state.oos_id)
    st.session_state.client_name = st.text_input("Client Name", st.session_state.client_name)
    st.session_state.sample_id = st.text_input("Sample ID (ETX Format)", st.session_state.sample_id)
with col2:
    st.session_state.test_date = st.text_input("Test Date (e.g., 07Jan26)", st.session_state.test_date)
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
        pre_full = get_full_name(st.session_state.prepper_initial)
        st.session_state.prepper_name = st.text_input("Prepper Full Name", value=st.session_state.prepper_name if st.session_state.prepper_name else pre_full)
    with p2:
        st.session_state.analyst_initial = st.text_input("Processor Initials", st.session_state.analyst_initial).upper()
        ana_full = get_full_name(st.session_state.analyst_initial)
        st.session_state.analyst_name = st.text_input("Processor Full Name", value=st.session_state.analyst_name if st.session_state.analyst_name else ana_full)
    with p3:
        st.session_state.changeover_initial = st.text_input("Changeover Initials", st.session_state.changeover_initial).upper()
        chg_full = get_full_name(st.session_state.changeover_initial)
        st.session_state.changeover_name = st.text_input("Changeover Full Name", value=st.session_state.changeover_name if st.session_state.changeover_name else chg_full)
    with p4:
        st.session_state.reader_initial = st.text_input("Reader Initials", st.session_state.reader_initial).upper()
        rea_full = get_full_name(st.session_state.reader_initial)
        st.session_state.reader_name = st.text_input("Reader Full Name", value=st.session_state.reader_name if st.session_state.reader_name else rea_full)

    e1, e2 = st.columns(2)
    bsc_list = ["1310", "1309", "1311", "1312", "1314", "1313", "1316", "1798", "Other"]
    with e1:
        st.session_state.bsc_id = st.selectbox("Select Processing BSC ID", bsc_list, index=bsc_list.index(st.session_state.bsc_id) if st.session_state.bsc_id in bsc_list else 0)
        p_room, p_suite, p_suffix, p_loc = get_room_logic(st.session_state.bsc_id)
        st.caption(f"Processor: Suite {p_suite}{p_suffix} ({p_loc}) [Room ID: {p_room}]")
    with e2:
        st.session_state.chgbsc_id = st.selectbox("Select Changeover BSC ID", bsc_list, index=bsc_list.index(st.session_state.chgbsc_id) if st.session_state.chgbsc_id in bsc_list else 0)
        c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
        st.caption(f"Changeover: Suite {c_suite}{c_suffix} ({c_loc}) [Room ID: {c_room}]")

    st.header("3. Findings & EM Data")
    f1, f2 = st.columns(2)
    with f1:
        scan_ids = ["1230", "2017", "1040", "1877", "2225", "2132"]
        st.session_state.scan_id = st.selectbox("ScanRDI ID", scan_ids, index=scan_ids.index(st.session_state.scan_id) if st.session_state.scan_id in scan_ids else 0)
        st.session_state.shift_number = st.text_input("Shift Number", st.session_state.shift_number)
        shape_opts = ["rod", "cocci", "Other"]
        st.session_state.org_choice = st.selectbox("Org Shape", shape_opts, index=shape_opts.index(st.session_state.org_choice) if st.session_state.org_choice in shape_opts else 0)
        if st.session_state.org_choice == "Other":
            st.session_state.manual_org = st.text_input("Enter Manual Org Shape", st.session_state.manual_org)
        try:
            date_part = datetime.strptime(st.session_state.test_date, "%d%b%y").strftime("%m%d%y")
            st.session_state.test_record = f"{date_part}-{st.session_state.scan_id}-{st.session_state.shift_number}"
        except: pass
        st.text_input("Record Ref", st.session_state.test_record, disabled=True)

    with f2:
        ctrl_opts = ["A. brasiliensis", "B. subtilis", "C. albicans", "C. sporogenes", "P. aeruginosa", "S. aureus"]
        st.session_state.control_pos = st.selectbox("Positive Control", ctrl_opts, index=ctrl_opts.index(st.session_state.control_pos) if st.session_state.control_pos in ctrl_opts else 0)
        st.session_state.control_lot = st.text_input("Control Lot", st.session_state.control_lot)
        st.session_state.control_exp = st.text_input("Control Exp Date", st.session_state.control_exp)

    st.header("4. EM Observations")
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
        st.session_state.obs_room = st.text_input("Weekly Surf Obs", st.session_state.obs_room)
        st.session_state.etx_room_weekly = st.text_input("Weekly Surf ETX #", st.session_state.etx_room_weekly)
        st.session_state.id_room_weekly = st.text_input("Weekly Surf ID", st.session_state.id_room_weekly)
        st.session_state.date_weekly = st.text_input("Date of Weekly Monitoring", st.session_state.date_weekly)

    st.header("5. Automated Summaries")
    h1, h2 = st.columns(2)
    with h1:
        try: curr_val = int(st.session_state.incidence_count)
        except: curr_val = 0
        st.session_state.incidence_count = st.number_input("Number of Prior Failures", value=curr_val, min_value=0, step=1)
    with h2:
        st.session_state.oos_refs = st.text_input("Related OOS IDs (e.g. OOS-25001, OOS-25002)", st.session_state.oos_refs)

    if st.button("🪄 Auto-Generate All"):
        # --- 1. SMART EQUIPMENT SUMMARY LOGIC ---
        t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
        c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
        
        # SCENARIO 1: SAME BSC (Implies Same Suite)
        if st.session_state.bsc_id == st.session_state.chgbsc_id:
            part1 = (
                f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: "
                f"the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), "
                f"and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite "
                f"to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
            )
            part2 = (
                f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both testing and changeover steps. "
                f"It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). "
                f"Additionally, BSC E00{st.session_state.bsc_id} was certified and approved by both the Engineering and Quality Assurance teams. "
                f"Sample processing and changeover were conducted in the ISO 5 BSC E00{st.session_state.bsc_id} in the {t_loc}, "
                f"(Suite {t_suite}{t_suffix}) by {st.session_state.analyst_name} on {st.session_state.test_date}."
            )
            st.session_state.equipment_summary = f"{part1}\n\n{part2}"

        # SCENARIO 2: DIFFERENT BSCs, BUT SAME SUITE
        elif t_suite == c_suite:
            part1 = (
                f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: "
                f"the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), "
                f"and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite "
                f"to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
            )
            part2 = (
                f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), and "
                f"ISO 5 BSC E00{st.session_state.chgbsc_id}, located in the {c_loc}, (Suite {c_suite}{c_suffix}), were thoroughly cleaned and disinfected prior to their respective procedures "
                f"in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Furthermore, the BSCs used throughout testing, "
                f"E00{st.session_state.bsc_id} for sample processing and E00{st.session_state.chgbsc_id} for the changeover step, were certified and approved by both the Engineering and Quality Assurance teams. "
                f"Sample processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{st.session_state.bsc_id}) by {st.session_state.analyst_name} "
                f"and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{st.session_state.chgbsc_id}) "
                f"by {st.session_state.changeover_name} on {st.session_state.test_date}."
            )
            st.session_state.equipment_summary = f"{part1}\n\n{part2}"

        # SCENARIO 3: DIFFERENT BSCs AND DIFFERENT SUITES
        else:
            part1 = (
                f"The cleanroom used for testing (E00{t_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), "
                f"which opens into the middle ISO 7 buffer room ({t_suite}A), and then into the outermost ISO 8 anteroom ({t_suite}). "
                f"A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
            )
            part2 = (
                f"The cleanroom used for changeover (E00{c_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({c_suite}B), "
                f"which opens into the middle ISO 7 buffer room ({c_suite}A), and then into the outermost ISO 8 anteroom ({c_suite}). "
                f"A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {c_suite}B through {c_suite}A and into {c_suite}."
            )
            part3 = (
                f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), and "
                f"ISO 5 BSC E00{st.session_state.chgbsc_id}, located in the {c_loc}, (Suite {c_suite}{c_suffix}), were thoroughly cleaned and disinfected prior to their respective procedures "
                f"in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Furthermore, the BSCs used throughout testing, "
                f"E00{st.session_state.bsc_id} for sample processing and E00{st.session_state.chgbsc_id} for the changeover step, were certified and approved by both the Engineering and Quality Assurance teams. "
                f"Sample processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{st.session_state.bsc_id}) by {st.session_state.analyst_name} "
                f"and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{st.session_state.chgbsc_id}) "
                f"by {st.session_state.changeover_name} on {st.session_state.test_date}."
            )
            st.session_state.equipment_summary = f"{part1}\n\n{part2}\n\n{part3}"

        # 2. History Logic
        if st.session_state.incidence_count == 0:
            hist_phrase = "no prior failures"
        elif st.session_state.incidence_count == 1:
            hist_phrase = f"1 incident ({st.session_state.oos_refs})"
        else:
            hist_phrase = f"{st.session_state.incidence_count} incidents ({st.session_state.oos_refs})"
        st.session_state.sample_history_paragraph = f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte “{st.session_state.sample_name}” has had {hist_phrase} using the Scan RDI method during this period."

        # 3. Narrative/EM Logic
        growth_sources = []
        if st.session_state.obs_pers.strip(): growth_sources.append(("Personnel Sampling", st.session_state.obs_pers, st.session_state.etx_pers, st.session_state.id_pers))
        if st.session_state.obs_surf.strip(): growth_sources.append(("Surface Sampling", st.session_state.obs_surf, st.session_state.etx_surf, st.session_state.id_surf))
        if st.session_state.obs_sett.strip(): growth_sources.append(("Settling Plates", st.session_state.obs_sett, st.session_state.etx_sett, st.session_state.id_sett))
        if st.session_state.obs_air.strip(): growth_sources.append(("Weekly Active Air Sampling", st.session_state.obs_air, st.session_state.etx_air_weekly, st.session_state.id_air_weekly))
        if st.session_state.obs_room.strip(): growth_sources.append(("surface sampling of cleanroom during weekly room surface sampling", st.session_state.obs_room, st.session_state.etx_room_weekly, st.session_state.id_room_weekly))

        if not growth_sources:
            st.session_state.narrative_summary = "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, or settling plates. Weekly active air sampling and weekly surface sampling showed no microbial growth."
            st.session_state.em_details = ""
        else:
            em_clean = [p for p, obs in [("personal sampling (left touch and right touch)", st.session_state.obs_pers), ("surface sampling", st.session_state.obs_surf), ("settling plates", st.session_state.obs_sett)] if not obs.strip()]
            wk_clean = [p for p, obs in [("weekly active air sampling", st.session_state.obs_air), ("weekly surface sampling", st.session_state.obs_room)] if not obs.strip()]
            summary_text = "Upon analyzing the environmental monitoring results, "
            if em_clean:
                summary_text += "no microbial growth was observed in " + ", ".join(em_clean[:-1]) + (", and " if len(em_clean) > 1 else "") + em_clean[-1] + ". "
            else: summary_text += "microbial growth was observed during the testing period. "
            if wk_clean: summary_text += "Additionally, " + " and ".join(wk_clean) + " showed no microbial growth."
            st.session_state.narrative_summary = summary_text

            details_list = []
            for category, obs, etx, org_id in growth_sources:
                is_sing = ("1" in obs and "CFU" in obs.upper() and "11" not in obs and "21" not in obs)
                growth_term, plate_term = ("growth was", "plate was") if is_sing else ("growths were", "plates were")
                id_label = "sample IDs" if ("," in etx or "AND" in etx.upper()) else "sample ID"
                org_verb = "organisms identified included" if ("," in org_id or "AND" in org_id.upper()) else "organism identified was"
                details_list.append(f"However, microbial {growth_term} observed in {category} the week of testing. Specifically, {obs}. The {plate_term} submitted for microbial identification under {id_label} {etx}. The {org_verb} {org_id}.")
            st.session_state.em_details = "\n\n".join(details_list)
        
        save_current_state()
        st.rerun()

    st.session_state.sample_history_paragraph = st.text_area("Sample History (Full Paragraph Editable)", value=st.session_state.sample_history_paragraph, height=100)
    st.session_state.narrative_summary = st.text_area("Narrative Summary (Editable)", value=st.session_state.narrative_summary, height=120)
    st.session_state.em_details = st.text_area("EM Growth Details (Editable)", value=st.session_state.em_details, height=150)

# --- FINAL GENERATION ---
if st.button("🚀 GENERATE FINAL REPORT"):
    template_name = f"{st.session_state.active_platform} OOS template.docx"
    if os.path.exists(template_name):
        doc = DocxTemplate(template_name)
        final_data = {k: v for k, v in st.session_state.items()}
        final_data["oos_full"] = f"OOS-{st.session_state.oos_id}"
        
        if st.session_state.active_platform == "ScanRDI":
            t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
            final_data["cr_suit"] = t_suite
            final_data["cr_id"] = t_room
            final_data["suit"] = t_suffix
            final_data["bsc_location"] = t_loc
            
            c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
            final_data["changeover_id"] = c_room
            final_data["changeover_suit"] = c_suite
            final_data["changeoversuit"] = c_suffix
            final_data["changeover_location"] = c_loc
            final_data["changeoverbsc_id"] = st.session_state.chgbsc_id
            
            final_data["changeover_name"] = st.session_state.changeover_name
            final_data["analyst_name"] = st.session_state.analyst_name
            
            final_data["control_positive"] = st.session_state.control_pos
            final_data["control_data"] = st.session_state.control_exp
            
            if st.session_state.org_choice == "Other":
                final_data["organism_morphology"] = st.session_state.manual_org
            else:
                final_data["organism_morphology"] = st.session_state.org_choice

            final_data["obs_pers_dur"] = st.session_state.obs_pers
            final_data["etx_pers_dur"] = st.session_state.etx_pers
            final_data["id_pers_dur"] = st.session_state.id_pers
            
            final_data["obs_surf_dur"] = st.session_state.obs_surf
            final_data["etx_surf_dur"] = st.session_state.etx_surf
            final_data["id_surf_dur"] = st.session_state.id_surf
            
            final_data["obs_sett_dur"] = st.session_state.obs_sett
            final_data["etx_sett_dur"] = st.session_state.etx_sett
            final_data["id_sett_dur"] = st.session_state.id_sett
            
            final_data["obs_air_wk_of"] = st.session_state.obs_air
            final_data["etx_air_wk_of"] = st.session_state.etx_air_weekly
            final_data["id_air_wk_of"] = st.session_state.id_air_weekly
            
            final_data["obs_room_wk_of"] = st.session_state.obs_room
            final_data["etx_room_wk_of"] = st.session_state.etx_room_weekly
            final_data["id_room_wk_of"] = st.session_state.id_room_weekly
            
            final_data["weekly_initial"] = st.session_state.weekly_init
            final_data["date_of_weekly"] = st.session_state.date_weekly

        for key in ["obs_pers", "obs_surf", "obs_sett", "obs_air", "obs_room"]:
            if not final_data[key].strip(): final_data[key] = "No Growth"
        try:
            dt_obj = datetime.strptime(st.session_state.test_date, "%d%b%y")
            final_data["date_before_test"] = (dt_obj - timedelta(days=1)).strftime("%d%b%y")
            final_data["date_after_test"] = (dt_obj + timedelta(days=1)).strftime("%d%b%y")
        except: pass
        doc.render(final_data)
        out_name = f"OOS-{st.session_state.oos_id} {st.session_state.client_name} ({st.session_state.sample_id}) - {st.session_state.active_platform}.docx"
        doc.save(out_name)
        
        save_current_state()
        
        with open(out_name, "rb") as f:
            st.download_button("📂 Download Document", f, file_name=out_name)
