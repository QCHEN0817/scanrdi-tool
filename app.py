import streamlit as st
from docxtpl import DocxTemplate
import os
import re
import json
import io
from datetime import datetime, timedelta

# --- PAGE CONFIG ---
st.set_page_config(page_title="Sterility Investigation Tool", layout="wide")

# --- CUSTOM STYLING ---
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #f0f2f6; }
    .main { background-color: #ffffff; }
    .stTextArea textarea { background-color: #ffffff; color: #31333F; border: 1px solid #d6d6d6; }
    </style>
    """, unsafe_allow_html=True)

# --- FILE PERSISTENCE (MEMORY) ---
STATE_FILE = "investigation_state.json"

# All keys that we want to save/load and use in the report
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
    "id_air_weekly", "obs_room", "etx_room_weekly", "id_room_wk_of", "weekly_init", 
    "date_weekly", "equipment_summary", "narrative_summary", "em_details", 
    "sample_history_paragraph", "incidence_count", "oos_refs",
    "other_positives", "cross_contamination_summary",
    "total_pos_count_num", "current_pos_order",
    "diff_changeover_bsc", "has_prior_failures",
    "em_growth_observed", "diff_changeover_analyst"
]
for i in range(10):
    field_keys.append(f"other_id_{i}")
    field_keys.append(f"other_order_{i}")

def load_saved_state():
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
    data_to_save = {k: v for k, v in st.session_state.items() if k in field_keys}
    try:
        with open(STATE_FILE, "w") as f:
            json.dump(data_to_save, f)
    except Exception as e:
        st.error(f"Could not save state: {e}")

# --- HELPER FUNCTIONS ---
def clean_filename(text):
    if not text: return ""
    clean = re.sub(r'[\\/*?:"<>|]', '_', str(text))
    return clean.strip()

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

def num_to_words(n):
    mapping = {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six", 7: "seven", 8: "eight", 9: "nine", 10: "ten"}
    return mapping.get(n, str(n))

def ordinal(n):
    try: n = int(n)
    except: return str(n)
    if 11 <= (n % 100) <= 13: suffix = 'th'
    else: suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"

def get_room_logic(bsc_id):
    try:
        num = int(bsc_id)
        if num % 2 == 0: suffix, location = "B", "innermost ISO 7 room"
        else: suffix, location = "A", "middle ISO 7 buffer room"
    except: suffix, location = "B", "innermost ISO 7 room"
    
    if bsc_id in ["1310", "1309"]: suite = "117"
    elif bsc_id in ["1311", "1312"]: suite = "116"
    elif bsc_id in ["1314", "1313"]: suite = "115"
    elif bsc_id in ["1316", "1798"]: suite = "114"
    else: suite = "Unknown"
    
    room_map = {"117": "1739", "116": "1738", "115": "1737", "114": "1736"}
    room_id = room_map.get(suite, "Unknown")
    return room_id, suite, suffix, location

# --- INIT STATE ---
def init_state(key, default_value=""):
    if key not in st.session_state: st.session_state[key] = default_value

for k in field_keys:
    if k == "incidence_count": init_state(k, 0)
    elif k == "shift_number": init_state(k, "1")
    elif "etx" in k or "id" in k: init_state(k, "N/A")
    elif k == "active_platform": init_state(k, "ScanRDI")
    elif k == "other_positives": init_state(k, "No")
    elif k == "total_pos_count_num": init_state(k, 2)
    elif k == "current_pos_order": init_state(k, 1) 
    elif k == "diff_changeover_bsc": init_state(k, "No")
    elif k == "has_prior_failures": init_state(k, "No")
    elif k == "em_growth_observed": init_state(k, "No")
    elif k == "diff_changeover_analyst": init_state(k, "No")
    else: init_state(k, "")

if "data_loaded" not in st.session_state:
    load_saved_state()
    st.session_state.data_loaded = True

# --- EMAIL PARSER ---
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
        try: d_obj = datetime.strptime(date_match.group(1).strip(), "%d %b %Y"); st.session_state.test_date = d_obj.strftime("%d%b%y")
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
    if st.button("💾 Save Current Inputs"): save_current_state()
    st.success(f"Active: {st.session_state.active_platform}")

st.title(f"Sterility Investigation & Reporting: {st.session_state.active_platform}")

# --- SMART PARSER ---
st.header("📧 Smart Email Import")
email_input = st.text_area("Paste the OOS Notification email here to auto-fill fields:", height=150)
if st.button("🪄 Parse Email & Auto-Fill"):
    if email_input: parse_email_text(email_input); st.success("Fields updated!"); st.rerun()

# --- SECTION 1 ---
st.header("1. General Test Details")
col1, col2, col3 = st.columns(3)
with col1:
    st.text_input("OOS Number (Numbers only)", key="oos_id")
    st.text_input("Client Name", key="client_name")
    st.text_input("Sample ID (ETX Format)", key="sample_id")
with col2:
    st.text_input("Test Date (e.g., 07Jan26)", key="test_date")
    st.text_input("Sample / Active Name", key="sample_name")
    st.text_input("Lot Number", key="lot_number")
with col3:
    dosage_options = ["Injectable", "Aqueous Solution", "Liquid", "Solution"]
    st.selectbox("Dosage Form", dosage_options, key="dosage_form", index=0 if st.session_state.dosage_form not in dosage_options else dosage_options.index(st.session_state.dosage_form))
    st.text_input("Monthly Cleaning Date", key="monthly_cleaning_date")

# --- SECTION 2 ---
if st.session_state.active_platform == "ScanRDI":
    st.header("2. Personnel")
    p1, p2, p3 = st.columns(3)
    with p1:
        st.text_input("Prepper Initials", key="prepper_initial")
        pre_full = get_full_name(st.session_state.prepper_initial)
        st.text_input("Prepper Full Name", key="prepper_name", value=st.session_state.prepper_name if st.session_state.prepper_name else pre_full)
    with p2:
        st.text_input("Processor Initials", key="analyst_initial")
        ana_full = get_full_name(st.session_state.analyst_initial)
        st.text_input("Processor Full Name", key="analyst_name", value=st.session_state.analyst_name if st.session_state.analyst_name else ana_full)
    with p3:
        st.text_input("Reader Initials", key="reader_initial")
        rea_full = get_full_name(st.session_state.reader_initial)
        st.text_input("Reader Full Name", key="reader_name", value=st.session_state.reader_name if st.session_state.reader_name else rea_full)

    st.radio("Was the Changeover performed by a different analyst?", ["No", "Yes"], key="diff_changeover_analyst", horizontal=True)
    
    if st.session_state.diff_changeover_analyst == "Yes":
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Changeover Initials", key="changeover_initial")
        with c2:
            chg_full = get_full_name(st.session_state.changeover_initial)
            st.text_input("Changeover Full Name", key="changeover_name", value=st.session_state.changeover_name if st.session_state.changeover_name else chg_full)
    else:
        # Auto-sync
        st.session_state.changeover_initial = st.session_state.analyst_initial
        st.session_state.changeover_name = st.session_state.analyst_name

    st.divider()

    e1, e2 = st.columns(2)
    bsc_list = ["1310", "1309", "1311", "1312", "1314", "1313", "1316", "1798", "Other"]
    with e1:
        st.selectbox("Select Processing BSC ID", bsc_list, key="bsc_id", index=0 if st.session_state.bsc_id not in bsc_list else bsc_list.index(st.session_state.bsc_id))
        p_room, p_suite, p_suffix, p_loc = get_room_logic(st.session_state.bsc_id)
        st.caption(f"Processor: Suite {p_suite}{p_suffix} ({p_loc}) [Room ID: {p_room}]")
    with e2:
        st.radio("Was the Changeover performed in a different BSC?", ["No", "Yes"], key="diff_changeover_bsc", horizontal=True)
        if st.session_state.diff_changeover_bsc == "Yes":
            st.selectbox("Select Changeover BSC ID", bsc_list, key="chgbsc_id", index=0 if st.session_state.chgbsc_id not in bsc_list else bsc_list.index(st.session_state.chgbsc_id))
            c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
            st.caption(f"Changeover: Suite {c_suite}{c_suffix} ({c_loc}) [Room ID: {c_room}]")
        else:
            st.session_state.chgbsc_id = st.session_state.bsc_id
            st.info(f"Changeover BSC set to {st.session_state.bsc_id}")

    st.header("3. Findings & EM Data")
    f1, f2 = st.columns(2)
    with f1:
        scan_ids = ["1230", "2017", "1040", "1877", "2225", "2132"]
        st.selectbox("ScanRDI ID", scan_ids, key="scan_id", index=0 if st.session_state.scan_id not in scan_ids else scan_ids.index(st.session_state.scan_id))
        st.text_input("Shift Number", key="shift_number")
        shape_opts = ["rod", "cocci", "Other"]
        st.selectbox("Org Shape", shape_opts, key="org_choice", index=0 if st.session_state.org_choice not in shape_opts else shape_opts.index(st.session_state.org_choice))
        if st.session_state.org_choice == "Other": 
            st.text_input("Enter Manual Org Shape", key="manual_org")
        try: d_obj = datetime.strptime(st.session_state.test_date, "%d%b%y").strftime("%m%d%y"); st.session_state.test_record = f"{d_obj}-{st.session_state.scan_id}-{st.session_state.shift_number}"
        except: pass
        st.text_input("Record Ref", st.session_state.test_record, disabled=True)
    with f2:
        ctrl_opts = ["A. brasiliensis", "B. subtilis", "C. albicans", "C. sporogenes", "P. aeruginosa", "S. aureus"]
        st.selectbox("Positive Control", ctrl_opts, key="control_pos", index=0 if st.session_state.control_pos not in ctrl_opts else ctrl_opts.index(st.session_state.control_pos))
        st.text_input("Control Lot", key="control_lot")
        st.text_input("Control Exp Date", key="control_exp")

    # --- SECTION 4: EM OBSERVATIONS ---
    st.header("4. EM Observations")
    
    st.radio("Was microbial growth observed in Environmental Monitoring?", ["No", "Yes"], key="em_growth_observed", horizontal=True)
    
    if st.session_state.em_growth_observed == "Yes":
        st.caption("Please enter the growth details:")
        e1, e2, e3 = st.columns(3)
        with e1:
            st.text_input("Personnel Obs", key="obs_pers")
            st.text_input("Pers ETX #", key="etx_pers")
            st.text_input("Pers ID", key="id_pers")
        with e2:
            st.text_input("Surface Obs", key="obs_surf")
            st.text_input("Surf ETX #", key="etx_surf")
            st.text_input("Surf ID", key="id_surf")
        with e3:
            st.text_input("Settling Obs", key="obs_sett")
            st.text_input("Sett ETX #", key="etx_sett")
            st.text_input("Sett ID", key="id_sett")
        
        st.subheader("Weekly Bracketing (Growth Observations)")
        w1, w2 = st.columns(2)
        with w1:
            st.text_input("Weekly Air Obs", key="obs_air")
            st.text_input("Weekly Air ETX #", key="etx_air_weekly")
            st.text_input("Weekly Air ID", key="id_air_weekly")
        with w2:
            st.text_input("Weekly Surf Obs", key="obs_room")
            st.text_input("Weekly Surf ETX #", key="etx_room_weekly")
            st.text_input("Weekly Surf ID", key="id_room_wk_of")
    
    st.divider()
    st.caption("Weekly Bracketing (Date & Initials Required)")
    m1, m2 = st.columns(2)
    with m1:
        st.text_input("Weekly Monitor Initials", key="weekly_init")
    with m2:
        st.text_input("Date of Weekly Monitoring", key="date_weekly")

    if st.session_state.em_growth_observed == "Yes":
        if st.button("🔄 Generate Narrative & Details"):
            # Clean lists
            em_clean = []
            if not st.session_state.obs_pers.strip(): em_clean.append("personal sampling (left touch and right touch)")
            if not st.session_state.obs_surf.strip(): em_clean.append("surface sampling")
            if not st.session_state.obs_sett.strip(): em_clean.append("settling plates")
            wk_clean = []
            if not st.session_state.obs_air.strip(): wk_clean.append("weekly active air sampling")
            if not st.session_state.obs_room.strip(): wk_clean.append("weekly surface sampling")
            
            # Narrative
            narr = "Upon analyzing the environmental monitoring results, "
            if em_clean:
                if len(em_clean) == 1: clean_str = em_clean[0]
                elif len(em_clean) == 2: clean_str = f"{em_clean[0]} and {em_clean[1]}"
                else: clean_str = f"{em_clean[0]}, {em_clean[1]}, and {em_clean[2]}"
                narr += f"no microbial growth was observed in {clean_str}. "
            else: narr += "microbial growth was observed during the testing period. "
            if wk_clean:
                if len(wk_clean) == 1: wk_str = wk_clean[0]
                elif len(wk_clean) == 2: wk_str = f"{wk_clean[0]} and {wk_clean[1]}"
                else: wk_str = ", ".join(wk_clean[:-1]) + ", and " + wk_clean[-1]
                narr += f"Additionally, {wk_str} showed no microbial growth."
            
            st.session_state.narrative_summary = narr

            # Details
            sources = [
                ("personnel sampling", st.session_state.obs_pers, st.session_state.etx_pers, st.session_state.id_pers, "on the date of testing"),
                ("surface sampling", st.session_state.obs_surf, st.session_state.etx_surf, st.session_state.id_surf, "on the date of testing"),
                ("settling plates", st.session_state.obs_sett, st.session_state.etx_sett, st.session_state.id_sett, "on the date of testing"),
                ("weekly active air sampling", st.session_state.obs_air, st.session_state.etx_air_weekly, st.session_state.id_air_weekly, "the week of testing"),
                ("surface sampling of cleanroom during weekly room surface sampling", st.session_state.obs_room, st.session_state.etx_room_weekly, st.session_state.id_room_wk_of, "the week of testing")
            ]
            d_list = []
            for cat, obs, etx, oid, tctx in sources:
                if obs.strip():
                    is_sing = ("1" in obs and "CFU" in obs.upper() and "11" not in obs and "21" not in obs)
                    growth_term, plate_term = ("growth was", "plate was") if is_sing else ("growths were", "plates were")
                    id_label = "sample IDs" if ("," in etx or "AND" in etx.upper()) else "sample ID"
                    org_verb = "organisms identified included" if ("," in oid or "AND" in oid.upper()) else "organism identified was"
                    d_list.append(f"However, microbial {growth_term} observed in {cat} {tctx}. Specifically, {obs}. The {plate_term} submitted for microbial identification under {id_label} {etx}. The {org_verb} {oid}.")
            
            st.session_state.em_details = "\n\n".join(d_list)
            st.rerun()

        st.subheader("Narrative Summary (Editable)")
        st.text_area("Narrative Content", key="narrative_summary", height=120, label_visibility="collapsed")
        
        st.subheader("EM Growth Details (Editable)")
        st.text_area("Details Content", key="em_details", height=200, label_visibility="collapsed")
    else:
        # Default No Growth Logic
        st.session_state.narrative_summary = "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, and settling plates. Additionally, weekly active air sampling and weekly surface sampling showed no microbial growth."
        st.session_state.em_details = ""

    # --- SECTION 5 ---
    st.header("5. Automated Summaries & Analysis")
    
    st.subheader("Sample History")
    st.radio("Were there any prior failures in the last 6 months?", ["No", "Yes"], key="has_prior_failures", horizontal=True)
    if st.session_state.has_prior_failures == "Yes":
        st.number_input("Number of Prior Failures", min_value=1, step=1, key="incidence_count")
        st.text_input("Related OOS IDs (e.g. OOS-25001, OOS-25002)", key="oos_refs")
        
        if st.button("🔄 Generate History Text"):
            if st.session_state.incidence_count == 1: phrase = f"1 incident ({st.session_state.oos_refs})"
            else: phrase = f"{st.session_state.incidence_count} incidents ({st.session_state.oos_refs})"
            st.session_state.sample_history_paragraph = f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte “{st.session_state.sample_name}” has had {phrase} using the Scan RDI method during this period."
            st.rerun()
            
        st.text_area("History Text", key="sample_history_paragraph", height=120, label_visibility="collapsed")
    else:
        st.session_state.sample_history_paragraph = f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte “{st.session_state.sample_name}” has had no prior failures using the Scan RDI method during this period."

    st.divider()

    st.subheader("Cross-Contamination Analysis")
    st.radio("Did other samples test positive on the same day?", ["No", "Yes"], key="other_positives", horizontal=True)
    if st.session_state.other_positives == "Yes":
        st.number_input("Total # of Positive Samples that day", min_value=2, step=1, key="total_pos_count_num")
        st.number_input(f"Order of THIS Sample ({st.session_state.sample_id})", min_value=1, step=1, key="current_pos_order")
        
        num_others = st.session_state.total_pos_count_num - 1
        st.caption(f"Details for {num_others} other positive(s):")
        for i in range(num_others):
            sub_c1, sub_c2 = st.columns(2)
            with sub_c1:
                st.text_input(f"Other Sample #{i+1} ID", key=f"other_id_{i}")
            with sub_c2:
                st.number_input(f"Other Sample #{i+1} Order", min_value=1, step=1, key=f"other_order_{i}")
        
        if st.button("🔄 Generate Cross-Contam Text"):
            other_list_ids = []
            detail_sentences = []
            for i in range(num_others):
                oid = st.session_state.get(f"other_id_{i}", "")
                oord_num = st.session_state.get(f"other_order_{i}", 1)
                oord_text = ordinal(oord_num)
                if oid:
                    other_list_ids.append(oid)
                    detail_sentences.append(f"{oid} was the {oord_text} sample processed")
            all_ids = other_list_ids + [st.session_state.sample_id]
            if len(all_ids) == 2: ids_str = f"{all_ids[0]} and {all_ids[1]}"
            else: ids_str = ", ".join(all_ids[:-1]) + ", and " + all_ids[-1]
            count_word = num_to_words(st.session_state.total_pos_count_num)
            cur_ord_text = ordinal(st.session_state.current_pos_order)
            current_detail = f"while {st.session_state.sample_id} was the {cur_ord_text}"
            if len(detail_sentences) == 1: details_str = f"{detail_sentences[0]}, {current_detail}"
            else: details_str = ", ".join(detail_sentences) + f", {current_detail}"
            
            st.session_state.cross_contamination_summary = f"{ids_str} were the {count_word} samples tested positive for microbial growth. The analyst confirmed that these samples were not processed concurrently, sequentially, or within the same manifold run. Specifically, {details_str}. The analyst also verified that gloves were thoroughly disinfected between samples. Furthermore, all other samples processed by the analyst that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
            st.rerun()

        st.text_area("Cross-Contam Text", key="cross_contamination_summary", height=250, label_visibility="collapsed")
    else:
        st.session_state.cross_contamination_summary = "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."

    save_current_state()

# --- FINAL GENERATION ---
st.divider()
if st.button("🚀 GENERATE FINAL REPORT"):
    # Generate Equipment Summary (Hidden)
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
    
    if st.session_state.bsc_id == st.session_state.chgbsc_id:
        part1 = f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        part2 = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both testing and changeover steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{st.session_state.bsc_id} was certified and approved by both the Engineering and Quality Assurance teams. Sample processing and changeover were conducted in the ISO 5 BSC E00{st.session_state.bsc_id} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {st.session_state.analyst_name} on {st.session_state.test_date}."
        st.session_state.equipment_summary = f"{part1}\n\n{part2}"
    elif t_suite == c_suite:
        part1 = f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        part2 = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), and ISO 5 BSC E00{st.session_state.chgbsc_id}, located in the {c_loc}, (Suite {c_suite}{c_suffix}), were thoroughly cleaned and disinfected prior to their respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Furthermore, the BSCs used throughout testing, E00{st.session_state.bsc_id} for sample processing and E00{st.session_state.chgbsc_id} for the changeover step, were certified and approved by both the Engineering and Quality Assurance teams. Sample processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{st.session_state.bsc_id}) by {st.session_state.analyst_name} and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{st.session_state.chgbsc_id}) by {st.session_state.changeover_name} on {st.session_state.test_date}."
        st.session_state.equipment_summary = f"{part1}\n\n{part2}"
    else:
        part1 = f"The cleanroom used for testing (E00{t_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which opens into the middle ISO 7 buffer room ({t_suite}A), and then into the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        part2 = f"The cleanroom used for changeover (E00{c_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({c_suite}B), which opens into the middle ISO 7 buffer room ({c_suite}A), and then into the outermost ISO 8 anteroom ({c_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {c_suite}B through {c_suite}A and into {c_suite}."
        part3 = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), and ISO 5 BSC E00{st.session_state.chgbsc_id}, located in the {c_loc}, (Suite {c_suite}{c_suffix}), were thoroughly cleaned and disinfected prior to their respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Furthermore, the BSCs used throughout testing, E00{st.session_state.bsc_id} for sample processing and E00{st.session_state.chgbsc_id} for the changeover step, were certified and approved by both the Engineering and Quality Assurance teams. Sample processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{st.session_state.bsc_id}) by {st.session_state.analyst_name} and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{st.session_state.chgbsc_id}) by {st.session_state.changeover_name} on {st.session_state.test_date}."
        st.session_state.equipment_summary = f"{part1}\n\n{part2}\n\n{part3}"
    
    # Merge Narrative if EM details exist
    final_narrative = st.session_state.narrative_summary
    if st.session_state.em_growth_observed == "Yes" and st.session_state.em_details.strip():
        final_narrative += f"\n\n{st.session_state.em_details}"
    
    # Prepare Data
    safe_oos = clean_filename(st.session_state.oos_id)
    safe_client = clean_filename(st.session_state.client_name)
    safe_sample = clean_filename(st.session_state.sample_id)
    
    final_data = {k: v for k, v in st.session_state.items()}
    final_data["narrative_summary"] = final_narrative 
    final_data["em_details"] = "" # Clear to prevent double print
    final_data["oos_full"] = f"OOS-{safe_oos}"
    
    # Logic Re-run for safety
    if st.session_state.active_platform == "ScanRDI":
        t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
        final_data["cr_suit"] = t_suite; final_data["cr_id"] = t_room; final_data["suit"] = t_suffix; final_data["bsc_location"] = t_loc
        c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
        final_data["changeover_id"] = c_room; final_data["changeover_suit"] = c_suite; final_data["changeoversuit"] = c_suffix; final_data["changeover_location"] = c_loc; final_data["changeoverbsc_id"] = st.session_state.chgbsc_id
        final_data["changeover_name"] = st.session_state.changeover_name; final_data["analyst_name"] = st.session_state.analyst_name
        final_data["control_positive"] = st.session_state.control_pos; final_data["control_data"] = st.session_state.control_exp
        if st.session_state.org_choice == "Other": final_data["organism_morphology"] = st.session_state.manual_org
        else: final_data["organism_morphology"] = st.session_state.org_choice
        final_data["obs_pers_dur"] = st.session_state.obs_pers; final_data["etx_pers_dur"] = st.session_state.etx_pers; final_data["id_pers_dur"] = st.session_state.id_pers
        final_data["obs_surf_dur"] = st.session_state.obs_surf; final_data["etx_surf_dur"] = st.session_state.etx_surf; final_data["id_surf_dur"] = st.session_state.id_surf
        final_data["obs_sett_dur"] = st.session_state.obs_sett; final_data["etx_sett_dur"] = st.session_state.etx_sett; final_data["id_sett_dur"] = st.session_state.id_sett
        final_data["obs_air_wk_of"] = st.session_state.obs_air; final_data["etx_air_wk_of"] = st.session_state.etx_air_weekly; final_data["id_air_wk_of"] = st.session_state.id_air_weekly
        final_data["obs_room_wk_of"] = st.session_state.obs_room; final_data["etx_room_wk_of"] = st.session_state.etx_room_weekly; final_data["id_room_wk_of"] = st.session_state.id_room_wk_of
        final_data["weekly_initial"] = st.session_state.weekly_init; final_data["date_of_weekly"] = st.session_state.date_weekly

    for key in ["obs_pers", "obs_surf", "obs_sett", "obs_air", "obs_room"]:
        if not final_data[key].strip(): final_data[key] = "No Growth"
    try:
        dt_obj = datetime.strptime(st.session_state.test_date, "%d%b%y")
        final_data["date_before_test"] = (dt_obj - timedelta(days=1)).strftime("%d%b%y")
        final_data["date_after_test"] = (dt_obj + timedelta(days=1)).strftime("%d%b%y")
    except: pass
    
    template_name = f"{st.session_state.active_platform} OOS template.docx"
    if os.path.exists(template_name):
        doc = DocxTemplate(template_name)
        doc.render(final_data)
        out_name = f"OOS-{safe_oos} {safe_client} ({safe_sample}) - {st.session_state.active_platform}.docx"
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button(label="📂 Download Document", data=buf, file_name=out_name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
