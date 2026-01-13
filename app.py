import streamlit as st
from docxtpl import DocxTemplate
import os
import re
import json
import io
from datetime import datetime, timedelta

# --- PAGE CONFIG ---
st.set_page_config(page_title="LabOps Report Tool", layout="wide")

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
    "em_growth_observed", "diff_changeover_analyst",
    "diff_reader_analyst",
    "em_growth_count" 
]
# Dynamic keys
for i in range(20):
    field_keys.append(f"other_id_{i}")
    field_keys.append(f"other_order_{i}")
    field_keys.append(f"prior_oos_{i}")
    field_keys.append(f"em_cat_{i}")
    field_keys.append(f"em_obs_{i}")
    field_keys.append(f"em_etx_{i}")
    field_keys.append(f"em_id_{i}")

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

# --- GENERATE LIVE TEXTS ---
def generate_equipment_text():
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
    
    if st.session_state.bsc_id == st.session_state.chgbsc_id:
        part1 = f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        part2 = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both testing and changeover steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{st.session_state.bsc_id} was certified and approved by both the Engineering and Quality Assurance teams. Sample processing and changeover were conducted in the ISO 5 BSC E00{st.session_state.bsc_id} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {st.session_state.analyst_name} on {st.session_state.test_date}."
        return f"{part1}\n\n{part2}"
    elif t_suite == c_suite:
        part1 = f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        part2 = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), and ISO 5 BSC E00{st.session_state.chgbsc_id}, located in the {c_loc}, (Suite {c_suite}{c_suffix}), were thoroughly cleaned and disinfected prior to their respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Furthermore, the BSCs used throughout testing, E00{st.session_state.bsc_id} for sample processing and E00{st.session_state.chgbsc_id} for the changeover step, were certified and approved by both the Engineering and Quality Assurance teams. Sample processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{st.session_state.bsc_id}) by {st.session_state.analyst_name} and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{st.session_state.chgbsc_id}) by {st.session_state.changeover_name} on {st.session_state.test_date}."
        return f"{part1}\n\n{part2}"
    else:
        part1 = f"The cleanroom used for testing (E00{t_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which opens into the middle ISO 7 buffer room ({t_suite}A), and then into the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        part2 = f"The cleanroom used for changeover (E00{c_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({c_suite}B), which opens into the middle ISO 7 buffer room ({c_suite}A), and then into the outermost ISO 8 anteroom ({c_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {c_suite}B through {c_suite}A and into {c_suite}."
        part3 = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), and ISO 5 BSC E00{st.session_state.chgbsc_id}, located in the {c_loc}, (Suite {c_suite}{c_suffix}), were thoroughly cleaned and disinfected prior to their respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Furthermore, the BSCs used throughout testing, E00{st.session_state.bsc_id} for sample processing and E00{st.session_state.chgbsc_id} for the changeover step, were certified and approved by both the Engineering and Quality Assurance teams. Sample processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{st.session_state.bsc_id}) by {st.session_state.analyst_name} and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{st.session_state.chgbsc_id}) by {st.session_state.changeover_name} on {st.session_state.test_date}."
        return f"{part1}\n\n{part2}\n\n{part3}"

def generate_history_text():
    if st.session_state.incidence_count == 0: 
        hist_phrase = "no prior failures"
    else:
        prior_ids = []
        for i in range(st.session_state.incidence_count):
            pid = st.session_state.get(f"prior_oos_{i}", "").strip()
            if pid: prior_ids.append(pid)
        
        if not prior_ids:
            refs_str = "..."
        elif len(prior_ids) == 1:
            refs_str = prior_ids[0]
        elif len(prior_ids) == 2:
            refs_str = f"{prior_ids[0]} and {prior_ids[1]}"
        else:
            refs_str = ", ".join(prior_ids[:-1]) + f", and {prior_ids[-1]}"
        
        if st.session_state.incidence_count == 1: 
            hist_phrase = f"1 incident ({refs_str})"
        else: 
            hist_phrase = f"{st.session_state.incidence_count} incidents ({refs_str})"
            
    return f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte “{st.session_state.sample_name}” has had {hist_phrase} using the Scan RDI method during this period."

def generate_cross_contam_text():
    if st.session_state.other_positives == "No":
        return "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
    else:
        num_others = st.session_state.total_pos_count_num - 1
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
        
        if not all_ids: ids_str = ""
        elif len(all_ids) == 1: ids_str = all_ids[0]
        elif len(all_ids) == 2: ids_str = f"{all_ids[0]} and {all_ids[1]}"
        else: ids_str = ", ".join(all_ids[:-1]) + f", and {all_ids[-1]}"
        
        count_word = num_to_words(st.session_state.total_pos_count_num)
        cur_ord_text = ordinal(st.session_state.current_pos_order)
        current_detail = f"while {st.session_state.sample_id} was the {cur_ord_text}"
        
        if len(detail_sentences) == 1: details_str = f"{detail_sentences[0]}, {current_detail}"
        else: details_str = ", ".join(detail_sentences) + f", {current_detail}"

        return f"{ids_str} were the {count_word} samples tested positive for microbial growth. The analyst confirmed that these samples were not processed concurrently, sequentially, or within the same manifold run. Specifically, {details_str}. The analyst also verified that gloves were thoroughly disinfected between samples. Furthermore, all other samples processed by the analyst that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."

def generate_narrative_and_details():
    # 1. Identify Failures FROM DYNAMIC FIELDS
    failures = []
    count = st.session_state.get("em_growth_count", 1)
    
    cat_map = {
        "Personnel Obs": "personnel sampling",
        "Surface Obs": "surface sampling",
        "Settling Obs": "settling plates",
        "Weekly Air Obs": "weekly active air sampling",
        "Weekly Surf Obs": "weekly surface sampling"
    }
    
    for i in range(count):
        cat_friendly = st.session_state.get(f"em_cat_{i}", "Personnel Obs")
        obs_val = st.session_state.get(f"em_obs_{i}", "")
        etx_val = st.session_state.get(f"em_etx_{i}", "")
        id_val = st.session_state.get(f"em_id_{i}", "")
        
        category = cat_map.get(cat_friendly, "personnel sampling")
        if "weekly" in category: time_ctx = "weekly"
        else: time_ctx = "daily"
        
        if obs_val.strip():
            failures.append({"cat": category, "obs": obs_val, "etx": etx_val, "id": id_val, "time": time_ctx})

    # 2. Build "Pass" Narrative
    failed_cats = [f["cat"] for f in failures]
    all_daily = ["personnel sampling", "surface sampling", "settling plates"]
    all_weekly = ["weekly active air sampling", "weekly surface sampling"]
    
    pass_em_clean = [c.replace("personnel sampling", "personal sampling (left touch and right touch)") for c in all_daily if c not in failed_cats]
    pass_wk_clean = [c for c in all_weekly if c not in failed_cats]

    narr = "Upon analyzing the environmental monitoring results, "
    has_clean_daily = False
    if pass_em_clean:
        if len(pass_em_clean) == 1: clean_str = pass_em_clean[0]
        elif len(pass_em_clean) == 2: clean_str = f"{pass_em_clean[0]} and {pass_em_clean[1]}"
        else: clean_str = f"{pass_em_clean[0]}, {pass_em_clean[1]}, and {pass_em_clean[2]}"
        narr += f"no microbial growth was observed in {clean_str}. "
        has_clean_daily = True
    
    if not pass_em_clean and not pass_wk_clean:
        narr = "Upon analyzing the environmental monitoring results, microbial growth was observed. "

    if pass_wk_clean:
        if len(pass_wk_clean) == 1: wk_str = pass_wk_clean[0]
        elif len(pass_wk_clean) == 2: wk_str = f"{pass_wk_clean[0]} and {pass_wk_clean[1]}"
        else: wk_str = ", ".join(pass_wk_clean)
        
        if has_clean_daily:
            narr += f"Additionally, {wk_str} showed no microbial growth."
        elif not pass_em_clean:
             narr += f"However, {wk_str} showed no microbial growth."

    # 3. Build "Fail" Narrative (Combined Paragraph)
    det = ""
    if failures:
        # Build Intro
        daily_fails = sorted(list(set([f["cat"] for f in failures if f['time'] == 'daily'])))
        weekly_fails = sorted(list(set([f["cat"] for f in failures if f['time'] == 'weekly'])))
        
        intro_parts = []
        if daily_fails:
            if len(daily_fails) == 1: d_str = daily_fails[0]
            elif len(daily_fails) == 2: d_str = f"{daily_fails[0]} and {daily_fails[1]}"
            else: d_str = ", ".join(daily_fails[:-1]) + f", and {daily_fails[-1]}"
            intro_parts.append(f"{d_str} on the date")
            
        if weekly_fails:
            if len(weekly_fails) == 1: w_str = weekly_fails[0]
            elif len(weekly_fails) == 2: w_str = f"{weekly_fails[0]} and {weekly_fails[1]}"
            else: w_str = ", ".join(weekly_fails[:-1]) + f", and {weekly_fails[-1]}"
            intro_parts.append(f"{w_str} from week of testing")
            
        if len(intro_parts) == 2:
            fail_intro = f"However, microbial growth was observed during both {intro_parts[0]} and {intro_parts[1]}."
        else:
            fail_intro = f"However, microbial growth was observed during {intro_parts[0]}."
        
        # Build Details - SPLIT SENTENCES LOGIC
        detail_sentences = []
        for i, f in enumerate(failures):
            base_sentence = f"{f['obs']} was detected during {f['cat']} and was submitted for microbial identification under sample ID {f['etx']}, where the organism was identified as {f['id']}"
            if i == 0: full_sent = f"Specifically, {base_sentence}."
            elif i == 1: full_sent = f"Additionally, {base_sentence}."
            elif i == 2: full_sent = f"Furthermore, {base_sentence}."
            else: full_sent = f"Also, {base_sentence}."
            detail_sentences.append(full_sent)
            
        det = f"{fail_intro} {' '.join(detail_sentences)}"

    return narr, det, failures

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
    elif k == "diff_reader_analyst": init_state(k, "No") 
    elif k == "em_growth_count": init_state(k, 1) 
    elif k.startswith("other_order_"): init_state(k, 1)
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
    
    # LOT NUMBER FIX: Greedily capture everything on the line after 'Lot:'
    lot_match = re.search(r"(?:Lot|Batch)\s*[:\.]?\s*([^\n\r]+)", text, re.IGNORECASE)
    if lot_match: st.session_state.lot_number = lot_match.group(1).strip()
    
    # MORPHOLOGY PARSER
    morph_match = re.search(r"exhibiting\s*[\W]*\s*(\w+)\s*[\W]*-shaped\s*morphology", text, re.IGNORECASE)
    if morph_match:
        shape = morph_match.group(1).lower()
        if "cocci" in shape: st.session_state.org_choice = "cocci"
        elif "rod" in shape: st.session_state.org_choice = "rod"
        else: 
            st.session_state.org_choice = "Other"
            st.session_state.manual_org = shape

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
    st.title("QC Platforms")
    if st.button("ScanRDI"): st.session_state.active_platform = "ScanRDI"
    if st.button("Celsis"): st.session_state.active_platform = "Celsis"
    if st.button("USP 71"): st.session_state.active_platform = "USP 71"
    st.divider()
    if st.button("💾 Save Current Inputs"): save_current_state()
    st.success(f"Active: {st.session_state.active_platform}")

st.title(f"LabOps Report Tool: {st.session_state.active_platform}")

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
    p1, p2 = st.columns(2)
    with p1:
        st.text_input("Prepper Initials", key="prepper_initial")
        if st.session_state.prepper_initial and not st.session_state.prepper_name:
            st.session_state.prepper_name = get_full_name(st.session_state.prepper_initial)
        st.text_input("Prepper Full Name", key="prepper_name")
    with p2:
        st.text_input("Processor Initials", key="analyst_initial")
        if st.session_state.analyst_initial and not st.session_state.analyst_name:
            st.session_state.analyst_name = get_full_name(st.session_state.analyst_initial)
        st.text_input("Processor Full Name", key="analyst_name")

    # READER LOGIC (NEW)
    st.session_state.diff_reader_analyst = st.radio("Was the Reading performed by a different analyst?", ["No", "Yes"], index=0 if st.session_state.diff_reader_analyst == "No" else 1, horizontal=True)
    if st.session_state.diff_reader_analyst == "Yes":
        c1, c2 = st.columns(2)
        with c1: st.text_input("Reader Initials", key="reader_initial")
        with c2: 
            if st.session_state.reader_initial and not st.session_state.reader_name:
                st.session_state.reader_name = get_full_name(st.session_state.reader_initial)
            st.text_input("Reader Full Name", key="reader_name")
    else:
        st.session_state.reader_initial = st.session_state.analyst_initial
        st.session_state.reader_name = st.session_state.analyst_name

    # CHANGEOVER LOGIC
    st.session_state.diff_changeover_analyst = st.radio("Was the Changeover performed by a different analyst?", ["No", "Yes"], index=0 if st.session_state.diff_changeover_analyst == "No" else 1, horizontal=True)
    if st.session_state.diff_changeover_analyst == "Yes":
        c1, c2 = st.columns(2)
        with c1: st.text_input("Changeover Initials", key="changeover_initial")
        with c2: 
            if st.session_state.changeover_initial and not st.session_state.changeover_name:
                st.session_state.changeover_name = get_full_name(st.session_state.changeover_initial)
            st.text_input("Changeover Full Name", key="changeover_name")
    else:
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
        if st.session_state.em_growth_count < 1: st.session_state.em_growth_count = 1
        count = st.number_input("Number of EM Failures", min_value=1, step=1, key="em_growth_count")
        
        cat_options = ["Personnel Obs", "Surface Obs", "Settling Obs", "Weekly Air Obs", "Weekly Surf Obs"]
        
        for i in range(count):
            st.subheader(f"Growth #{i+1}")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.selectbox(f"Category", cat_options, key=f"em_cat_{i}")
            with c2:
                st.text_input(f"Observation (e.g. 1 CFU...)", key=f"em_obs_{i}")
            with c3:
                st.text_input(f"ETX #", key=f"em_etx_{i}")
            with c4:
                st.text_input(f"Microbial ID", key=f"em_id_{i}")
    
    st.divider()
    st.caption("Weekly Bracketing (Date & Initials Required)")
    m1, m2 = st.columns(2)
    with m1:
        st.text_input("Weekly Monitor Initials", key="weekly_init")
    with m2:
        st.text_input("Date of Weekly Monitoring", key="date_weekly")

    if st.session_state.em_growth_observed == "Yes":
        if st.button("🔄 Generate Narrative & Details"):
            n, d, failures = generate_narrative_and_details()
            st.session_state.narrative_summary = n
            st.session_state.em_details = d
            
            # Reset
            st.session_state.obs_pers = ""; st.session_state.etx_pers = ""; st.session_state.id_pers = ""
            st.session_state.obs_surf = ""; st.session_state.etx_surf = ""; st.session_state.id_surf = ""
            st.session_state.obs_sett = ""; st.session_state.etx_sett = ""; st.session_state.id_sett = ""
            st.session_state.obs_air = ""; st.session_state.etx_air_weekly = ""; st.session_state.id_air_weekly = ""
            st.session_state.obs_room = ""; st.session_state.etx_room_weekly = ""; st.session_state.id_room_wk_of = ""
            
            def join_val(old, new): return f"{old}, {new}" if old else new
            
            for f in failures:
                if f['cat'] == "personnel sampling":
                    st.session_state.obs_pers = join_val(st.session_state.obs_pers, f['obs'])
                    st.session_state.etx_pers = join_val(st.session_state.etx_pers, f['etx'])
                    st.session_state.id_pers = join_val(st.session_state.id_pers, f['id'])
                elif f['cat'] == "surface sampling":
                    st.session_state.obs_surf = join_val(st.session_state.obs_surf, f['obs'])
                    st.session_state.etx_surf = join_val(st.session_state.etx_surf, f['etx'])
                    st.session_state.id_surf = join_val(st.session_state.id_surf, f['id'])
                elif f['cat'] == "settling plates":
                    st.session_state.obs_sett = join_val(st.session_state.obs_sett, f['obs'])
                    st.session_state.etx_sett = join_val(st.session_state.etx_sett, f['etx'])
                    st.session_state.id_sett = join_val(st.session_state.id_sett, f['id'])
                elif f['cat'] == "weekly active air sampling":
                    st.session_state.obs_air = join_val(st.session_state.obs_air, f['obs'])
                    st.session_state.etx_air_weekly = join_val(st.session_state.etx_air_weekly, f['etx'])
                    st.session_state.id_air_weekly = join_val(st.session_state.id_air_weekly, f['id'])
                elif f['cat'] == "weekly surface sampling":
                    st.session_state.obs_room = join_val(st.session_state.obs_room, f['obs'])
                    st.session_state.etx_room_weekly = join_val(st.session_state.etx_room_weekly, f['etx'])
                    st.session_state.id_room_wk_of = join_val(st.session_state.id_room_wk_of, f['id'])
                    
            st.rerun()

        st.subheader("Narrative Summary (Editable)")
        st.text_area("Narrative Content", key="narrative_summary", height=120, label_visibility="collapsed")
        
        st.subheader("EM Growth Details (Editable)")
        st.text_area("Details Content", key="em_details", height=200, label_visibility="collapsed")
    else:
        st.session_state.narrative_summary = "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, and settling plates. Additionally, weekly active air sampling and weekly surface sampling showed no microbial growth."
        st.session_state.em_details = ""

    # --- SECTION 5 ---
    st.header("5. Automated Summaries & Analysis")
    
    st.subheader("Sample History")
    st.radio("Were there any prior failures in the last 6 months?", ["No", "Yes"], key="has_prior_failures", horizontal=True)
    if st.session_state.has_prior_failures == "Yes":
        if st.session_state.incidence_count < 1: st.session_state.incidence_count = 1
        count = st.number_input("Number of Prior Failures", min_value=1, step=1, key="incidence_count")
        
        for i in range(count):
            if f"prior_oos_{i}" not in st.session_state: st.session_state[f"prior_oos_{i}"] = ""
            st.text_input(f"Prior Failure #{i+1} OOS ID", key=f"prior_oos_{i}")
        
        if st.button("🔄 Generate History Text"):
            st.session_state.sample_history_paragraph = generate_history_text()
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
                if f"other_order_{i}" not in st.session_state: st.session_state[f"other_order_{i}"] = 1
                st.number_input(f"Other Sample #{i+1} Order", min_value=1, step=1, key=f"other_order_{i}")
        
        if st.button("🔄 Generate Cross-Contam Text"):
            st.session_state.cross_contamination_summary = generate_cross_contam_text()
            st.rerun()

        st.text_area("Cross-Contam Text", key="cross_contamination_summary", height=250, label_visibility="collapsed")
    else:
        st.session_state.cross_contamination_summary = "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."

    save_current_state()

# --- FINAL GENERATION ---
st.divider()
if st.button("🚀 GENERATE FINAL REPORT"):
    st.session_state.equipment_summary = generate_equipment_text()
    
    if st.session_state.em_growth_observed == "No":
        n, d, _ = generate_narrative_and_details()
        st.session_state.narrative_summary = n
        st.session_state.em_details = d
        st.session_state.obs_pers = ""; st.session_state.etx_pers = ""; st.session_state.id_pers = ""
        st.session_state.obs_surf = ""; st.session_state.etx_surf = ""; st.session_state.id_surf = ""
        st.session_state.obs_sett = ""; st.session_state.etx_sett = ""; st.session_state.id_sett = ""
        st.session_state.obs_air = ""; st.session_state.etx_air_weekly = ""; st.session_state.id_air_weekly = ""
        st.session_state.obs_room = ""; st.session_state.etx_room_weekly = ""; st.session_state.id_room_wk_of = ""

    if st.session_state.has_prior_failures == "No":
        st.session_state.sample_history_paragraph = generate_history_text()
    if st.session_state.other_positives == "No":
        st.session_state.cross_contamination_summary = generate_cross_contam_text()

    final_narrative = st.session_state.narrative_summary
    if st.session_state.em_growth_observed == "Yes" and st.session_state.em_details.strip():
        final_narrative += f"\n\n{st.session_state.em_details}"
    
    safe_oos = clean_filename(st.session_state.oos_id)
    safe_client = clean_filename(st.session_state.client_name)
    safe_sample = clean_filename(st.session_state.sample_id)
    
    final_data = {k: v for k, v in st.session_state.items()}
    final_data["narrative_summary"] = final_narrative 
    final_data["em_details"] = "" 
    final_data["oos_full"] = f"OOS-{safe_oos}"
    
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
