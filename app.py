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

field_keys = [
    "oos_id", "client_name", "sample_id", "test_date", "sample_name", "lot_number", 
    "dosage_form", "monthly_cleaning_date", "prepper_initial", "analyst_initial", 
    "changeover_initial", "reader_initial", "bsc_id", "chgbsc_id", "scan_id", 
    "org_choice", "manual_org", "test_record", "control_pos", "control_lot", 
    "control_exp", "obs_pers", "etx_pers", "id_pers", "obs_surf", "etx_surf", 
    "id_surf", "obs_sett", "etx_sett", "id_sett", "obs_air", "etx_air_weekly", 
    "id_air_weekly", "obs_room", "etx_room_weekly", "id_room_weekly", "weekly_init", 
    "date_weekly", "equipment_summary", "narrative_summary", "em_details", 
    "sample_history_para", "incidence_count", "oos_refs"
]

for k in field_keys:
    if k == "oos_id": init_state(k, "OOS-252503")
    elif k == "client_name": init_state(k, "Northmark Pharmacy")
    elif k == "sample_id": init_state(k, "E12955")
    elif k == "test_date": init_state(k, datetime.now().strftime("%d%b%y"))
    elif k == "dosage_form": init_state(k, "Injectable")
    elif k == "incidence_count": init_state(k, 0)
    elif "etx" in k or "id" in k: init_state(k, "N/A")
    else: init_state(k, "")

# --- SIDEBAR ---
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
    st.session_state.dosage_form = st.selectbox("Dosage Form", dosage_options, index=dosage_options.index(st.session_state.dosage_form) if st.session_state.dosage_form in dosage_options else 0)
    st.session_state.monthly_cleaning_date = st.text_input("Monthly Cleaning Date", st.session_state.monthly_cleaning_date)

# --- SECTION 2: SCANRDI PERSONNEL/FACILITY ---
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

    def get_room_logic(bsc_id):
        try:
            num = int(bsc_id)
            suffix = "B" if num % 2 == 0 else "A"
            location = "innermost ISO 7 room" if suffix == "B" else "middle ISO 7 buffer room"
        except: suffix, location = "B", "innermost ISO 7 room"
        if bsc_id in ["1310", "1309"]: room, suite = "117", "117"
        elif bsc_id in ["1311", "1312"]: room, suite = "116", "116"
        elif bsc_id in ["1314", "1313"]: room, suite = "115", "115"
        elif bsc_id in ["1316", "1798"]: room, suite = "114", "11
