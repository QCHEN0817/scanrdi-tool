import streamlit as st
from docxtpl import DocxTemplate
from datetime import datetime, timedelta
import io
import os

# --- PAGE SETUP ---
st.set_page_config(page_title="ScanRDI Report Generator", page_icon="🦠", layout="wide")

st.title("🦠 ScanRDI OOS Report Generator")
st.markdown("---")

# --- SIDEBAR: TEMPLATE UPLOAD ---
st.sidebar.header("📂 Step 1: Upload Template")
uploaded_file = st.sidebar.file_uploader("Upload 'ScanRDI OOS template.docx'", type="docx")

if not uploaded_file:
    st.info("👈 Please upload your Word template in the sidebar to begin.")
    st.stop()

# --- MAIN FORM ---
st.header("📝 Step 2: Fill Report Details")

# 1. General Info
st.subheader("1. General Info")
col1, col2 = st.columns(2)
with col1:
    oos_id = st.text_input("OOS Number", value="OOS-252503")
    client_name = st.text_input("Client Name", value="Northmark Pharmacy (E12955)")
    sample_id = st.text_input("Sample ID", value="SCAN-251029-001")
    sample_name = st.text_input("Sample Name", value="Sterile Saline")
with col2:
    lot_number = st.text_input("Lot Number", value="Lot-ABC-123")
    dosage_form = st.text_input("Dosage Form", value="Liquid")
    test_record = st.text_input("Record Ref", value="Record #12345")
    test_date = st.text_input("Test Date (e.g. 29Oct25)", value="29Oct25")
    monthly_cleaning_date = st.text_input("Cleaning Date", value="01Oct25")

# 2. People
st.subheader("2. People")
c1, c2, c3, c4 = st.columns(4)
with c1:
    prepper_name = st.text_input("Prepper Name", value="John Prepper")
    prepper_init = st.text_input("Prepper Initials", value="JP")
with c2:
    analyst_name = st.text_input("Processor Name", value="Gabrielle Surber")
    analyst_init = st.text_input("Processor Initials", value="GS")
with c3:
    change_name = st.text_input("Changeover Name", value="Jane Changeover")
    change_init = st.text_input("Changeover Initials", value="JC")
with c4:
    reader_name = st.text_input("Reader Name", value="Alex Reader")
    reader_init = st.text_input("Reader Initials", value="AR")

weekly_init = st.text_input("Weekly Monitor Initials", value="GS")

# 3. Equipment
st.subheader("3. Equipment (Smart Mode)")
st.caption("💡 **Logic:** If you leave the **Changeover** fields empty, the report will generate the 'Same Room' description. If you fill them, it generates the 'Two Rooms' description.")

e1, e2, e3 = st.columns(3)
with e1:
    scan_id = st.text_input("ScanRDI ID", value="1234")
    bsc_id = st.text_input("Process BSC ID", value="1314")
    chgbsc_id = st.text_input("Changeover BSC ID", value="", placeholder="Leave blank if same")
with e2:
    cr_id = st.text_input("Process Room ID", value="115")
    cr_suit = st.text_input("Process Suite #", value="115")
with e3:
    chg_id = st.text_input("Changeover Room ID", value="", placeholder="Leave blank if same")
    chg_suit = st.text_input("Changeover Suite #", value="", placeholder="Leave blank if same")

# 4. Results
st.subheader("4. Results")
r1, r2 = st.columns(2)
with r1:
    pos_ctrl = st.text_input("Pos Control", value="B. subtilis")
    control_lot = st.text_input("Control Lot", value="Lot-99")
with r2:
    control_data = st.text_input("Control Data", value="Exp: 01Jan26")
    org_morph = st.text_input("Org Shape", value="rod")

# 5. Observations
st.subheader("5. Observations (Table 1)")
def make_obs_row(label, prefix):
    c_obs, c_etx, c_id = st.columns([2, 1, 1])
    with c_obs: obs = st.text_input(f"{label} Obs", value="No Growth", key=f"{prefix}_obs")
    with c_etx: etx = st.text_area(f"{label} ETX", value="N/A", height=70, key=f"{prefix}_etx")
    with c_id: ids = st.text_area(f"{label} ID", value="N/A", height=70, key=f"{prefix}_id")
    return obs, etx, ids

obs_pers, etx_pers, id_pers = make_obs_row("Personnel", "p")
obs_surf, etx_surf, id_surf = make_obs_row("BSC Surface", "s")
obs_sett, etx_sett, id_sett = make_obs_row("Settling", "st")
obs_air, etx_air, id_air = make_obs_row("Weekly Air", "a")
obs_room, etx_room, id_room = make_obs_row("Weekly Room", "r")

# --- SMART LOGIC SECTION ---
st.markdown("### 6. Smart Generators")

# A. Narrative Generator
if 'narrative_text' not in st.session_state: st.session_state['narrative_text'] = ""
if 'details_text' not in st.session_state: st.session_state['details_text'] = ""

def run_smart_narrative():
    items = [
        ("personal sampling", obs_pers, etx_pers, id_pers, "Personnel Sampling", "during personnel monitoring"),
        ("surface sampling", obs_surf, etx_surf, id_surf, "Surface Sampling", "during BSC surface monitoring"),
        ("settling plates", obs_sett, etx_sett, id_sett, "Settling Sampling", "during BSC settling plate monitoring"),
        ("weekly air", obs_air, etx_air, id_air, "Active Air Sampling of Cleanrooms", "during weekly active air sampling"),
        ("weekly surface", obs_room, etx_room, id_room, "Surface Sampling of Cleanrooms", "during weekly room surface sampling")
    ]
    # Summary
    passed_daily = [i[0] for i in items[:3] if "no growth" in i[1].lower()]
    failures = [i for i in items if "no growth" not in i[1].lower()]

    if len(passed_daily) == 3: sent1 = "Upon reviewing the data presented in Table 1, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, or settling plates."
    elif len(passed_daily) > 0:
        list_str = ", ".join(passed_daily[:-1]) + ", or " + passed_daily[-1] if len(passed_daily) > 1 else passed_daily[0]
        sent1 = f"Upon reviewing the data presented in Table 1, no microbial growth was observed in {list_str}."
    else: sent1 = "Upon reviewing the data presented in Table 1, microbial growth was observed in all daily monitoring points."

    air_passed = "no growth" in obs_air.lower()
    room_passed = "no growth" in obs_room.lower()
    sent2 = ""
    if air_passed and room_passed: sent2 = " The weekly air and surface samplings also showed no growth."
    elif air_passed: sent2 = " The weekly air samplings also showed no growth."
    elif room_passed: sent2 = " The weekly surface samplings also showed no growth."
    st.session_state['narrative_text'] = sent1 + sent2

    # Details
    detail_text = ""
    for idx, f in enumerate(failures):
        obs_c = f[1].replace("\n", ", ")
        etx_c = f[2].replace("\n", ", ")
        ids_c = f[3].replace("\n", ", ")
        transition = "However" if idx == 0 else "Additionally"
        story = f"{transition}, microbial growths were observed in {f[4]} {f[5]} the week of testing. Specifically, {obs_c}. The plates were submitted for microbial identification under sample IDs {etx_c}. The organisms identified included {ids_c}.\n\n"
        detail_text += story
    st.session_state['details_text'] = detail_text.strip()

if st.button("🪄 1. Generate Smart Summary"):
    run_smart_narrative()

final_narrative = st.text_area("Narrative Summary", value=st.session_state['narrative_text'], height=100)
final_details = st.text_area("EM Details", value=st.session_state['details_text'], height=150)

# B. Equipment Logic (UPDATED)
def generate_equipment_text():
    # Detect if user filled in Changeover details
    has_changeover_info = chg_id.strip() != ""
    
    if not has_changeover_info:
        # --- SCENARIO 1: SAME ROOM ---
        text = f"""The cleanroom used for testing and changeover (E00{cr_id}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({cr_suit}B), which opens into the middle ISO 7 buffer room ({cr_suit}A), and then into the outermost ISO 8 anteroom ({cr_suit}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {cr_suit}B through {cr_suit}A and into {cr_suit}.

The ISO 5 BSC E00{bsc_id}, located in the innermost ISO 7 cleanroom ({cr_suit}B) was thoroughly cleaned and disinfected prior to their respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{bsc_id} was certified and approved by both the Engineering and Quality Assurance teams. Sample processing and changeover were conducted in the ISO 5 BSC E00{bsc_id} in the innermost ISO 7 room, Suite {cr_suit}B, by {change_name} on {test_date}.

The analyst, {analyst_name}, confirmed that the equipment was set up as per SOP 2.700.004 (Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup and Maintenance), and the negative control and the positive control for the analyst {analyst_name} yielded expected results."""
    
    else:
        # --- SCENARIO 2: DIFFERENT ROOMS (Uses your specific new text) ---
        text = f"""The cleanroom used for testing (E00{cr_id}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({cr_suit}B), which opens into the middle ISO 7 buffer room ({cr_suit}A), and then into the outermost ISO 8 anteroom ({cr_suit}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {cr_suit}B through {cr_suit}A and into {cr_suit}.

The cleanroom used for changeover (E00{chg_id}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({chg_suit}B), which opens into the middle ISO 7 buffer room ({chg_suit}A), and then into the outermost ISO 8 anteroom ({chg_suit}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {chg_suit}B through {chg_suit}A and into {chg_suit}.

The ISO 5 BSC E00{bsc_id}, located in the innermost ISO 7 cleanroom ({cr_suit}B), and ISO 5 BSC E00{chgbsc_id}, located in the middle ISO 7 buffer room ({chg_suit}A), were thoroughly cleaned and disinfected prior to their respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Furthermore, the BSCs used throughout testing, E00{bsc_id} for sample processing and E00{chgbsc_id} for the changeover step, were certified and approved by both the Engineering and Quality Assurance teams. Sample processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {cr_suit}B, BSC E00{bsc_id}) by {analyst_name} and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {chg_suit}A, BSC E00{chgbsc_id}) by {change_name} on {test_date}."""
    
    return text

# --- GENERATION ---
st.markdown("### 🚀 Final Step")

def generate_file():
    # Calculate dates
    try:
        dt = datetime.strptime(test_date, "%d%b%y")
        if dt.weekday() == 0: day_bef = dt - timedelta(days=3)
        elif dt.weekday() == 6: day_bef = dt - timedelta(days=2)
        else: day_bef = dt - timedelta(days=1)
        if dt.weekday() == 4: day_aft = dt + timedelta(days=3)
        else: day_aft = dt + timedelta(days=1)
        date_bef_str = day_bef.strftime("%d%b%y")
        date_aft_str = day_aft.strftime("%d%b%y")
    except:
        date_bef_str = "Unknown"
        date_aft_str = "Unknown"

    # Generate Equipment Text dynamically
    equipment_text = generate_equipment_text()
    
    # Fill defaults if blank (for the header fields)
    final_chg_name = change_name if change_name else analyst_name
    final_chg_init = change_init if change_init else analyst_init

    context = {
        'oos_id': oos_id, 'client_name': client_name, 'sample_id': sample_id,
        'sample_name': sample_name, 'lot_number': lot_number, 'dosage_form': dosage_form,
        'test_record': test_record, 'test_date': test_date, 'monthly_cleaning_date': monthly_cleaning_date,
        'prepper_name': prepper_name, 'prepper_initial': prepper_init,
        'analyst_name': analyst_name, 'analyst_initial': analyst_init,
        'changeover_name': final_chg_name, 'changeover_initial': final_chg_init,
        'reader_name': reader_name, 'reader_initial': reader_init,
        'weekly_initial': weekly_init, 'date_of_weekly': test_date,
        'scan_id': scan_id, 'cr_id': cr_id, 'cr_suit': cr_suit, 'bsc_id': bsc_id,
        'changeover_id': chg_id if chg_id else cr_id,
        'changeover_suit': chg_suit if chg_suit else cr_suit,
        'changeoverbsc_id': chgbsc_id if chgbsc_id else bsc_id,
        'control_positive': pos_ctrl, 'control_lot': control_lot, 'control_data': control_data,
        'organism_morphology': org_morph,
        'obs_pers_dur': obs_pers, 'etx_pers_dur': etx_pers, 'id_pers_dur': id_pers,
        'obs_surf_dur': obs_surf, 'etx_surf_dur': etx_surf, 'id_surf_dur': id_surf,
        'obs_sett_dur': obs_sett, 'etx_sett_dur': etx_sett, 'id_sett_dur': id_sett,
        'obs_air_wk_of': obs_air, 'etx_air_wk_of': etx_air, 'id_air_wk_of': id_air,
        'obs_room_wk_of': obs_room, 'etx_room_wk_of': etx_room, 'id_room_wk_of': id_room,
        'date_before_test': date_bef_str, 'date_after_test': date_aft_str,
        'narrative_summary': final_narrative, 'em_details': final_details,
        'equipment_summary': equipment_text
    }

    doc = DocxTemplate(uploaded_file)
    doc.render(context)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio

if uploaded_file:
    clean_id = oos_id.strip()
    if not clean_id.upper().startswith("OOS-"):
        clean_id = f"OOS-{clean_id}"
    
    out_filename = f"{clean_id} {client_name} -Scan RDI.docx"
    
    if st.button("⬇️ Download Final Report"):
        st.download_button(
            label="Download Now",
            data=generate_file(),
            file_name=out_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
