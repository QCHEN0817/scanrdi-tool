# --- PART 1: INSTALL & SETUP ---
!pip install docxtpl ipywidgets
from docxtpl import DocxTemplate
import ipywidgets as widgets
from IPython.display import display, clear_output
from datetime import datetime, timedelta
import os
import json
from google.colab import files

# --- PART 2: DATA & HISTORY ---
CURRENT_HISTORY_FILE = 'oos_wizard_history.json'

def load_history():
    if os.path.exists(CURRENT_HISTORY_FILE):
        try:
            with open(CURRENT_HISTORY_FILE, 'r') as f: return json.load(f)
        except: return {}
    return {}

def save_history(data_dict):
    with open(CURRENT_HISTORY_FILE, 'w') as f: json.dump(data_dict, f)

defaults = load_history()
def get_val(key, fallback): return defaults.get(key, fallback)

# --- PART 3: UI WIDGETS DEFINITION ---
style = {'description_width': 'initial'}
layout_full = widgets.Layout(width='98%')
layout_half = widgets.Layout(width='48%')

# Platform Selection
platform_selector = widgets.ToggleButtons(
    options=['ScanRDI', 'Celsis', 'USP 71'],
    description='Select OS:',
    button_style='info',
    tooltips=['Rapid Sterility', 'ATP Bioluminescence', 'Traditional Sterility'],
)

# Shared Widgets
w_oos_id = widgets.Text(description="OOS Number:", value=get_val('oos_id', "OOS-250000"), style=style, layout=layout_full)
w_client = widgets.Text(description="Client Name:", value=get_val('client_name', "Pharmacy Name"), style=style, layout=layout_full)
w_sample_id = widgets.Text(description="Sample ID:", value=get_val('sample_id', ""), style=style, layout=layout_full)
w_sample_name = widgets.Text(description="Sample Name:", value=get_val('sample_name', ""), style=style, layout=layout_full)
w_lot = widgets.Text(description="Lot Number:", value=get_val('lot_number', ""), style=style, layout=layout_full)
w_test_date = widgets.Text(description="Test Date (DDMonYY):", value=datetime.now().strftime("%d%b%y"), style=style, layout=layout_full)

# ScanRDI Specific Widgets
w_prepper = widgets.Text(description="Prepper Name:", value=get_val('prepper_name', ""), style=style, layout=layout_half)
w_prepper_init = widgets.Text(description="Prepper Initials:", value=get_val('prepper_initial', ""), style=style, layout=layout_half)
w_analyst = widgets.Text(description="Processor Name:", value=get_val('analyst_name', ""), style=style, layout=layout_half)
w_analyst_init = widgets.Text(description="Processor Initials:", value=get_val('analyst_initial', ""), style=style, layout=layout_half)
w_changeover = widgets.Text(description="Changeover Name:", value=get_val('changeover_name', ""), style=style, layout=layout_half)
w_changeover_init = widgets.Text(description="Changeover Initials:", value=get_val('changeover_initial', ""), style=style, layout=layout_half)
w_reader = widgets.Text(description="Reader Name:", value=get_val('reader_name', ""), style=style, layout=layout_half)
w_reader_init = widgets.Text(description="Reader Initials:", value=get_val('reader_initial', ""), style=style, layout=layout_half)
w_scan_id = widgets.Text(description="ScanRDI ID (E00...):", value=get_val('scan_id', ""), style=style, layout=layout_half)
w_cr_suit = widgets.Text(description="Suite/Room #:", value=get_val('cr_suit', ""), style=style, layout=layout_half)
w_bsc_id = widgets.Text(description="BSC ID (E00...):", value=get_val('bsc_id', ""), style=style, layout=layout_half)
w_org_morph = widgets.Text(description="Org Shape:", value=get_val('organism_morphology', "rod"), style=style, layout=layout_half)

# --- PART 4: DYNAMIC UI RENDERER ---
output_form = widgets.Output()

def render_form(change=None):
    with output_form:
        clear_output()
        selection = platform_selector.value
        
        display(widgets.HTML(f"<h2>{selection} Investigation Data Entry</h2>"))
        
        # Section 1: Common Info
        display(widgets.HTML("<b>1. General Information</b>"))
        display(w_oos_id, w_client, w_sample_id, w_sample_name, w_lot, w_test_date)
        
        # Section 2: Platform Specific
        if selection == 'ScanRDI':
            display(widgets.HTML("<b>2. Personnel (ScanRDI)</b>"))
            display(widgets.HBox([w_prepper, w_prepper_init]))
            display(widgets.HBox([w_analyst, w_analyst_init]))
            display(widgets.HBox([w_changeover, w_changeover_init]))
            display(widgets.HBox([w_reader, w_reader_init]))
            display(widgets.HTML("<b>3. Equipment & Findings</b>"))
            display(widgets.HBox([w_scan_id, w_cr_suit]))
            display(widgets.HBox([w_bsc_id, w_org_morph]))
            
        elif selection == 'Celsis':
            display(widgets.HTML("<b>2. Celsis Specific Fields</b>"))
            display(widgets.Text(description="Luminometer ID:", style=style, layout=layout_full))
            display(widgets.Text(description="Reagent Kit Lot:", style=style, layout=layout_full))
            
        elif selection == 'USP 71':
            display(widgets.HTML("<b>2. USP 71 Specific Fields</b>"))
            display(widgets.Text(description="FTM Lot #:", style=style, layout=layout_half))
            display(widgets.Text(description="TSB Lot #:", style=style, layout=layout_half))
            display(widgets.Dropdown(description="Subculture?", options=['No', 'Yes'], style=style, layout=layout_full))

        display(widgets.HTML("<br>"), btn_generate)

btn_generate = widgets.Button(description="GENERATE & DOWNLOAD DOCUMENT", button_style='success', layout=widgets.Layout(width='100%', height='50px'))

# --- PART 5: GENERATION LOGIC ---
def on_click_generate(b):
    selection = platform_selector.value
    template_map = {
        'ScanRDI': 'ScanRDI OOS template.docx',
        'Celsis': 'Celsis OOS template.docx',
        'USP 71': 'USP71 OOS template.docx'
    }
    
    template_file = template_map[selection]
    
    if not os.path.exists(template_file):
        print(f"❌ Error: {template_file} not found. Please upload it to your folder!")
        return

    # Gather data based on widgets
    data = {
        'oos_id': w_oos_id.value,
        'client_name': w_client.value,
        'sample_id': w_sample_id.value,
        'sample_name': w_sample_name.value,
        'lot_number': w_lot.value,
        'test_date': w_test_date.value,
        # ScanRDI Specific Mappings 
        'prepper_name': w_prepper.value,
        'prepper_initial': w_prepper_init.value,
        'analyst_name': w_analyst.value,
        'analyst_initial': w_analyst_init.value,
        'changeover_name': w_changeover.value,
        'changeover_initial': w_changeover_init.value,
        'reader_name': w_reader.value,
        'reader_initial': w_reader_init.value,
        'scan_id': w_scan_id.value,
        'cr_suit': w_cr_suit.value,
        'bsc_id': w_bsc_id.value,
        'organism_morphology': w_org_morph.value,
    }

    try:
        doc = DocxTemplate(template_file)
        doc.render(data)
        save_history(data)
        
        output_name = f"{w_oos_id.value} {w_client.value} - {selection}.docx"
        doc.save(output_name)
        print(f"✅ Created: {output_name}")
        files.download(output_name)
    except Exception as e:
        print(f"❌ Error: {e}")

# Wire up the events
platform_selector.observe(render_form, names='value')
btn_generate.on_click(on_click_generate)

# Initial Display
display(platform_selector)
display(output_form)
render_form()
