import os
import json
from datetime import datetime, timedelta
import ipywidgets as widgets
from docxtpl import DocxTemplate
from IPython.display import display, clear_output

# --- PART 1: DATA & HISTORY ---
CURRENT_HISTORY_FILE = 'oos_wizard_history.json'

def load_history():
    if os.path.exists(CURRENT_HISTORY_FILE):
        try:
            with open(CURRENT_HISTORY_FILE, 'r') as f: return json.load(f)
        except: return {}
    return {}

def save_history(data_dict):
    with open(CURRENT_HISTORY_FILE, 'w') as f: 
        json.dump(data_dict, f)

defaults = load_history()
def get_val(key, fallback): 
    return defaults.get(key, fallback)

# --- PART 2: UI WIDGETS ---
style = {'description_width': 'initial'}
layout_full = widgets.Layout(width='98%')
layout_half = widgets.Layout(width='48%')

# Platform Selection (Left-side Sidebar logic)
platform_selector = widgets.ToggleButtons(
    options=['ScanRDI', 'Celsis', 'USP 71'],
    description='Select OS:',
    button_style='info'
)

# General Info (Common to all)
w_oos_id = widgets.Text(description="OOS Number:", value=get_val('oos_id', "OOS-250000"), style=style, layout=layout_full)
w_client = widgets.Text(description="Client Name:", value=get_val('client_name', ""), style=style, layout=layout_full)
w_sample_id = widgets.Text(description="Sample ID:", value=get_val('sample_id', ""), style=style, layout=layout_full)
w_test_date = widgets.Text(description="Test Date (DDMonYY):", value=datetime.now().strftime("%d%b%y"), style=style, layout=layout_full)

# ScanRDI Specific Widgets [cite: 181, 182]
w_prepper = widgets.Text(description="Prepper:", value=get_val('prepper_name', ""), style=style, layout=layout_half)
w_analyst = widgets.Text(description="Processor:", value=get_val('analyst_name', ""), style=style, layout=layout_half)
w_scan_id = widgets.Text(description="ScanRDI ID:", value=get_val('scan_id', ""), style=style, layout=layout_half)
w_org_morph = widgets.Text(description="Org Shape:", value=get_val('organism_morphology', ""), style=style, layout=layout_half)

# --- PART 3: DYNAMIC UI RENDERING ---
output_form = widgets.Output()

def render_form(change=None):
    with output_form:
        clear_output()
        selection = platform_selector.value
        display(widgets.HTML(f"<h2>{selection} Investigation Wizard</h2>"))
        
        # Section 1: General Information
        display(widgets.HTML("<h3>1. General Details</h3>"), w_oos_id, w_client, w_sample_id, w_test_date)
        
        # Section 2: Platform Specific Fields
        if selection == 'ScanRDI':
            display(widgets.HTML("<h3>2. ScanRDI Specifics</h3>"))
            display(widgets.HBox([w_prepper, w_analyst]))
            display(widgets.HBox([w_scan_id, w_org_morph]))
        elif selection == 'Celsis':
            display(widgets.HTML("<h3>2. Celsis Specifics</h3>"))
            display(widgets.Text(description="Luminometer ID:", style=style, layout=layout_full))
        elif selection == 'USP 71':
            display(widgets.HTML("<h3>2. USP 71 Specifics</h3>"))
            display(widgets.Text(description="Incubation Room:", style=style, layout=layout_full))
        
        display(widgets.HTML("<br>"), btn_generate)

btn_generate = widgets.Button(description="GENERATE & DOWNLOAD", button_style='success', layout=widgets.Layout(width='100%', height='50px'))

# --- PART 4: GENERATION LOGIC ---
def on_click_generate(b):
    selection = platform_selector.value
    template_file = f"{selection} OOS template.docx" [cite: 182]
    
    if not os.path.exists(template_file):
        print(f"❌ Error: {template_file} not found! Upload the template to GitHub.")
        return

    # Data Mapping [cite: 182, 184]
    data = {
        'oos_id': w_oos_id.value,
        'client_name': w_client.value,
        'sample_id': w_sample_id.value,
        'test_date': w_test_date.value,
        'prepper_name': w_prepper.value,
        'analyst_name': w_analyst.value,
        'scan_id': w_scan_id.value,
        'organism_morphology': w_org_morph.value
    }

    try:
        doc = DocxTemplate(template_file)
        doc.render(data)
        output_name = f"{w_oos_id.value}_{selection}.docx"
        doc.save(output_name)
        print(f"✅ Created: {output_name}")
        save_history(data)
    except Exception as e:
        print(f"❌ Error: {e}")

platform_selector.observe(render_form, names='value')
btn_generate.on_click(on_click_generate)

display(platform_selector, output_form)
render_form()
