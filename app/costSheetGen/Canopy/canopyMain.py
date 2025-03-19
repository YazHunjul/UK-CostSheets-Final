import streamlit as st
import zipfile
from io import BytesIO
from costSheetGen.Canopy import canopyExcel as CE
from costSheetGen.Canopy import canopyWord as CW
import time
from openpyxl import load_workbook
import io
import os
from costSheetGen.Canopy.canopyUtils import extract_canopy_prices, convert_formulas_to_values, run_excel_script
import pyperclip
import pandas as pd
import openpyxl
import math
from ..config import TEMPLATES
import json
from datetime import datetime


def get_initials(name):
    """Extract initials from a name"""
    if name:
        words = name.split()
        initials = ''.join(word[0].upper() for word in words if word)
        return initials
    return ''

def get_delivery_install_details(floor_name, key_suffix):
    """
    Collect delivery and installation details for a specific floor
    
    Args:
        floor_name: Name of the floor for labeling
        key_suffix: Unique suffix for Streamlit keys to prevent conflicts
    """
    with st.expander(f"🚚 Delivery & Installation Details - {floor_name}", expanded=False):
        st.markdown('<div class="section-header"><h3>📍 Location & Plant Hire</h3></div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([1, 3])
        with col1:
            delivery_lift_qty = st.number_input(
                "DELIVERY 1 x 7.5T TAIL LIFT",
                min_value=0,
                value=1,
                step=1,
                key=f"delivery_lift_qty_{key_suffix}"
            )

        with col2:
            delivery_locations = [
                "",
                "ABERDEEN 590",
                "ABINGDON 110",
                "ALDEBURGH 112",
                "ALDERSHOT 110",
                "ALNWICK 342",
                "ANDOVER 110",
                "ASHFORD 25",
                "AYLESBURY 86",
                "BANBURY 102",
                "BANGOR 324",
                "BARKING 32",
                "BARNET 55",
                "BARNSLEY 209",
                "BARNSTABLE 227",
                "BARROW-IN-FURNESS 348",
                "BASILDON 38",
                "BASINGSTOKE 82",
                "BATH 154",
                "BEDFORD 103",
                "BERWICK-UPON-TWEED 371",
                "BILLERICAY 37",
                "BIRKENHEAD 277",
                "BIRMINGHAM 168",
                "BLACKBURN 283",
                "BLACKPOOL 289",
                "BLANDFORD FORUM 144",
                "BODMIN 273",
                "BOGNOR REGIS 88",
                "BOLTON 259",
                "BOOTLE 272",
                "BOURNEMOUTH 140",
                "BRADFORD 234",
                "BRAINTREE 60",
                "BRIDGEND 205",
                "BRIDLINGTON 244",
                "BRIGHTON 68",
                "BRISTOL 157",
                "BUCKINGHAMSHIRE 109",
                "BURNLEY 296",
                "BURTON UPON TRENT 175",
                "BURY ST EDMUNDS 98",
                "CAMBRIDGE 85",
                "CANNOCK 175",
                "CANTERBURY 30",
                "CARDIFF 192",
                "CARLISLE 356",
                "CARMARTHEN 252",
                "CHELTENHAM 148",
                "CHESTER 268",
                "COVENTRY 146",
                "CHIPPENHAM 136",
                "COLCHESTER 78",
                "CORBY 128",
                "DARTMOUTH 245",
                "DERBY 178",
                "DONCASTER 203",
                "DORCHESTER 160",
                "DORKING 46",
                "DOVER 45",
                "DURHAM 299",
                "EASTBOURNE 57",
                "EASTLEIGH 109",
                "EDINBURGH 428",
                "ENFIELD 49",
                "EXETER 205",
                "EXMOUTH 207",
                "FELIXSTOWE 103",
                "GATWICK 44",
                "GLASGOW 456",
                "GLASTONBURY 164",
                "GLOUCESTER 151",
                "GRANTHAM 143",
                "GREAT YARMOUTH 147",
                "GRIMSBY 215",
                "GUILDFORD 59",
                "HARLOW 47",
                "HARROGATE 236",
                "HARTLEPOOL 286",
                "HASTINGS 40",
                "HEXHAM 325",
                "HEREFORD 184",
                "HIGH WYCOMBE 80",
                "HIGHBRIDGE 187",
                "HONITON 190",
                "HORSHAM 55",
                "HOUNSLOW 55",
                "HUDDERSFIELD 239",
                "HULL 247",
                "HUNTINGDON 94",
                "INVERNESS 619",
                "IPSWICH 94",
                "IRELAND",
                "KENDAL 321",
                "KETTERING 127",
                "KIDDERMINSTER 179",
                "KILMARNOCK 449",
                "KINGSTON UPON HULL 220",
                "KINGSTON UPON THAMES 52",
                "LANCASTER 290",
                "LAUNCESTON 251",
                "LEAMINGTON SPA 146",
                "LEEDS 231",
                "LEICESTER 151",
                "LEIGH ON SEA 45",
                "LEWISHAM 29",
                "LINCOLN 179",
                "LIVERPOOL 258",
                "LLANDUDNO 309",
                "LONDON in FORS GOLD (varies)",
                "LUTON 80",
                "MABLETHORPE 182",
                "MACCLESFIELD 244",
                "MANCHESTER 251",
                "MARGATE 46",
                "MIDDLESBROUGH 286",
                "MILFORD HAVEN 289",
                "MILTON KEYNES 101",
                "MORPETH 327",
                "NANTWICH 232",
                "NEWBURY 101",
                "NEWCASTLE 308",
                "NEWPORT 178",
                "NEWQUAY 178",
                "NORTHAMPTON 116",
                "NORTHUMBERLAND 341",
                "NORWICH 136",
                "NOTTINGHAM 177",
                "OKEHAMPTON 232",
                "OXFORD 106",
                "PENRITH 316",
                "PENZANCE 318",
                "PERTH 477",
                "PETERBOROUGH 124",
                "PETERSFIELD 87",
                "PETWORTH 71",
                "PLYMOUTH 247",
                "PONTEFRACT 221",
                "POOLE 144",
                "PORTSMOUTH 102",
                "READING 88",
                "REIGATE 39",
                "RINGWOOD 130",
                "ROSS-ON-WYE 171",
                "ROTHERHAM 203",
                "SALISBURY 120",
                "SCARBOROUGH 277",
                "SCUNTHORPE 204",
                "SHEFFIELD 205",
                "SHREWSBURY 207",
                "SHROPSHIRE 218",
                "SLOUGH 72",
                "SOUTH SHIELDS 310",
                "SOUTHAMPTON 112",
                "SOUTHEND 52",
                "SOUTHPORT 279",
                "SPALDING 143",
                "ST ALBANS 62",
                "ST IVES 317",
                "STAFFORD 187",
                "STAINES 61",
                "STEVENAGE 72",
                "STIRLING 445",
                "STOCKPORT 257",
                "STOCKTON 278",
                "STOKE-ON-TRENT 205",
                "STRATFORD UPON AVON 151",
                "SUNDERLAND 309",
                "SWINDON 121",
                "TAMWORTH 180",
                "TAUNTON 185",
                "TELFORD 193",
                "TILBURY 34",
                "TORQUAY 227",
                "TUNBRIDGE WELLS 26",
                "UXBRIDGE 74",
                "WAKEFIELD 214",
                "WARMISTER 137",
                "WARWICK 148",
                "WATFORD 67",
                "WELSHPOOL 238",
                "WEMBLEY 55",
                "WEYMOUTH 173",
                "WHITBY 282",
                "WIGAN 252",
                "WINCANTON 149",
                "WINCHESTER 100",
                "WOKING 60",
                "WOLVERHAMPTON 175",
                "WORCESTER 160",
                "WREXHAM 250",
                "YEOVIL 163",
                "YORK 243"
            ]
            
            delivery_location = st.selectbox(
                "SELECT LOCATION",
                options=delivery_locations,
                help="Select delivery location",
                key=f"delivery_location_{key_suffix}"
            )

        # Plant Hire Section
        plant_hires = st.multiselect(
            "Select Plant Hires (max 2)",
            options=["Plant Hire 1", "Plant Hire 2"],
            max_selections=2,
            key=f"plant_hires_{key_suffix}"
        )

        plant_selections = {}
        quantities = {}
        for plant in plant_hires:
            col1, col2 = st.columns([3, 1])
            with col1:
                plant_selections[plant] = st.selectbox(
                    f"PLANT SELECTION (weekly) for {plant}",
                    ["", "SL10 GENIE", "EXTENSION FORKS", "2.5M COMBI LADDER", 
                     "1.5M PODIUM", "3M TOWER", "COMBI LADDER", "PECO LIFT", 
                     "3M YOUNGMAN BOARD", "GS1930 SCISSOR LIFT", "4-6 SHERASCOPIC", 
                     "7-9 SHERASCOPIC"],
                    key=f"plant_selection_{plant}_{key_suffix}"
                )
            
            with col2:
                if plant_selections[plant]:
                    quantities[plant] = st.number_input(
                        "QTY",
                        min_value=1,
                        value=1,
                        step=1,
                        key=f"qty_{plant}_{key_suffix}"
                    )

        # Installation details in columns
        col1, col2 = st.columns(2)
        
        with col1:
            strip_out = st.number_input("STRIP OUT (PER DAY)", min_value=0.0, value=0.0, step=1.0, key=f"strip_out_{key_suffix}")
            #consumables = st.number_input("CONSUMABLES 15(P) + 19(H)", min_value=0.0, value=0.0, step=1.0, key=f"consumables_{key_suffix}")
            installation_normal = st.number_input("INSTALLATION NORMAL HOURS", min_value=0.0, value=0.0, step=1.0, key=f"installation_normal_{key_suffix}")
            installation_after = st.number_input("INSTALLATION AFTER HOURS", min_value=0.0, value=0.0, step=1.0, key=f"installation_after_{key_suffix}")
            wall_cladding = st.number_input("WALL CLADDING INSTALLATION", min_value=0.0, value=0.0, step=1.0, key=f"wall_cladding_{key_suffix}")
            
        with col2:
            overnight_expenses = st.number_input("OVERNIGHT/TRAVEL EXPENSES", min_value=0.0, value=0.0, step=1.0, key=f"overnight_expenses_{key_suffix}")
            test_commission = st.number_input("TEST & COMMISSION", min_value=0.0, value=0.0, step=1.0, key=f"test_commission_{key_suffix}")
            gas_interlock = st.number_input("GAS INTERLOCK (INSTALLED)", min_value=0.0, value=0.0, step=1.0, key=f"gas_interlock_{key_suffix}")
            co_sensor = st.number_input("CO SENSOR (SOLID FUEL)", min_value=0.0, value=0.0, step=1.0, key=f"co_sensor_{key_suffix}")
            co2_sensor = st.number_input("CO2 SENSOR (DCK)", min_value=0.0, value=0.0, step=1.0, key=f"co2_sensor_{key_suffix}")
            bms_interface = st.number_input("BMS FAULT INTERFACE", min_value=0.0, value=0.0, step=1.0, key=f"bms_interface_{key_suffix}")

        return {
            "delivery_location": delivery_location,
            "delivery_lift_qty": delivery_lift_qty,
            "plant_hires": plant_selections,
            "quantities": quantities,
            "strip_out": strip_out,
            # "consumables": consumables,
            "installation_normal": installation_normal,
            "installation_after": installation_after,
            "wall_cladding": wall_cladding,
            "overnight_expenses": overnight_expenses,
            "test_commission": test_commission,
            "gas_interlock": gas_interlock,
            "co_sensor": co_sensor,
            "co2_sensor": co2_sensor,
            "bms_interface": bms_interface,
        }

def load_session_state(json_data):
    """Load session state from JSON data"""
    try:
        data = json.loads(json_data)
        
        # Only restore session state values, skip general info
        if 'session_state' in data:
            # First clear existing session state (except system keys)
            keys_to_keep = {k for k in st.session_state.keys() if k.startswith('_')}
            for k in list(st.session_state.keys()):
                if k not in keys_to_keep and not k.startswith('customer'):  # Skip customer and other genInfo fields
                    del st.session_state[k]
            
            # Then restore session state values
            for key, value in data['session_state'].items():
                if not key.startswith('_') and not key.startswith('customer'):  # Skip system keys and customer fields
                    st.session_state[key] = value
        
        # Force a rerun to ensure all widgets update
        st.rerun()
        return True
        
    except Exception as e:
        st.error(f"Error loading project data: {str(e)}")
        return False

def save_session_state():
    """Save current session state to a JSON file"""
    # Get all relevant session state data
    save_data = {
        'session_state': {
            k: v for k, v in st.session_state.items() 
            if not k.startswith('_') and not k.startswith('customer')  # Exclude internal keys and customer fields
        },
        'timestamp': datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    }
    
    # Convert to JSON string
    json_str = json.dumps(save_data, indent=2)
    
    return json_str

def add_save_load_section():
    """Add save/load UI section"""
    with st.expander("💾 Save/Load Project Data"):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### Save Project")
            if st.button("Generate Project File"):
                json_str = save_session_state()
                project_name = st.session_state.get('projectNum', 'project')
                
                st.download_button(
                    label="⬇️ Download Project File",
                    data=json_str,
                    file_name=f"{project_name}_data.json",
                    mime="application/json"
                )
        
        with col2:
            st.markdown("### Load Project")
            uploaded_file = st.file_uploader(
                "Upload Project File", 
                type=['json'],
                help="Upload a previously saved project file"
            )
            
            if uploaded_file is not None:
                json_str = uploaded_file.read().decode()
                if load_session_state(json_str):
                    st.success("Project data loaded successfully!")
                    st.rerun()  # Refresh the page to show loaded data

def main(genInfo):
    st.markdown('<hr>', unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center;'>Canopy Cost Sheet</h2>", unsafe_allow_html=True)
    
    # Add save/load section at the top
    add_save_load_section()
    
    st.markdown("---")
    
    # Add anchors for main sections
    st.markdown("<div id='general_info'></div>", unsafe_allow_html=True)
    
    # Get Kitchen Count first
    num_kitchens = st.number_input("Enter Number of Levels", min_value=1, key='num_kitchens')
    kitchen_info = []

    # Navigation sidebar
    with st.sidebar:
        st.markdown("### 🔍 Quick Navigation")
        
        # Main sections navigation
        st.markdown("#### Main Sections:")
        st.markdown("[📋 General Information](#general_info)")
        st.markdown("[📦 Delivery & Installation](#delivery)")
        st.markdown("[💾 Generate Files](#generate)")
        
        # Floor-Area navigation
        st.markdown("#### Jump to Floor - Area:")
        
        for i in range(num_kitchens):
            kitchen_name = st.session_state.get(f'kitchen_name_{i}', f'Floor {i + 1}')
            num_floors = st.session_state.get(f'floors_input_{i}', 1)
            
            if i > 0:
                st.markdown("---")
            
            for floor in range(num_floors):
                floor_name = st.session_state.get(f'floor_name_{i}_{floor}', f'Area {floor + 1}')
                
                # Get canopy info for this floor
                canopy_summary = []
                canopy_key = f'canopies_input_{i}_{floor}'
                if canopy_key in st.session_state:
                    num_canopies = st.session_state[canopy_key]
                    models = {}
                    
                    for c in range(num_canopies):
                        model_key = f'model_{i}_{floor}_{c}'
                        if model_key in st.session_state:
                            model = st.session_state[model_key]
                            models[model] = models.get(model, 0) + 1
                    
                    for model, count in models.items():
                        if model:
                            canopy_summary.append(f"{count}x {model}")
                
                summary_text = f"{kitchen_name} - {floor_name}"
                if canopy_summary:
                    summary_text += f"\n({', '.join(canopy_summary)})"
                
                st.markdown(f"[{summary_text}](#floor_{i}_{floor})")

        # Add separator for Canopy Details
        st.markdown("---")
        
        # Make Canopy Details expandable
        with st.expander("🏗️ Current Canopy Details", expanded=False):
            has_canopies = False
            
            # Get current canopy details from session state
            for i in range(num_kitchens):
                num_floors = st.session_state.get(f'floors_input_{i}', 1)
                for floor in range(num_floors):
                    canopy_key = f'canopies_input_{i}_{floor}'
                    if canopy_key in st.session_state:
                        num_canopies = st.session_state[canopy_key]
                        for c in range(num_canopies):
                            kitchen_name = st.session_state.get(f'kitchen_name_{i}', f'Floor {i + 1}')
                            floor_name = st.session_state.get(f'floor_name_{i}_{floor}', f'Area {floor + 1}')
                            
                            # Use collapsible markdown for each canopy
                            has_canopies = True
                            st.markdown(f"<details><summary><b>{kitchen_name} - {floor_name} (Canopy {c + 1})</b></summary>", unsafe_allow_html=True)
                            
                            # Define all possible fields with their correct session state keys
                            fields = {
                                'itemNum': ('📌 Reference Number', 'Reference Number'),
                                'model': ('🔧 Model', 'Model'),
                                'config': ('⚙️ Configuration', 'Configuration'),
                                'width': ('📏 Width', 'Width'),
                                'length': ('📏 Length', 'Length'),
                                'height': ('📏 Height', 'Height'),
                                'section': ('🔢 Section', 'Section'),
                                'flowRate': ('💨 Flow Rate', 'Flow Rate'),
                                'light_type': ('💡 Light Type', 'Light Type'),
                                'cladding': ('🧱 Wall Cladding', 'Wall Cladding'),
                                'control_panel': ('🎛️ Control Panel', 'Control Panel'),
                                'WW_pods': ('💧 WW Pods', 'WW Pods'),
                                'WW_pods_quantity': ('🔢 WW Pods Quantity', 'WW Pods Quantity'),
                                'pipework': ('🔧 Pipework', 'Pipework')
                            }
                            
                            # Track filled and missing fields
                            filled_fields = []
                            missing_fields = []
                            
                            # Get model to check for CMWI/CMWF fields
                            model = st.session_state.get(f'model_{i}_{floor}_{c}')
                            
                            # Check each field
                            for field_key, (icon, label) in fields.items():
                                value = st.session_state.get(f'{field_key}_{i}_{floor}_{c}')
                                
                                # Skip CMWI/CMWF specific fields if not applicable
                                if field_key in ['control_panel', 'WW_pods', 'WW_pods_quantity', 'pipework']:
                                    if model not in ['CMWI', 'CMWF']:
                                        continue
                                
                                # Consider a field filled if it has any value (including 0) or is False
                                if value is not None and value != '':
                                    # Format the value display
                                    if field_key in ['width', 'length', 'height']:
                                        display_value = f"{value} mm"
                                    elif field_key == 'flowRate':
                                        display_value = f"{value} m³/s"
                                    elif field_key == 'cladding':
                                        display_value = 'Yes' if value else 'No'
                                    else:
                                        display_value = str(value)
                                    
                                    filled_fields.append((icon, label, display_value))
                                else:
                                    missing_fields.append((icon, label))
                            
                            # Show filled fields
                            if filled_fields:
                                st.markdown("<div style='margin-left: 20px'>", unsafe_allow_html=True)
                                st.markdown("**✅ Filled Fields:**")
                                for icon, label, value in filled_fields:
                                    st.markdown(f"{icon} {label}: {value}")
                                st.markdown("</div>", unsafe_allow_html=True)
                            
                            # Show missing fields
                            if missing_fields:
                                st.markdown("<div style='margin-left: 20px'>", unsafe_allow_html=True)
                                st.markdown("**❌ Missing Fields:**")
                                for icon, label in missing_fields:
                                    st.markdown(f"{icon} {label}")
                                st.markdown("</div>", unsafe_allow_html=True)
                            
                            st.markdown("</details>", unsafe_allow_html=True)
                            st.markdown("<br>", unsafe_allow_html=True)  # Add some spacing
            
            if not has_canopies:
                st.markdown("No canopies created yet.")

        # Add final separator before Project Summary
        st.markdown("---")
        
        # Project Summary
        st.markdown("### 📊 Project Summary")
        
        # Check for missing general info
        missing_info = []
        required_fields = {
            'projectName': 'Project Name',
            'projectNum': 'Project Number',
            'customer': 'Customer',
            'sales_contact': 'Sales Contact',
            'estimator': 'Estimator',
            'location': 'Location',
            'address': 'Address'
        }
        
        for field, display_name in required_fields.items():
            if not genInfo.get(field):
                missing_info.append(display_name)
        
        if missing_info:
            st.markdown("#### ⚠️ Missing Information:")
            for field in missing_info:
                st.markdown(f"- {field}")
        
        # Canopy Summary
        st.markdown("#### 🏗️ Canopy Overview:")
        total_canopies = 0
        model_counts = {}
        
        # Iterate through session state to count canopies
        for i in range(num_kitchens):
            num_floors = st.session_state.get(f'floors_input_{i}', 1)
            for floor in range(num_floors):
                canopy_key = f'canopies_input_{i}_{floor}'
                if canopy_key in st.session_state:
                    num_canopies = st.session_state[canopy_key]
                    total_canopies += num_canopies
                    
                    # Count models
                    for c in range(num_canopies):
                        model_key = f'model_{i}_{floor}_{c}'
                        if model_key in st.session_state:
                            model = st.session_state[model_key]
                            if model:  # Only count if model is selected
                                model_counts[model] = model_counts.get(model, 0) + 1
        
        st.markdown(f"Total Canopies: {total_canopies}")
        if model_counts:
            st.markdown("Models:")
            for model, count in model_counts.items():
                st.markdown(f"- {model}: {count}")
        
        # Add any warnings about missing canopy data
        st.markdown("#### ⚠️ Missing Canopy Data:")
        missing_data = False
        
        # Check session state for missing canopy data
        for i in range(num_kitchens):
            kitchen_name = st.session_state.get(f'kitchen_name_{i}', f'Floor {i + 1}')
            num_floors = st.session_state.get(f'floors_input_{i}', 1)
            
            for floor in range(num_floors):
                floor_name = st.session_state.get(f'floor_name_{i}_{floor}', f'Area {floor + 1}')
                canopy_key = f'canopies_input_{i}_{floor}'
                
                if canopy_key in st.session_state:
                    num_canopies = st.session_state[canopy_key]
                    for c in range(num_canopies):
                        missing_fields = []
                        
                        # Check all required fields
                        field_checks = {
                            'itemNum': 'Reference Number',
                            'model': 'Model',
                            'config': 'Configuration',
                            'width': 'Width',
                            'length': 'Length',
                            'section': 'Section',
                            'height': 'Height',
                            'flowRate': 'Flow Rate',
                            'light_type': 'Light Type'
                        }
                        
                        for field, display_name in field_checks.items():
                            field_key = f'{field}_{i}_{floor}_{c}'
                            value = st.session_state.get(field_key)
                            if not value and value != 0:  # Check for empty or None, but allow 0
                                missing_fields.append(display_name)
                        
                        # Check light quantity if light type is selected and not a strip light
                        light_type = st.session_state.get(f'light_type_{i}_{floor}_{c}')
                        if light_type and not any(x in light_type for x in ['L6', 'L12', 'L18']):
                            light_qty_key = f'light_qty_{i}_{floor}_{c}'
                            if not st.session_state.get(light_qty_key):
                                missing_fields.append('Light Quantity')
                        
                        # Check CMWI/CMWF specific fields
                        model = st.session_state.get(f'model_{i}_{floor}_{c}')
                        if model in ['CMWI', 'CMWF']:
                            cmwi_fields = {
                                'control_panel': 'Control Panel',
                                'WW_pods': 'WW Pods',
                                'WW_pods_quantity': 'WW Pods Quantity',
                                'pipework': 'Pipework'
                            }
                            for field, display_name in cmwi_fields.items():
                                field_key = f'{field}_{i}_{floor}_{c}'
                                if not st.session_state.get(field_key):
                                    missing_fields.append(display_name)
                        
                        if missing_fields:
                            missing_data = True
                            st.markdown(f"**{kitchen_name} - {floor_name} (Canopy {c + 1}):**")
                            st.markdown(f"- Missing: {', '.join(missing_fields)}")
        
        if not missing_data:
            st.markdown("✅ All required canopy data complete")

        # Add CMWF/CMWI specific summary
        cmwf_count = 0
        cmwi_count = 0
        ww_pods_total = 0
        
        for i in range(num_kitchens):
            num_floors = st.session_state.get(f'floors_input_{i}', 1)
            for floor in range(num_floors):
                canopy_key = f'canopies_input_{i}_{floor}'
                if canopy_key in st.session_state:
                    num_canopies = st.session_state[canopy_key]
                    for c in range(num_canopies):
                        model = st.session_state.get(f'model_{i}_{floor}_{c}')
                        if model == 'CMWF':
                            cmwf_count += 1
                            pods_qty = st.session_state.get(f'WW_pods_quantity_{i}_{floor}_{c}', 0)
                            ww_pods_total += pods_qty
                        elif model == 'CMWI':
                            cmwi_count += 1
                            pods_qty = st.session_state.get(f'WW_pods_quantity_{i}_{floor}_{c}', 0)
                            ww_pods_total += pods_qty
        
        if cmwf_count > 0 or cmwi_count > 0:
            st.markdown("#### 💧 Water Wash Details:")
            if cmwf_count > 0:
                st.markdown(f"- CMWF Canopies: {cmwf_count}")
            if cmwi_count > 0:
                st.markdown(f"- CMWI Canopies: {cmwi_count}")
            st.markdown(f"- Total WW Pods: {ww_pods_total}")

        # Add delivery and installation checks
        st.markdown("#### 🚚 Delivery & Installation Status:")
        missing_delivery = False
        
        for i in range(num_kitchens):
            kitchen_name = st.session_state.get(f'kitchen_name_{i}', f'Floor {i + 1}')
            num_floors = st.session_state.get(f'floors_input_{i}', 1)
            
            for floor in range(num_floors):
                floor_name = st.session_state.get(f'floor_name_{i}_{floor}', f'Area {floor + 1}')
                key_suffix = f"kitchen_{i}_floor_{floor}"
                
                # Check required delivery & installation fields
                missing_fields = []
                
                # Check delivery fields
                location = st.session_state.get(f'location_{key_suffix}')
                plant_hire = st.session_state.get(f'plant_hire_{key_suffix}')
                plant_selection = st.session_state.get(f'plant_selection_{key_suffix}')
                strip_out = st.session_state.get(f'strip_out_{key_suffix}')
                overnight = st.session_state.get(f'overnight_{key_suffix}')
                normal_hours = st.session_state.get(f'normal_hours_{key_suffix}')
                after_hours = st.session_state.get(f'after_hours_{key_suffix}')
                wall_cladding_install = st.session_state.get(f'wall_cladding_install_{key_suffix}')
                gas_interlock = st.session_state.get(f'gas_interlock_{key_suffix}')
                co_sensor = st.session_state.get(f'co_sensor_{key_suffix}')
                co2_sensor = st.session_state.get(f'co2_sensor_{key_suffix}')
                bms_interface = st.session_state.get(f'bms_interface_{key_suffix}')
                
                # Add missing fields to list with more descriptive labels
                if not location:
                    missing_fields.append("Location")
                if plant_hire and not plant_selection:  # Only check plant selection if plant hire is selected
                    missing_fields.append("Plant Selection")
                if strip_out is None:  # Check if strip out value is set
                    missing_fields.append("Strip Out Cost")
                if overnight is None:
                    missing_fields.append("Overnight/Travel Expenses")
                if normal_hours is None:
                    missing_fields.append("Installation Normal Hours")
                if after_hours is None:
                    missing_fields.append("Installation After Hours")
                if wall_cladding_install is None:
                    missing_fields.append("Wall Cladding Installation")
                if gas_interlock is None:
                    missing_fields.append("Gas Interlock")
                if co_sensor is None:
                    missing_fields.append("CO Sensor")
                if co2_sensor is None:
                    missing_fields.append("CO2 Sensor")
                if bms_interface is None:
                    missing_fields.append("BMS Interface")
                
                if missing_fields:
                    missing_delivery = True
                    st.markdown(f"**{kitchen_name} - {floor_name}:**")
                    st.markdown("Missing:")
                    for field in missing_fields:
                        st.markdown(f"❌ {field}")
                    st.markdown("---")
        
        if not missing_delivery:
            st.markdown("✅ All delivery & installation details complete")

    # Main form content
    for i in range(num_kitchens):
        st.markdown(f"<div id='kitchen_{i}'></div>", unsafe_allow_html=True)
        kitchen_name = st.text_input(f"Enter Level {i + 1} Name", key=f'kitchen_name_{i}')
        if kitchen_name:
            kitchen_data = {
                "kitchen_name": kitchen_name,
                "floors": []
            }

            with st.expander(f'{kitchen_name.title()} Floor Information', expanded=False):
                num_floors = st.number_input(
                    f"Enter the number of areas in {kitchen_name} Floor", 
                    min_value=1, 
                    key=f'floors_input_{i}'
                )
                for floor in range(num_floors):
                    # Add anchor for floor
                    st.markdown(f"<div id='floor_{i}_{floor}'></div>", unsafe_allow_html=True)
                    floor_name = st.text_input(
                        f"Enter area {floor + 1} Name", 
                        key=f'floor_name_{i}_{floor}'
                    )
                    if floor_name:
                        # Create a dictionary for this floor
                        floor_data = {
                            "floor_name": floor_name,
                            "canopies": []
                        }

                        num_canopies = st.number_input(
                            f"Enter Number of Canopies in {floor_name}",
                            min_value=1, 
                            key=f'canopies_input_{i}_{floor}'
                        )
                        for canopy in range(num_canopies):
                            st.markdown(f"<h4 style='text-align:center;'>Canopy {canopy + 1} - Floor: ({floor_name})</h4>", unsafe_allow_html=True)

                            coll1, coll2, coll3, coll4 = st.columns(4)
                            with coll1:
                                item_number = st.text_input('Reference Number', key=f'itemNum_{i}_{floor}_{canopy}')
                                length = st.number_input("Length", min_value=0, key=f'length_{i}_{floor}_{canopy}')
                                section = st.number_input('Sections', min_value=0, key=f'section_{i}_{floor}_{canopy}')

                                # Get configuration value from session state since it's defined in coll2
                                config_key = f'config_{i}_{floor}_{canopy}'
                                if config_key in st.session_state and st.session_state[config_key] == "ISLAND":
                                    st.caption(f"Total sections for ISLAND configuration: {section * 2}")

                                light_type = st.selectbox(
                                    'Light Type',
                                    ['','LED STRIP L6 Inc DALI', 'LED STRIP L12 inc DALI', 'LED STRIP L18 Inc DALI', 'Small LED Spots inc DALI', 'LARGE LED Spots inc DALI'],
                                    key=f'light_type_{i}_{floor}_{canopy}'
                                )
                                
                                # Check if it's a strip light (L6, L12, L18)
                                is_strip_light = any(x in light_type for x in ['L6', 'L12', 'L18'])
                                
                                # Set quantity to sections for strip lights, otherwise show input
                                light_quantity = None
                                if light_type:  # Only if a light type is selected
                                    if is_strip_light:
                                        light_quantity = section
                                        st.text(f"Quantity: {section} (Based on sections)")
                                    else:
                                        light_quantity = st.number_input(
                                            'Light Quantity', 
                                            min_value=0, 
                                            key=f'light_qty_{i}_{floor}_{canopy}'
                                        )

                            with coll2:
                                configuration = st.selectbox('Configuration', ['WALL', "ISLAND"], key=f'config_{i}_{floor}_{canopy}')
                                width = st.number_input("Width", min_value=0, key=f'width_{i}_{floor}_{canopy}')
                                special_works = st.multiselect(
                                    'Special Works (Max 2)',
                                    ['ROUND CORNERS', 'CUT OUT', 'CASTELLE LOCKING ', 'HEADER DUCT S/S', 'HEADER DUCT ', 'PAINT FINSH', 'UV ON DEMAND', 'E/over for emergency strip light', 'E/over for small emer. spot light', 'E/over for large emer. spot light', 'COLD MIST ON DEMAND', 'CMW  PIPEWORK HWS/CWS', 'CANOPY GROUND SUPPORT', ' 2nd EXTRACT PLENUM', 'SUPPLY AIR PLENUM', 'CAPTUREJET PLENUM', 'COALESCER'],
                                    key=f'specialWorks_{i}_{floor}_{canopy}',
                                    max_selections=2
                                )
                                
                                # Warn if trying to select more than 2
                                if len(special_works) > 2:
                                    st.warning("Only the first 2 special works will be included")
                                    special_works = special_works[:2]
                                
                                # Initialize special works dictionary
                                special_works_dict = {}
                                
                                # For each selected special work (max 2), add a quantity input
                                for work in special_works:
                                    quantity = st.number_input(
                                        f'{work} Quantity',
                                        min_value=1,
                                        value=1,
                                        key=f'specialWorks_qty_{i}_{floor}_{canopy}_{work}'
                                    )
                                    special_works_dict[work] = quantity

                            # Initialize cladding variables with defaults
                            cladding_height = None
                            cladding_width = None
                            description = None

                            with coll3:
                                model = st.selectbox(
                                    'Model', 
                                    ['KVF', 'KVX-M', "KVI", "UVX", "UVX-M", "UVI", "UVF", "UV-C POD", "CMWI", "CMWF", "CXW", "CXW-M", "KVV"], 
                                    key=f'model_{i}_{floor}_{canopy}'
                                )
                                height = st.number_input("Height", min_value=0, value=555, key=f'height_{i}_{floor}_{canopy}')
                                cladding = st.selectbox(
                                    "Wall Cladding",
                                    ['', '2M² (HFL)'],
                                    key=f'cladding_{i}_{floor}_{canopy}'
                                )
                                if cladding:
                                    cladding_height = st.number_input("Cladding Height", key=f'cladding_Height{i}_{floor}_{canopy}', min_value=0)
                                    cladding_width = st.number_input("Cladding Length", key=f'CladdingLength_{i}_{floor}_{canopy}', min_value=0)
                                    description = st.multiselect('Cladding Description', ['','Rear', 'Left', "Right" ], key=f'cladding_desc_{i}_{floor}_{canopy}')

                            # Initialize CMWI/CMWF specific variables with defaults
                            control_panel = None
                            WW_pods = None
                            CWS_HWS_pipework = None
                            WW_pods_quantity = 0
                            with coll4:
                                flowrate = st.number_input('Enter Flow Rate', min_value=0.0, key=f'flowRate_{i}_{floor}_{canopy}')
                                if model in ['CMWI', 'CMWF']:
                                    control_panel = st.selectbox('Select Control Panel', ['CP1S', 'CP2S', 'CP3S', 'CP4S'], key=f'CP_{i}_{floor}_{canopy}')
                                    WW_pods = st.selectbox("W/W Pods", ['1000-S', '1500-S', '2000-S', '2500-S', '3000-S', '1000-D', '1500-D', '2000-D', '2000-D', '2500-D', '3000-D'], key=f'WW_{i}_{floor}_{canopy}')
                                    
                                    if WW_pods:  # Only show quantity if a W/W pod is selected
                                        WW_pods_quantity = st.number_input(
                                        f"{WW_pods} Quantity",
                                        min_value=0,
                                        value=0,
                                        step=1,
                                        key=f'WW_qty_{i}_{floor}_{canopy}'
                                    ) 

                                        
                                    CWS_HWS_pipework = st.selectbox("CWS/HWS Pipework", [1,2,3,4,5], key=f'pipework_{i}_{floor}_{canopy}')

                            # Create a dictionary for this canopy
                            canopy_data = {
                                'item_number': item_number,
                                'model': model,
                                'configuration': configuration,
                                'section': section,
                                'height': height,
                                'width': width,
                                'length': length,
                                'lights': light_type,
                                'light_quantity': light_quantity,
                                'flowrate': flowrate,
                                'specialWorks': special_works_dict,
                                'wallCladding': cladding,
                                'control_panel': control_panel,
                                'WW_pods': WW_pods,
                                'pipework': CWS_HWS_pipework,
                                'cladding_width': cladding_width,
                                'cladding_height': cladding_height,
                                'cladding_desc': description,
                                'WW_pods_quantity': WW_pods_quantity
                            }

                            # Append canopy data to the floor
                            floor_data["canopies"].append(canopy_data)

                        # Append floor data to the kitchen
                        kitchen_data["floors"].append(floor_data)

            # Append kitchen data to the main list
            kitchen_info.append(kitchen_data)

    # Delivery & Installation Section
    st.markdown("<div id='delivery'></div>", unsafe_allow_html=True)
    st.markdown("## 🚚 Delivery & Installation Details")
    
    # Iterate through all kitchens and their floors
    for kitchen_idx, kitchen in enumerate(kitchen_info):
        for floor_idx, floor in enumerate(kitchen['floors']):
            floor_name = floor['floor_name']
            kitchen_name = kitchen['kitchen_name']
            
            # Get delivery & installation details for this floor
            delivery_install_data = get_delivery_install_details(
                floor_name=f"{kitchen_name} - {floor_name}",
                key_suffix=f"kitchen_{kitchen_idx}_floor_{floor_idx}"
            )
            
            # Store delivery & installation data including P182 value
            floor['delivery_install_data'] = delivery_install_data

    st.markdown('<hr>', unsafe_allow_html=True)

    # Add Email Summary Section
    col1, col2 = st.columns([3, 1])
    with col1:
        generate_email = st.checkbox("Generate Email Summary", value=False)
    
    # Email settings if enabled
    if generate_email:
        with st.expander("📧 Email Summary Settings", expanded=True):
            st.markdown("### Email Template Settings")
            
            col1, col2 = st.columns(2)
            with col1:
                email_tone = st.selectbox(
                    "Email Style",
                    [
                        "Professional",
                        "Friendly Professional",
                        "Casual Professional",
                        "Enthusiastic",
                        "Direct and Brief"
                    ],
                    index=0
                )
            
            with col2:
                email_focus = st.selectbox(
                    "Email Focus",
                    [
                        "Standard Proposal",
                        "Building Relationship",
                        "Quick Update",
                        "Technical Details",
                        "Project Timeline"
                    ],
                    index=0
                )
            
            # Split additional notes and closing message
            col1, col2 = st.columns(2)
            with col1:
                additional_notes = st.text_area(
                    "Project Context",
                    placeholder="Any specific details about the project or client relationship...",
                    height=100
                )
            
            with col2:
                closing_message = st.text_area(
                    "Personal Touch",
                    placeholder="e.g., 'Let's grab coffee to discuss this further' or 'Looking forward to our site visit next week'",
                    height=100
                )

            # Store email preferences in genInfo
            genInfo.update({
                'generate_email': generate_email,
                'email_tone': email_tone,
                'email_focus': email_focus,
                'additional_notes': additional_notes,
                'closing_message': closing_message
            })

            # Create container for email content
            email_container = st.container()

            # Separate button for email generation
            if st.button("Generate Email Summary"):
                scope_work = CW.scope_of_works({'kitchens': kitchen_info})
                email_summary = CW.generate_email_summary(genInfo, kitchen_info, scope_work)
                if email_summary:
                    with email_container:
                        st.markdown("### Generated Email Summary")
                        text_area = st.text_area(
                            "Email Content",
                            value=email_summary,
                            height=300,
                            key="email_text_area"
                        )
                        
                        col1, col2 = st.columns([1, 4])
                        with col1:
                            if st.button("📋 Copy"):
                                pyperclip.copy(text_area)
                                st.success("Copied!")

    st.markdown('<hr>', unsafe_allow_html=True)

    # Document generation section
    st.markdown("<div id='generate'></div>", unsafe_allow_html=True)
    st.markdown("## 💾 Generate Files")
    st.markdown("### Step 1: Generate and Download Excel")
    if st.button("Generate Excel File"):
        try:
            excel_bytes = CE.generate_sheet(kitchen_info, genInfo)
            if excel_bytes is None:
                st.error("Error generating Excel sheet: No data returned.")
                return
            
            st.download_button(
                label="⬇️ Download Excel File",
                data=excel_bytes.getvalue(),
                file_name=f"{genInfo['projectNum']} Cost Sheet {genInfo['date']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.info("1. Download the Excel file\n2. Open it in Excel\n3. Let Excel calculate the values\n4. Make Sure To Save the file and upload below")
            
        except Exception as e:
            st.error(f"Error generating Excel: {str(e)}")
    
    # Upload section
    st.markdown("### Step 2: Upload Processed Excel")
    uploaded_file = st.file_uploader("Upload Excel file after calculations", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            st.write("Loading Excel file...")
            # Create a copy of the uploaded file in memory
            excel_data = BytesIO(uploaded_file.getvalue())
            
            # Load workbook with data_only=True to get calculated values
            wb_data = openpyxl.load_workbook(excel_data, data_only=True)
            
            # Initialize total P182
            total_p182 = 0

            # Initialize grouped_canopy_data first
            grouped_canopy_data = {}
            for kitchen in kitchen_info:
                kitchen_name = kitchen['kitchen_name']
                grouped_canopy_data[kitchen_name] = {}
                for floor in kitchen['floors']:
                    floor_name = floor['floor_name']
                    floor['p182_value'] = 0.0  # Initialize p182_value in the floor data
                    grouped_canopy_data[kitchen_name][floor_name] = {
                        'canopies': floor['canopies'],
                        'floor_name': floor['floor_name'],
                        'delivery_install_data': floor.get('delivery_install_data', {}),
                        'p182': 0.0,  # Initialize p182 with 0
                        'floor_data': {
                            'cladding_total': 0,
                            'uv_total': 0
                        }
                    }
            
            # First pass: Collect P182 values and canopy prices from each canopy sheet
            p182_values = {}
            canopy_prices = {}  # New dictionary to store canopy prices
            for sheet_name in wb_data.sheetnames:
                if sheet_name.startswith('CANOPY - '):
                    item_sheet = wb_data[sheet_name]
                    
                    # Get P182 value
                    p182_value = item_sheet['P182'].value
                    if p182_value is not None:
                        try:
                            if isinstance(p182_value, str):
                                p182_value = p182_value.replace('£', '').replace(',', '').strip()
                            p182_values[sheet_name] = math.ceil(float(p182_value))  # Round up P182
                            st.write(f"Found P182 value in sheet {sheet_name}: {p182_value}")
                        except (ValueError, TypeError) as e:
                            st.error(f"Error processing P182 value from sheet {sheet_name}: {e}")
                    
                    # Get canopy prices and extract static values
                    current_row = 12  # Start row for prices
                    extract_row = 22  # Start row for extract static
                    sheet_prices = []
                    extract_statics = []  # List to store extract static values
                    
                    while True:
                        # Get price
                        price_cell = item_sheet[f'P{current_row}'].value
                        if price_cell is None:
                            break
                        
                        # Get extract static from column F
                        extract_static = item_sheet[f'F{extract_row}'].value
                        
                        try:
                            if isinstance(price_cell, str):
                                price_cell = price_cell.replace('£', '').replace(',', '').strip()
                            sheet_prices.append(float(price_cell))
                            
                            # Process extract static value
                            if extract_static is not None:
                                if isinstance(extract_static, str):
                                    extract_static = extract_static.replace('Pa', '').strip()
                                extract_statics.append(float(extract_static))
                            else:
                                extract_statics.append(0.0)
                            
                            st.write(f"Found in {sheet_name}:")
                            st.write(f"  Price at P{current_row}: {price_cell}")
                            st.write(f"  Extract Static at F{extract_row}: {extract_static}")
                        except (ValueError, TypeError) as e:
                            st.error(f"Error processing values at row {current_row}: {e}")
                        
                        current_row += 17  # Move to next canopy section
                        extract_row += 17  # Move to next extract static value
                    
                    if sheet_prices:
                        canopy_prices[sheet_name] = sheet_prices
                        # Store extract static values with the same sheet name
                        canopy_prices[f"{sheet_name}_extract_static"] = extract_statics

            # Update floor data with both P182 and canopy prices
            for kitchen_name, kitchen in grouped_canopy_data.items():
                for floor_name, floor_data in kitchen.items():
                    # Try different sheet name formats
                    possible_sheet_names = [
                        f"CANOPY - {kitchen_name} ({floor_name})",
                        f"CANOPY - {kitchen_name} ({floor_name[:8]}",
                        f"CANOPY - {floor_name} ({kitchen_name})",
                        f"CANOPY - {floor_name} ({kitchen_name[:8]}"
                    ]
                    
                    # Find matching sheet
                    matching_sheet = None
                    for sheet_name in p182_values.keys():
                        if any(possible_name in sheet_name for possible_name in possible_sheet_names):
                            matching_sheet = sheet_name
                            break
                    
                    if matching_sheet:
                        # Add P182 value
                        p182_value = p182_values[matching_sheet]
                        floor_data['p182'] = p182_value
                        
                        # Store at both levels to ensure consistency
                        for kitchen in kitchen_info:
                            for floor in kitchen['floors']:
                                if floor['floor_name'] == floor_name:
                                    floor['p182_value'] = p182_value
                                    break
                        
                        # Add canopy prices and extract static values if available
                        if matching_sheet in canopy_prices:
                            prices = canopy_prices[matching_sheet]
                            extract_statics = canopy_prices.get(f"{matching_sheet}_extract_static", [])
                            
                            # Assign prices and extract static values to canopies in order
                            for canopy, price, extract_static in zip(floor_data['canopies'], prices, extract_statics):
                                canopy['total_price'] = math.ceil(price)  # Round up canopy price
                                canopy['extract_static'] = extract_static
                            
                            st.write(f"For {kitchen_name} - {floor_name}:")
                            st.write(f"  P182: {math.ceil(p182_value)}")  # Show rounded P182

            # Create a dictionary for kitchen totals
            kitchen_totals = {}
            
            # Before generating Word document, restructure the data
            grouped_canopy_data = {}
            for kitchen in kitchen_info:
                kitchen_name = kitchen['kitchen_name']
                grouped_canopy_data[kitchen_name] = {}
                kitchen_total = 0
                
                # First get all sheets for this kitchen to sum N9
                kitchen_sheets = []
                for sheet_name in wb_data.sheetnames:
                    if sheet_name.startswith('CANOPY - ') and kitchen['kitchen_name'][:8] in sheet_name:
                        kitchen_sheets.append(sheet_name)
                
                # Sum N9 from all kitchen sheets
                for sheet_name in kitchen_sheets:
                    item_sheet = wb_data[sheet_name]
                    n9_value = item_sheet['N9'].value
                    if n9_value is not None:
                        try:
                            if isinstance(n9_value, str):
                                n9_value = n9_value.replace('£', '').replace(',', '').strip()
                            # Round up the N9 value
                            kitchen_total += math.ceil(float(n9_value))
                        except (ValueError, TypeError) as e:
                            st.error(f"Error processing N9 value from {sheet_name}: {e}")
                
                # Process each floor
                for floor in kitchen['floors']:
                    floor_name = floor['floor_name']
                    
                    # Find matching sheet for this floor
                    for sheet_name in wb_data.sheetnames:
                        if sheet_name.startswith('CANOPY - '):
                            if floor['floor_name'] in sheet_name and kitchen['kitchen_name'][:8] in sheet_name:
                                item_sheet = wb_data[sheet_name]
                                
                                # Process each canopy
                                current_row = 19  # Start at N19 for cladding
                                uv_row = 24      # Start at N24 for UV components
                                
                                for canopy in floor['canopies']:
                                    # Get cladding price if applicable
                                    if canopy.get('wallCladding'):
                                        price_cell = item_sheet[f'N{current_row}'].value
                                        if price_cell is not None:
                                            try:
                                                if isinstance(price_cell, str):
                                                    price_cell = price_cell.replace('£', '').replace(',', '').strip()
                                                # Round up the cladding price
                                                canopy['cladding_price'] = math.ceil(float(price_cell))
                                            except (ValueError, TypeError) as e:
                                                st.error(f"Error processing cladding price for {canopy.get('item_number')}: {e}")
                                    
                                    # Get UV component prices if it's a UV canopy
                                    if 'UV' in canopy.get('model', ''):
                                        uv_prices = []
                                        for i in range(3):  # Get 3 component prices
                                            price_cell = item_sheet[f'N{uv_row + i}'].value
                                            if price_cell is not None:
                                                try:
                                                    if isinstance(price_cell, str):
                                                        price_cell = price_cell.replace('£', '').replace(',', '').strip()
                                                    # Round up UV prices
                                                    uv_prices.append(math.ceil(float(price_cell)))
                                                except (ValueError, TypeError) as e:
                                                    st.error(f"Error processing UV price from N{uv_row + i}: {e}")
                                        canopy['uv_component_prices'] = uv_prices
                                    
                                    current_row += 17    # Move to next canopy's cladding price
                                    uv_row += 17        # Move to next canopy's UV prices
                                
                                break
                    
                    # Store floor data with all required fields
                    grouped_canopy_data[kitchen_name][floor_name] = {
                        'canopies': floor['canopies'],
                        'floor_name': floor['floor_name'],
                        'delivery_install_data': floor.get('delivery_install_data', {}),
                        'p182': floor['p182_value'],  # Use the stored value directly
                        'floor_data': {
                            'cladding_total': 0,
                            'uv_total': 0
                        }
                    }
                
                # Store kitchen total after processing all floors
                kitchen_totals[kitchen_name] = kitchen_total

            # First calculate all totals
            try:
                # Sum up all kitchen totals for T16
                total_job_price = sum(math.ceil(total) for total in kitchen_totals.values())
                
                # Get K9 totals from each sheet using data_only workbook
                total_k9_cost = 0
                for sheet_name in wb_data.sheetnames:  # Use wb_data instead of wb
                    if sheet_name.startswith('CANOPY - '):
                        sheet = wb_data[sheet_name]  # Use wb_data to get calculated values
                        k9_value = sheet['K9'].value
                        if k9_value is not None:
                            try:
                                if isinstance(k9_value, str):
                                    k9_value = k9_value.replace('£', '').replace(',', '').strip()
                                total_k9_cost += math.ceil(float(k9_value))
                                st.write(f"K9 value from {sheet_name}: £{math.ceil(float(k9_value))}")
                            except (ValueError, TypeError) as e:
                                st.error(f"Error processing K9 value from {sheet_name}: {e}")
                
                # Now create the word context with the calculated totals
                word_context = {
                    'kitchens': kitchen_info,
                    'grouped_canopy_data': grouped_canopy_data,
                    'kitchen_totals': kitchen_totals,  # Add the totals dictionary
                    'total_k9_cost': total_k9_cost,    # Add S16 total
                    'total_job_price': total_job_price  # Add T16 total
                }

                # Generate Word document
                word_file = CW.generate_word(word_context, genInfo)
                
                # Write to JOB TOTAL sheet in the formula workbook
                if 'JOB TOTAL' in wb_data.sheetnames:  # Use wb_data for writing
                    job_total_sheet = wb_data['JOB TOTAL']
                    # Write total job price to T16
                    job_total_sheet['T16'] = total_job_price
                    # Write total costs to S16
                    job_total_sheet['S16'] = total_k9_cost
                    
                    st.write(f"Total Job Price (T16): £{total_job_price}")
                    st.write(f"Total Costs (S16): £{total_k9_cost}")
                    
                    # Save the workbook with updated values
                    modified_excel = BytesIO()
                    wb_data.save(modified_excel)
                    modified_excel.seek(0)
                    
                    # Extract N193 and N19 values before creating ZIP
                    for kitchen in kitchen_info:
                        for floor in kitchen.get('floors', []):
                            # Find matching sheet for this floor
                            for sheet_name in wb_data.sheetnames:
                                if sheet_name.startswith('CANOPY - '):
                                    if floor['floor_name'] in sheet_name and kitchen['kitchen_name'][:8] in sheet_name:
                                        item_sheet = wb_data[sheet_name]
                                        
                                        # Get N193 value (test & commission price)
                                        n193_cell = item_sheet['N193']
                                        if n193_cell.value is not None:
                                            try:
                                                if isinstance(n193_cell.value, str):
                                                    n193_value = n193_cell.value.replace('£', '').replace(',', '').strip()
                                                else:
                                                    n193_value = n193_cell.value
                                                floor['test_commission'] = math.ceil(float(n193_value))
                                                st.write(f"N193 value for {sheet_name}: {n193_value}")  # Debug print
                                            except (ValueError, TypeError) as e:
                                                st.error(f"Error processing N193 value: {e}")
                                                floor['test_commission'] = 0
                                        
                                        # Get N19 value (total price)
                                        n19_cell = item_sheet['N19']
                                        if n19_cell.value is not None:
                                            try:
                                                if isinstance(n19_cell.value, str):
                                                    n19_value = n19_cell.value.replace('£', '').replace(',', '').strip()
                                                else:
                                                    n19_value = n19_cell.value
                                                floor['total_price'] = math.ceil(float(n19_value))
                                                st.write(f"N19 value for {sheet_name}: {n19_value}")  # Debug print
                                            except (ValueError, TypeError) as e:
                                                st.error(f"Error processing N19 value: {e}")
                                                floor['total_price'] = 0
                    
                    # Now create ZIP package with updated data
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w') as zf:
                        # Add the modified Excel file
                        modified_excel.seek(0)
                        zf.writestr(uploaded_file.name, modified_excel.getvalue())
                        
                        # Add Word file with updated test_commission values
                        word_file = CW.generate_word(word_context, genInfo)
                        zf.writestr("Halton Quotation.docx", word_file.getvalue())
                        
                        # Add JSON file
                        project_data = {
                            'session_state': {
                                k: v for k, v in st.session_state.items() 
                                if not k.startswith('_') and not k.startswith('customer')
                            },
                            'kitchen_info': kitchen_info,
                            'grouped_canopy_data': grouped_canopy_data,
                            'timestamp': datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                        }
                        
                        json_str = json.dumps(project_data, indent=2)
                        project_name = st.session_state.get('projectNum', 'project')
                        zf.writestr(
                            f"{project_name}_data.json",
                            json_str.encode('utf-8')
                        )
                    
                    zip_buffer.seek(0)
                    
                    # Provide download button for the ZIP file
                    st.download_button(
                        label="⬇️ Download Final Package",
                        data=zip_buffer,
                        file_name=f"{genInfo['projectNum']}_Package.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("JOB TOTAL sheet not found in workbook")

            except Exception as e:
                st.error(f"Error processing files: {str(e)}")

        except Exception as e:
            st.error(f"Error generating files: {str(e)}")

    # Use the names from general_info
  
def fill_dummy_kitchen_data():
    """Fill kitchen form with dummy data for testing"""
    # Set kitchen name
    st.session_state['kitchen_name'] = "Main Kitchen"
    
    # Set floor name
    st.session_state['floor_name'] = "Ground Floor"
    
    # Set canopy data
    st.session_state['itemNum'] = "KM123-1"
    st.session_state['model'] = "KVF"  # One of the models that needs Supply Air calc
    st.session_state['width'] = 1000
    st.session_state['length'] = 2000
    st.session_state['height'] = 600
    st.session_state['section'] = 2
    st.session_state['flowRate'] = 0.5
    
    # Calculate Supply Air for specific models
    if st.session_state['model'] in ['UVX-M', 'KVX-M', 'KVF', 'CMWF', 'UVF']:
        # Supply Air calculation
        supply_air = st.session_state['flowRate'] * 0.85  # 85% of extract rate
        st.session_state['supply_air'] = supply_air
    
    # Set lights
    st.session_state['lights'] = "LED Strip Light"
    st.session_state['light_quantity'] = 2
    
    # Set wall cladding
    st.session_state['cladding'] = True
    st.session_state['cladding_desc'] = ["Rear", "Left"]

def canopy_main():
    excel_template_path = TEMPLATES['EXCEL']
    
    if not os.path.exists(excel_template_path):
        st.error(f"Excel template not found at: {excel_template_path}")
        return
    
    st.title("Canopy Cost Sheet Generator")
    
    # Add dummy data button
    if st.button("Fill Test Kitchen Data"):
        fill_dummy_kitchen_data()
    
    # Rest of your existing code...
  