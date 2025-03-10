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
            "bms_interface": bms_interface
        }

def main(genInfo):
    st.markdown('<hr>', unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center;'>Canopy Cost Sheet</h2>", unsafe_allow_html=True)
    
    # Get Kitchen Count
    num_kitchens = st.number_input("Enter Number of Levels", min_value=1, key='num_kitchens')
    kitchen_info = []

    for i in range(num_kitchens):
        kitchen_name = st.text_input(f"Enter Level {i + 1} Name", key=f'kitchen_name_{i}')
        if kitchen_name:
            # Create a dictionary for this kitchen
            kitchen_data = {
                "kitchen_name": kitchen_name,
                "floors": []
            }

            with st.expander(f'{kitchen_name.title()} Floor Information', expanded=True):
                num_floors = st.number_input(
                    f"Enter the number of areas in {kitchen_name} Floor", 
                    min_value=1, 
                    key=f'floors_input_{i}'
                )
                for floor in range(num_floors):
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
                                    ['ROUND CORNERS', 'CUT OUT', 'CASTELLE LOCKING', 'HEADER DUCT S/S', 'HEADER DUCT', 'PAINT FINISH'],
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
                                height = st.number_input("Height", min_value=0, key=f'height_{i}_{floor}_{canopy}')
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

                            # Assuming WW_pods selection is already defined
                            

                            # Create a dictionary for this canopy
                            canopy_data = {
                                'itemNum' : item_number,
                                "model": model,
                                "configuration": configuration,
                                "section": section,
                                "height": height,
                                "width": width,
                                "length": length,
                                'lights': light_type,
                                'light_quantity': light_quantity,
                                'specialWorks': special_works_dict,
                                'wallCladding': cladding,
                                'flowrate' : flowrate,
                                'control_panel' : control_panel,
                                'WW_pods' : WW_pods,
                                'pipework' : CWS_HWS_pipework,
                                'cladding_width' : cladding_width,
                                'cladding_height': cladding_height,
                                'cladding_desc' : description,
                                'WW_pods_quantity' : WW_pods_quantity
                            }

                            # Append canopy data to the floor
                            floor_data["canopies"].append(canopy_data)

                        # Append floor data to the kitchen
                        kitchen_data["floors"].append(floor_data)

            # Append kitchen data to the main list
            kitchen_info.append(kitchen_data)
    
    st.markdown('<hr>', unsafe_allow_html=True)

    # Delivery & Installation Section
    st.markdown("## 🚚 Delivery & Installation Details")
    
    # Iterate through all kitchens and their floors
    for kitchen_idx, kitchen in enumerate(kitchen_info):
        for floor_idx, floor in enumerate(kitchen['floors']):
            floor_name = floor['floor_name']
            kitchen_name = kitchen['kitchen_name']
            
            # Get delivery & installation details for this floor
            floor['delivery_install_data'] = get_delivery_install_details(
                floor_name=f"{kitchen_name} - {floor_name}",  # Include kitchen name for clarity
                key_suffix=f"kitchen_{kitchen_idx}_floor_{floor_idx}"  # Unique key for each floor
            )

    st.markdown('<hr>', unsafe_allow_html=True)

    # After delivery & installation section
    st.markdown('<hr>', unsafe_allow_html=True)

    # Add Email Summary Section first
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
            # Load workbooks
            wb_data = openpyxl.load_workbook(uploaded_file, data_only=True)  # For reading values
            wb = openpyxl.load_workbook(uploaded_file, keep_vba=True, data_only=False)  # For preserving formulas and data validation
            
            # First collect all prices and P182 values
            prices_by_item = {}
            current_row = 12
            sheet = wb_data['CANOPY']
            
            # First pass: Collect all prices and find corresponding P182 values
            while sheet[f'C{current_row}'].value:
                item_num = sheet[f'C{current_row}'].value
                model = sheet[f'D{current_row + 2}'].value
                
                st.write(f"\nCollecting prices for item {item_num}")
                
                # Get canopy price
                price = sheet[f'N{current_row}'].value
                if price is not None:
                    try:
                        st.write(f"Raw price value: {price} (type: {type(price)})")
                        if isinstance(price, str):
                            price = price.replace('£', '').replace(',', '').strip()
                        rounded_price = float(f"{math.ceil(float(price))}.00")
                        
                        # Find corresponding sheet for this item
                        p182_value = 0.0
                        for sheet_name in wb_data.sheetnames:
                            if sheet_name.startswith('CANOPY - '):
                                item_sheet = wb_data[sheet_name]
                                # Debug sheet contents
                                st.write(f"\nChecking sheet: {sheet_name}")
                                st.write("Sheet type:", type(item_sheet))
                                st.write("All sheet names:", wb_data.sheetnames)
                                
                                # Get P182 value from this sheet
                                try:
                                    p182_cell = item_sheet['P182'].value
                                    st.write("P182 cell raw value:", p182_cell)
                                    st.write("P182 cell type:", type(p182_cell))
                                    
                                    if p182_cell is not None:
                                        if isinstance(p182_cell, str):
                                            p182_cell = p182_cell.replace('£', '').replace(',', '').strip()
                                        p182_value = float(f"{math.ceil(float(p182_cell))}.00")
                                        st.write(f"Processed P182 value: £{p182_value:.2f}")
                                        break  # Found the value, no need to check other sheets
                                except Exception as e:
                                    st.error(f"Error accessing P182 in {sheet_name}: {str(e)}")
                                    st.write("Full error:", e)
                        
                        prices_by_item[item_num] = {
                            'base_price': rounded_price,
                            'p182_value': p182_value
                        }
                        st.write(f"Stored price: £{rounded_price:.2f} and P182: £{p182_value:.2f}")
                    except (ValueError, TypeError) as e:
                        st.error(f"Error processing price for {item_num}: {e}")
                else:
                    st.warning(f"No price found in cell N{current_row}")
                
                current_row += 17
            
            # Debug the collected prices and P182 values
            st.write("\nAll collected prices:", prices_by_item)
            
            # Calculate total P182
            total_p182 = sum(item['p182_value'] for item in prices_by_item.values())
            st.write(f"Total P182 value: £{total_p182:.2f}")
            
            # Write total P182 to CANOPY sheet
            sheet = wb['CANOPY']
            sheet['N182'] = total_p182
            
            # Save the modified workbook
            modified_excel = BytesIO()
            wb.save(modified_excel)
            modified_excel.seek(0)
            
            # Second pass: Add prices and P182 to kitchen_info structure
            st.write("\nAssigning prices to canopies...")
            for kitchen in kitchen_info:
                st.write(f"\nKitchen: {kitchen['kitchen_name']}")
                kitchen['total_p182'] = 0
                
                for floor in kitchen['floors']:
                    st.write(f"Floor: {floor['floor_name']}")
                    
                    # Find P182 value for this floor
                    for sheet_name in wb_data.sheetnames:
                        if sheet_name.startswith('CANOPY - '):
                            if floor['floor_name'] in sheet_name and kitchen['kitchen_name'][:8] in sheet_name:
                                item_sheet = wb_data[sheet_name]
                                p182_cell = item_sheet['P182'].value
                                if p182_cell is not None:
                                    try:
                                        if isinstance(p182_cell, str):
                                            p182_cell = p182_cell.replace('£', '').replace(',', '').strip()
                                        floor['p182_value'] = float(f"{math.ceil(float(p182_cell))}.00")
                                        kitchen['total_p182'] += floor['p182_value']
                                        st.write(f"Found P182 value for {floor['floor_name']}: £{floor['p182_value']:.2f}")
                                        break
                                    except (ValueError, TypeError) as e:
                                        st.error(f"Error processing P182 for {floor['floor_name']}: {e}")
                    
                    # Add prices to canopies
                    for canopy in floor['canopies']:
                        item_num = canopy.get('itemNum')
                        if item_num in prices_by_item:
                            canopy['total_price'] = prices_by_item[item_num]['base_price']
                            st.write(f"Added price £{canopy['total_price']:.2f} to canopy {item_num}")
                        else:
                            st.warning(f"No price found for canopy {item_num}")

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
                            kitchen_total += float(f"{math.ceil(float(n9_value))}.00")
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
                                                canopy['cladding_price'] = float(f"{math.ceil(float(price_cell))}.00")
                                            except (ValueError, TypeError) as e:
                                                st.error(f"Error processing cladding price for {canopy.get('itemNum')}: {e}")
                                    
                                    # Get UV component prices if it's a UV canopy
                                    if 'UV' in canopy.get('model', ''):
                                        uv_prices = []
                                        for i in range(3):  # Get 3 component prices
                                            price_cell = item_sheet[f'N{uv_row + i}'].value
                                            if price_cell is not None:
                                                try:
                                                    if isinstance(price_cell, str):
                                                        price_cell = price_cell.replace('£', '').replace(',', '').strip()
                                                    uv_prices.append(float(f"{math.ceil(float(price_cell))}.00"))
                                                except (ValueError, TypeError) as e:
                                                    st.error(f"Error processing UV price from N{uv_row + i}: {e}")
                                        canopy['uv_component_prices'] = uv_prices
                                    
                                    current_row += 17    # Move to next canopy's cladding price
                                    uv_row += 17        # Move to next canopy's UV prices
                                
                                break
                    
                    # Store floor data after processing all canopies
                    grouped_canopy_data[kitchen_name][floor_name] = {
                        'canopies': floor['canopies'],
                        'delivery_install_data': floor['delivery_install_data'],
                        'floor_name': floor['floor_name'],
                        'p182': floor['p182_value']
                    }
                
                # Store kitchen total after processing all floors
                kitchen_totals[kitchen_name] = kitchen_total

            # Add kitchen_totals to word context
            word_context = {
                'kitchens': kitchen_info,
                'grouped_canopy_data': grouped_canopy_data,
                'kitchen_totals': kitchen_totals  # Add the totals dictionary
            }

            # Generate Word document with the restructured data
            try:
                word_file = CW.generate_word(word_context, genInfo)
                
                if word_file is None:
                    st.error("Error: Word document generation returned None")
                    return
                
                # Create ZIP with both files
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zf:
                    # Add modified uploaded Excel file
                    excel = f"{genInfo['projectNum']} Cost Sheet {genInfo['date']}.xlsx"
                    modified_excel.seek(0)
                    zf.writestr("Cost Sheet.xlsx", modified_excel.getvalue())
                    
                    # Add Word file
                    zf.writestr("Halton Quotation.docx", word_file.getvalue())
                
                zip_buffer.seek(0)
                
                # Provide download button for the ZIP file
                st.download_button(
                    label="⬇️ Download Final Package",
                    data=zip_buffer,
                    file_name="Cost_Sheet_and_Quotation.zip",
                    mime="application/zip"
                )
            
            except Exception as e:
                st.error(f"Error generating Word document: {str(e)}")
                st.write("Word context:", word_context)

        except Exception as e:
            st.error(f"An error occurred processing the file: {str(e)}")

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
    st.session_state['flowrate'] = 0.5
    
    # Calculate Supply Air for specific models
    if st.session_state['model'] in ['UVX-M', 'KVX-M', 'KVF', 'CMWF', 'UVF']:
        # Supply Air calculation
        supply_air = st.session_state['flowrate'] * 0.85  # 85% of extract rate
        st.session_state['supply_air'] = supply_air
    
    # Set lights
    st.session_state['lights'] = "LED Strip Light"
    st.session_state['light_quantity'] = 2
    
    # Set wall cladding
    st.session_state['wallCladding'] = True
    st.session_state['cladding_desc'] = ["Rear", "Left"]

def canopy_main():
    # Get the absolute path to the resources directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    resources_dir = os.path.join(current_dir, '..', 'costSheetResources', 'templates')
    
    # Use the path for loading the Excel template
    excel_template_path = os.path.join(resources_dir, 'Halton Cost Sheet Jan 2025.xlsx')
    
    if not os.path.exists(excel_template_path):
        st.error(f"Excel template not found at: {excel_template_path}")
        return
    
    st.title("Canopy Cost Sheet Generator")
    
    # Add dummy data button
    if st.button("Fill Test Kitchen Data"):
        fill_dummy_kitchen_data()
    
    # Rest of your existing code...
  