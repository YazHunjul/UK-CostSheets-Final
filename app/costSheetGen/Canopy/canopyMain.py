import streamlit as st
import zipfile
from io import BytesIO
from costSheetGen.Canopy import canopyExcel as CE
from costSheetGen.Canopy import canopyWord as CW
import time
from openpyxl import load_workbook
import io
import os
from costSheetGen.Canopy.canopyUtils import extract_canopy_prices


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

                            with coll4:
                                flowrate = st.number_input('Enter Flow Rate', min_value=0.0, key=f'flowRate_{i}_{floor}_{canopy}')
                                if model in ['CMWI', 'CMWF']:
                                    control_panel = st.selectbox('Select Control Panel', ['CP1S', 'CP2S', 'CP3S', 'CP4S'], key=f'CP_{i}_{floor}_{canopy}')
                                    WW_pods = st.selectbox("W/W Pods", ['1000-S', '1500-S', '2000-S', '2500-S', '3000-S', '1000-D', '1500-D', '2000-D', '2000-D', '2500-D', '3000-D'], key=f'WW_{i}_{floor}_{canopy}')
                                    CWS_HWS_pipework = st.selectbox("CWS/HWS Pipework", [1,2,3,4,5], key=f'pipework_{i}_{floor}_{canopy}')

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
                                'cladding_desc' : description
                            }

                            # Append canopy data to the floor
                            floor_data["canopies"].append(canopy_data)

                        # Append floor data to the kitchen
                        kitchen_data["floors"].append(floor_data)

            # Append kitchen data to the main list
            kitchen_info.append(kitchen_data)
    
    st.markdown('<hr>', unsafe_allow_html=True)

    # Import and call your Excel generation module
    if st.button("Generate Documents"):
        # Generate Excel and store the BytesIO object
        excel_bytes = CE.generate_sheet(kitchen_info, genInfo)
        
        # Save to a temporary file for reading values
        temp_path = "temp_cost_sheet.xlsx"
        with open(temp_path, "wb") as f:
            f.write(excel_bytes.getvalue())
        
        # Add a longer delay to ensure Excel is fully saved and calculated
        time.sleep(5)  # Increased to 5 seconds
        
        # Extract prices from the temporary file
        canopy_prices = extract_canopy_prices(temp_path)
        
        # Display the values
        if canopy_prices:
            st.write("Canopy Prices (Rounded Up):")
            total_sum = 0
            for price in canopy_prices:
                st.write(f"Canopy {price['canopy_number']}: £{price['total_price']:,}")
                total_sum += price['total_price']
            
            st.write(f"\nTotal Project Cost: £{total_sum:,}")
        
        # Clean up temporary file
        if os.path.exists(temp_path):
            os.remove(temp_path)
        
        # Generate Word document
        word_context = {'kitchens': kitchen_info}  # Example context for the Word document
        word_file = CW.generate_word(word_context, genInfo)

        # Zip the files
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zf:
            zf.writestr("Modified_Cost_Sheet.xlsx", excel_bytes.getvalue())
            zf.writestr("Halton_Quotation.docx", word_file.getvalue())
        zip_buffer.seek(0)

        # Provide download button for the ZIP file
        st.download_button(
            label="Download ZIP File",
            data=zip_buffer,
            file_name="Cost_Sheet_and_Quotation.zip",
            mime="application/zip"
        )