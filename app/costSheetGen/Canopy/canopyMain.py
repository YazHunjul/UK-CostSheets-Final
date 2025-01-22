import streamlit as st
import zipfile
from io import BytesIO
from costSheetGen.Canopy import canopyExcel as CE
from costSheetGen.Canopy import canopyWord as CW


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

                                model = st.selectbox(
                                    'Model', 
                                    ['KVF', 'KVX-M', "KVI", "UVX", "UVX-M", "UVI", "UVF", "UV-C POD", "CMWI", "CMWF", "CXW", "CXW-M", "KVV"], 
                                    key=f'model_{i}_{floor}_{canopy}'
                                )
                                
                                control_panel = st.selectbox('Select Control Panel', ['CP1S', 'CP2S', 'CP3S', 'CP4S'], key=f'CP_{i}_{floor}_{canopy}') if (model == 'CMWI' or model =='CMWF') else ''
                                
                                
                                
                                flowrate = st.number_input('Entert Flow Rate', min_value=0.0, key=f'flowRate_{i}_{floor}_{canopy}')
                                
                    
                                
                            with coll2:
                                

                                height = st.number_input(
                                    "Height", min_value=0,
                                    key=f'height_{i}_{floor}_{canopy}'
                                )
                                
                                section = st.number_input(
                                    'Sections', 
                                    min_value=0, 
                                    key=f'section_{i}_{floor}_{canopy}'
                                )
                                
                                WW_pods = st.selectbox("W/W Pods", ['1000-S', '1500-S', '2000-S', '2500-S', '3000-S', '1000-D', '1500-D', '2000-D', '2000-D', '2500-D', '3000-D'], key=f'WW_{i}_{floor}_{canopy}') if (model == 'CMWI' or model =='CMWF') else ''
                                light_selection = st.selectbox(
                                    'Light Selection',
                                    ['LED STRIP L6 Inc DALI', 'LED STRIP L12 inc DALI', 'LED STRIP L18 Inc DALI', 'Small LED Spots inc DALI', 'LARGE LED Spots inc DALI'],
                                    key=f'light_{i}_{floor}_{canopy}'
                                )
                                
                            with coll3:
                                model = st.selectbox(
                                    'Model', 
                                    ['KVF', 'KVX-M', "KVI", "UVX", "UVX-M", "UVI", "UVF", "UV-C POD", "CMWI", "CMWF", "CXW", "CXW-M", "KVV"], 
                                    key=f'model_{i}_{floor}_{canopy}'
                                )
                                width = st.number_input(
                                    "Width", min_value=0,
                                    key=f'width_{i}_{floor}_{canopy}'
                                )
                                
                                CWS_HWS_pipework = st.selectbox("CWS/HWS Pipework", [1,2,3,4,5], key=f'pipework_{i}_{floor}_{canopy}') if model in ['CMWF', 'CMWI'] else ''
                                
                                special_works = st.selectbox(
                                    'Special Works',
                                    ['', 'ROUND CORNERS', 'CUT OUT', 'CASTELLE LOCKING', 'HEADER DUCT S/S', 'HEADER DUCT', 'PAINT FINISH'],
                                    key=f'specialWorks_{i}_{floor}_{canopy}'
                                )
                                
                                configuration = st.selectbox(
                                    'Configuration', ['WALL', "ISLAND"], 
                                    key=f'config_{i}_{floor}_{canopy}'
                                )
                                
                            with coll4:
                                length = st.number_input(
                                    "Length", min_value=0,
                                    key=f'length_{i}_{floor}_{canopy}'
                                )
                                cladding = st.selectbox(
                                    "Wall Cladding",
                                    ['', '2M² (HFL)'],
                                    key=f'cladding_{i}_{floor}_{canopy}'
                                )
                                
                                cladding_height = st.number_input("Cladding Height", key=f'cladding_Height{i}_{floor}_{canopy}', min_value=0) if cladding else ''
                                cladding_width = st.number_input("Cladding Length", key=f'CladdingLength_{i}_{floor}_{canopy}', min_value=0) if cladding else ''
                                description = st.multiselect('Cladding Description', ['','Rear', 'Left', "Right" ], key=f'cladding_desc_{i}_{floor}_{canopy}') if cladding else ''

                            # Create a dictionary for this canopy
                            canopy_data = {
                                'itemNum' : item_number,
                                "model": model,
                                "configuration": configuration,
                                "section": section,
                                "height": height,
                                "width": width,
                                "length": length,
                                'lights': light_selection,
                                'specialWorks': special_works,
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
    if st.button("Generate Sheet"):
        # Generate Excel file
        excel_file = CE.generate_sheet(info=kitchen_info, genInfo=genInfo)

        # Generate Word document
        word_context = {'kitchens': kitchen_info}  # Example context for the Word document
        word_file = CW.generate_word(word_context, genInfo)

        # Zip the files
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zf:
            zf.writestr("Modified_Cost_Sheet.xlsx", excel_file.getvalue())
            zf.writestr("Halton_Quotation.docx", word_file.getvalue())
        zip_buffer.seek(0)

        # Provide download button for the ZIP file
        st.download_button(
            label="Download ZIP File",
            data=zip_buffer,
            file_name="Cost_Sheet_and_Quotation.zip",
            mime="application/zip"
        )