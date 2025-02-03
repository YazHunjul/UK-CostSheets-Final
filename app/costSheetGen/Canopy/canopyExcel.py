from openpyxl import load_workbook
import streamlit as st
import io
import os
import openpyxl
from io import BytesIO

def generate_sheet(kitchen_info, genInfo, delivery_install_data):
    """
    Generate an Excel sheet based on kitchen information and return it as a BytesIO object.
    """
    try:
        # Load the workbook
        wb = load_workbook('app/costSheetGen/costSheetResources/costSheetTest.xlsx')
        ws = wb['CANOPY']

        # Fill out General Information
        ws['C3'] = genInfo.get('projectNum').title()
        ws['C5'] = genInfo.get('customer').title()
        ws['C7'] = genInfo.get('combined_initials')
        ws['G3'] = genInfo.get('projectName').title()
        ws['G5'] = genInfo.get('location').title()
        ws['G7'] = genInfo.get('date')
        
        # Fill in canopy data
        itemNum = 12
        canopyInfo = 14
        lights = 15
        specialWorks = itemNum + 5
        wallCladding = itemNum + 7
        control_panel = itemNum + 13
        ww_pods = itemNum + 14
        pipework = itemNum + 15

        for kitchen in kitchen_info:
            for floor in kitchen['floors']:
                for canopy in floor['canopies']:
                    ws[f'C{itemNum}'] = canopy.get('itemNum')
                    ws[f'D{canopyInfo}'] = canopy.get('model')
                    ws[f'E{canopyInfo}'] = canopy.get('width')
                    ws[f'F{canopyInfo}'] = canopy.get('length')
                    ws[f'G{canopyInfo}'] = canopy.get('height')
                    ws[f'H{canopyInfo}'] = canopy.get('section')
                    ws[f'I{canopyInfo}'] = canopy.get('flowrate')
                    
                    ws[f'C{lights}'] = canopy.get('lights')
                    ws[f'D{lights}'] = canopy.get('light_quantity')
                    
                    special_works_dict = canopy.get('specialWorks', {})
                    for i, (work, qty) in enumerate(list(special_works_dict.items())[:2]):
                        current_row = specialWorks + i
                        ws[f'C{current_row}'] = work
                        ws[f'D{current_row}'] = qty
                    
                    ws[f'C{wallCladding}'] = "2M² (HFL)" if canopy.get('wallCladding') else ''
                    if canopy.get('wallCladding') is None:
                        ws[f'D{wallCladding}'] = 0 
                    
                    ws[f'C{control_panel}'] = canopy.get('control_panel') if canopy.get('model') in ['CMWI', 'CMWF'] else ''
                    ws[f'C{ww_pods}'] = canopy.get('WW_pods') if canopy.get('model') in ['CMWI', 'CMWF'] else ''
                    ws[f'D{ww_pods}'] = canopy.get('WW_pods_quantity') if canopy.get('model') in ['CMWI', 'CMWF'] else ''
                    ws[f'C{pipework}'] = canopy.get('pipework') if canopy.get('model') in ['CMWI', 'CMWF'] else ''
                    
                    itemNum += 17
                    canopyInfo += 17
                    lights += 17
                    specialWorks = itemNum + 5
                    wallCladding += 17
                    control_panel += 17
                    ww_pods += 17
                    pipework += 17

        # Fill delivery and installation data
        ws['D183'] = delivery_install_data['delivery_location']
        ws['C183'] = delivery_install_data['delivery_lift_qty']
        
        if delivery_install_data['plant_hires']:
            if "Plant Hire 1" in delivery_install_data['plant_hires']:
                ws['D184'] = delivery_install_data['plant_hires']["Plant Hire 1"]
                ws['C184'] = delivery_install_data['quantities'].get("Plant Hire 1", 0)
            if "Plant Hire 2" in delivery_install_data['plant_hires']:
                ws['D185'] = delivery_install_data['plant_hires']["Plant Hire 2"]
                ws['C185'] = delivery_install_data['quantities'].get("Plant Hire 2", 0)

        ws['C187'] = delivery_install_data['strip_out']
        ws['C188'] = delivery_install_data['consumables']
        ws['C189'] = delivery_install_data['installation_normal']
        ws['C190'] = delivery_install_data['installation_after']
        ws['C191'] = delivery_install_data['wall_cladding']
        ws['C192'] = delivery_install_data['overnight_expenses']
        ws['C193'] = delivery_install_data['test_commission']
        ws['C194'] = delivery_install_data['gas_interlock']
        ws['C195'] = delivery_install_data['co_sensor']
        ws['C196'] = delivery_install_data['co2_sensor']
        ws['C197'] = delivery_install_data['bms_interface']

        # Save to BytesIO for return
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        wb.close()
        
        return output

    except Exception as e:
        st.error(f"Error generating Excel sheet: {str(e)}")
        return None