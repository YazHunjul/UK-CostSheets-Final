from openpyxl import load_workbook
import streamlit as st
import io
import os

def generate_sheet(info, genInfo):
    try:
        # Load the workbook
        wb = load_workbook('/Users/yazan/Desktop/Efficiency/UK-CostSheets-Final/app/costSheetGen/costSheetResources/Cost Sheet R18.5 Sep 2024 (7).xlsx')
        ws = wb['CANOPY']

        # Fill out General Information
        ws['C3'] = genInfo.get('projectNum').title()
        ws['C5'] = genInfo.get('customer').title()
        ws['C7'] = genInfo.get('salesContact').title()
        ws['G3'] = genInfo.get('projectName').title()
        ws['G5'] = genInfo.get('location').title()
        ws['G7'] = genInfo.get('date')
        
        # Fill in the rest of your data without modifying any formulas
        itemNum = 12
        canopyInfo = 14
        lights = 15
        specialWorks = itemNum + 5
        wallCladding = itemNum + 7
        control_panel = itemNum + 9
        ww_pods = itemNum + 10
        pipework = itemNum + 11

        for kitchen in info:
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
                    
                    ws[f'C{control_panel}'] = canopy.get('control_panel') if canopy.get('model') in ['CMWI', 'CMWF'] else ''
                    ws[f'C{ww_pods}'] = canopy.get('WW_pods') if canopy.get('model') in ['CMWI', 'CMWF'] else ''
                    ws[f'C{pipework}'] = canopy.get('pipework') if canopy.get('model') in ['CMWI', 'CMWF'] else ''
                    
                    itemNum += 17
                    canopyInfo += 17
                    lights += 17
                    specialWorks = itemNum + 5
                    wallCladding += 17
                    control_panel += 17
                    ww_pods += 17
                    pipework += 17

        # First save to a real file to ensure Excel can calculate it
        temp_path = "temp_for_calc.xlsx"
        wb.save(temp_path)
        
        # Now read it back with data_only=True
        wb_data = load_workbook(temp_path)
        
        # Save to BytesIO for return
        output = io.BytesIO()
        wb_data.save(output)
        
        # Clean up temp file
        os.remove(temp_path)
        
        return output

    except Exception as e:
        st.error(f"Error generating Excel sheet: {str(e)}")
        return None