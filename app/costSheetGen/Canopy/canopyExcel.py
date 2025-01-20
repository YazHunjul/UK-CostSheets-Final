import streamlit as st
from openpyxl import Workbook, load_workbook
import pandas as pd
from io import BytesIO

def generate_sheet(info, genInfo):
    try:
        # Load the existing workbook
        wb = load_workbook('app/costSheetGen/costSheetResources/Cost Sheet R18.5 Sep 2024 (7).xlsx')
        ws = wb['CANOPY']

        #Fill out General Information
        ws['C3'] = genInfo.get('projectNum').title()
        ws['C5'] = genInfo.get('customer').title()
        ws['C7'] = genInfo.get('salesContact').title()
        ws['G3'] = genInfo.get('projectName').title()
        ws['G5'] = genInfo.get('location').title()
        ws['G7'] = genInfo.get('date')
        
                    
        count = 1
        itemNum = 12
        canopyInfo = 14
        lights = 15
        specialWorks = 17
        wallCladding = 19
        
        
        control_panel = 25
        ww_pods = 26
        pipework = 27
        
        #Now, we need to dynamically fill out the Canopy information. Starting at index 12, adjusting the item number
        for kitchen in info:
            for floor in kitchen['floors']:
                for canopy in floor['canopies']:
                    ws[f'C{itemNum}'] = count
                    ws[f'D{canopyInfo}'] = canopy.get('model')
                    ws[f'E{canopyInfo}'] = canopy.get('width')
                    ws[f'F{canopyInfo}'] = canopy.get('length')
                    ws[f'G{canopyInfo}'] = canopy.get('height')
                    ws[f'H{canopyInfo}'] = canopy.get('section')
                    ws[f'I{canopyInfo}'] = canopy.get('flowrate')
                    ws[f'C{lights}'] = canopy.get('lights')
                    ws[f'C{specialWorks}'] = canopy.get('specialWorks') 
                    ws[f'C{wallCladding}'] = "2M² (HFL)" if canopy.get('wallCladding') else ''
                    
                    # Water Wash Hoods
                    ws[f'C{control_panel}'] = canopy.get('control_panel') if canopy.get('model') in ['CMWI', 'CMWF'] else ''
                    ws[f'C{ww_pods}'] = canopy.get('WW_pods') if canopy.get('model') in ['CMWI', 'CMWF'] else ''
                    ws[f'C{pipework}'] = canopy.get('pipework') if canopy.get('model') in ['CMWI', 'CMWF'] else ''
                    
                    
                    count+=1
                    itemNum += 17
                    canopyInfo += 17
                    
                    # Water Wash Hoods
                    control_panel += 17
                    ww_pods += 17
                    pipework += 17
                    


        # Save workbook to an in-memory BytesIO buffer
        output = BytesIO()
        wb.save(output)
        output.seek(0)  # Reset the pointer to the beginning of the buffer
        return output
    except Exception as e:
        st.error(f"Error generating the sheet: {e}")
        return None