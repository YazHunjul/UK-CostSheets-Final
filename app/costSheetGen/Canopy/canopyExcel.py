from openpyxl import load_workbook
import streamlit as st
import io
import os
import openpyxl
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.formatting.rule import CellIsRule

def generate_sheet(kitchen_info, genInfo):
    """
    Generate an Excel sheet based on kitchen information and return it as a BytesIO object.
    """
    try:
        # Load the workbook
        excel_path = '/Users/yazan/Desktop/Efficiency/UK-CostSheets-Final/app/costSheetGen/costSheetResources/Halton Cost Sheet Jan 2025.xlsx'
        wb = load_workbook(excel_path)
        
        # Keep a clean template
        template_ws = wb['CANOPY']
        clean_template = wb.copy_worksheet(template_ws)
        clean_template.title = 'TEMPLATE_TEMP'
        
        # Write general info to main CANOPY sheet
        template_ws['C3'] = genInfo.get('projectNum', '').title()
        template_ws['C5'] = genInfo.get('customer', '').title()
        template_ws['C7'] = genInfo.get('combined_initials', '')
        template_ws['G3'] = genInfo.get('projectName', '').title()
        template_ws['G5'] = genInfo.get('location', '').title()
        template_ws['G7'] = genInfo.get('date', '')
        
        # First write all data to main CANOPY sheet
        current_row = 12
        
        for kitchen in kitchen_info:
            for floor in kitchen.get('floors', []):
                for canopy in floor.get('canopies', []):
                    # Write to main CANOPY sheet
                    template_ws[f'C{current_row + 2}'] = canopy.get('configuration', '')
                    template_ws[f'C{current_row}'] = canopy.get('itemNum', '')
                    template_ws[f'D{current_row + 2}'] = canopy.get('model', '')
                    template_ws[f'E{current_row + 2}'] = canopy.get('width', '')
                    template_ws[f'F{current_row + 2}'] = canopy.get('length', '')
                    template_ws[f'G{current_row + 2}'] = canopy.get('height', '')
                    template_ws[f'H{current_row + 2}'] = canopy.get('section', '')
                    template_ws[f'I{current_row + 2}'] = canopy.get('flowrate', '')
                    
                    template_ws[f'C{current_row + 3}'] = canopy.get('lights', '')
                    template_ws[f'D{current_row + 3}'] = canopy.get('light_quantity', '')
                    
                    # Special works
                    special_works_dict = canopy.get('specialWorks', {})
                    for i, (work, qty) in enumerate(list(special_works_dict.items())[:2]):
                        template_ws[f'C{current_row + 5 + i}'] = work
                        template_ws[f'D{current_row + 5 + i}'] = qty
                    
                    # Wall cladding
                    template_ws[f'C{current_row + 7}'] = "2M² (HFL)" if canopy.get('wallCladding') else ''
                    
                    # CMWI/CMWF specific fields
                    if canopy.get('model') in ['CMWI', 'CMWF']:
                        template_ws[f'C{current_row + 13}'] = canopy.get('control_panel', '')
                        template_ws[f'C{current_row + 14}'] = canopy.get('WW_pods', '')
                        template_ws[f'D{current_row + 14}'] = canopy.get('WW_pods_quantity', '')
                        template_ws[f'C{current_row + 15}'] = canopy.get('pipework', '')
                    
                    current_row += 17
        
        # Now create individual floor sheets from clean template
        for kitchen in kitchen_info:
            kitchen_name = kitchen.get('kitchen_name', 'Unknown').title()
            
            for floor in kitchen.get('floors', []):
                floor_name = floor.get('floor_name', 'Unknown').title()
                sheet_name = f"CANOPY - {floor_name} ({kitchen_name})"[:31]
                
                new_sheet = wb.copy_worksheet(clean_template)
                new_sheet.title = sheet_name
                
                # Add conditional formatting
                light_blue_fill = PatternFill(
                    start_color='DEF2F7',  # RGB(222, 237, 242)
                    end_color='DEF2F7',
                    fill_type='solid'
                )
                
                dark_blue_font = Font(
                    color='9FCBDA'  # RGB(159, 203, 218)
                )
                
                for start_row in range(14, 200, 17):
                    for col in ['J', 'K', 'N', 'O']:
                        cell_range = f"{col}{start_row}:{col}{start_row + 13}"
                        rule = CellIsRule(
                            operator='greaterThan',
                            formula=['0'],
                            stopIfTrue=True,
                            fill=light_blue_fill,
                            font=dark_blue_font
                        )
                        new_sheet.conditional_formatting.add(cell_range, rule)
                
                # Fill general info
                new_sheet['C3'] = genInfo.get('projectNum', '').title()
                new_sheet['C5'] = genInfo.get('customer', '').title()
                new_sheet['C7'] = genInfo.get('combined_initials', '')
                new_sheet['G3'] = genInfo.get('projectName', '').title()
                new_sheet['G5'] = genInfo.get('location', '').title()
                new_sheet['G7'] = genInfo.get('date', '')
                
                # Fill only this floor's data
                current_row = 12
                for canopy in floor.get('canopies', []):
                    new_sheet[f'C{current_row + 2}'] = canopy.get('configuration', '')
                    new_sheet[f'C{current_row}'] = canopy.get('itemNum', '')
                    new_sheet[f'D{current_row + 2}'] = canopy.get('model', '')
                    new_sheet[f'E{current_row + 2}'] = canopy.get('width', '')
                    new_sheet[f'F{current_row + 2}'] = canopy.get('length', '')
                    new_sheet[f'G{current_row + 2}'] = canopy.get('height', '')
                    new_sheet[f'H{current_row + 2}'] = canopy.get('section', '')
                    new_sheet[f'I{current_row + 2}'] = canopy.get('flowrate', '')
                    
                    new_sheet[f'C{current_row + 3}'] = canopy.get('lights', '')
                    new_sheet[f'D{current_row + 3}'] = canopy.get('light_quantity', '')
                    
                    special_works_dict = canopy.get('specialWorks', {})
                    for i, (work, qty) in enumerate(list(special_works_dict.items())[:2]):
                        new_sheet[f'C{current_row + 5 + i}'] = work
                        new_sheet[f'D{current_row + 5 + i}'] = qty
                    
                    new_sheet[f'C{current_row + 7}'] = "2M² (HFL)" if canopy.get('wallCladding') else ''
                    
                    if canopy.get('model') in ['CMWI', 'CMWF']:
                        new_sheet[f'C{current_row + 13}'] = canopy.get('control_panel', '')
                        new_sheet[f'C{current_row + 14}'] = canopy.get('WW_pods', '')
                        new_sheet[f'D{current_row + 14}'] = canopy.get('WW_pods_quantity', '')
                        new_sheet[f'C{current_row + 15}'] = canopy.get('pipework', '')
                    
                    current_row += 17
                
                # Fill floor-specific delivery data from the floor's own data
                delivery_data = floor.get('delivery_install_data', {})
                
                # Only write to input cells, preserve calculation cells
                new_sheet['D183'] = delivery_data.get('delivery_location', '')  # Location (input)
                new_sheet['C183'] = delivery_data.get('delivery_lift_qty', '')  # Quantity (input)
                
                if delivery_data.get('plant_hires'):
                    if "Plant Hire 1" in delivery_data['plant_hires']:
                        new_sheet['D184'] = delivery_data['plant_hires']["Plant Hire 1"]  # Plant type (input)
                        new_sheet['C184'] = delivery_data['quantities'].get("Plant Hire 1", 0)  # Quantity (input)
                    if "Plant Hire 2" in delivery_data['plant_hires']:
                        new_sheet['D185'] = delivery_data['plant_hires']["Plant Hire 2"]  # Plant type (input)
                        new_sheet['C185'] = delivery_data['quantities'].get("Plant Hire 2", 0)  # Quantity (input)

                # Don't overwrite calculation cells C187-C197
                # These cells contain formulas that should be preserved
                
        # Remove temporary template
        wb.remove(clean_template)
        
        # Save workbook
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        wb.close()
        
        return output

    except Exception as e:
        st.error(f"Error generating Excel sheet: {str(e)}")
        return None