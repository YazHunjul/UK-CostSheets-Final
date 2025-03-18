from openpyxl import load_workbook
import streamlit as st
import io
import os
import openpyxl
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.formatting.rule import CellIsRule

def adjust_gp_to_target(wb, target_gp=44.0):
    """
    Adjust the GP to target by modifying C9
    Args:
        wb: Workbook object
        target_gp: Target GP percentage (default 44.0)
    """
    try:
        sheet = wb['CANOPY']
        current_gp = float(sheet['K14'].value or 0)  # Current GP percentage
        
        if current_gp < target_gp:
            # Calculate how much we need to increase by
            # Start with a small increment and adjust C9 until we hit target
            c9_value = 0
            step = 0.1
            
            while current_gp < target_gp and c9_value <= 100:
                c9_value += step
                sheet['C9'] = c9_value
                # Recalculate the GP value
                current_gp = float(sheet['K14'].value or 0)
            
            return c9_value
    except Exception as e:
        st.error(f"Error adjusting GP: {str(e)}")
        return None

def generate_sheet(kitchen_info, genInfo):
    """
    Generate an Excel sheet based on kitchen information and return it as a BytesIO object.
    """
    try:
        # Load the workbook
        excel_path = 'app/costSheetGen/costSheetResources/Halton Cost Sheet Jan 2025.xlsx'
        wb = load_workbook(excel_path, data_only=False)  # Keep formulas
        
        # Keep a clean template
        template_ws = wb['CANOPY']
        clean_template = wb.copy_worksheet(template_ws)
        clean_template.title = 'TEMPLATE_TEMP'
        
        # Adjust GP to 44% by modifying C9
        c9_value = adjust_gp_to_target(wb)
        if c9_value:
            template_ws['C9'] = c9_value
        
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
                    # Get the reference number (try both possible keys)
                    ref_num = canopy.get('itemNum') or canopy.get('item_number') or canopy.get('reference_number', '')
                    
                    # Write reference number ONLY to B column
                    template_ws[f'B{current_row}'] = ref_num
                    
                    # Write to main CANOPY sheet (skip C12)
                    template_ws[f'C{current_row + 2}'] = canopy.get('configuration', '')
                    template_ws[f'D{current_row + 2}'] = canopy.get('model', '')
                    template_ws[f'E{current_row + 2}'] = canopy.get('width', '')
                    template_ws[f'F{current_row + 2}'] = canopy.get('length', '')
                    template_ws[f'G{current_row + 2}'] = canopy.get('height', '')
                    template_ws[f'H{current_row + 2}'] = canopy.get('section', '')
                    template_ws[f'I{current_row + 2}'] = canopy.get('flowrate', '')
                    
                    # Set lights and quantities
                    template_ws[f'C{current_row + 3}'] = canopy.get('lights', '')
                    template_ws[f'D{current_row + 3}'] = canopy.get('light_quantity', '')
                    
                    # Add light indicator in column C for main sheet
                    if canopy.get('lights') or canopy.get('light_type'):
                        light_indicator_row = current_row + 11  # This will be 23 for first canopy
                        template_ws[f'C{light_indicator_row}'] = 1
                    
                    # Set special works in template sheet
                    special_works_dict = canopy.get('specialWorks', {})
                    special_works_items = list(special_works_dict.items())[:2]
                    
                    # First special work
                    if len(special_works_items) > 0:
                        work, qty = special_works_items[0]
                        template_ws[f'C{current_row + 4}'] = work
                        template_ws[f'D{current_row + 4}'] = qty
                    
                    # Second special work - default to BIM/REVIT unless user specified something else
                    if len(special_works_items) > 1:
                        # User specified a second special work
                        work, qty = special_works_items[1]
                        template_ws[f'C{current_row + 5}'] = work
                        template_ws[f'D{current_row + 5}'] = qty if qty else 1
                    else:
                        # Set default BIM/REVIT
                        template_ws[f'C{current_row + 5}'] = 'BIM/ REVIT per CANOPY'
                        template_ws[f'D{current_row + 5}'] = 1
                    
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
                
                # Fill general info
                new_sheet['C3'] = genInfo.get('projectNum', '').title()
                new_sheet['C5'] = genInfo.get('customer', '').title()
                new_sheet['C7'] = genInfo.get('combined_initials', '')
                new_sheet['G3'] = genInfo.get('projectName', '').title()
                new_sheet['G5'] = genInfo.get('location', '').title()
                new_sheet['G7'] = genInfo.get('date', '')
                
                # Start data at row 12
                current_row = 12
                
                # Start color at row 22 and increment
                color_row = 22
                # Create font with bold and teal color
                teal_bold_font = Font(
                    color='1F9D9D',  # Teal/greenish-blue color
                    bold=True
                )
                
                for canopy in floor.get('canopies', []):
                    # Get the reference number (try both possible keys)
                    ref_num = canopy.get('itemNum') or canopy.get('item_number') or canopy.get('reference_number', '')
                    
                    # Write reference number ONLY to B column
                    new_sheet[f'B{current_row}'] = ref_num
                    
                    # Set configuration first (matches template)
                    new_sheet[f'C{current_row + 2}'] = canopy.get('configuration', '')
                    new_sheet[f'D{current_row + 2}'] = canopy.get('model', '')
                    
                    # Set item number and model
                    new_sheet[f'E{current_row + 2}'] = canopy.get('width', '')
                    new_sheet[f'F{current_row + 2}'] = canopy.get('length', '')
                    
                    # Apply bold teal font to E and F at color_row and color_row + 1
                    for col in ['E', 'F']:
                        cell = new_sheet[f'{col}{color_row}']
                        next_cell = new_sheet[f'{col}{color_row + 1}']
                        cell.font = teal_bold_font
                        next_cell.font = teal_bold_font
                    
                    # Set remaining dimensions
                    new_sheet[f'G{current_row + 2}'] = canopy.get('height', '')
                    new_sheet[f'H{current_row + 2}'] = canopy.get('section', '')
                    new_sheet[f'I{current_row + 2}'] = canopy.get('flowrate', '')
                    
                    # Set lights and quantities
                    new_sheet[f'C{current_row + 3}'] = canopy.get('lights', '')
                    new_sheet[f'D{current_row + 3}'] = canopy.get('light_quantity', '')
                    
                    # Add light indicator in column C for individual sheet
                    if canopy.get('lights') or canopy.get('light_type'):
                        light_indicator_row = current_row + 11  # This will be 23 for first canopy
                        new_sheet[f'C{light_indicator_row}'] = 1
                    
                    # Set special works
                    special_works_dict = canopy.get('specialWorks', {})
                    special_works_items = list(special_works_dict.items())[:2]
                    
                    # First special work
                    if len(special_works_items) > 0:
                        work, qty = special_works_items[0]
                        new_sheet[f'C{current_row + 4}'] = work
                        new_sheet[f'D{current_row + 4}'] = qty
                    
                    # Second special work - default to BIM/REVIT unless user specified something else
                    if len(special_works_items) > 1:
                        # User specified a second special work
                        work, qty = special_works_items[1]
                        new_sheet[f'C{current_row + 5}'] = work
                        new_sheet[f'D{current_row + 5}'] = qty if qty else 1
                    else:
                        # Set default BIM/REVIT
                        new_sheet[f'C{current_row + 5}'] = 'BIM/ REVIT per CANOPY'
                        new_sheet[f'D{current_row + 5}'] = 1
                    
                    # Set wall cladding
                    new_sheet[f'C{current_row + 7}'] = "2M² (HFL)" if canopy.get('wallCladding') else ''
                    
                    # Set CMWI/CMWF specific fields
                    if canopy.get('model') in ['CMWI', 'CMWF']:
                        new_sheet[f'C{current_row + 13}'] = canopy.get('control_panel', '')
                        new_sheet[f'C{current_row + 14}'] = canopy.get('WW_pods', '')
                        new_sheet[f'D{current_row + 14}'] = canopy.get('WW_pods_quantity', '')
                        new_sheet[f'C{current_row + 15}'] = canopy.get('pipework', '')
                    
                    current_row += 17  # Move to next canopy data section
                    color_row += 17    # Move to next color section
                
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

def generate_excel(kitchen_info, genInfo):
    """
    Generate an Excel file with canopy data.
    
    Args:
        kitchen_info (dict): Dictionary containing kitchen and canopy data.
        genInfo (dict): General information for the document.
        
    Returns:
        BytesIO: Excel file as a BytesIO object.
    """
    # Load the template
    template_path = TEMPLATES['EXCEL']
    wb = load_workbook(template_path)
    
    # Create a new workbook for each kitchen and floor
    for kitchen in kitchen_info:
        kitchen_name = kitchen['kitchen_name']
        
        for floor in kitchen['floors']:
            floor_name = floor['floor_name']
            sheet_name = f"CANOPY - {kitchen_name[:8]} {floor_name[:8]}"
            
            # Copy the template sheet
            template_sheet = wb['CANOPY']
            new_sheet = wb.copy_worksheet(template_sheet)
            new_sheet.title = sheet_name[:31]  # Excel limits sheet names to 31 chars
            
            # Set header information
            new_sheet['C3'] = genInfo.get('projectName', '')
            new_sheet['C4'] = genInfo.get('projectNum', '')
            new_sheet['C5'] = genInfo.get('projectManager', '')
            new_sheet['C6'] = genInfo.get('salesPerson', '')
            new_sheet['C7'] = f"{kitchen_name} - {floor_name}"
            
            # Set canopy data
            current_row = 22  # Starting row for canopy data
            
            # Define single color for E and F columns
            text_color = '0000FF'  # Blue - or whichever single color you want
            
            for canopy in floor['canopies']:
                # Set cell values
                new_sheet[f'C{current_row}'] = canopy.get('itemNum', '')
                new_sheet[f'D{current_row}'] = canopy.get('model', '')
                new_sheet[f'E{current_row}'] = canopy.get('width', 0)
                new_sheet[f'F{current_row}'] = canopy.get('length', 0)

                # Set font color for E and F for both current row and row below
                for col in ['E', 'F']:
                    # Set font color for current row and row below
                    for row_offset in [0, 1]:
                        cell = new_sheet[f'{col}{current_row + row_offset}']
                        # Preserve existing font properties, only change color
                        old_font = cell.font
                        new_font = Font(
                            name=old_font.name,
                            size=old_font.size,
                            bold=old_font.bold,
                            italic=old_font.italic,
                            vertAlign=old_font.vertAlign,
                            underline=old_font.underline,
                            strike=old_font.strike,
                            color=text_color  # Use the single consistent color
                        )
                        cell.font = new_font

                # Set other canopy data
                new_sheet[f'G{current_row}'] = canopy.get('height', 0)
                new_sheet[f'H{current_row}'] = canopy.get('section', 0)
                new_sheet[f'I{current_row}'] = canopy.get('flowrate', 0)
                
                # Move to next canopy row
                current_row += 17  # Each canopy takes 17 rows
    
    # Remove the template sheet
    wb.remove(wb['CANOPY'])
    
    # Save to BytesIO
    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    
    return excel_buffer