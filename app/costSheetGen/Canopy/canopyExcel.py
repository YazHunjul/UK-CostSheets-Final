from openpyxl import load_workbook
import streamlit as st
import io
import os
import openpyxl
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import CellIsRule
from copy import copy
import math

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
        
        # Write general info to main CANOPY sheet
        template_ws['C3'] = genInfo.get('projectNum', '').title()
        template_ws['C5'] = genInfo.get('customer', '').title()
        template_ws['C7'] = genInfo.get('combined_initials', '')
        template_ws['G3'] = genInfo.get('projectName', '').title()
        template_ws['G5'] = genInfo.get('location', '').title()
        template_ws['G7'] = genInfo.get('date', '')
        
        # First write all data to main CANOPY sheet
        current_row = 12
        
        # Create a dark blue font style
        dark_blue_font = Font(name='Calibri', size=11, color='000080')  # 000080 is dark blue
        
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
                    
                    # Apply dark blue font to J14-J27 and K14-K27
                    for row in range(14, 28):
                        # Format J column
                        j_cell = template_ws[f'J{row}']
                        if j_cell.value is not None:
                            j_cell.font = dark_blue_font
                        
                        # Format K column
                        k_cell = template_ws[f'K{row}']
                        if k_cell.value is not None:
                            k_cell.font = dark_blue_font
                    
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
                    
                    # Apply formatting to value cells
                    format_value_cells(new_sheet, 14)  # Start at row 14
                    
                    current_row += 17  # Move to next canopy section
                    color_row += 17    # Move to next color section
                
                # Fill floor-specific delivery data from the floor's own data
                delivery_data = floor.get('delivery_install_data', {})
                
                # Only write to input cells, preserve calculation cells
                new_sheet['D183'] = delivery_data.get('delivery_location', '')  # Location (input)
                new_sheet['C183'] = delivery_data.get('delivery_lift_qty', '')  # Quantity (input)
                
                # Only write the input value to C193
                new_sheet['C193'] = delivery_data.get('test_commission', 0)  # Write the test & commission value
                
                if delivery_data.get('plant_hires'):
                    if "Plant Hire 1" in delivery_data['plant_hires']:
                        new_sheet['D184'] = delivery_data['plant_hires']["Plant Hire 1"]  # Plant type (input)
                        new_sheet['C184'] = delivery_data['quantities'].get("Plant Hire 1", 0)  # Quantity (input)
                    if "Plant Hire 2" in delivery_data['plant_hires']:
                        new_sheet['D185'] = delivery_data['plant_hires']["Plant Hire 2"]  # Plant type (input)
                        new_sheet['C185'] = delivery_data['quantities'].get("Plant Hire 2", 0)  # Quantity (input)

                # Don't overwrite calculation cells C187-C197
                # These cells contain formulas that should be preserved
                
                # Apply formatting after writing values
                format_currency_cells(new_sheet, current_row)
                
                # Get commissioning price from N193
                commission_cell = new_sheet['N193']
                try:
                    # Try to get the value directly first
                    commission_price = float(commission_cell.value if commission_cell.value else 0)
                except (ValueError, TypeError):
                    # If it's a formula, try to get the calculated value
                    wb_data = load_workbook(excel_path, data_only=True)  # Load with data_only=True to get calculated values
                    sheet_name = new_sheet.title
                    if sheet_name in wb_data.sheetnames:
                        data_sheet = wb_data[sheet_name]
                        commission_price = float(data_sheet['N193'].value if data_sheet['N193'].value else 0)
                    else:
                        commission_price = 0
                    wb_data.close()
                
                # Store commissioning price in floor data
                floor['commission_price'] = commission_price
                
                # Calculate floor subtotal and cladding total
                floor_subtotal, cladding_total = calculate_floor_subtotal(
                    floor.get('canopies', []),
                    floor.get('delivery_install_data', {}),
                    commission_price
                )
                
                # Store both totals in floor data
                floor['subtotal'] = floor_subtotal
                floor['cladding_total'] = cladding_total
                
                # Write subtotal to Excel
                subtotal_row = current_row + 15
                new_sheet[f'N{subtotal_row}'] = floor_subtotal
                
                # Write cladding total somewhere if needed
                # new_sheet[f'M{subtotal_row}'] = cladding_total  # Adjust cell reference as needed
                
                # Apply formatting to subtotal
                subtotal_cell = new_sheet[f'N{subtotal_row}']
                apply_cell_formatting(subtotal_cell)
                
                # Move to next section
                current_row += 17
        
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

def apply_cell_formatting(cell):
    """Apply standard cell formatting for currency values"""
    # Light blue background - exact color from image
    cell.fill = PatternFill(start_color='B8CCE4', end_color='B8CCE4', fill_type='solid')
    
    # Thin borders on all sides
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    cell.border = thin_border
    
    # Calibri font
    cell.font = Font(name='Calibri', size=11)
    
    # Currency format with 2 decimal places
    cell.number_format = '£#,##0.00'
    
    # Alignment
    cell.alignment = Alignment(horizontal='right')

def format_currency_cells(worksheet, current_row):
    """Apply currency formatting to specific cells"""
    # List of cells that need currency formatting
    currency_cells = [
        'K9', 'N9',  # Top cells
        f'K{current_row}', f'N{current_row}',  # Current row cells
        'P182'  # Bottom cell
    ]
    
    for cell_ref in currency_cells:
        cell = worksheet[cell_ref]
        apply_cell_formatting(cell)

def apply_value_cell_formatting(cell, add_border=False):
    """Apply formatting for value cells in J, K, N, and O columns"""
    # Only add border if specified (for J column)
    if add_border:
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        cell.border = thin_border
    
    # Bold Calibri font in dark blue
    cell.font = Font(name='Calibri', size=11, bold=True, color='000080')  # 000080 is dark blue
    
    # Currency format with 2 decimal places
    cell.number_format = '£#,##0.00'
    
    # Right alignment
    cell.alignment = Alignment(horizontal='right')

def format_value_cells(worksheet, start_row=14):
    """Format the J, K, N, and O columns for the canopy section"""
    # Format J14-J27, K14-K27, N14-N27, O14-O27
    for row in range(start_row, start_row + 14):  # 14 rows from start_row
        # Format J column with border
        j_cell = worksheet[f'J{row}']
        # Always apply font formatting
        j_cell.font = Font(name='Calibri', size=11, bold=True, color='9FCBDA')  # Changed color to #9FCBDA
        j_cell.number_format = '£#,##0.00'
        j_cell.alignment = Alignment(horizontal='right')
        # Add border for J column
        j_cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Format K, N, O columns without border
        for col in ['K', 'N', 'O']:
            cell = worksheet[f'{col}{row}']
            # Always apply font formatting
            cell.font = Font(name='Calibri', size=11, bold=True, color='9FCBDA')  # Changed color to #9FCBDA
            cell.number_format = '£#,##0.00'
            cell.alignment = Alignment(horizontal='right')

def get_wall_cladding(kitchen_info):
    """
    Extracts wall cladding information only for canopies that have cladding.
    Returns None if no canopies have wall cladding.
    """
    kitchens = kitchen_info.get("kitchens", [])
    wall_cladding_data = []

    for kitchen in kitchens:
        for floor in kitchen["floors"]:
            for canopy in floor["canopies"]:
                # Only process canopies that have wallCladding set to True and have cladding_desc
                if canopy.get('wallCladding') and canopy.get('cladding_desc'):
                    # Get the canopy's item number
                    item_num = canopy.get('itemNum') or canopy.get('item_number', '1')
                    
                    # Get the selected walls
                    selected_walls = canopy.get('cladding_desc', [])
                    
                    if selected_walls:
                        wall_parts = []
                        if 'Rear' in selected_walls:  # Put Rear first
                            wall_parts.append('Rear')
                        if 'Left' in selected_walls:
                            wall_parts.append('Left')
                        if 'Right' in selected_walls:
                            wall_parts.append('Right')
                        
                        if wall_parts:
                            # Format as "Rear, Left & Right-hand Walls"
                            if len(wall_parts) > 1:
                                wall_description = f"Cladding to {', '.join(wall_parts[:-1])} & {wall_parts[-1]}-hand Walls"
                            else:
                                wall_description = f"Cladding to {wall_parts[0]}-hand Wall"
                        else:
                            continue  # Skip if no walls selected
                    else:
                        continue  # Skip if no cladding description

                    # Get cladding price and round up to nearest integer
                    cladding_price = canopy.get('cladding_price', 0)
                    if isinstance(cladding_price, (int, float)):
                        cladding_price = math.ceil(cladding_price)

                    wall_cladding_data.append({
                        "item_no": str(item_num),  # Use the canopy's item number
                        "description": wall_description,
                        "width": canopy.get("cladding_width", 0),
                        "height": canopy.get("cladding_height", 0),
                        "price": cladding_price
                    })

    # Only return data if we found canopies with cladding
    return wall_cladding_data if wall_cladding_data else None

def calculate_cladding_subtotal(canopies):
    """Calculate total cladding price for a floor's canopies"""
    cladding_total = 0
    for canopy in canopies:
        if canopy.get('wallCladding'):
            cladding_price = float(canopy.get('cladding_price', 0))
            cladding_total += cladding_price
    return round(cladding_total, 2)

def calculate_floor_subtotal(floor_canopies, delivery_install_data, commission_price=0):
    """Calculate subtotal for a floor including all canopies, P182, cladding, delivery/installation, and commissioning"""
    subtotal = 0
    
    # Sum all canopy prices and P182
    for canopy in floor_canopies:
        canopy_price = float(canopy.get('total_price', 0))
        p182_value = float(canopy.get('p182_value', 0))
        subtotal += canopy_price + p182_value
    
    # Calculate cladding total
    cladding_total = calculate_cladding_subtotal(floor_canopies)
    subtotal += cladding_total
    
    # Add delivery and installation costs
    delivery_price = float(delivery_install_data.get('delivery_price', 0))
    install_price = float(delivery_install_data.get('install_price', 0))
    
    # Add commissioning price
    subtotal += delivery_price + install_price + commission_price
    
    return round(subtotal, 2), round(cladding_total, 2)