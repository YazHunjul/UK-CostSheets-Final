from openpyxl import load_workbook
import streamlit as st
import math
import time

def extract_canopy_prices(excel_path):
    """
    Extract prices by reading individual components
    """
    try:
        time.sleep(5)  # Wait for file to be ready
        wb = load_workbook(excel_path, data_only=True)
        ws = wb['CANOPY']
        
        canopy_prices = []
        row = 14  # Start at first canopy's details
        
        while row <= ws.max_row:
            # Read all the component values that make up the total
            base_price = ws[f'K{row}'].value or 0  # Base price
            light_price = ws[f'K{row+1}'].value or 0  # Lighting
            special_works_1 = ws[f'K{row+3}'].value or 0  # First special work
            special_works_2 = ws[f'K{row+4}'].value or 0  # Second special work
            wall_cladding = ws[f'K{row+5}'].value or 0  # Wall cladding
            control_panel = ws[f'K{row+7}'].value or 0  # Control panel
            ww_pods = ws[f'K{row+8}'].value or 0  # WW pods
            pipework = ws[f'K{row+9}'].value or 0  # Pipework
            
            # Sum all components
            total_value = sum([
                base_price,
                light_price,
                special_works_1,
                special_works_2,
                wall_cladding,
                control_panel,
                ww_pods,
                pipework
            ])
            
            st.write(f"\nCanopy at row {row}:")
            st.write(f"Base Price: £{base_price}")
            st.write(f"Light Price: £{light_price}")
            st.write(f"Special Works 1: £{special_works_1}")
            st.write(f"Special Works 2: £{special_works_2}")
            st.write(f"Wall Cladding: £{wall_cladding}")
            st.write(f"Control Panel: £{control_panel}")
            st.write(f"WW Pods: £{ww_pods}")
            st.write(f"Pipework: £{pipework}")
            st.write(f"Calculated Total: £{total_value}")
            
            if total_value > 0:  # Only add if we found values
                rounded_value = math.ceil(total_value)
                canopy_prices.append({
                    'canopy_number': len(canopy_prices) + 1,
                    'total_price': rounded_value
                })
            else:
                break  # Stop if we don't find any values
            
            row += 17  # Move to next canopy section
        
        wb.close()
        return canopy_prices
        
    except Exception as e:
        st.error(f"Error extracting prices: {str(e)}")
        st.error(f"At row: {row}")  # Show which row caused the error
        return [] 