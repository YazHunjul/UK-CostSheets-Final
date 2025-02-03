import openpyxl
import math
import streamlit as st
import os
import pandas as pd
from os import path
from xlcalculator import ModelCompiler
from xlcalculator import Model
from xlcalculator import Evaluator
import xlwings as xw
from io import BytesIO

def extract_canopy_prices(excel_file):
    """Extract prices from Excel file"""
    try:
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        ws = wb['CANOPY']
        
        # Get prices
        prices = {}
        current_row = 12
        
        while ws[f'C{current_row}'].value:  # While there's an item number
            item_num = ws[f'C{current_row}'].value
            model = ws[f'D{current_row + 2}'].value
            
            # Get canopy price
            price = ws[f'P{current_row}'].value
            if price:
                prices[item_num] = float(f"{math.ceil(float(price))}.00")
            
            # Get cladding price if exists
            cladding_price = ws[f'N{current_row + 7}'].value
            if cladding_price:
                prices[f"{item_num}_cladding"] = float(f"{math.ceil(float(cladding_price))}.00")
            
            # Get UV price if it's a UV model
            if model and 'UV' in str(model):
                uv_price = ws[f'N{current_row + 12}'].value
                if uv_price:
                    prices[f"{item_num}_uv"] = float(f"{math.ceil(float(uv_price))}.00")
            
            current_row += 17
        
        return prices
        
    except Exception as e:
        st.error(f"Error extracting prices: {str(e)}")
        return None

def convert_formulas_to_values(excel_file):
    """Convert formulas to values in Excel file"""
    try:
        # Load workbook with data_only=True to get values
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        return wb
    except Exception as e:
        st.error(f"Error converting formulas: {str(e)}")
        return None

def get_excel_calculated_values(excel_path):
    """
    Reads Excel file and saves as CSV for inspection.
    """
    try:
        # Read Excel file
        df = pd.read_excel(excel_path, sheet_name='CANOPY', header=None)
        
        # Save as CSV in the same directory
        csv_path = excel_path.replace('.xlsx', '.csv')
        df.to_csv(csv_path, index=False)
        
        # Print file location and first few rows
        print(f"\nCSV saved at: {csv_path}")
        print("\nFirst 20 rows of CSV:")
        print(df.head(20).to_string())
        
        # Print specific rows we're interested in
        print("\nSpecific Rows:")
        print("\nRow 183 (Delivery):")
        print(df.iloc[182].to_string())
        print("\nRow 189 (Installation):")
        print(df.iloc[188].to_string())
        print("\nRow 200 (Total):")
        print(df.iloc[199].to_string())
        
        # Print canopy rows
        print("\nCanopy Rows (every 17 rows):")
        for row in range(10, len(df), 17):
            print(f"\nRow {row+1}:")
            print(df.iloc[row].to_string())
        
        return csv_path
        
    except Exception as e:
        print(f"Error reading Excel values: {str(e)}")
        print(f"At path: {excel_path}")
        return None

def run_excel_script(excel_file):
    """Run Excel script to update values"""
    try:
        # Load workbook
        wb = openpyxl.load_workbook(excel_file)
        ws = wb['CANOPY']
        
        # Save and return
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error running Excel script: {str(e)}")
        return None

def convert_to_macro_enabled(excel_bytes):
    """
    Converts Excel bytes to a macro-enabled workbook
    """
    try:
        st.write("Converting to macro-enabled workbook...")
        
        # Create a temporary file to write the bytes
        temp_path = "temp_original.xlsx"
        with open(temp_path, "wb") as f:
            f.write(excel_bytes.getvalue())
        
        # Load the workbook
        wb = openpyxl.load_workbook(temp_path)
        
        # Save as macro-enabled workbook
        macro_path = "temp_macro.xlsm"
        wb.save(macro_path)
        wb.close()
        
        # Read the macro-enabled file back
        with open(macro_path, "rb") as f:
            macro_bytes = BytesIO(f.read())
        
        # Clean up temp files
        os.remove(temp_path)
        os.remove(macro_path)
        
        st.write("Conversion complete")
        return macro_bytes
        
    except Exception as e:
        st.error(f"Error converting to macro-enabled: {str(e)}")
        return None

def read_excel_value(excel_path):
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb[wb.sheetnames[0]]
        value = ws['N12'].value
        wb.close()
        return value
    except Exception as e:
        st.error(f"Error reading Excel: {str(e)}")
        return None