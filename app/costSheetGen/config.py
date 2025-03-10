import os
import streamlit as st

# Get template paths from Streamlit secrets or use defaults
TEMPLATES = {
    'EXCEL': st.secrets.get('EXCEL_TEMPLATE_PATH', 
        '/Users/yazan/Desktop/Efficiency/UK-CostSheets-Final/app/costSheetGen/costSheetResources/Halton Cost Sheet Jan 2025.xlsx'),
    'WORD': st.secrets.get('WORD_TEMPLATE_PATH',
        '/Users/yazan/Desktop/Efficiency/UK-CostSheets-Final/app/costSheetGen/costSheetResources/costSheet_canopy.docx')
}

# Debug output
print("EXCEL path:", TEMPLATES['EXCEL'])
print("WORD path:", TEMPLATES['WORD']) 