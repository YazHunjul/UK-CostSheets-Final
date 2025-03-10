import os
import streamlit as st
from pkg_resources import resource_filename

# Get template paths using pkg_resources
TEMPLATES = {
    'EXCEL': resource_filename('app.costSheetGen.costSheetResources', 'Halton Cost Sheet Jan 2025.xlsx'),
    'WORD': resource_filename('app.costSheetGen.costSheetResources', 'costSheet_canopy.docx')
}

# Debug output
st.write("EXCEL path:", TEMPLATES['EXCEL'])
st.write("WORD path:", TEMPLATES['WORD'])

# Verify files exist
if not os.path.exists(TEMPLATES['EXCEL']):
    st.error(f"Excel template not found! Looking in: {TEMPLATES['EXCEL']}")
if not os.path.exists(TEMPLATES['WORD']):
    st.error(f"Word template not found! Looking in: {TEMPLATES['WORD']}") 