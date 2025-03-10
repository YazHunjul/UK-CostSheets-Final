import os

# Template paths - using relative paths from the config.py location
current_dir = os.path.dirname(os.path.abspath(__file__))
TEMPLATES = {
    'EXCEL': os.path.join(current_dir, 'costSheetResources', 'Halton Cost Sheet Jan 2025.xlsx'),
    'WORD': os.path.join(current_dir, 'costSheetResources', 'costSheet_canopy.docx')
}

# Debug output
print("EXCEL path:", TEMPLATES['EXCEL'])
print("WORD path:", TEMPLATES['WORD'])

# Verify files exist
if not os.path.exists(TEMPLATES['EXCEL']):
    print(f"Excel template not found! Looking in: {TEMPLATES['EXCEL']}")
if not os.path.exists(TEMPLATES['WORD']):
    print(f"Word template not found! Looking in: {TEMPLATES['WORD']}") 