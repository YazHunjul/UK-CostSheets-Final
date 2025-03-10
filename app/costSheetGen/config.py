import os

# Base directory of the package
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Resources directory
RESOURCES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'costSheetResources')

# Template paths
TEMPLATES = {
    'EXCEL': os.path.join(RESOURCES_DIR, 'Halton Cost Sheet Jan 2025.xlsx'),
    'WORD': os.path.join(RESOURCES_DIR, 'costSheet_canopy.docx')
}

# Debug output
print("BASE_DIR:", BASE_DIR)
print("RESOURCES_DIR:", RESOURCES_DIR)
print("EXCEL path:", TEMPLATES['EXCEL'])
print("WORD path:", TEMPLATES['WORD']) 