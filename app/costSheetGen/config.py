import os

# Base directory of the package
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Template paths
TEMPLATES = {
    'EXCEL': os.path.join(BASE_DIR, 'costSheetGen', 'costSheetResources', 'Halton Cost Sheet Jan 2025.xlsx'),
    'WORD': os.path.join(BASE_DIR, 'costSheetGen', 'costSheetResources', 'costSheet_canopy.docx')
} 