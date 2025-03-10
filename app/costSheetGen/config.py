import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get template paths from environment variables or use defaults
TEMPLATES = {
    'EXCEL': os.getenv('EXCEL_TEMPLATE_PATH', 
        '/Users/yazan/Desktop/Efficiency/UK-CostSheets-Final/app/costSheetGen/costSheetResources/Halton Cost Sheet Jan 2025.xlsx'),
    'WORD': os.getenv('WORD_TEMPLATE_PATH',
        '/Users/yazan/Desktop/Efficiency/UK-CostSheets-Final/app/costSheetGen/costSheetResources/costSheet_canopy.docx')
}

# Debug output
print("EXCEL path:", TEMPLATES['EXCEL'])
print("WORD path:", TEMPLATES['WORD']) 