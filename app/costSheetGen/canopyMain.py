# Upload section
st.markdown("### Step 2: Upload Processed Excel")
uploaded_file = st.file_uploader("Upload Excel file after calculations", type=['xlsx'])

if uploaded_file is not None:
    try:
        st.write("Loading Excel file...")
        # Create a copy of the uploaded file in memory
        excel_data = BytesIO(uploaded_file.getvalue())
        
        # Load workbook
        wb_data = openpyxl.load_workbook(excel_data, data_only=True)
        excel_data.seek(0)
        wb = openpyxl.load_workbook(excel_data, data_only=False)
        
        # Save the Excel file to BytesIO
        modified_excel = BytesIO()
        wb.save(modified_excel)
        modified_excel.seek(0)
        
        # Initialize total P182
        total_p182 = 0
        
        # Iterate over each floor sheet to collect P182 values
        for sheet in wb.worksheets:
            if sheet.title != 'CANOPY' and sheet.title.startswith('CANOPY - '):
                p182_value = sheet['P182'].value
                
                if p182_value is not None:
                    try:
                        if isinstance(p182_value, str):
                            p182_value = float(p182_value.replace('£', '').replace(',', '').strip())
                        elif isinstance(p182_value, (int, float)):
                            p182_value = float(p182_value)
                        total_p182 += p182_value
                    except (ValueError, TypeError) as e:
                        st.error(f"Error processing P182: {e}")
        
        # Write total to main CANOPY sheet
        main_sheet = wb['CANOPY']
        main_sheet['P182'] = total_p182
        
        # Save the modified workbook before creating ZIP
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Create ZIP with updated Excel file
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zf:
            # Add updated Excel file
            excel = f"{genInfo['projectNum']} Cost Sheet {genInfo['date']}.xlsx"
            output.seek(0)
            zf.writestr("Cost Sheet.xlsx", output.getvalue())
            
            # Add Word file
            word_context = {'kitchens': kitchen_info}
            word_file = CW.generate_word(word_context, genInfo)
            zf.writestr("Halton Quotation.docx", word_file.getvalue())
        
        zip_buffer.seek(0)
        
        # Provide download button for the ZIP file
        st.download_button(
            label="⬇️ Download Final Package",
            data=zip_buffer,
            file_name="Cost_Sheet_and_Quotation.zip",
            mime="application/zip"
        )

    except Exception as e:
        st.error(f"An error occurred processing the file: {str(e)}") 