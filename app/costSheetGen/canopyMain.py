if st.button("Generate Excel File"):
    try:
        excel_bytes = CE.generate_sheet(kitchen_info, genInfo)
        if excel_bytes is None:
            st.error("Error generating Excel sheet: No data returned.")
            return
        
        st.download_button(
            label="⬇️ Download Excel File",
            data=excel_bytes.getvalue(),
            file_name=f"{genInfo['projectNum']} Cost Sheet {genInfo['date']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("1. Download the Excel file\n2. Open it in Excel\n3. Let Excel calculate the values\n4. Make Sure To Save the file and upload below")
        
    except Exception as e:
        st.error(f"Error generating Excel: {str(e)}")

if uploaded_file is not None:
    try:
        st.write("Loading Excel file...")
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # Initialize total P182
        total_p182 = 0
        st.write("\n🧮 Calculating delivery & installation total:")
        
        # Iterate over each floor sheet to collect P182 values
        for sheet in wb.worksheets:
            if sheet.title != 'CANOPY' and sheet.title.startswith('CANOPY - '):
                st.write(f"\nChecking sheet: {sheet.title}")
                p182_value = sheet['P182'].value
                
                if p182_value is not None:
                    try:
                        if isinstance(p182_value, str):
                            p182_value = float(p182_value.replace('£', '').replace(',', '').strip())
                        elif isinstance(p182_value, (int, float)):
                            p182_value = float(p182_value)
                        total_p182 += p182_value
                        st.write(f"  Found value: £{p182_value:.2f}")
                        st.write(f"  Running total: £{total_p182:.2f}")
                    except (ValueError, TypeError) as e:
                        st.error(f"Error processing P182: {e}")
                else:
                    st.warning(f"No P182 value found in {sheet.title}")
        
        st.write(f"\n💰 Total delivery & installation: £{total_p182:.2f}")
        
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
        import traceback
        st.write("Full error:", traceback.format_exc()) 