from docxtpl import DocxTemplate
from io import BytesIO
import os
from collections import Counter
import math
import requests
import streamlit as st
import toml

def generate_word(context, genInfo):
    """
    Generates a Word document using the provided context.

    Args:
        context (dict): A dictionary containing the data to render in the Word template.

    Returns:
        BytesIO: The generated Word document as a BytesIO object.
    """
    try:
        # Format the reference number for the Word document
        ref_num = genInfo.get('projectNum', '')
        genInfo['referenceNum'] = f"{ref_num}/{genInfo['combined_initials']}"

        # Path to the Word template
        template_path = 'app/costSheetGen/costSheetResources/costSheet_canopy.docx'
        
        # Load the Word template
        template = DocxTemplate(template_path)

        # Extract and process canopy data grouped by kitchens and floors
        grouped_canopy_data = extract_canopy_info_grouped(context)
        cmwi_data = extract_cmwi_canopies(context)
        scope_work = scope_of_works(context)
        wall_cladding_data = get_wall_cladding(context)

        print("DEBUG - Context data:", context)  # Debug print
        print("DEBUG - Wall cladding data:", wall_cladding_data)  # Debug print
        print("DEBUG - Scope of work:", scope_work)  # Debug print

        # Validate data
        if not grouped_canopy_data:
            raise ValueError("No grouped canopy data available.")
        if not cmwi_data:
            cmwi_data = []  # Ensure it's at least an empty list
        if not scope_work:
            scope_work = []

        # Update genInfo with processed data
        genInfo["grouped_canopy_data"] = grouped_canopy_data
        genInfo["cmwi_canopies"] = cmwi_data
        genInfo["scope_of_work"] = scope_work
        if wall_cladding_data:
            genInfo["wall_cladding_info"] = wall_cladding_data
            print("Adding wall cladding to template")
        else:
            genInfo["wall_cladding_info"] = None
            print("No wall cladding data to add")

        print("DEBUG - Final genInfo:", genInfo)  # Debug print

        # Calculate totals for each floor and kitchen
        kitchen_totals = {}
        grand_total = 0
        
        for kitchen_name, floors in grouped_canopy_data.items():
            kitchen_total = 0
            for floor_name, floor_data in floors.items():
                floor_total = 0
                cladding_total = 0
                has_uv = False
                
                # Calculate floor totals
                for canopy in floor_data['canopies']:
                    if 'total_price' in canopy:
                        floor_total += canopy['total_price']
                    if canopy.get('wallCladding'):
                        cladding_total += 3520.00  # Price per cladding
                    if 'UV' in canopy.get('model', ''):
                        has_uv = True
                
                # Add delivery and commissioning to floor total
                floor_total += 13863.00  # Example delivery price
                floor_total += 1502.00   # Example commissioning price
                
                # Add UV-c price if applicable
                if has_uv:
                    floor_total += 1040.00
                
                # Store totals in floor data
                floor_data['floor_total'] = floor_total
                floor_data['cladding_total'] = cladding_total
                
                # Add to kitchen total
                kitchen_total += floor_total + cladding_total
            
            kitchen_totals[kitchen_name] = kitchen_total
            grand_total += kitchen_total

        # Add pricing information to genInfo
        genInfo.update({
            'delivery_price': 13863.00,
            'commissioning_price': 1502.00,
            'uvc_price': 1040.00,
            'cladding_price': 3520.00,
            'kitchen_totals': kitchen_totals,
            'grand_total': grand_total
        })

        # Render the template with the given context
        template.render(genInfo)

        # Save to BytesIO
        word_buffer = BytesIO()
        template.save(word_buffer)
        
        # Important: Seek to start of buffer
        word_buffer.seek(0)
        
        # Read the entire content into a new buffer to ensure it's complete
        final_buffer = BytesIO(word_buffer.read())
        word_buffer.close()
        
        return final_buffer

    except Exception as e:
        st.error(f"Error generating Word document: {str(e)}")
        return None

def scope_of_works(context):
    """
    Generates the Scope of Work text as a list based on the context.
    """
    canopy_counts = Counter()
    cladding_count = 0  # Back to simple counter
    
    kitchens = context.get('kitchens', [])
    for kitchen in kitchens:
        for floor in kitchen.get('floors', []):
            for canopy in floor.get('canopies', []):
                model = canopy.get('model', 'Unknown Model')
                canopy_counts[model] += 1
                
                if canopy.get("wallCladding") and canopy.get('cladding_desc'):
                    cladding_count += 1

    scope_lines = []
    
    for model, count in canopy_counts.items():
        if 'CXW' in model:
            scope_lines.append(f" {count}X {model} Condense Canopies")
        else:
            scope_lines.append(f" {count}X {model} Ventilation Canopies")
    
    # Simple cladding line if any exist
    if cladding_count > 0:
        scope_lines.append(f" {cladding_count}X Areas with Stainless Steel Cladding")

    print(f"Final scope lines: {scope_lines}")
    return scope_lines

def extract_canopy_info_grouped(kitchen_info):
    """
    Groups processed canopy data by kitchens and floors and adds Important Note calculations.

    Args:
        kitchen_info (dict): Dictionary containing the 'kitchens' key with a list of kitchen data.

    Returns:
        dict: A dictionary organized by kitchens and floors with processed canopy data.
    """
    kitchens = kitchen_info.get("kitchens", [])
    grouped_canopies = {}

    for kitchen in kitchens:
        kitchen_name = kitchen["kitchen_name"].title()
        grouped_canopies[kitchen_name] = {}

        for floor in kitchen["floors"]:
            floor_name = floor["floor_name"].title()
            display_name = f"{kitchen_name} – {floor_name}"
            
            grouped_canopies[kitchen_name][display_name] = {
                "canopies": [],
                "important_note": "",
                "cladding_total": 0,
                "uv_total": 0  # Add UV total
            }

            total_extract_volume = 0
            total_mua_volume = 0
            cladding_total = 0

            for canopy in floor["canopies"]:
                model = canopy.get("model", "")
                length = canopy.get("length", 0)
                sections = canopy.get("section", 0)
                flowrate = round(canopy.get("flowrate", 0), 3)

                # Calculate Supply Air for specific models
                supply_air = None
                if model in ['UVX-M', 'KVX-M', 'KVF', 'CMWF', 'UVF']:
                    # MUA Volume = (L - 100) × 0.225 / 1000
                    supply_air = round(((length - 100) * 0.225) / 1000, 3)
                    total_mua_volume += supply_air

                # Add to total extract volume (the full amount)
                total_extract_volume += flowrate

                if canopy.get('wallCladding'):
                    cladding_total += 3520.00

                # Add processed data to the current floor's list
                canopy_data = {
                    'item_number': canopy.get('itemNum', ''),
                    "model": model,
                    "length": length,
                    "width": canopy.get("width", 0),
                    "height": canopy.get("height", 0),
                    "sections": sections,
                    "flowrate": flowrate,
                    "supply_air": supply_air,
                    "f12": calculate_f12(length, sections),
                    "grease_filters": calculate_grease_filters(model, calculate_f12(length, sections), length, sections),
                    "extract_static_pa": calculate_extract_static_pa(calculate_grease_filters(model, calculate_f12(length, sections), length, sections), flowrate, model),
                    "wallCladding": canopy.get("wallCladding", ""),
                    "total_price": canopy.get("total_price", 0),
                    "cladding_price": canopy.get("cladding_price", 0),
                    "uv_price": canopy.get("uv_price", 0)  # Add UV price
                }

                # Add lights information if present
                if canopy.get('lights'):
                    canopy_data["lights"] = "LED Strip"  # Just set it to "LED Strip" instead of the full value
                    canopy_data["light_quantity"] = canopy.get('light_quantity')

                # Add to UV total if it's a UV model
                if 'UV' in model:
                    grouped_canopies[kitchen_name][display_name]["uv_total"] += canopy_data["uv_price"]

                grouped_canopies[kitchen_name][display_name]["canopies"].append(canopy_data)

            # Store the cladding total for this floor
            grouped_canopies[kitchen_name][display_name]["cladding_total"] = cladding_total

            # Calculate the important note with rounded values
            total_extract_volume = round(total_extract_volume, 3)
            total_mua_volume = round(total_mua_volume, 3)
            required_mua = round(total_extract_volume * 0.85, 3)
            shortfall = round(required_mua - total_mua_volume, 3)

            important_note = (
                f"The make-up air flows shown above are the maximum that we can introduce through the canopy. "
                f"This should be equal to approximately 85% of the extract i.e. {required_mua}m³/s. "
                f"In this instance it only totals {total_mua_volume}m³/s, therefore the shortfall of {shortfall}m³/s "
                f"must be introduced through ceiling grilles or diffusers, by others."
            )

            grouped_canopies[kitchen_name][display_name]["important_note"] = important_note

    return grouped_canopies

def extract_cmwi_canopies(kitchen_info):
    """
    Extracts data for CMWI canopies and calculates CWS and HWS requirements.

    Args:
        kitchen_info (dict): Dictionary containing the 'kitchens' key with a list of kitchen data.

    Returns:
        list: A list of dictionaries containing CMWI canopy data with calculated CWS and HWS values.
    """
    kitchens = kitchen_info.get("kitchens", [])
    cmwi_canopies = []

    for kitchen in kitchens:
        for floor in kitchen["floors"]:
            for canopy in floor["canopies"]:
                if "CMW" in canopy.get("model", ""):  # Filter for CMWI canopies
                    length = canopy.get("length", 0)

                    # Perform calculations
                    cws_continuous = round(length / 1000 * 0.02, 2)  # CWS @ 2 Bar (L/s)
                    hws_wash_cycle = round(length / 1000 * 0.103, 3)  # HWS @ 2 Bar (L/s)
                    hws_storage = round(hws_wash_cycle * 180, 3)  # HWS Storage (Litres)

                    cmwi_canopies.append({
                        "item_no": f"{canopy.get('itemNum', '')}",
                        "model": canopy["model"],
                        "cws_continuous": f"{cws_continuous} L/s",
                        "hws_wash_cycle": f"{hws_wash_cycle} L/s",
                        "hws_storage": f"{hws_storage} L",
                    })

    return cmwi_canopies

def calculate_f12(length, sections):
    """Calculate F12 (First Calculation)."""
    if sections < 1:
        return 0
    return math.ceil((length - 100) / sections / 250) * 250

def calculate_grease_filters(model, f12, length, sections):
    """Calculate the number of grease filters."""
    print(f"Inputs -> Model: {model}, F12: {f12}, Length: {length}, Sections: {sections}")
    
    if "CMW" in model:
        print("Model is CMW. Returning 1.")
        return 1
    
    if f12 == 0:
        print("F12 is 0. Returning 0.")
        return 0
    
    if sections < 1 or length < (100 + 50 * sections):
        print("Invalid sections or length. Returning 0.")
        return 0

    adjusted_length = (length - (100 + (50 * sections))) / sections
    print(f"Adjusted Length: {adjusted_length}")

    if adjusted_length < 500:
        print("Adjusted length is less than 500. Returning 0.")
        return 0

    filters_per_section = math.floor(adjusted_length / 500)
    print(f"Filters Per Section: {filters_per_section}")

    return filters_per_section * sections

def calculate_extract_static_pa(grease_filters, flow_rate, model):
    """Calculate Supply Static Pressure (Pa)."""
    i14 = 49.7 if "UV" in model else 71.75
    if grease_filters == 0:
        return "0 Pa"
    return f"{round((((flow_rate * 3600) / (grease_filters * i14)) ** 2) + 20, 1)}"

def get_wall_cladding(kitchen_info):
    """
    Extracts wall cladding information based on canopy data.
    Returns None if no canopies have wall cladding.
    """
    print("DEBUG - Starting wall cladding check")
    
    kitchens = kitchen_info.get("kitchens", [])
    wall_cladding_data = []

    for kitchen in kitchens:
        for floor in kitchen["floors"]:
            for canopy in floor["canopies"]:
                print("Canopy data:", canopy.get('wallCladding'), canopy.get('cladding_desc'))
                
                if canopy.get('wallCladding') and canopy.get('wallCladding') != '':
                    # Get the selected walls
                    selected_walls = canopy.get('cladding_desc', [])
                    
                    # Create description based on selected walls
                    if selected_walls:
                        wall_parts = []
                        if 'Rear' in selected_walls:  # Put Rear first
                            wall_parts.append('Rear')
                        if 'Left' in selected_walls:
                            wall_parts.append('Left')
                        if 'Right' in selected_walls:
                            wall_parts.append('Right')
                        
                        if wall_parts:
                            # Format as "Rear, Left & Right-hand Walls"
                            if len(wall_parts) > 1:
                                wall_description = f"Cladding to {', '.join(wall_parts[:-1])} & {wall_parts[-1]}-hand Walls"
                            else:
                                wall_description = f"Cladding to {wall_parts[0]}-hand Wall"
                        else:
                            wall_description = "Cladding below item, supplied and installed"
                    else:
                        wall_description = "Cladding below item, supplied and installed"

                    wall_cladding_data.append({
                        "item_no": str(canopy.get('itemNum', '1')),
                        "description": wall_description,
                        "width": canopy.get("cladding_width", 0),
                        "height": canopy.get("cladding_height", 0),
                    })

    print(f"DEBUG - Final wall_cladding_data: {wall_cladding_data}")
    return wall_cladding_data if wall_cladding_data else None

def generate_email_summary(genInfo, kitchen_info, scope_work):
    """
    Generates an email summary using Deepseek API based on the project information.
    """
    try:
        # Try to get API key from different locations
        api_key = None
        
        # Try environment variable first
        api_key = os.environ.get('DEEPSEEK_API_KEY')
        
        if not api_key:
            # List of possible paths for secrets.toml
            possible_paths = [
                '.streamlit/secrets.toml',
                'app/.streamlit/secrets.toml',
                '../.streamlit/secrets.toml',
                '../../.streamlit/secrets.toml',
                os.path.expanduser('~/.streamlit/secrets.toml'),
                os.path.join(os.getcwd(), '.streamlit/secrets.toml')
            ]
            
            for path in possible_paths:
                try:
                    if os.path.exists(path):
                        st.write(f"Found secrets file at: {path}")  # Debug output
                        secrets = toml.load(path)
                        if 'api_keys' in secrets and 'deepseek' in secrets['api_keys']:
                            api_key = secrets['api_keys']['deepseek']
                            break
                except Exception as e:
                    st.write(f"Error loading {path}: {str(e)}")  # Debug output
                    continue
        
        if not api_key:
            st.error("""
            API key not found. Please ensure one of the following:
            1. Create .streamlit/secrets.toml with your API key
            2. Set DEEPSEEK_API_KEY environment variable
            """)
            return None
        
        # Construct the project summary
        project_details = {
            "project_name": genInfo.get('projectName'),
            "project_number": genInfo.get('projectNum'),
            "customer": genInfo.get('customer'),
            "company": genInfo.get('company'),
            "location": genInfo.get('location'),
            "address": genInfo.get('address'),
            "date": genInfo.get('date'),
            "sales_contact": genInfo.get('salesContact'),
            "scope_of_work": scope_work,
            "email_tone": genInfo.get('email_tone', 'Professional'),
            "include_pricing": genInfo.get('include_pricing', False),
            "additional_notes": genInfo.get('additional_notes', ''),
            "closing_message": genInfo.get('closing_message', '')
        }

        # Create a prompt for Deepseek
        prompt = f"""
        Create a {project_details['email_tone'].lower()} email about a kitchen ventilation project. 
        Use your creativity but ensure you include these key details:

        Key Information to Include:
        - Project Name: {project_details['project_name']}
        - Location: {project_details['address']}, {project_details['location']}
        - Reference: {project_details['project_number']}
        - Sales Contact: {project_details['sales_contact']}
        - Customer: {project_details['customer']}

        Scope of Work (integrate naturally):
        {chr(10).join(project_details['scope_of_work'])}

        Additional Context (integrate naturally):
        {project_details['additional_notes'] if project_details['additional_notes'] else 'No additional notes provided.'}

        Personal Touch:
        {project_details.get('closing_message', 'Keep it professional and straightforward.')}

        Guidelines:
        1. Write in a {project_details['email_tone'].lower()} tone
        2. Include a clear subject line
        3. Naturally incorporate the scope of work
        4. Add the personal message in a natural way
        5. Keep the email concise but engaging
        6. Ensure all key project details are included
        7. Use plain text only, no special characters or formatting

        Feel free to be creative with the structure and wording while maintaining professionalism.
        """

        # Deepseek API endpoint and headers
        url = "https://api.deepseek.com/v1/chat/completions"
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }

        # API request payload
        payload = {
            "model": "deepseek-chat",
            "messages": [
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "max_tokens": 1000,
            "temperature": 0.7
        }

        # Make the API request with error handling
        try:
            response = requests.post(url, json=payload, headers=headers)
            response.raise_for_status()
            
            # Print response for debugging
            # st.write("API Response Status:", response.status_code)
            # st.write("API Response Headers:", dict(response.headers))
            # st.write("API Response:", response.text)
            
            # Extract and return the generated email
            email_summary = response.json()['choices'][0]['message']['content']
            
            # Clean up and format the response
            if email_summary:
                # Remove any markdown formatting
                email_summary = email_summary.replace('*', '').replace('_', '')
                
                # Ensure proper spacing between sections
                sections = email_summary.split('\n\n')
                formatted_sections = []
                
                for section in sections:
                    if section.strip():
                        # Properly format list items
                        if any(line.strip().startswith(('-', '•', '1.', '2.', '3.')) for line in section.splitlines()):
                            formatted_lines = []
                            for line in section.splitlines():
                                if line.strip():
                                    if line.strip().startswith(('-', '•')):
                                        formatted_lines.append(line.replace('•', '-').strip())
                                    else:
                                        formatted_lines.append(line.strip())
                            formatted_sections.append('\n'.join(formatted_lines))
                        else:
                            formatted_sections.append(section.strip())
                
                # Join sections with double newlines
                email_summary = '\n\n'.join(formatted_sections)
                
                # Ensure proper spacing after colons
                email_summary = email_summary.replace(':', ': ')
                
                # Remove any triple or more newlines
                while '\n\n\n' in email_summary:
                    email_summary = email_summary.replace('\n\n\n', '\n\n')
            
            return email_summary
            
        except requests.exceptions.RequestException as e:
            st.error(f"API Request Error: {str(e)}")
            if hasattr(e.response, 'json'):
                st.error(f"API Response: {e.response.json()}")
            return None

    except Exception as e:
        st.error(f"Error generating email summary: {str(e)}")
        return None