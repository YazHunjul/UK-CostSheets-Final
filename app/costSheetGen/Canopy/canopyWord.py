from docxtpl import DocxTemplate
from io import BytesIO
import os
from collections import Counter
import math
import requests
import streamlit as st
import toml
from ..config import TEMPLATES

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

<<<<<<< HEAD
        # Get just the first name of the customer
        full_name = genInfo.get('customer', '')
        first_name = full_name.split()[0] if full_name else ''  # Take first word as first name
        genInfo['customer'] = first_name  # Update customer to just first name

        template_path = '/Users/yazan/Desktop/Efficiency/UK-CostSheets-Final/app/costSheetGen/costSheetResources/costSheet_canopy.docx'
=======
        template_path = 'app/costSheetGen/costSheetResources/costSheet_canopy.docx'
>>>>>>> 74c3922b8157d83ba4464e01abbfc30bf30c2903
        
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Word template not found at: {template_path}")
            
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

        # Use the kitchen_totals passed from canopyMain.py
        kitchen_totals = context.get('kitchen_totals', {})
        # Round each kitchen total
        kitchen_totals = {k: math.ceil(v) for k, v in kitchen_totals.items()}
        grand_total = math.ceil(sum(kitchen_totals.values()))

        # Get the totals from context
        total_cost = math.ceil(sum(kitchen_totals.values()))  # Total from T16
        total_k9_cost = context.get('total_k9_cost', 0)  # Total from S16

        # Add pricing information to genInfo
        genInfo.update({
            'kitchen_totals': kitchen_totals,  # Already rounded when extracted from Excel
            'grand_total': math.ceil(sum(kitchen_totals.values())),
            'total_job_price': total_cost,
            'total_cost': total_k9_cost
        })

        # Round prices in grouped_canopy_data
        for kitchen_name, kitchen in grouped_canopy_data.items():
            for floor_name, floor_data in kitchen.items():
                for canopy in floor_data['canopies']:
                    if 'total_price' in canopy:
                        canopy['total_price'] = math.ceil(canopy['total_price'])
                    if 'cladding_price' in canopy:
                        canopy['cladding_price'] = math.ceil(canopy['cladding_price'])

        # Update genInfo with rounded data
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

def extract_canopy_info_grouped(context):
    kitchens = context.get("kitchens", [])
    grouped_canopies = {}

    for kitchen in kitchens:
        kitchen_name = kitchen["kitchen_name"].title()
        grouped_canopies[kitchen_name] = {}

        for floor in kitchen["floors"]:
            floor_name = floor["floor_name"].title()
            display_name = f"{kitchen_name} – {floor_name}"
            
            # Calculate total extract and makeup air for the floor
            total_extract = 0
            total_makeup = 0
            has_fresh_air_canopy = False

            for canopy in floor.get('canopies', []):
                model = canopy.get('model', '')
                # Check if model ends with 'F' for Fresh Air
                if model.endswith('F'):
                    has_fresh_air_canopy = True
                    flow_rate = canopy.get('flowRate', 0)
                    total_extract += flow_rate
                    # Makeup air is 85% of extract
                    makeup_air = round(flow_rate * 0.85, 2)
                    total_makeup += makeup_air
                    # Add makeup air to canopy data
                    canopy['makeup_air'] = makeup_air

            # Calculate shortfall
            shortfall = round(total_extract - total_makeup, 2)
            
            # Always create important note if there's a Fresh Air canopy
            important_note = ""
            if has_fresh_air_canopy:
                important_note = (
                    f"The makeup air flows shown above are the maximum that we can introduce through the "
                    f"canopy. This should be equal to approximately 85% of the extract the shortfall of "
                    f"{shortfall} m3/s must be introduced through ceiling grilles or diffusers, by others.\n"
                    f"If you require further guidance on this, please do not hesitate to contact us."
                )
            
            # Calculate floor subtotal
            floor_subtotal = 0
            for canopy in floor.get('canopies', []):
                floor_subtotal += float(canopy.get('total_price', 0))
            
            # Add delivery/installation costs
            delivery_data = floor.get('delivery_install_data', {})
            floor_subtotal += float(delivery_data.get('delivery_price', 0))
            floor_subtotal += float(delivery_data.get('install_price', 0))
            
            # Calculate cladding total and store cladding info
            cladding_total = 0
            cladding_canopies = []  # List of canopies with cladding
            
            for canopy in floor.get('canopies', []):
                # Debug prints
                st.write(f"Checking canopy {canopy.get('item_number')}:")
                st.write(f"- wallCladding: {canopy.get('wallCladding')}")
                st.write(f"- cladding_desc: {canopy.get('cladding_desc')}")
                
                # Check if canopy has wallCladding and walls selected
                if canopy.get('wallCladding') and canopy.get('cladding_desc'):
                    st.write("Adding to cladding_canopies")
                    cladding_price = float(canopy.get('cladding_price', 0))
                    cladding_total += cladding_price
                    cladding_canopies.append({
                        'item_number': canopy.get('item_number', ''),
                        'model': canopy.get('model', ''),
                        'cladding_desc': canopy.get('cladding_desc', []),
                        'cladding_price': cladding_price,
                        'width': canopy.get('cladding_length', 0),
                        'height': canopy.get('cladding_height', 0)
                    })

            # Get commissioning price
            commission_price = float(floor.get('commission_price', 0))
            
            # Get test & commission value from floor data
            test_commission = float(floor.get('test_commission', 0))
            
            # Store in grouped_canopies
            grouped_canopies[kitchen_name][display_name] = {
                "canopies": floor.get('canopies', []),
                "important_note": important_note,
                "floor_name": floor['floor_name'],
                "p182": floor.get('p182_value', 0),
                "test_commission": test_commission,
                "commission_price": commission_price,
                "subtotal": round(floor_subtotal, 2),
                "cladding_total": round(cladding_total, 2),
                "floor_data": {
                    "cladding_total": round(cladding_total, 2),
                    "cladding_canopies": cladding_canopies,
                    "test_commission": test_commission,
                    "uv_total": 0,
                    "total_extract": total_extract,
                    "total_makeup": total_makeup,
                    "shortfall": shortfall,
                    "has_makeup_air": has_fresh_air_canopy,
                    "important_note": important_note,
                    "subtotal": round(floor_subtotal, 2)
                }
            }
    
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

                    # Calculate values
                    cws_continuous = round(length / 1000 * 0.02, 2)
                    hws_wash_cycle = round(length / 1000 * 0.103, 3)
                    hws_storage = round(hws_wash_cycle * 180, 3)

                    # Try multiple possible keys for item number
                    item_number = (
                        canopy.get('itemNum') or 
                        canopy.get('item_number') or 
                        canopy.get('reference_number') or 
                        ''
                    )

                    cmwi_canopies.append({
                        "item_no": str(item_number),  # Convert to string and ensure not None
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
    Extracts wall cladding information only for canopies that have cladding.
    Returns None if no canopies have wall cladding.
    """
    kitchens = kitchen_info.get("kitchens", [])
    wall_cladding_data = []

    for kitchen in kitchens:
        for floor in kitchen["floors"]:
            for canopy in floor["canopies"]:
                # Only process canopies that have wallCladding and cladding_desc
                if (canopy.get('wallCladding') and 
                    canopy.get('cladding_desc')):  # Check if walls were selected
                    
                    walls = canopy['cladding_desc']
                    if len(walls) > 1:
                        wall_description = f"Cladding to {', '.join(walls[:-1])} & {walls[-1]}-hand Walls"
                    else:
                        wall_description = f"Cladding to {walls[0]}-hand Wall"
                        
                    wall_cladding_data.append({
                        "item_no": str(canopy.get('item_number', '')),
                        "description": wall_description,
                        "width": canopy.get("cladding_length", 0),
                        "height": canopy.get("cladding_height", 0),
                        "price": math.ceil(canopy.get('cladding_price', 0))
                    })

    # Only return data if we found canopies with cladding
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