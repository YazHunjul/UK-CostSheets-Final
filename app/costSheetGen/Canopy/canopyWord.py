from docxtpl import DocxTemplate
from io import BytesIO
import os
from collections import Counter
import math

def generate_word(context, genInfo):
    """
    Generates a Word document using the provided context.

    Args:
        context (dict): A dictionary containing the data to render in the Word template.

    Returns:
        BytesIO: The generated Word document as a BytesIO object.
    """
    # Path to the Word template
    template_path = 'app/costSheetGen/costSheetResources/costSheet_canopy.docx'
    
    # Load the Word template
    template = DocxTemplate(template_path)

    # Extract and process canopy data grouped by kitchens and floors
    grouped_canopy_data = extract_canopy_info_grouped(context)
    cmwi_data = extract_cmwi_canopies(context)
    scope_work = scope_of_works(context)

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
    wall_cladding_info = get_wall_cladding(context)
    genInfo["wall_cladding_info"] = wall_cladding_info
    print(genInfo)

    # Render the template with the given context
    template.render(genInfo)

    # Save the rendered document to a BytesIO buffer
    word_buffer = BytesIO()
    template.save(word_buffer)
    word_buffer.seek(0)
    
    return word_buffer

def scope_of_works(context):
    """
    Generates the Scope of Work text as a list based on the context.

    Args:
        context (dict): Contains kitchen, floor, and canopy data.

    Returns:
        list: A list of formatted strings for each canopy model and count.
    """
    canopy_counts = Counter()
    kitchens = context.get('kitchens', [])
    for kitchen in kitchens:
        for floor in kitchen.get('floors', []):
            for canopy in floor.get('canopies', []):
                model = canopy.get('model', 'Unknown Model')
                canopy_counts[model] += 1

    scope_lines = [f"{count}X {model}" for model, count in canopy_counts.items()]
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
        kitchen_name = kitchen["kitchen_name"]
        grouped_canopies[kitchen_name] = {}

        for floor in kitchen["floors"]:
            floor_name = floor["floor_name"]
            grouped_canopies[kitchen_name][floor_name] = {
                "canopies": [],
                "important_note": "",
            }

            total_extract_volume = 0
            total_mua_volume = 0

            for canopy in floor["canopies"]:
                model = canopy.get("model", "")
                length = canopy.get("length", 0)
                sections = canopy.get("section", 0)
                flowrate = canopy.get("flowrate", 0)

                # Perform calculations
                f12 = calculate_f12(length, sections)
                grease_filters = calculate_grease_filters(model, f12, length, sections)
                extract_static_pa = 50 if model == 'CXW' else calculate_extract_static_pa(grease_filters, flowrate, model)
                print(extract_static_pa)
                
                # Add to total EXT.VOL and MUA VOL
                total_extract_volume += flowrate
                if "MUA VOL" in canopy:
                    total_mua_volume += canopy["MUA VOL"]  # Ensure this field exists

                # Add processed data to the current floor's list
                grouped_canopies[kitchen_name][floor_name]["canopies"].append({
                    'item_number' : canopy.get('itemNum', ''),
                    "model": model,
                    "length": length,
                    "width": canopy.get("width", 0),
                    "height": canopy.get("height", 0),
                    "sections": sections,
                    "flowrate": flowrate,
                    "f12": f12,
                    "grease_filters": grease_filters,
                    "extract_static_pa": extract_static_pa,
                    "lights": canopy.get("lights", ""),
                })

            # Calculate the important note
            required_mua = round(0.85 * total_extract_volume, 3)  # 85% of total extract volume
            shortfall = round(required_mua - total_mua_volume, 3)

            important_note = (
                f"The make-up air flows shown above are the maximum that we can introduce through the canopy. "
                f"This should be equal to approximately 85% of the extract i.e. {total_extract_volume} m³/s. "
                f"In this instance, it only totals {total_mua_volume} m³/s, therefore the shortfall of {shortfall} m³/s "
                f"must be introduced through ceiling grilles or diffusers, by others."
            )

            grouped_canopies[kitchen_name][floor_name]["important_note"] = important_note

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
                        "item_no": f"{len(cmwi_canopies) + 1:.2f}",
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
    return f"{round((((flow_rate * 3600) / (grease_filters * i14)) ** 2) + 20, 1)} Pa"


# TODO: Fill out wall cladding Info
def get_wall_cladding(kitchen_info):
    """
    Extracts wall cladding information based on canopy data.

    Args:
        kitchen_info (dict): Contains kitchen, floor, and canopy data.

    Returns:
        list: A list of dictionaries containing wall cladding data.
    """
    kitchens = kitchen_info.get("kitchens", [])
    wall_cladding_data = []

    for kitchen in kitchens:
        for floor in kitchen["floors"]:
            for canopy in floor["canopies"]:
                # exists = True if canopy.get('wallCladding') else False
                item_number = canopy.get("itemNum", "Unknown")
                width = canopy.get("cladding_width", 0)
                height = canopy.get("cladding_height", 0)

                # Add wall cladding info for this canopy
                wall_cladding_data.append({
                    "item_no": item_number,
                    "description": "Cladding to Rear, Left & Right-hand Walls",
                    "width": width,
                    "height": height,
                    # 'exists' : exists
                })

    return wall_cladding_data