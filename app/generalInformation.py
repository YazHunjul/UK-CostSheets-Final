import streamlit as st
import uuid

key = uuid.uuid4()

contacts = {
    'Marc Byford' : '(07974 403322)',
    'Karl Nicholson' : '(07791 397866)',
    'Dan Butler' :  '(07703 729686)',
    'Chris Mannus' : '(07870 263280)',
    'Dean Griffiths' :  '(07814 784352)', 
    'Kent Phillips' : '(07949 016501)'
}

estimators = {
    'Simon Still': 'Lead Estimator',
    'Nick Soton': 'Estimator',
    'Chris Davis': 'Estimator',
}

def get_initials(name):
    """Extract initials from a name"""
    if name:
        words = name.split()
        initials = ''.join(word[0].upper() for word in words if word)
        return initials
    return ''

def getInfo():
    col1, col2, col3 = st.columns(3)
    with col1:
        proj_name = st.text_input('Project Name', placeholder='', key='projName')
        customer = st.text_input('Customer', placeholder='Azzam hunjul', key='customer')
        address = st.text_input('Address', placeholder='123 Main St, London', key='address')
        
    with col2:
        proj_num = st.text_input('Project Number', placeholder='', key='projNum')
        company = st.text_input('Company', placeholder='Halton', key='company')
        sales_contact = st.selectbox("Sales Contact", ['', 'Marc Byford', 'Karl Nicholson', 'Dan Butler', 'Chris Mannus', 'Dean Griffiths', 'David Stewart'])
        
    with col3:
        date = st.date_input(
            "Date",
            format="DD/MM/YYYY"  # This will display in UK format
        )
        location = st.text_input('Location', placeholder='London', key='location')
        estimator = st.selectbox(
            "Estimator",
            options=[''] + list(estimators.keys())
        )
    
       
        
    col1, col2, col3 = st.columns(3)
   
        
    with col2:
        reference_num = ''
    
    # Calculate combined initials
    combined_initials = f"{get_initials(sales_contact)}/{get_initials(estimator)}"
    
    # Get estimator role
    estimator_role = estimators.get(estimator, '')
    
    # Retrieve User Information
    return {
        'projectName': proj_name,
        'location': location,
        'projectNum': proj_num,
        'salesContact': f'{sales_contact} {contacts.get(sales_contact, "")}',
        'customer': customer,
        'date': date.strftime("%d/%m/%Y"),  # This will store as DD/MM/YYYY
        'estimator': estimator,
        'estimator_role': estimator_role,
        'referenceNum': reference_num,
        'sales_contact': sales_contact,
        'combined_initials': combined_initials,
        'company': company.title(),
        'address': address.title(),
    }

# Marc Byford (07974 403322)
# Karl Nicholson (07791 397866)
# Dan Butler (07703 729686)
# Chris Mannus (07870 263280)
# Dean Griffiths (07814 784352)                 David Stewart (07989 185991)          
