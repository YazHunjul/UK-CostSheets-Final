import streamlit as st
import uuid

key = uuid.uuid4()

def getInfo():
    col1, col2, col3 = st.columns(3)
    with col1:
        proj_name = st.text_input('Project Name', placeholder='', key='projName')
        location = st.text_input('Location', placeholder='London', key='location')
    with col2:
        proj_num = st.text_input('Project Number', placeholder='', key='projNum')
        sales_contact = st.selectbox("Sales Contact", ['', 'Marc Byford', 'Karl Nicholson', 'Dan Butler', 'Chris Mannus', 'Dean Griffiths', 'David Stewart'])
    with col3:
        customer = st.text_input('Customer', placeholder='Azzam hunjul', key='customer')
        date = st.date_input('Date', 'today')
        
    reference_num = st.text_input('Enter Reference Number', placeholder='(e.g., 12345/01/23)', key='refNum')
    
    # Retrieve User Information
    return {
        'projectName' : proj_name,
        'location' : location,
        'projectNum' : proj_num,
        'salesContant' : sales_contact,
        'customer' : customer,
        'date' : date,
        'referenceNum' : reference_num,
    }