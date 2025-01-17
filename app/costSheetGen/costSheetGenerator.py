import streamlit as st
from generalInformation import getInfo

def main():
    st.markdown("<h1 style= 'text-align: center; margin-bottom: 40px;'>Cost Sheet Generator </h1>", unsafe_allow_html=True)
    # Get General Information
    info = getInfo()
    #Get Report Type
    user_selection = st.selectbox('Select Cost Sheet to Report', ['', 'Canopy', 'AHU'])
    # Run canopy program to get Details
    if user_selection.lower() == 'canopy':
        from costSheetGen.Canopy import canopyMain
        canopyMain.main(info)
        return