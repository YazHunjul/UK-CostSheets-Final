import streamlit as st
from generalInformation import getInfo

def main():
    st.markdown("<h1 style= 'text-align: center; margin-bottom: 40px;'>Cost Sheet Generator </h1>", unsafe_allow_html=True)
    # Get General Information
    info = getInfo()
    print(info)
    