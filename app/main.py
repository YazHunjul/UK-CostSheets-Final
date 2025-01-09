import streamlit as st


sideBar = st.sidebar.selectbox('Navigation', ['Home','Cost Sheet Generator'])
if sideBar == 'Cost Sheet Generator':
    from costSheetGen import costSheetGenerator
    costSheetGenerator.main()