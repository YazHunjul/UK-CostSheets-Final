import streamlit as st
import json
from datetime import datetime, timedelta
import pandas as pd

def render_history_section(db_manager):
    with st.expander("📋 Previous Submissions", expanded=False):
        # Search and filter controls
        col1, col2, col3 = st.columns(3)
        with col1:
            search_term = st.text_input("🔍 Search projects", "")
        with col2:
            date_from = st.date_input("From date", 
                                    value=datetime.now() - timedelta(days=30))
        with col3:
            date_to = st.date_input("To date", 
                                  value=datetime.now())

        # Export button
        if st.button("📥 Export All to Excel"):
            export_file = db_manager.export_to_excel()
            st.success(f"Data exported to {export_file}")

        # Get filtered submissions
        submissions = db_manager.get_submissions(
            search_term=search_term,
            date_from=date_from.strftime("%Y-%m-%d"),
            date_to=date_to.strftime("%Y-%m-%d")
        )

        # Display submissions
        for _, row in submissions.iterrows():
            with st.container():
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"**Project:** {row['project_name']}")
                    st.write(f"**Date:** {row['timestamp']}")
                    st.write(f"**Customer:** {row['customer_name']}")
                    st.write(f"**Total Cost:** £{row['total_cost']:,.2f}")
                
                with col2:
                    with st.expander("View Details"):
                        st.write("### Customer Details")
                        st.write(f"Contact: {row['contact_name']}")
                        st.write(f"Phone: {row['contact_number']}")
                        st.write(f"Email: {row['email']}")
                        st.write(f"Site Address: {row['site_address']}")
                        
                        st.write("### Canopy Details")
                        st.write(f"Type: {row['canopy_type']}")
                        st.write(f"Dimensions: {row['canopy_length']}m x {row['canopy_width']}m x {row['canopy_height']}m")
                        st.write(f"Color: {row['canopy_color']}")
                        
                        st.write("### Additional Options")
                        st.write(f"Guttering: {row['guttering']}")
                        st.write(f"Side Panels: {row['side_panels']}")
                        st.write(f"Door Panels: {row['door_panels']}")
                        
                        st.write("### Installation Details")
                        st.write(f"Location: {row['delivery_location']}")
                        st.write("Plant Hires:", json.loads(row['plant_hires']))
                        st.write("Quantities:", json.loads(row['quantities']))
                        st.write(f"Strip Out: {row['strip_out']}")
                
                if st.button("🗑️ Delete", key=f"del_{row['id']}"):
                    db_manager.delete_submission(row['id'])
                    st.experimental_rerun()
                
                st.divider() 