# interface.py
# Streamlit app interface

import streamlit as st
import pandas as pd
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from scripts.process_calendar import process_calendar_data
import datetime
import datetime

# Title and instructions
st.title("Engagement Reporting App")
st.markdown("""
**Please Upload your calendar CSV or Excel file here.**

To download your calendar from Outlook, follow these steps:
1. Open Outlook
2. Go to the Calendar view.
3. Click File → Open & Export → Import/Export.
4. Choose Export to a file → Next.
5. Select Microsoft Excel (or Comma Separated Values (CSV) if Excel isn’t available) → Next.
6. Select the calendar folder you want to export → Next.
7. Choose a location and filename for your exported file → Finish.
8. Outlook may ask you for a date range; set the period you want → OK.
""")

# Upload calendar CSV or Excel
calendar_file = st.file_uploader("Upload your calendar CSV or Excel file", type=["csv", "xlsx"])

import io
from openpyxl import load_workbook

if calendar_file:
    calendar_df = process_calendar_data(calendar_file, None)
    st.write("Preview of uploaded calendar data:")
    st.dataframe(calendar_df.head())

    # --- Excel download using template ---
    template_path = "template.xlsx"
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            writer.book = load_workbook(template_path)
            calendar_df.to_excel(writer, index=False, sheet_name='Processed Data')
            writer.save()
        output.seek(0)
        st.download_button(
            label="Download processed data as Excel (template)",
            data=output,
            file_name="processed_calendar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error exporting to template: {e}")
