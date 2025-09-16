import func.report_type as rt
import streamlit as st
import pandas as pd
import tempfile
import os
import io

report_type = ["گزارش پرداختی", "براساس نوع عضویت", "report3"]

st.title("Report Generator")

file_select = st.file_uploader("Select File", type=["accdb"])
drop_report = st.selectbox("report: ", report_type)

df = pd.DataFrame()

if file_select is not None:
    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
        tmp_file.write(file_select.getvalue())
        tmp_db_path = tmp_file.name

    if st.button("Start"):
        # The function now returns the formatted stats DataFrame
        data_df = rt.report_type(tmp_db_path, drop_report)

        if data_df is not None and not data_df.empty:
            # Create an in-memory Excel file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                data_df.to_excel(writer, index=False, sheet_name='Sheet1')

            # Reset the buffer to the beginning
            output.seek(0)

            st.success("Report generated successfully! Click to download.")

            # Use the in-memory file for the download button
            st.download_button("Download Excel Report",
                               data=output,
                               file_name="گزارش پرداختی.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("Could not generate report. Please check the file and try again.")

    try:
        os.unlink(tmp_db_path)
    except Exception as e:
        st.error(f"Error deleting temporary file: {e}")
else:
    st.info("Please upload an Access database file to begin.")