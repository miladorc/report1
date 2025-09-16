import func.report_type as rt
import streamlit as st
import pandas as pd
import tempfile
import os

report_type = ["گزارش پرداختی", "report2", "report3"]

st.title("Report Generator")

file_select = st.file_uploader("Select File", type=["accdb"])
drop_report = st.selectbox("report: ", report_type)

# Initialize df outside the conditional block
df = pd.DataFrame()

if file_select is not None:
    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
        tmp_file.write(file_select.getvalue())
        tmp_db_path = tmp_file.name

    if st.button("Start"):
        data = rt.report_type(tmp_db_path, drop_report)

        if data is not None:
            df = pd.DataFrame(data)
            csv_data = df.to_csv(index=False)
            st.download_button("download",
                               data=csv_data,
                               file_name="output.csv",
                               mime="text/csv")
        else:
            st.error("Could not generate report. Please check the file and try again.")

    # Always try to delete the temporary file
    try:
        os.unlink(tmp_db_path)
    except Exception as e:
        st.error(f"Error deleting temporary file: {e}")
else:
    st.info("Please upload an Access database file to begin.")