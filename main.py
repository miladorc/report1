import func.report_type as rt
import streamlit as st
import pandas as pd


report_type = ["گزارش پرداختی", "report2", "report3"]

file_select = st.file_uploader("Select File", type=["accdb"])
drop_report = st.selectbox("report: ", report_type)

data = rt.report_type()
df = pd.DataFrame(data)

csv_data = df.to_csv(index=False)


save_dir = st.download_button("download",
                              data=csv_data,
                              file_name="output.csv",
                              mime="text/csv")
button1 = st.button("Start")








