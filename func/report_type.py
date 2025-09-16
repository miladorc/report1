import os
import streamlit as st
import pandas as pd
import pyodbc
import xlsxwriter
import sys
import json


def report_type(db_path, selected_report):
    if not os.path.exists(db_path):
        st.error(f"Error: Database file not found at {db_path}")
        return None

    conn = None  # Initialize conn to None
    try:
        # Create the connection
        conn = pyodbc.connect(
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={db_path};'
        )

        query = """
                SELECT * FROM [nedaja-hog]
                WHERE CMO IS NULL AND CVLK2F <= 55
                """

        df = pd.read_sql(query, conn)

        if selected_report == "گزارش پرداختی":
            filtered_df = df[
                (df['CMO'].isnull()) &
                (df['G_SHO'] != 90)
                ]

            # ... (your stats calculations)
            fma_stats = {
                'تعداد': int(filtered_df[filtered_df['FMA'] > 0]['FMA'].count()),
                'حداقل': int(filtered_df[filtered_df['FMA'] > 0]['FMA'].min()),
                'میانگین': int(filtered_df['FMA'].mean()),
                'حداکثر': int(filtered_df['FMA'].max()),
                'جمع': int(filtered_df['FMA'].sum())
            }

            tlbk_stats = {
                'تعداد': int(filtered_df[filtered_df['TLBK'] > 0]['TLBK'].count()),
                'حداقل': int(filtered_df[filtered_df['TLBK'] > 0]['TLBK'].min()),
                'میانگین': int(filtered_df['TLBK'].mean()),
                'حداکثر': int(filtered_df['TLBK'].max()),
                'جمع': int(filtered_df['TLBK'].sum())
            }

            hto_stats = {
                'تعداد': int(filtered_df[filtered_df['HTO'] > 0]['HTO'].count()),
                'حداقل': int(filtered_df[filtered_df['HTO'] > 0]['HTO'].min()),
                'میانگین': int(filtered_df['HTO'].mean()),
                'حداکثر': int(filtered_df['HTO'].max()),
                'جمع': int(filtered_df['HTO'].sum())
            }

            mkol_final_stats = {
                'تعداد': int(filtered_df[filtered_df['mkol-final'] > 0]['mkol-final'].count()),
                'حداقل': int(filtered_df[filtered_df['mkol-final'] > 0]['mkol-final'].min()),
                'میانگین': int(filtered_df['mkol-final'].mean()),
                'حداکثر': int(filtered_df['mkol-final'].max()),
                'جمع': int(filtered_df['mkol-final'].sum())
            }

            st.toast("گزارش با موفقیت ایجاد شد.", icon='✅')

            # The function should return data, not an Excel file
            return filtered_df
        else:
            st.warning("Selected report type is not yet implemented.")
            return None

    except pyodbc.Error as ex:
        st.error(f"A database connection error occurred: {ex}")
        return None
    finally:
        # Ensure the connection is closed, even if an error occurs
        if conn:
            conn.close()