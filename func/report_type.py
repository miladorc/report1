import os
import streamlit as st
import pandas as pd
import pyodbc
import xlsxwriter
import sys
import json



def report_type():


    if getattr(sys, 'frozen', False):

        application_path = os.path.dirname(sys.executable)
    else:

        application_path = os.path.dirname(os.path.abspath(__file__))

    config_file = os.path.join(application_path, "search_app_config.json")

    def load_config():
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print("خطا در خواندن تنظیمات: {e}".format(e))
        return {}

    conn = pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={load_config().get("access_file_path")};'
        'charset=utf8;'
    )

    query = """
            SELECT * FROM [nedaja-hog]
            WHERE CMO IS NULL AND CVLK2F <= 55
            """

    df = pd.read_sql(query, conn)

    if report_type == "گزارش پرداختی":
        # Filter and calculate stats remain the same
        filtered_df = df[
            (df['CMO'].isnull()) &
            (df['G_SHO'] != 90)
            ]

        # Calculate statistics for FMA
        fma_stats = {
            'تعداد': int(filtered_df[filtered_df['FMA'] > 0]['FMA'].count()),
            'حداقل': int(filtered_df[filtered_df['FMA'] > 0]['FMA'].min()),
            'میانگین': int(filtered_df['FMA'].mean()),
            'حداکثر': int(filtered_df['FMA'].max()),
            'جمع': int(filtered_df['FMA'].sum())
        }

        # Calculate statistics for TLBK
        tlbk_stats = {
            'تعداد': int(filtered_df[filtered_df['TLBK'] > 0]['TLBK'].count()),
            'حداقل': int(filtered_df[filtered_df['TLBK'] > 0]['TLBK'].min()),
            'میانگین': int(filtered_df['TLBK'].mean()),
            'حداکثر': int(filtered_df['TLBK'].max()),
            'جمع': int(filtered_df['TLBK'].sum())
        }

        # Calculate statistics for HTO
        hto_stats = {
            'تعداد': int(filtered_df[filtered_df['HTO'] > 0]['HTO'].count()),
            'حداقل': int(filtered_df[filtered_df['HTO'] > 0]['HTO'].min()),
            'میانگین': int(filtered_df['HTO'].mean()),
            'حداکثر': int(filtered_df['HTO'].max()),
            'جمع': int(filtered_df['HTO'].sum())
        }

        # Calculate statistics for MKOL-final
        mkol_final_stats = {
            'تعداد': int(filtered_df[filtered_df['mkol-final'] > 0]['mkol-final'].count()),
            'حداقل': int(filtered_df[filtered_df['mkol-final'] > 0]['mkol-final'].min()),
            'میانگین': int(filtered_df['mkol-final'].mean()),
            'حداکثر': int(filtered_df['mkol-final'].max()),
            'جمع': int(filtered_df['mkol-final'].sum())
        }

        # Create Excel workbook
        workbook = xlsxwriter.Workbook('گزارش پرداختی.xlsx')
        worksheet = workbook.add_worksheet()

        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'font_name': 'B Nazanin',
            'font_size': 12,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 2,
            'bg_color': '#D9E1F2',
            'pattern': 1
        })

        row_header_format = workbook.add_format({
            'bold': True,
            'font_name': 'B Nazanin',
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': '#E2EFDA'
        })

        cell_format = workbook.add_format({
            'font_name': 'B Nazanin',
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#,##0',
            'border': 1
        })

        cell_format_2 = workbook.add_format({
            'font_name': 'B Nazanin',
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })

        sum_format = workbook.add_format({
            'bold': True,
            'font_name': 'B Nazanin',
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#,##0',
            'border': 1,
            'bg_color': '#FFE699'
        })

        sum_format_2 = workbook.add_format({
            'bold': True,
            'font_name': 'B Nazanin',
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1,
            'bg_color': '#FFE699'
        })

        # Write column headers
        headers = ['نوع', 'تعداد', 'حداقل', 'میانگین', 'حداکثر', 'جمع']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
            worksheet.set_column(col, col, 15)

        # Write MKOL row
        worksheet.write(1, 0, 'فوق العاده عملیاتی', row_header_format)
        for col, key in enumerate(fma_stats.keys(), 1):
            if col != 1:
                worksheet.write(1, col, fma_stats[key], cell_format)
            else:
                worksheet.write(1, col, fma_stats[key], cell_format_2)

        # Write TLBK row
        worksheet.write(2, 0, 'طلب کار', row_header_format)
        for col, key in enumerate(tlbk_stats.keys(), 1):
            if col != 1:
                worksheet.write(2, col, tlbk_stats[key], cell_format)
            else:
                worksheet.write(2, col, tlbk_stats[key], cell_format_2)

        # Write HTO row
        worksheet.write(3, 0, 'حق التدریس', row_header_format)
        for col, key in enumerate(hto_stats.keys(), 1):
            if col != 1:
                worksheet.write(3, col, hto_stats[key], cell_format)
            else:
                worksheet.write(3, col, hto_stats[key], cell_format_2)

        # Write MKOL-final row
        worksheet.write(4, 0, 'حقوق و مزایا', row_header_format)
        for col, key in enumerate(mkol_final_stats.keys(), 1):
            if col != 1:
                worksheet.write(4, col, mkol_final_stats[key], cell_format)
            else:
                worksheet.write(4, col, mkol_final_stats[key], cell_format_2)

        # Write totals row
        worksheet.write(5, 0, 'جمع کل', sum_format)
        for col in range(1, len(headers)):
            if headers[col] == 'تعداد':
                total = fma_stats['تعداد'] + tlbk_stats['تعداد'] + hto_stats['تعداد'] + mkol_final_stats['تعداد']
            elif headers[col] == 'میانگین':
                total = (fma_stats['میانگین'] + tlbk_stats['میانگین'] + hto_stats['میانگین'] + mkol_final_stats[
                    'میانگین']) / 4
            elif headers[col] == 'حداقل':
                total = min(fma_stats['حداقل'], tlbk_stats['حداقل'], hto_stats['حداقل'], mkol_final_stats['حداقل'])
            elif headers[col] == 'حداکثر':
                total = max(fma_stats['حداکثر'], tlbk_stats['حداکثر'], hto_stats['حداکثر'], mkol_final_stats['حداکثر'])
            elif headers[col] == 'جمع':
                total = fma_stats['جمع'] + tlbk_stats['جمع'] + hto_stats['جمع'] + mkol_final_stats['جمع']
            if col != 1:
                worksheet.write(5, col, total, sum_format)
            else:
                worksheet.write(5, col, total, sum_format_2)

        workbook.close()

    st.toast.showinfo("گزارش موفق", "گزارش در فایل 'گزارش پرداختی.xlsx' ذخیره شد.")