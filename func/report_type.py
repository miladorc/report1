import os
import streamlit as st
import pandas as pd
import pyodbc
import io
import json


def report_type(db_path, selected_report):
    if not os.path.exists(db_path):
        st.error(f"Error: Database file not found at {db_path}")
        return None

    conn = None
    try:
        conn = pyodbc.connect(
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={db_path};'
        )

        query = """
                SELECT * FROM [nedaja-hog]
                WHERE CMO IS NULL
                """

        df = pd.read_sql(query, conn)
        # convert cell in database to number for mathematics
        numeric_cols = ['CVLK2F', 'G_SHO', 'CDYR', 'CHK21', 'FMA', 'HTO', 'EZF', 'FBK', 'PPK', 'HEZ', 'KHF',
                        'ID_SHOGHL1',
                        'MKOL', 'TLBK', 'HTO', 'mkol-final',
                        'EMZSHA', 'EMZSHO', 'EMZMOD', 'EMZMGH', 'EMZBAH', 'EMZESAR', 'EMZKMJZ', 'EMZSKG', 'EMZAO',
                        'EMZFSMTM', 'EMZFV', 'CHK21',
                        'TRM401', 'TTH', 'TTG', 'T97_PF', 'MKHM', 'KHR']

        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        # separate each group by job title
        heyat_elmi = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CHK21'] == 12)  # 'CHK21' is 12 , هیئت علمی
            ]

        janbaz_halat_eshtaghal = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CHK21'] == 21)  # 'CHK21' == 21 , جانباز حالت اشتغال
            ]

        karmand_elmi = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CHK21'] == 18)  # 'CHK21' == 18 , کارمند علمی
            ]

        karmand_tajrobi = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CHK21'] == 16)  # 'CHK21' == 16 , کارمند تجربی
            ]

        karmand_peymani = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CHK21'].isin([15, 17]))  # 'CHK21' is 15 or 17, کارمند پیمانی
            ]

        nezami = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CHK21'] == 20)  # 'CHK21' is 20, نظامی
            ]

        nezami_peymani = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CHK21'] == 19)  # 'CHK21' is 19, نظامی پیمانی
            ]

        mohasel = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CHK21'].isin([7, 10, 8]))  # 'CHK21' is 07 ,10 or 08, محصل نظامی
            ]

        karmandan = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CDYR'].between(601, 699, inclusive='both'))  # 'CDYR' is between 601 and 699, کارمندان دانشجو
            ]

        daraje_dar = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CDYR'].isin([706, 707, 708, 709,
                              806, 807, 808, 809,
                              906, 907, 908, 909]))  # 'CDYR' is 706, 707, 708, 709, درجات دانشجو
            ]

        afsar_joz = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CDYR'].isin([710, 711, 712, 713,
                              810, 811, 812, 813,
                              910, 911, 912, 913]))  # 'CHK21' is 711 ,712 ,713 , افسر جز
            ]

        afsar_arshad = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CDYR'].isin([714, 715, 716,
                              814, 815, 816,
                              914, 915, 916]))  # 'CHK21' is 714, 715, 716 , افسر ارشد
            ]

        amiri = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CDYR'].isin([717, 718, 719, 720,
                              817, 818, 819, 820,
                              917, 918, 919, 920]))  # 'CHK21' is 717 ,718 ,719 , امیری
            ]

        tenyear = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['sanavat'] <= 10)  # 'sanavat' is <=10 , سنوات تا 10 سال
            ]

        Twenty = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['sanavat'].between(11, 20, inclusive='both'))  # 'sanavat' is between 20 and 11 , سنوات از 11 سال تا 20
            ]

        thirty = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['sanavat'].between(21, 30, inclusive='both'))  # 'sanavat' is between 20 and 11 , سنوات از 21 سال تا 30

            ]

        morethanthirty = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['sanavat'] > 30)  # 'sanavat' is between 20 and 11 , سنوات بالای 30 سال
            ]

        razm = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['ID_SHOGHL1'].astype(str).str[5:7].isin(['21', '22', '23']))
            # Condition on digits 6 and 7 of 'ID_SHOGHL1'
            ]

        razmsupport = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['ID_SHOGHL1'].astype(str).str[5:7].isin(['24', '25', '26']))
            # Condition on digits 6 and 7 of 'ID_SHOGHL1'
            ]

        razmsupport_khadamat = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['ID_SHOGHL1'].astype(str).str[5:7].isin(['11', '12']))  # Condition on digits 6 and 7 of 'ID_SHOGHL1'
            ]

        sum_all = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CHK21'].isin([12, 18, 16, 15, 17, 20, 19, 10, 7, 8]))
            ]

        sum_all_without_mohasel = df[
            ((df['CMO'].isnull()) | (df['CMO'] == '')) &  # 'CMO' is null or blank
            (df['CVLK2F'] <= 55) &  # 'CVLK2F' is 55 or less
            (df['G_SHO'] != 90) &  # 'G_SHO' is not 90
            (df['CHK21'].isin([12, 18, 16, 15, 17, 20, 19]))
            ]

        afsar_vazife = df[
            (df['RANK'].isin([601, 602, 603, 604,
                              701, 702, 703, 704, 791, 792,
                              801, 802, 803, 804, 805,
                              901, 902, 903, 904]))
        ]

        darajedar_vazife = df[
            (df['RANK'].isin([605, 606,
                              705, 706, 707, 708, 709, 793,
                              806,
                              905, 906, 907, 993]))
        ]

        vazife_adi = df[
            (df['RANK'].isin([607, 608, 609, 610, 611, 612,
                              710, 711, 712, 794,
                              807, 810, 809, 811, 812, 894,
                              910, 911, 912, 994]))
        ]



        if report_type == "براساس نوع عضویت":  # Check for the specific report type
            # Create a list to store all results
            all_results = []

            # Process each group and add to results list
            groups = [
                (heyat_elmi, 'اعضاء هیئت علمی'),
                (janbaz_halat_eshtaghal, 'جانبازان حالت اشتغال'),
                (karmand_elmi, 'کارمندان علمی'),
                (karmand_tajrobi, 'کارمند تجربی'),
                (karmand_peymani, 'کارمند پیمانی'),
                (nezami, 'نظامی'),
                (nezami_peymani, 'نظامی پیمانی'),
                (mohasel, 'محصلین'),
                (sum_all, 'کل با محصلین'),
                (sum_all_without_mohasel, 'کل بدون محصلین')
            ]

            for df_group, group_name in groups:
                if not df_group.empty and 'mkol-final' in df_group.columns:
                    # Calculate min value excluding zeros
                    non_zero_mkol = df_group[df_group['mkol-final'] > 0]['mkol-final']
                    min_value = int(non_zero_mkol.min()) if not non_zero_mkol.empty else None
                    max_value = int(df_group['mkol-final'].max())
                    average_value = int(df_group['mkol-final'].mean())
                    count_value = int(df_group['mkol-final'].count())
                    total_mkol_final = int(df_group['mkol-final'].sum())

                    # Calculate FMA statistics excluding zeros
                    non_zero_fma = df_group[df_group['FMA'] > 0]['FMA']  # Changed from != 0 to > 0
                    min_fogh = non_zero_fma.min() if not non_zero_fma.empty else None
                    max_fogh = int(non_zero_fma.max()) if not non_zero_fma.empty else None
                    average_fogh = int(non_zero_fma.mean()) if not non_zero_fma.empty else None
                    count_fogh = int(non_zero_fma.count())
                    total_fogh = int(df_group['FMA'].sum())

                    # Calculate other allowances
                    columns_to_check = ['HTO', 'EZF', 'FBK', 'PPK', 'HEZ', 'KHF']
                    filtered_records = df_group[df_group[columns_to_check].gt(0).any(axis=1)]
                    count_other = filtered_records.shape[0]

                    total_other = df_group[columns_to_check].sum().sum()
                    average_other = total_other / count_other if count_other != 0 else 0

                    # Calculate total sum
                    columns_to_sum = ['mkol-final', 'FMA', 'HTO', 'EZF', 'FBK', 'PPK', 'HEZ', 'KHF']
                    total_sum = df_group[columns_to_sum].sum().sum()

                    # Add results to list
                    all_results.append({
                        'نوع عضویت': group_name,
                        'تعداد': count_value,
                        'حداقل': min_value,
                        'میانگین': average_value,
                        'حداکثر': max_value,
                        'تعداد فوق العاده عملیاتی': count_fogh,
                        'حداقل فوق العاده عملیاتی': min_fogh,
                        'میانگین فوق العاده عملیاتی': average_fogh,
                        'حداکثر فوق العاده عملیاتی': max_fogh,
                        'تعداد سایر': count_other,
                        'میانگین سایر': average_other,
                        'مجموع مبالغ پرداختی': total_sum,
                        'کل پرداختی حقوق و مزایا': total_mkol_final,
                        'کل پرداختی فوق العاده عملیاتی': total_fogh
                    })




        if selected_report == "گزارش پرداختی":
            filtered_df = df[
                (df['CMO'].isnull()) &
                (df['G_SHO'] != 90)
                ].copy()  # Using .copy() to avoid SettingWithCopyWarning

            # A helper function to safely calculate stats for a column
            def calculate_stats(series):
                if series.empty:
                    return {
                        'تعداد': 0, 'حداقل': 0, 'میانگین': 0, 'حداکثر': 0, 'جمع': 0
                    }
                return {
                    'تعداد': int(series.count()),
                    'حداقل': int(series.min()),
                    'میانگین': int(series.mean()),
                    'حداکثر': int(series.max()),
                    'جمع': int(series.sum())
                }

            fma_stats = calculate_stats(filtered_df[filtered_df['FMA'] > 0]['FMA'])
            tlbk_stats = calculate_stats(filtered_df[filtered_df['TLBK'] > 0]['TLBK'])
            hto_stats = calculate_stats(filtered_df[filtered_df['HTO'] > 0]['HTO'])
            mkol_final_stats = calculate_stats(filtered_df[filtered_df['mkol-final'] > 0]['mkol-final'])

            # Create a DataFrame from the calculated stats
            stats_df = pd.DataFrame([
                {'نوع': 'فوق العاده عملیاتی', **fma_stats},
                {'نوع': 'طلب کار', **tlbk_stats},
                {'نوع': 'حق التدریس', **hto_stats},
                {'نوع': 'حقوق و مزایا', **mkol_final_stats}
            ])

            # Calculate the totals row
            totals = {
                'نوع': 'جمع کل',
                'تعداد': stats_df['تعداد'].sum(),
                'حداقل': stats_df['حداقل'].min(),
                'میانگین': stats_df['میانگین'].mean(),
                'حداکثر': stats_df['حداکثر'].max(),
                'جمع': stats_df['جمع'].sum()
            }
            stats_df = pd.concat([stats_df, pd.DataFrame([totals])], ignore_index=True)

            st.toast("گزارش با موفقیت ایجاد شد.", icon='✅')

            return stats_df
        else:
            st.warning("Selected report type is not yet implemented.")
            return None

    except pyodbc.Error as ex:
        st.error(f"A database connection error occurred: {ex}")
        return None
    finally:
        if conn:
            conn.close()