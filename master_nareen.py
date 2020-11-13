"""
This file takes in a 6a absence sheet for a given day and produces a summary sheet that takes multiple types of absence
data.
"""

import pandas as pd
import numpy as np


def build_WTE_lookup():
    print(f'Taking in Staff Download...')
    df = pd.read_excel('W:/Staff Downloads/2020-09 - Staff Download.xlsx')

    print(f'File downloaded (len={len(df)}).')
    dept_lookup = df[['WTE', 'department']].groupby(['department']).sum().round(1)
    sector_lookup = df[['WTE', 'Sector/Directorate/HSCP']].groupby(['Sector/Directorate/HSCP']).sum().round(1)
    subdir1_lookup = df[['WTE', 'Sub-Directorate 1']].groupby(['Sub-Directorate 1']).sum().round()
    return df, dept_lookup, sector_lookup, subdir1_lookup

def take_in_6a():
    print(f'Taking in 6a Absence file.')
    newpath = 'W:/Daily_Absence/' + (pd.Timestamp.now()).strftime("%Y-%m-%d") + '.xls'
    df = pd.read_excel(newpath, skiprows=4)
    df = df[df['Absence Type'] != 'Annual leave']
    print(f'6a file loaded (len={len(df)})')
    return df

def merge_data(abs, staff_download):
    print(abs.columns)
    abs.rename(columns={'Pay No':'Pay_Number', 'AbsenceReason Description':'Absence_Reason',
                        'Absence Type':'Absence_Type'}, inplace=True)
    abs = abs[['Pay_Number', 'Absence_Reason', 'Absence_Type']]
    merged = abs.merge(staff_download[['department', 'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Cost_Centre', 'Pay_Number', 'WTE']])
    return merged

def abs_type_pivot(merged_data):
    type_piv = pd.pivot_table(merged_data, index=['Absence_Type', 'Absence_Reason'], values='WTE', aggfunc=np.sum)
    type_piv.reset_index(inplace=True)
    type_piv['WTE'] = type_piv['WTE'].round(1)
    return type_piv

def sector_fraction_piv(merged_data, sector_lookup):
    sector_piv = pd.pivot_table(merged_data, index=['Sector/Directorate/HSCP', 'Absence_Type', 'Absence_Reason'],
                                values='WTE', aggfunc=np.sum)
    sector_piv['WTE'] = sector_piv['WTE'].round(1)
    sector_piv.reset_index(inplace=True)
    sector_lookup.rename(inplace=True, columns={'WTE':'Sector WTE'})
    sector_piv = sector_piv.merge(sector_lookup, on='Sector/Directorate/HSCP', how='left')
    sector_piv['% Absent'] = ((sector_piv['WTE'] / sector_piv['Sector WTE']) * 100).round(2)
    sector_piv.drop(columns=['Sector WTE'], inplace=True)

    return sector_piv

def build_output_file(merged_data, dept_lookup, sector_lookup, type_piv, sector_piv, subdir1_lookup):
    print(f'Building excel sheet')
    dept_lookup.rename(columns={'WTE':'Department WTE'}, inplace=True)



    with pd.ExcelWriter('C:/tong/nareen_test_file.xlsx', engine='xlsxwriter') as writer:
        workbook = writer.book
        header_format = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 16})
        subheader = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 14})
        subheadernoBold = workbook.add_format({'font_name': 'Arial', 'font_size': 14})
        table_format_ul = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white', 'underline': True})
        table_format = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white'})
        back_button = table_format_ul = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white',
                                                             'underline': True, 'align': 'center', 'valign': 'vcenter'})
        # create cover page
        worksheet =workbook.add_worksheet('Contents')
        worksheet.hide_gridlines(2)
        worksheet.set_column('C:C', 26)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('B:B', 36)
        worksheet.set_column('A:A', 5)
        worksheet.set_column('E:E', 20)

        # write headers
        worksheet.write(0, 1, 'Daily Absence - Summary Report', header_format)
        worksheet.write(1, 1, 'Date:', subheader)
        worksheet.write(1, 2, f'{pd.Timestamp.now().strftime("%d %B %Y")}', subheadernoBold)
        worksheet.insert_image('D1', 'W:/Danny/ggclogo.jpg', {'x_scale':0.25, 'y_scale':0.25})

        worksheet.write(3, 1, f'Sector/Directorate/HSCP', subheader)
        worksheet.write(3, 2, f'Absence Episodes', subheader)
        worksheet.write(3, 3, f'Absence WTE', subheader)
        worksheet.write(3, 4, f'Sector WTE', subheader)
        worksheet.write(3, 5, f'Absence %', subheader)

        current_row = 4
        sector_lookup.rename(columns={'WTE':'Sector WTE'}, inplace=True)
        secwte = merged_data.merge(sector_lookup, on='Sector/Directorate/HSCP', how='left')
        for i in merged_data['Sector/Directorate/HSCP'].unique():

            sector_wte = secwte[secwte['Sector/Directorate/HSCP'] == i]['Sector WTE'].min()
            abs_wte = merged_data[merged_data['Sector/Directorate/HSCP'] == i]['WTE'].sum().round(1)
            if len(i) > 31:
                worksheet.write_url(current_row, 1, "internal:'"+i[:30]+"'!A1", string=i, cell_format=table_format_ul)
            else:
                worksheet.write_url(current_row, 1, "internal:'"+i+"'!A1", string=i, cell_format=table_format_ul)
            worksheet.write(current_row, 2, len(merged_data[merged_data["Sector/Directorate/HSCP"] == i]),
                            table_format)
            worksheet.write(current_row, 3, abs_wte,  table_format)
            worksheet.write(current_row, 4, sector_wte, table_format)
            worksheet.write(current_row, 5, ((abs_wte/sector_wte)*100).round(1), table_format)

            current_row +=1
        worksheet.conditional_format(f'F4:F{current_row + 1}',
                                     {"type": "3_color_scale", "min_color": 'green', 'min_value': 0, 'max_value': 30,
                                      'max_color': 'red'})
        # write a sheet for each sector
        for i in merged_data['Sector/Directorate/HSCP'].unique():
            df = merged_data[merged_data['Sector/Directorate/HSCP'] == i]

            if len(i) > 31:
                sheetname = i[:30]
            else:sheetname = i
            piv_subdir = pd.pivot_table(df, index=['Sub-Directorate 1'], values='WTE', aggfunc=np.sum).round(1)
            piv_subdir.reset_index(inplace=True)
            piv_subdir.rename(columns={'WTE':'Absence WTE'}, inplace=True)
            piv_subdir = piv_subdir.merge(subdir_lookup, on='Sub-Directorate 1', how='left')
            piv_subdir['Absence %'] = ((piv_subdir['Absence WTE'] / piv_subdir['WTE'])*100).round(1)
            piv_subdir.to_excel(writer, sheet_name=sheetname, startrow=2, index=False)


            piv = pd.pivot_table(df, index=['Sub-Directorate 1', 'department', 'Absence_Type', 'Absence_Reason'], values='WTE',
                                 aggfunc=np.sum)
            piv.reset_index(inplace=True)
            piv = piv.merge(dept_lookup, on='department', how='left')
            print(piv.columns)
            piv['% Absent'] = ((piv['WTE'] / piv['Department WTE']) * 100).round(2)
            piv['WTE'] = piv['WTE'].round(1)
            piv['Size of Dept'] = pd.cut(piv['Department WTE'], bins=[0, 15, 30, 60, 10000],
                                          labels=['Small', 'Medium', 'Large', 'Extra Large'])
            piv.sort_values(['Size of Dept', 'department'], ascending=[False, True], inplace=True)
            piv.rename(columns={'WTE':'Absence WTE', 'Department WTE':'Dept WTE'}, inplace=True)
            print(piv.columns)
            piv = piv[['Sub-Directorate 1', 'department','Size of Dept', 'Absence_Type', 'Absence_Reason', 'Absence WTE', 'Dept WTE',
       '% Absent']]
            piv.to_excel(writer, sheet_name=sheetname, index=False, startrow=len(piv_subdir) + 4)
            sheet = writer.sheets[sheetname]
            sheet.write('A1', f'{i} - Absence Summary - {pd.Timestamp.now().strftime("%d %B %Y")}', header_format)
            sheet.conditional_format(f'H{3 + len(piv_subdir)}:H{len(piv) + len(piv_subdir) + 3}',
                                     {"type":"3_color_scale", "min_color":'green', 'max_color':'red'})
            sheet.conditional_format(f'D4:D{len(piv_subdir) + 3}',
                                     {"type": "3_color_scale", "min_color": 'green', 'max_color': 'red'})

            sheet.set_column('A:A', 25)
            sheet.set_column('D:D', 15)
            sheet.set_column('E:E', 72.43)
            sheet.set_column('C:C', 24.86)
            sheet.set_column('B:B', 28.14)
            sheet.set_column('F:F', 15)
            sheet.set_column('G:H', 20)

            sheet.autofilter(f'A{len(piv_subdir) + 5}:H{len(piv) + len(piv_subdir) + 5}')
            sheet.freeze_panes(3, 0)
            sheet.write_url('D2', "internal:'Contents'!A1",
                            string='Back', cell_format=back_button)
        # merged_data.to_excel(writer, sheet_name="Merged Data", index=False)
        # dept_lookup.to_excel(writer, sheet_name="Department Lookup")
        # sector_piv.to_excel(writer, sheet_name='Sector pivot', index=False)
        # sector_lookup.to_excel(writer, sheet_name="Sector Lookup")
        # type_piv.to_excel(writer, sheet_name="Type Pivot", index=False)




sd, dept_lookup, sector_lookup, subdir_lookup = build_WTE_lookup()
abs_6a = take_in_6a()
merged = merge_data(abs_6a, sd)
type_piv = abs_type_pivot(merged)
sector_piv = sector_fraction_piv(merged, sector_lookup)
build_output_file(merged, dept_lookup, sector_lookup, type_piv, sector_piv, subdir_lookup)

