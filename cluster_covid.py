'''Takes a long term 6a report and tries to locate clusters of workplace infections by department'''
import pandas as pd
from tkinter.filedialog import askopenfilename


absence_data = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Daily_Absence',
                               filetypes=(("Excel File", "*.xls"), ("All Files", "*.*")),
                               title="Choose the relevant absence extract."
                               )
df = pd.read_excel(absence_data, skiprows=4)
print(df.columns)
sd = pd.read_excel(
    'W:/Workforce Monthly Reports/Monthly_Reports/Oct-20 Snapshot/Staff Download/2020-10 - Staff Download - GGC.xls')
print(sd.columns)

df = df.rename(columns={'Pay No': 'Pay_Number', 'Absence Episode Start Date':'AbsStart', 'Absence Episode End Date':'AbsEnd'})
df = df.merge(sd[['Cost_Centre', 'Pay_Number']], on='Pay_Number', how='left')
df = df[~(df['Cost_Centre'].isna())]
print(len(df))
df = df[df['AbsenceReason Description'].str.contains('Coronavirus – Covid 19 Positive')]
df = df[~(df['AbsenceReason Description'].str.contains('Household'))]
df = df[~(df['AbsenceReason Description'].str.contains('Underlying'))]
print(len(df))

clusters = []
for i in df['Cost_Centre'].unique():
    df_all_cases = df[df['Cost_Centre'] == i]
    print(i)
    if len(df_all_cases) == 1:
        print(f'{i} - single case')
        continue
    else:
        df_all_cases.sort_values(by='AbsStart')
        print(f'{i} - {len(df_all_cases)} cases')
        curr_start_date = df_all_cases['AbsStart'].iloc[0]
        print(curr_start_date)
        df_all_cases.set_index(['AbsStart'])

        for row in df_all_cases.itertuples():
            curr_start_date = row.AbsStart
            start_14 = curr_start_date + pd.DateOffset(days=14)
            df2 = df_all_cases[(df_all_cases['AbsStart'] >= curr_start_date) & (df_all_cases['AbsStart'] < start_14)]
            df2.drop_duplicates(subset='Pay_Number', inplace=True)
            if len(df2) == 1:
                print(f" 1 case only")
                continue
            else:
                print(f'Cluster found - {len(df2)} within 2 weeks.')
                print(df2[['AbsStart']])
                clusters.append(i)
#df = df[df['Cost_Centre'].isin(clusters)]
df.drop(columns='Cost_Centre', inplace=True)
df = df.merge(sd, on='Pay_Number')
with pd.ExcelWriter('C:/Tong/Cluster_counter.xlsx', engine='xlsxwriter') as writer:
    # cover sheet
    workbook = writer.book
    header_format = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 16})
    subheader = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 14})
    subheadernoBold = workbook.add_format({'font_name': 'Arial', 'font_size': 14})
    table_format_ul = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white', 'underline': True})
    table_format = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white'})
    back_button = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white',
                                                         'underline': True, 'align': 'center', 'valign': 'vcenter'})
    cluster_format = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'red'})
    # create cover page
    worksheet = workbook.add_worksheet('Contents')
    worksheet.hide_gridlines(2)
    worksheet.set_column('C:C', 26)
    worksheet.set_column('D:D', 17)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('A:A', 33)
    worksheet.set_column('E:E', 20)

    # write headers
    worksheet.write(0, 0, 'Potential Covid Clusters - Summary Report', header_format)
    worksheet.write(1, 0, 'Date:', subheader)
    worksheet.write(1, 1, f'{pd.Timestamp.now().strftime("%d %B %Y")}', subheadernoBold)
    worksheet.insert_image('D1', 'W:/Danny/ggclogo.jpg', {'x_scale': 0.25, 'y_scale': 0.25})

    worksheet.write(3, 0, f'Sector/Directorate/HSCP', subheader)
    worksheet.write(3, 1, f'Cost Centre', subheader)
    worksheet.write(3, 2, f'Department Name', subheader)
    worksheet.write(3, 3, f'Positives', subheader)
    worksheet.write(3, 4, f'Cluster', subheader)


    current_row = 4
    dept_lookup = dict(zip(df['Cost_Centre'], df['department']))
    sector_lookup = dict(zip(df['Cost_Centre'], df['Sector/Directorate/HSCP']))
    df.sort_values(by='Sector/Directorate/HSCP', inplace=True)
    for i in df['Cost_Centre'].unique():
        worksheet.write(current_row, 0, sector_lookup.get(i), table_format)
        worksheet.write_url(current_row, 1, "internal:'" + i + "'!A1", string=i, cell_format=table_format_ul)
        worksheet.write(current_row, 2, dept_lookup.get(i), table_format)
        worksheet.write(current_row, 3, len(df[df['Cost_Centre'] == i]), table_format)
        if i in clusters:
            worksheet.write(current_row, 4, 'Potential cluster', cluster_format)
        else:
            worksheet.write(current_row, 4, '', cluster_format)

        current_row += 1



    for dept in df['Cost_Centre'].unique():
        df[df['Cost_Centre'] == dept][['Pay_Number', 'Roster Location','department', 'Cost_Centre',
                                       'AbsenceReason Description', 'AbsStart', 'AbsEnd','Sub-Directorate 1',
                                       'Sub-Directorate 2', 'Job_Family', 'Sub_Job_Family'
                                       ]].to_excel(sheet_name=dept, startrow=2, excel_writer=writer, index=False)
        sheet = writer.sheets[dept]
        sheet.write('A1', f'{dept_lookup.get(dept)} - Covid Case Summary - {pd.Timestamp.now().strftime("%d %B %Y")}', header_format)
        sheet.write_url('D2', "internal:'Contents'!A1",
                        string='Back', cell_format=back_button)
        sheet.set_column('A:F', 15)
        sheet.set_column('G:I', 20)
        sheet.set_column('J:M', 15)




