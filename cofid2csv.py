from openpyxl import load_workbook
import csv

columns = []
total_rows = 0
parent_columns = ['Food Name', 'Food Code', 'Description', 'Group', 'Previous', 'Main data references', 'Footnote']

sheet_file = 'CoFID_2019.xlsx'

print('Processing file "' + sheet_file + '"...')

cofid_sheet = load_workbook(sheet_file)
sheets_quantity = len(cofid_sheet.sheetnames)
parent_names = ['table_list', 'notes', 'factors', 'proximates', 'inorganics', 'vitamins', 'vitamin_fractions', 'SFA_FA',
                'SFA', 'MUFA_FA', 'MUFA', 'PUFA_FA', 'PUFA', 'phytosterols', 'organic_acids']

print('Processing ' + str(sheets_quantity) + ' worksheets...')

with open('columnNames.csv', 'r') as f:
    reader = csv.reader(f)
    next(reader)
    for row in reader:
        columns.append(row)


def format_data(sheet, file, parent):
    global total_rows
    ws = file[sheet]
    max_column = 0
    sheet_content = []
    row_content = []
    for cell in ws[1]:
        if cell.value:
            for column_index, field in enumerate(columns):
                cell_value = cell.value.strip()
                if cell_value == field[0].strip():
                    if cell_value in parent_columns:
                        row_content.append(field[2].strip())
                    else:
                        row_content.append(parent + '.' + field[2].strip())
            max_column = cell.column
    sheet_content.append(row_content)

    for sheet_row in ws.iter_rows(min_row=4, max_col=max_column):
        row_content = []
        for cell in sheet_row:
            cell_value = cell.value
            if cell.data_type == 's':
                cell_value = cell.value.strip()
                total_rows += 1
            row_content.append(cell_value)
        sheet_content.append(row_content)
    return sheet_content


def extract_and_save_csv(excel_file, sheet_name, csv_file, parent):
    file = open(csv_file, 'w')
    filewriter = csv.writer(file, delimiter=',', quoting=csv.QUOTE_MINIMAL)
    sheet_data = format_data(sheet_name, excel_file, parent)
    for sheet_row in sheet_data:
        filewriter.writerow(sheet_row)
    return


def format_number(number):
    return '%.0f' % number


for index, sheet_name in enumerate(cofid_sheet.sheetnames):
    newSheetName = sheet_name.replace('.', '_').replace(' ', '_').replace('(', '').replace(')', '')
    extract_and_save_csv(cofid_sheet, sheet_name, './csv_exports/cofid_' + newSheetName + '.csv', parent_names[index])
    print('...' + str(format_number((index + 1) * 100 / sheets_quantity) + '%'), end='')

print()
print('Completed processing ' + str(total_rows) + ' rows of data!!')
