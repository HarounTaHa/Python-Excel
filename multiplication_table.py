import openpyxl, sys
from pathlib import Path
from openpyxl.styles import Font

if len(sys.argv) == 2:
    try:
        number = int(sys.argv[1])
    except Exception as e:
        print(e)
    excel_file = openpyxl.Workbook()
    sheet = excel_file.active
    for i in range(number + 1):
        for x in range(number + 1):
            # Check if in header row or column.
            is_header = False
            if i == 0 and x == 0:
                is_header = True
                n = ''
            elif x == 0:
                is_header = True
                n = i
            elif i == 0:
                is_header = True
                n = x
            else:
                n = x * i
            cell = sheet.cell(row=i + 1, column=x + 1)
            if is_header:
                cell.font = Font(bold=True)
            cell.value = n
    save_file = Path('D:\_Python_Projects\Python-Excel', 'multiplication_table_' + str(number) + '.xlsx')
    excel_file.save(save_file)
    print('Save as', save_file)

else:
    print('Please enter only two arguments')
