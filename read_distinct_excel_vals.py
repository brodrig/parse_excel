import xlrd
import xlsxwriter
import os
# ==============================================
FILE_NAME = "file.xls"
COLUMN_VAL = 7 # Col to produce distinct values from
# ==============================================


def check_file_exists():
    if os.path.isfile(FILE_NAME):
        print "File exists."
        return True
    else:
        print "File does not exist!\nTerminating execution, please check file name and path."
        exit(1)


def distinct_parser(xls_file_name, col_val):
    # List for unique values
    distinct = []

    # Open MSExcel file/workbook for reading
    wb = xlrd.open_workbook(xls_file_name)
    workbook = xlsxwriter.Workbook("distinct_values_excel.xlsx")
    worksheet = workbook.add_worksheet()

    # Get the first worksheet in the workbook and print number of columns and rows
    sheet1 = wb.sheet_by_index(0)
    print "Number of Columns: {}".format(sheet1.ncols)
    print "Number of Rows: %d" % sheet1.nrows

    counter = 0
    for i in range(0, sheet1.nrows):
        if sheet1.cell(i, 7).value not in distinct:
            temp_val = str(sheet1.cell(i, col_val).value).strip()
            distinct.append(temp_val)
            worksheet.write(counter, 0, temp_val)
            counter += 1
    workbook.close()
    return distinct

if check_file_exists():
    # Pass filename to distinct parser
    unique_values = distinct_parser(FILE_NAME, COLUMN_VAL)
    # Print count of distinct values
    print "Count of distinct values: {}\n".format(len(unique_values))
    # Print distinct values
    for value in unique_values:
        print str(value)
