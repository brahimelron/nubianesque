import argparse
import os
import openpyxl

""" -------------------------------
  | I N I T I A L I Z A T I O N    |
    -------------------------------
"""
parser = argparse.ArgumentParser(description='Remove Work Orders off LAGAN Tracker Upon Resolution', prefix_chars='-+/')
parser.add_argument('--pswd', action='store', dest='usrpwd', help='Enter database login password')

userdomain = os.getenv('USERDOMAIN')
username = os.getenv('USERNAME')
userprof = os.getenv('USERPROFILE')
# Get the input file
inptBillingfl = 'P02 WO Completions Tracker.xlsx'
inptPath = os.path.join('K:\Projects\Work Order n Resource Mgmnt Sys', inptBillingfl)

""" -----------------------------------
  | F U N C T I O N: print_rows_of_data |
    -----------------------------------
"""

def print_rows_of_data():
    # ----------------------------------------------------
    # The function prints the contents of the sheet by row.|
    # ----------------------------------------------------
    book = openpyxl.load_workbook(inptPath)
    sheet = book.active
    format = "%m/%d/%Y"
    #Print rows of data
    for row in sheet.iter_rows(min_row=3, max_col=11, max_row=7):
        for cell in row:
            if (cell.col_idx == 1 and cell.column == 'A'):
                if cell.value == '=IF($A$2="","",$A$2)':
                    cell.value = sheet['A2'].value.strftime(format)
                    print(cell.value, end=" ")
                else:
                    print(cell.value, end=" ")
            else:
                print(cell.value, end=" ")
        print()


""" -------------------------------
  | F U N C T I O N: main          |
    -------------------------------
"""

def main():
    args = parser.parse_args()      # Get command line arguments
    print_rows_of_data()


""" -------------------------------
  | M A I N   P R O G R A M        |
    -------------------------------
"""
if __name__ == "__main__":
    main()
