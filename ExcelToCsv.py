#! python
import xlrd
import csv
import argparse
import glob
import os

parser = argparse.ArgumentParser(description='Convert File Excel File to CSV quickly')

parser.add_argument('--input_file', metavar='F', type=str, nargs='+',
                    help='File to convert')

parser.add_argument('--output_file', type=str,
    help='Output file name if different than input')

parser.add_argument('--sheet', type=int,
    help='sheet number where first sheet is 1')

parser.add_argument('--skiprows', type=int, 
    help='skip first n rows')

parser.add_argument('--verbose', action='store_const', const='verbose',
    help='asks for all arguments when not passed' )

args = parser.parse_args()


def get_sheet_from_input(sheets):
    sheet_choice = ''
    for i, val in enumerate(sheets):
        sheet_choice = sheet_choice + str(i + 1) + '. ' + val + '\n'
    sheet = wb.sheet_by_index(int(input(sheet_choice)) - 1)
    return sheet

def get_files_from_input(pattern='*.xl*'):
    print(os.getcwd())
    file_choice = ''
    files = ['Not Shown in list']
    for suggested_file in glob.glob(pattern):
        files.append(suggested_file)
    for i, file in enumerate(files):
        file_choice = file_choice + str(i) + '. ' + str(file) + '\n'

    file = files[int(input(file_choice))]
    if file == files[0]:
        file = input('Enter input filename ')
    return file

if args.input_file:
    input_file = args.input_file[0]
else:
    input_file = get_files_from_input()
    print(input_file)
    
if args.output_file:   
    output_file = args.output_file
else:
    output_file = input_file[: input_file.find('.')] + '.csv'
    if args.verbose:
        use_default = input(r'%s  filename  OK? y/n ' % output_file)
        if use_default == 'n':
            output_file = input('Enter destination filename: ')


if args.skiprows:
    start = args.skiprows
else:
    if args.verbose:
        start = int(input('skip first n rows: '))
    else:
        start = 0

with xlrd.open_workbook(input_file) as wb:
    if args.sheet:
        sh = wb.sheet_by_index(args.sheet - 1)
    else:
        if args.verbose:
            sh = get_sheet_from_input(wb.sheet_names())
        else:
            sh = wb.sheet_by_index(0)
    with open(output_file, 'w', encoding='utf8', newline='') as f:
        c = csv.writer(f)
        for r in range(start, sh.nrows):
#             values = [value.encode('utf-8') for value in sh.row_values(r)]
            c.writerow(sh.row_values(r))
print('converting %s' % sh.name)