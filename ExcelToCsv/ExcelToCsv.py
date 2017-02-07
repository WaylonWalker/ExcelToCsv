#! python
import glob
import csv
import argparse
import xlrd


def get_sheet_from_input(sheets):
    """
    returns a sheet index based on user input from a list of sheet names

    :list param sheets: List of sheet names
    :return: sheet index
    :rtype: int
    """
    sheet_choice = ''
    for i, val in enumerate(sheets):
        sheet_choice = sheet_choice + str(i + 1) + '. ' + val + '\n'
    sheet_idx = int(input(sheet_choice)) - 1
    return sheet_idx


def get_file_from_input(pattern='*.xl*'):
    """
    Gets a file name from user input

    :str param pattern:
    :return: filename
    :rtype: str
    """

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


def ExcelToCsv(input_file, sheet_idx=0, start=0, output_file=None):
    """
    Writes input_file to csv and saves it as output_file

    :str param input_file: Excel file to read
    :int param sheet_idx:  Sheet Number to read from the input file default 0
    :int param start: Line number to start parsing from the input file
    :str param output_file: csv file to save to
                            (if None input_file name will be used)

    """
    with xlrd.open_workbook(input_file) as wb:
        if not output_file:
            input_file[: input_file.find('.')] + '.csv'
        sheet = wb.sheet_by_index(sheet_idx)
        with open(output_file, 'w', encoding='utf8', newline='') as f:
            c = csv.writer(f)
            for r in range(start, sheet.nrows):
                c.writerow(sheet.row_values(r))


def console_tool():
    """
    Run ExcelToCsv from the shell using argparse

    """
    parser = argparse.ArgumentParser(description='Convert File Excel File to csv\
                                     quickly without loading excel')

    parser.add_argument('--input_file', metavar='F', type=str, nargs='+',
                        help='File to convert')
    parser.add_argument('--output_file', type=str,
                        help='Output file name if different than input')

    parser.add_argument('--sheet', type=int,
                        help='sheet number where first sheet is 1')

    parser.add_argument('--skiprows', type=int,
                        help='skip first n rows')

    parser.add_argument('--verbose', action='store_const', const='verbose',
                        help='asks for all arguments when not passed')

    args = parser.parse_args()

    print(r'''
  ______              _ _______     _____
 |  ____|            | |__   __|   / ____|
 | |__  __  _____ ___| |  | | ___ | |     _____   __
 |  __| \ \/ / __/ _ \ |  | |/ _ \| |    / __\ \ / /
 | |____ >  < (_|  __/ |  | | (_) | |____\__ \\ V /
 |______/_/\_\___\___|_|  |_|\___/ \_____|___/ \_/
    Efficiently Convert Excel Files to csv
    by Waylon S. Walker

                                                    ''')
    if args.input_file:
        input_file = args.input_file[0]
    else:
        input_file = get_file_from_input()

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
            sheet_idx = args.sheet - 1
        else:
            if args.verbose:
                sheet_idx = get_sheet_from_input(wb.sheet_names())
            else:
                sheet_idx = 0

        print('converting %s' % wb.sheet_by_index(sheet_idx).name)
    ExcelToCsv(input_file, sheet_idx=sheet_idx, start=start,
               output_file=output_file)


if __name__ == '__main__':
    console_tool()
