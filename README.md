# ExcelToCsv
A python package to convert excel spreadsheets to csv files without launching excel.  This works very well with large files that include many tabs which makes Excel very slow to open.

## requirements
xlrd

## Installation

```
git clone https://github.com/WaylonWalker/ExcelToCsv.git
cd ./ExcelToCsv
python setup.py install
```

## Usage
ExcelToCsv [-h] [--input_file F [F ...]] [--output_file OUTPUT_FILE]
                  [--sheet SHEET] [--skiprows SKIPROWS] [--verbose]

Convert File Excel File to csv quickly without loading excel

optional arguments:

  -h, --help            show this help message and exit
  
  --input_file F [F ...] File to convert
                        
  --output_file OUTPUT_FILE
                        Output file name if different than input
                        
  --sheet SHEET         sheet number where first sheet is 1
  
  --skiprows SKIPROWS   skip first n rows
  
  --verbose             asks for all arguments when not passed
  
## Example

```
ExcelToCvs --verbose
```

![example](example.gif)


