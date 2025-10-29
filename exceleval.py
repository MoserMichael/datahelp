import pandas as pd
import argparse
import sys
import re
import math

def parse_cmd_line():
    usage = '''
apply python expression on range of excel cells
Allows to apply regex substitution over a bunch of excel cells

The python code has access to packages re and math, the input value is in global variable arg

Example usage:

python3 exceleval.py --excel a.xlsx  --col_from=4 --row_from=1 --row_to=5 --py="re.sub(r'[0-9]','_', arg)"
'''

    parse = argparse.ArgumentParser(description=usage)

    parse.add_argument('--excel', 
                       '-i', 
                       required=True,
                       type=str,
                       dest='excel_file', 
                       help='file name of excel input file')

    parse.add_argument('--tab', 
                       '-t', 
                       type=str, 
                       dest='excel_tab',
                       help='tab name of excel tab')

    parse.add_argument('--row_from', 
                       '-x', 
                       required=True,
                       type=int, 
                       dest='row_from',
                       help='starting row of range')

    parse.add_argument('--col_from', 
                       '-y', 
                       required=True,
                       type=int, 
                       dest='col_from',
                       help='starting column of range')
    
    parse.add_argument('--row_to', 
                       type=int, 
                       dest='row_to',
                       help='ending row of range (must have at least --row_to or --col_to)')
    
    parse.add_argument('--col_to', 
                       type=int, 
                       dest='col_to',
                       help='ending row of range (must have at least --row_to or --col_to)')

    parse.add_argument('--py', 
                       '-p',
                       required=True,
                       type=str, 
                       dest='py_code',
                       help='python expression to transform each cell. The current cell value is in global variable arg')

    return parse.parse_args(), parse

def err(msg):
    print("Error: {msg}")
    sys.exit(1)

def check_vals(arg):
    if arg.col_from < 0:
        err("non negative --col_from expected")

    if arg.row_from < 0:
        err("non negative --row_from expected")

    if not arg.col_to and not arg.row_to:
        err("either --row_to or --col_to must be defined")


def process(arg, prs):
    check_vals(arg)

    if arg.excel_tab is not None:
        df = pd.read_excel(arg.excel_file, sheet_name=arg.excel_tab, header=None)
    else:
        df = pd.read_excel(arg.excel_file,header=None)
    
    x = to_x = arg.row_from
    y = to_y = arg.col_from
    if arg.row_to is not None:
        to_x = arg.row_to
    if arg.col_to is not None:
        to_y = arg.col_to


    for pos_x in range(x, to_x+1):
        for pos_y in range(y, to_y+1):
            cell = df.iat[pos_y, pos_x]
            if isinstance(cell,str):
                s_val = str(cell)

                r_val = eval(arg.py_code, {"re": re, "math": math, "arg": s_val})
                df.iat[pos_y, pos_x] = str(r_val)

    if arg.excel_tab is not None:
        df.to_excel(arg.excel_file, sheet_name=arg.excel_tab, header=None)
    else:
        df.to_excel(arg.excel_file, header=None)


def main():
    arg,prs = parse_cmd_line()
    process(arg, prs)

main()
