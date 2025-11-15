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

python3 excelsed.py --excel a.xlsx  --col_from=4 --row_from=1 --row_to=5 --py="re.sub(r'[0-9]','_', arg)"
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
                       help='starting row of range (one based)')

    parse.add_argument('--col_from', 
                       '-y', 
                       required=True,
                       type=int, 
                       dest='col_from',
                       help='starting column of range (one based)')
    
    parse.add_argument('--row_to', 
                       type=int, 
                       dest='row_to',
                       help='ending row of range (must have at least --row_to or --col_to) (one based)')
    
    parse.add_argument('--col_to', 
                       type=int, 
                       dest='col_to',
                       help='ending row of range (must have at least --row_to or --col_to) (one based)')

    parse.add_argument('--py', 
                       '-p',
                       required=True,
                       type=str, 
                       dest='py_code',
                       help='python expression to transform each cell. The current cell value is in global variable arg')

    return parse.parse_args(), parse

def err(msg):
    print(f"Error: {msg}")
    sys.exit(1)

def check_vals(arg):
    if arg.col_from < 0:
        err("non negative --col_from expected")

    if arg.row_from < 0:
        err("non negative --row_from expected")

    if not arg.col_to and not arg.row_to:
        err("either one or both of --row_to and --col_to must be defined")

    if arg.col_to and arg.col_from > arg.col_to:
        err("col_to must not be smaller than col_from")

    if arg.row_to and arg.row_from > arg.row_to:
        err("row_to must not be smaller than row_from")

def process(arg, prs):
    check_vals(arg)

    #if arg.excel_tab is not None:
    #    df = pd.read_excel(arg.excel_file, sheet_name=arg.excel_tab, header=None, keep_default_na=False)
    #else:
    #    df = pd.read_excel(arg.excel_file,header=None, keep_default_na=False)

    def convert(val):
        return eval(arg.py_code, {"re": re, "math": math, "arg": val})

    df_all = pd.read_excel(arg.excel_file,header=None, sheet_name=None, keep_default_na=False)
    
    for sheet_name, df in df_all.items():

        if arg.excel_tab is not None and sheet_name.strip() != arg.excel_tab.strip():
            continue

        num_columns = df.shape[1]
        num_rows = df.shape[0]

        x = to_x = arg.row_from
        y = to_y = arg.col_from
        if arg.row_to is not None:
            to_x = arg.row_to
        if arg.col_to is not None:
            to_y = arg.col_to

                
        for pos_x in range(x, to_x+1):

            if pos_x >= num_columns:
                continue 

            for pos_y in range(y, to_y+1):

                if pos_y >= num_rows:
                    continue 

                cell = df.iat[pos_y, pos_x]
                s_val = ""
                if isinstance(cell,str):
                    s_val = str(cell).strip()
                elif isinstance(cell, int) or isinstance(cell, pd.StringDtype):
                    s_val = str(cell)
                elif isinstance(cell, float):
                    if not math.isinf(cell):
                        s_val = str(cell)

                r_val = convert(s_val)
                df.iat[pos_y, pos_x] = str(r_val)

    #if arg.excel_tab is not None:
    #    df.to_excel(arg.excel_file, sheet_name=arg.excel_tab, header=None)
    #else:
    #    df.to_excel(arg.excel_file, header=None)
    
    df.to_excel(arg.excel_file, header=None)


def main():
    arg,prs = parse_cmd_line()
    process(arg, prs)

main()
