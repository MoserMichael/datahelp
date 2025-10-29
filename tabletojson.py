import pandas as pd
import argparse
import sys
import re
import json
import math

def parse_cmd_line():
    usage = '''
Extract a json table from given position in excel file
into json file

The initial psition is the start of he header line.
All data lines are taken, until end of excel, or until the first line that does not have any values

Example:

python3 tabletojson.py --excel=tbl.xlsx --col_from=3 --row_from=1 --json=out.json

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

    parse.add_argument('--json', 
                       '-o', 
                       required=True,
                       type=str,
                       dest='out_file', 
                       help='output file file')

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
    
    return parse.parse_args(), parse

def err(msg):
    print("Error: {msg}")
    sys.exit(1)

def check_vals(arg):
    if arg.col_from < 0:
        err("non negative --col_from expected")

    if arg.row_from < 0:
        err("non negative --row_from expected")

def parse_header(df, x, y):
    out_header = []

    num_columns = df.shape[1]
    #print(f"shape: {df.shape}")

    #num_rows = df.shape[0]

    x_cur = x
    while True:
        cell = df.iat[y, x_cur]
        
        s_val = ""
        if isinstance(cell,str):
            s_val = str(cell).strip()
        elif isinstance(cell, int) or isinstance(cell, pd.StringDtype):
            s_val = str(cell)
        elif isinstance(cell, float):
            if not math.isinf(cell):
                s_val = str(cell)

        if s_val == "":
            break

        s_val = s_val.replace(' ', '-')

        out_header.append(s_val)
        x_cur += 1
        
        if x_cur >= num_columns:
            break  
            
    return out_header

def process(arg, prs):
    check_vals(arg)

    if arg.excel_tab is not None:
        df = pd.read_excel(arg.excel_file, sheet_name=arg.excel_tab, header=None)
    else:
        df = pd.read_excel(arg.excel_file,header=None)

    header_names = parse_header(df, arg.row_from, arg.col_from)

    num_rows = df.shape[0]

    json_data = []
    x = arg.row_from
    y = arg.col_from + 1
    while True:
        if y >= num_rows:
            break  
        row_entry = {}
        all_empty_vals = True
        for x_cur in range(0, len(header_names)):
            cell = df.iat[y, x+x_cur]
            
            s_val = ""
            if isinstance(cell,str) or isinstance(cell, int) or isinstance(cell, pd.StringDtype):
                s_val = str(cell)
            elif isinstance(cell, float):
                if not math.isinf(cell):
                    s_val = str(cell)

            if s_val != "":
                all_empty_vals = False

            row_entry[ header_names[x_cur] ] = s_val

        if all_empty_vals:
                break 
        json_data.append(row_entry)
        
        y += 1


        # save it
        out_data = json.dumps(json_data, indent=4)
        with open(arg.out_file, 'w') as out_file:
            out_file.write(out_data)            

def main():
    arg,prs = parse_cmd_line()
    process(arg, prs)

main()
