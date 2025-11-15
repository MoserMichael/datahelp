import pandas as pd
import argparse
import sys
import re
import json
import math

def parse_cmd_line():
    usage = '''
Extract a table from an excel file into a json file.
The json file is an array of records, where each row stands for a record in the json.
The row-record consists of name-value pairs, where the name is the table header and the value is the cell value for this row.

The initial psition is the start of he header line.
All data lines are taken, until end of excel, or until the first line that does not have any values

Example:

python3 excel2json.py --excel=tbl.xlsx --col_from=3 --row_from=1 --json=out.json

Extract the table where the header line starts from column 3 (one is the first column) and row 1 (one is the first column) 

python3 excel2json.py --excel=tbl.xlsx  --json=out.json


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
                       help='starting row of range (one based)')

    parse.add_argument('--col_from', 
                       '-y', 
                       required=True,
                       type=int, 
                       dest='col_from',
                       help='starting column of range (one based)')

    parse.add_argument('--filter', 
                       '-f', 
                       required=False,
                       default="",
                       type=str,
                       dest='use_columns', 
                       help='filter a subset of column (comma delimited list of column names)')


    return parse.parse_args(), parse

def err(msg):
    print(f"Error: {msg}")
    sys.exit(1)

def check_vals(arg):
    if arg.col_from < 0:
        err("positive (greater equal to one) value for --col_from expected")

    if arg.row_from < 0:
        err("positive (greater equal to one) value for --row_from expected")

    if arg.use_columns == "":
        return []
    return list(map(lambda arg : arg.strip(), arg.use_columns.split(",")))

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
    filter_columns = check_vals(arg)

    if arg.excel_tab is not None:
        df = pd.read_excel(arg.excel_file, sheet_name=arg.excel_tab, header=None,keep_default_na=False)
    else:
        df = pd.read_excel(arg.excel_file,header=None,keep_default_na=False)

    header_names = parse_header(df, arg.row_from-1, arg.col_from-1)
    print(f"table headers: {header_names}")

    num_rows = df.shape[0]

    json_data = []
    x = arg.row_from - 1
    y = arg.col_from
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

            if len(filter_columns) != 0 and not header_names[x_cur] in filter_columns:
                print(f"skippping '{header_names[x_cur]}' {filter_columns}")
                continue

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
