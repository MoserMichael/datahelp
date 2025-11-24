import pandas as pd
import argparse
import sys
import re
import json
import math

def parse_cmd_line():
    usage = '''
Compare two excel files, sheet by sheet.

Ever had to modify an excel in a few places, then report back to the owner what changed?
This script compares the sheets of the excel, and reports the cells that changed!
'''

    parse = argparse.ArgumentParser(description=usage)

    parse.add_argument('--excel1', 
                       '-f', 
                       required=True,
                       type=str,
                       dest='excel_file1', 
                       help='file name of first excel input file')

    parse.add_argument('--excel2', 
                       '-t', 
                       type=str, 
                       dest='excel_file2',
                       help='file name of second excel input file')


    return parse.parse_args(), parse

def err(msg):
    print(f"Error: {msg}")
    sys.exit(1)

def sheet_list(df):
    lst = []
    for sheet_name, sh in df.items():
        lst.append( (sheet_name.strip(), sh) )
    return lst

def sheet_name_map(sheet_list):
    ret_map={}
    for elm in sheet_list:
        ret_map[ elm[0] ] = elm[1]
    return ret_map

def common_sheets(excel_file1, df1, excel_file2, df2):
    sheets1 = sheet_list(df1)
    sheets2 = sheet_list(df2)

    name_map1 = sheet_name_map(sheets1)
    name_map2 = sheet_name_map(sheets2)

    common_sheets = []

    for elm in sheets1:
        sheet_name = elm[0]
        if sheet_name in name_map2:
            common_sheets.append( (sheet_name, elm[1], name_map2[sheet_name]))

            del name_map1[sheet_name]
            del name_map2[sheet_name]

    if len(name_map1) != 0:
        print(f"sheets exclusive to {excel_file1} : {list(name_map1.keys())}\n")
    if len(name_map2) != 0:
        print(f"sheets exclusive to {excel_file2} : {list(name_map2.keys())}\n")

    return common_sheets

def row_to_excel(num):
    temp_column = num
    column_string = ""
    
    while temp_column > 0:
        remainder = (temp_column - 1) % 26
        
        # Add the character corresponding to the remainder, starting from 'A'
        column_string += chr(ord('A') + remainder)
        
        # Update the column number for the next iteration
        temp_column = (temp_column - 1) // 26

    # Reverse the string to get the correct column name
    return column_string[::-1]

def compare_sheet(sheet_name, sh1, sh2):
    num_columns1 = sh1.shape[1]
    num_rows1 = sh1.shape[0]

    num_columns2 = sh2.shape[1]
    num_rows2 = sh2.shape[0]

    diff_cells = []

    for idx_x in range(0, max(num_columns1, num_columns2)):
        for idx_y in range(0, max(num_rows1, num_rows2)):

            cell1 = ""
            if idx_x < num_columns1 and idx_y < num_rows1:
                cell1 = sh1.iat[idx_y, idx_x]

            cell2 = ""
            if idx_x < num_columns2 and idx_y < num_rows2:
                cell2 = sh2.iat[idx_y, idx_x]

            if cell1 != cell2:
                diff_cells.append( (idx_x, idx_y, cell1, cell2) )

    # report cell differences
    if len(diff_cells) != 0:
        print(f"Differences in sheet {sheet_name}\n")
        for elm in diff_cells:
            print(f"\tcell={row_to_excel(elm[0]+1)}{elm[1]+1} changed from='{elm[2]}' to='{elm[3]}'")
        print("")


def process(arg, excel_file1, excel_file2):
    df1 = pd.read_excel(excel_file1, header=None, keep_default_na=False, sheet_name=None)
    df2 = pd.read_excel(excel_file2, header=None, keep_default_na=False, sheet_name=None)

    common = common_sheets(excel_file1, df1, excel_file2, df2)
    for elm in common:
        compare_sheet(elm[0], elm[1], elm[2])


def main():
    arg,prs = parse_cmd_line()
    process(arg, arg.excel_file1, arg.excel_file2)

main()
