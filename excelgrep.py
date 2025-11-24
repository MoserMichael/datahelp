import pandas as pd
import argparse
import sys
import re
import os
import math

def parse_cmd_line():
    usage = '''
for all files (with extension .xlsx) in a given directory/suddirectory:
Search all cells in all tabs  for a given search term and report matching lines
'''

    parse = argparse.ArgumentParser(description=usage)

    parse.add_argument('--dir', 
                       '-d', 
                       required=True,
                       type=str,
                   
                       dest='search_dir', 
                       help='directory to search for excel files')

    parse.add_argument('--regex', 
                       '-r', 
                       required=False,
                       type=str,
                       dest='regex_search', 
                       help='Regular expression to search for')


    parse.add_argument('--search', 
                       '-s', 
                       required=False,
                       type=str,
                       dest='string_search', 
                       help='String to search for')
    
    return parse.parse_args(), parse

def err(msg):
    print(f"Error: {msg}")
    sys.exit(1)

class RegexMatcher:
    def __init__(self, regex_str):
        try:
            #print(f"regex {regex_str}")
            self.r = re.compile(regex_str)
        except Exception as e:
            err(f"regular expression errro: {e}")
         
    def match(self, value):
        #return self.r.fullmatch(value) is not None
        return self.r.search(value) is not None

class StringMatcher:
    def __init__(self, search_str):
        self.search_str = search_str

    def match(self, value):
        return value.find(self.search_str) != -1
    

def check_vals(arg):
    #print(f"regex_search {arg.regex_search} string_search {arg.string_search}")
    if arg.regex_search:
        return RegexMatcher(arg.regex_search) 
    elif arg.string_search:
        return StringMatcher(arg.string_search)
    else:   
        err("either one of -s or -e options must be specified")


def search_excel(fpath, file, matcher):
    if file.startswith("~"): # lock file?
        return
    #print(f"file: {fpath}")

    df = pd.read_excel(fpath, header=None, keep_default_na=False, sheet_name=None)
    
    for sheet_name, df in df.items():
        #print(f"checking: {fpath}:{sheet_name}")
        num_columns = df.shape[1]
        num_rows = df.shape[0]

        for row in range(num_rows):

            row_match = False
            for col in range(num_columns):
                val = str(df.iat[row, col])
                #print(f"col {col} row {row} val {val}")
                if matcher.match(val):
                    row_match = True
                    break

            if row_match:
                csv_row = ""
                for col in range(num_columns):
                    if col != 0:
                        csv_row += ","
                    csv_row += str(df.iat[row, col]) 

                print(f"{file}:{sheet_name}:{row+1}:{csv_row}")                            



def match_all(search_dir, matcher):
    for root, _, files in os.walk(search_dir):
        for file in files:
            if file.endswith('.xlsx'):
                fpath = os.path.join(root, file)
                search_excel(fpath, file, matcher)   

def process(arg, prs):
    matcher = check_vals(arg)
    match_all(arg.search_dir, matcher)


def main():
    arg, prs = parse_cmd_line()
    process(arg, prs)

main()
