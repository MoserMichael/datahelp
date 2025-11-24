

## intro

Some utility scripts for working with excel files; [pandas](https://pandas.pydata.org/) is used for accessing excel files, with [openpyxl](https://openpyxl.readthedocs.io/en/stable/) engine.

(works unexpected slowly on WSL2 but quickly on a MAC...)

## setup

```python

# first time: create virtual env
python3 -m venv .venv

# venv activate (each usage)
source .venv/bin/activate

# first time: install requirements
pip3 install -r requirements.txt
```

## difference between two excel files

```code
usage: exceldiff.py [-h] --excel1 EXCEL_FILE1 [--excel2 EXCEL_FILE2]

Compare two excel files, sheet by sheet. Ever had to modify an excel in a few places, then report back to the owner what changed?
This script compares the sheets of the excel, and reports the cells that changed!

options:
  -h, --help            show this help message and exit
  --excel1 EXCEL_FILE1, -f EXCEL_FILE1
                        file name of first excel input file
  --excel2 EXCEL_FILE2, -t EXCEL_FILE2
                        file name of second excel input file

```

## Extracting excel table to json


Usage:

```code

usage: excel2json.py [-h] --excel EXCEL_FILE [--tab EXCEL_TAB] --json OUT_FILE --row_from ROW_FROM --col_from COL_FROM
                     [--filter USE_COLUMNS]

Extract a table from an excel file into a json file. The json file is an array of records, where each row stands for a record in
the json. The row-record consists of name-value pairs, where the name is the table header and the value is the cell value for this
row. The initial psition is the start of he header line. All data lines are taken, until end of excel, or until the first line that
does not have any values Example: python3 excel2json.py --excel=tbl.xlsx --col_from=3 --row_from=1 --json=out.json Extract the
table where the header line starts from column 3 (one is the first column) and row 1 (one is the first column) python3
excel2json.py --excel=tbl.xlsx --json=out.json

options:
  -h, --help            show this help message and exit
  --excel EXCEL_FILE, -i EXCEL_FILE
                        file name of excel input file
  --tab EXCEL_TAB, -t EXCEL_TAB
                        tab name of excel tab
  --json OUT_FILE, -o OUT_FILE
                        output file file
  --row_from ROW_FROM, -x ROW_FROM
                        starting row of range (one based)
  --col_from COL_FROM, -y COL_FROM
                        starting column of range (one based)
  --filter USE_COLUMNS, -f USE_COLUMNS
                        filter a subset of column (comma delimited list of column names)
```

Example:

For an input table of the following form

<table>
    <th>
        <td>a col</td>
        <td>b col</td>
        <td>c col</td>
    <th>
    <tr>
        <td>1</td>
        <td>2</td>
        <td>3</td>
    </tr>
    <tr>
        <td>a</td>
        <td>b</td>
        <td>c</td>
    </tr>
    <tr>
        <td>4</td>
        <td>5</td>
        <td>6</td>
    </tr>
    <tr>
        <td>e</td>
        <td>f</td>
        <td>g</td>
    </tr>
</table>

This command 

```code
python3 excel2json.py --excel=tbl.xlsx --col_from=4 --row_from=2 --json=out.json
```

We will get a json of the following form

```code

[
    {
        "a-col": "1",
        "b-col": "2",
        "c-col": "3"
    },
    {
        "a-col": "a",
        "b-col": "b",
        "c-col": "c"
    },
    {
        "a-col": "4",
        "b-col": "5",
        "c-col": "6"
    },
    {
        "a-col": "e",
        "b-col": "f",
        "c-col": "g"
    }
]
```

The following command will filter a subset of the columns a-col and c-col

```code
python3 tabletojson.py --excel=tbl.xlsx --col_from=3 --row_from=1 --json=out.json --filter=a-col,c-col
```

And result in:


```code

[
    {
        "a-col": "1",
        "c-col": "3"
    },
    {
        "a-col": "a",
        "c-col": "c"
    },
    {
        "a-col": "4",
        "c-col": "6"
    },
    {
        "a-col": "e",
        "c-col": "g"
    }
]
```


## Searching for a value in all excel files in a subdirectory (and further down)

```code

python3 excelgrep.py --help

usage: excelgrep.py [-h] --dir SEARCH_DIR [--regex REGEX_SEARCH] [--search STRING_SEARCH]

Search cells in tabs of excel files (with extension .xlsx) in a given directory for a search term. Report matching lines

options:
  -h, --help            show this help message and exit
  --dir SEARCH_DIR, -d SEARCH_DIR
                        directory to search for excel files
  --regex REGEX_SEARCH, -r REGEX_SEARCH
                        Regular expression to search for
  --search STRING_SEARCH, -s STRING_SEARCH
                        String to search for
```

## Applying a python expression over a range of cells

Warning! This program uses pandas to read and write an excel workbook. 
This works fine for excel files with data only - without formatting / scripts.

Help text

Example usage

```code


python3 excelsed.py --excel a.xlsx  --col_from=4 --row_from=1 --row_to=5 --py="re.sub(r'[0-9]','_', arg)"
```

Help text 

````code 
python3 excelsed.py --help

usage: excelsed.py [-h] --excel EXCEL_FILE [--tab EXCEL_TAB] --row_from ROW_FROM --col_from COL_FROM [--row_to ROW_TO]
                   [--col_to COL_TO] --py PY_CODE

apply python expression on range of excel cells Allows to apply regex substitution over a bunch of excel cells The python code has
access to packages re and math, the input value is in global variable arg Example usage: python3 excelsed.py --excel a.xlsx
--col_from=4 --row_from=1 --row_to=5 --py="re.sub(r'[0-9]','_', arg)"

options:
  -h, --help            show this help message and exit
  --excel EXCEL_FILE, -i EXCEL_FILE
                        file name of excel input file
  --tab EXCEL_TAB, -t EXCEL_TAB
                        tab name of excel tab
  --row_from ROW_FROM, -x ROW_FROM
                        starting row of range (one based)
  --col_from COL_FROM, -y COL_FROM
                        starting column of range (one based)
  --row_to ROW_TO       ending row of range (must have at least --row_to or --col_to) (one based)
  --col_to COL_TO       ending row of range (must have at least --row_to or --col_to) (one based)
  --py PY_CODE, -p PY_CODE
                        python expression to transform each cell. The current cell value is in global variable arg
```

