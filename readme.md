
## setup

```python

# first time: create virtual env
python3 -m venv .venv

# venv activate (each usage)
source .venv/bin/activate

# first time: install requirements
pip3 install -r requirements.txt
```

## applying a python regex over a range of cells

Help text

Example usage

```code

python3 exceleval.py --excel a.xlsx  --col_from=4 --row_from=1 --row_to=5 --py="re.sub(r'[0-9]','_', arg)"
```

Help text 

````code 
usage: exceleval.py [-h] --excel EXCEL_FILE [--tab EXCEL_TAB] --row_from ROW_FROM --col_from COL_FROM
                    [--row_to ROW_TO] [--col_to COL_TO] --py PY_CODE

apply python expression on range of excel cells Allows to apply regex substitution over a bunch of excel cells The
python code has access to packages re and math, the input value is in global variable arg Example usage: python3
exceleval.py --excel a.xlsx --col_from=4 --row_from=1 --row_to=5 --py="re.sub(r'[0-9]','_', arg)"

options:
  -h, --help            show this help message and exit
  --excel EXCEL_FILE, -i EXCEL_FILE
                        file name of excel input file
  --tab EXCEL_TAB, -t EXCEL_TAB
                        tab name of excel tab
  --row_from ROW_FROM, -x ROW_FROM
                        starting row of range
  --col_from COL_FROM, -y COL_FROM
                        starting column of range
  --row_to ROW_TO       ending row of range (must have at least --row_to or --col_to)
  --col_to COL_TO       ending row of range (must have at least --row_to or --col_to)
  --py PY_CODE, -p PY_CODE
                        python expression to transform each cell. The current cell value is in global variable arg
```


## extracting excel table to json

Example 

```code
python3 tabletojson.py --excel=tbl.xlsx --col_from=3 --row_from=1 --json=out.json
```

Usage:

```code

usage: tabletojson.py [-h] --excel EXCEL_FILE [--tab EXCEL_TAB] --json OUT_FILE --row_from ROW_FROM --col_from
                      COL_FROM

Extract a json table from given position in excel file into json file The initial psition is the start of he header
line. All data lines are taken, until end of excel, or until the first line that does not have any values Example:
python3 tabletojson.py --excel=tbl.xlsx --col_from=3 --row_from=1 --json=out.json

options:
  -h, --help            show this help message and exit
  --excel EXCEL_FILE, -i EXCEL_FILE
                        file name of excel input file
  --tab EXCEL_TAB, -t EXCEL_TAB
                        tab name of excel tab
  --json OUT_FILE, -o OUT_FILE
                        output file file
  --row_from ROW_FROM, -x ROW_FROM
                        starting row of range
  --col_from COL_FROM, -y COL_FROM
                        starting column of range
```

for an input table of the following form

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

we will get a json of the following form

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