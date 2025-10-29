
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