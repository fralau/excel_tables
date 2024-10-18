#!/usr/bin/env python3
"Test the program"

import os


import click
import pandas as pd



from rich import print
from rich.panel import Panel
from icecream import ic

from excel_tables import ExcelReport, Worksheet, df_columns, ExcelDB 


# ------------------------
# Initialization
# ------------------------

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_FILENAME = os.path.join(CURRENT_DIR, "mountains.xlsx")
OUT_FILE = os.path.join(CURRENT_DIR, "output.xlsx")

def add_suffix(filename:str, suffix:str):
    "Add a suffix to the basename of a file"
    fn, extension = os.path.splitext(filename)
    return ''.join((fn, suffix, extension))

def title(s: str):
    "Print a title"
    print(Panel(f"[green]{s}", expand=False))


def compare_df(df1: pd.DataFrame, df2: pd.DataFrame):
    """
    Compare two dataframes 
    (it works as a left join, i.e. it checks only the columns in df1)
    """
    assert len(df1) == len(df2)
    df1_cols = list(df1.columns)
    df2_cols = list(df2.columns)
    assert all(col in df2_cols for col in df1_cols), "Columns are different"
    for col in df1.columns:
        print(f"Checking column {col}")
        assert df1[col].equals(df2[col]), "Column is different"

# ------------------------
# Tests
# ------------------------



def test_short(in_file:str = TEST_FILENAME, 
               out_file:str = OUT_FILE,
               open_file:bool=False,
                debug:bool=False):
    "Short version"
    title("Short version")
    xl = pd.ExcelFile(in_file)
    df = xl.parse(0)
    assert df_columns(df)['Ascension'] == 'date'
    report = ExcelReport(out_file, font_name='Times New Roman', 
                        df=df,
                        emphasize=lambda x: x[1] > 8200,
                        debug=debug)
    report.rich_print()
    if open_file:
        report.open()

    # Read back and make sure we have the same thing in the dataframes
    # from the original and destination files (which is what matters)
    # This is exceptionally important for dates
    # Note that ExcelDB converts all ISO dates (str) into datetime objects. 
    xldb = ExcelDB(out_file)
    df_back = xldb.table(0)
    compare_df(df, df_back)



def test_long(in_file:str = TEST_FILENAME, 
              out_file:str = OUT_FILE,
              open_file:bool=False,
              debug:bool=False):
    title("Long version")
    second_out_file = add_suffix(out_file, '_mult')
    # read the whole file into a db
    xldb = ExcelDB(in_file)
    print(f"  {second_out_file}")
    report = ExcelReport(second_out_file, 
                        font_name='Helvetica', 
                        format_int="[>=1000]#'##0;[<1000]0",
                        format_float="[>=1000]#'##0.00;[<1000]0.00",
                        format_date="DD-MM-YYYY",
                        debug=debug)
    try:
        print(report)
    except KeyError:
        # No report available yet.
        pass
    

    title("First worksheet")
    tab_name = 'Mountains'
    mountains = xldb.table(0)
    print(df_columns(mountains))
    assert df_columns(mountains)['Ascension'] == 'date'
    wks = report.add_sheet(tab_name, mountains, 
                            emphasize=lambda x: x[1] > 8500,
                            num_formats={'Feet': "#'##0"})
    assert tab_name in report.tabs
    print("Columns:", wks.columns)
    print(report)

    title("Second worksheet")
    # filter where height > 8000
    MAX_ALTITUDE = 8000
    ic(MAX_ALTITUDE)
    df2 = mountains[mountains['Metres']>MAX_ALTITUDE]
    # assert df_columns(df2)['Ascension'] == 'date'
    wks = Worksheet(f'Higher than {MAX_ALTITUDE}', df2, 
                    header_color='#A1CAF1')
    report.append(wks)
    print("Number formats:")
    print(report.number_formats)


    title("Third worksheet")
    TABLE = "Cities"
    df = xldb.table(TABLE)
    wks = Worksheet(TABLE, df)
    report.append(wks)

    title("Fifth worksheet (Filter)")
    MAX_ALTITUDE = 7900
    myquery = """
    SELECT * FROM Mountains
    WHERE Metres > :MAX_ALTITUDE
    """
    df = xldb.query(myquery)
    wks = report.add_sheet(f"Higher than {MAX_ALTITUDE}", df)  


    title("Fifth worksheet (JOIN)")
    TABLE = "Mountains Full"
    ic(MAX_ALTITUDE)
    myquery = """
    SELECT main.*, 
        city.Population as [City Population],
        city.Altitude as [City Altitude],
        city.Country as [City Country] 
    FROM Mountains as main 
    LEFT JOIN Cities as city
        ON main.[Closest City] = city.city
    """
    ic(myquery)
    df = xldb.query(myquery)
    wks = report.add_sheet(TABLE, df)
    
    title("Save")
    # no autosave by default:
    report.rich_print(1)
    report.save(open_file=open_file)
    print("Saved!")

    # Read back and make sure we have the same thing in the dataframes
    # from the original and destination files (which is what matters)
    xldb = ExcelDB(second_out_file)
    print("Tables:", xldb.tables)
    mountains_full = xldb.table(TABLE)
    compare_df(mountains, mountains_full)
    
# ------------------------
# Command line
# ------------------------



@click.command()
@click.argument('in_file', default=TEST_FILENAME)
@click.argument('out_file', default=OUT_FILE)
@click.option('-e', '--extended', default=False, is_flag=True,
              help='test several worksheets')
@click.option('-d', '--debug', default=False, is_flag=True,
              help='test several worksheets')
def command_line(in_file:str, out_file:str, extended:bool=False,
         debug:bool=False):
    "Manual Test procedure"

    if not extended:
        test_short(in_file, out_file, open_file=True,
                   debug=debug)
    else:
        test_long(in_file, out_file, open_file=True,
                  debug=debug)

if __name__ == '__main__':
    command_line()
