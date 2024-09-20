#!/usr/bin/env python3
"Test the program"

import os


import click
import pandas as pd
from rich import print
from rich.panel import Panel
from icecream import ic

from excel_tables import ExcelReport, Worksheet, df_columns, ExcelDB 
from excel_tables.df_util import is_iso_date


TEST_FILENAME = "mountains.xlsx"
OUT_FILE = "output.xlsx"

def add_suffix(filename:str, suffix:str):
    "Add a suffix to the basename of a file"
    fn, extension = os.path.splitext(filename)
    return ''.join((fn, suffix, extension))

def title(s: str):
    "Print a title"
    print(Panel(f"[green]{s}", expand=False))


# ------------------------
# Prepare rich output (terminal)
# ------------------------



@click.command()
@click.argument('in_file', default=TEST_FILENAME)
@click.argument('out_file', default=OUT_FILE)
@click.option('-e', '--extended', default=False, is_flag=True,
              help='test several worksheets')
@click.option('-d', '--debug', default=False, is_flag=True,
              help='test several worksheets')
def test(in_file:str, out_file:str, extended:bool=False,
         debug:bool=False):
    "Test procedure"


    # ------------------------
    # Prepare Excel
    # ------------------------
    if not extended:
        title("Short version")
        xl = pd.ExcelFile(in_file)
        df = xl.parse(0)
        assert df_columns(df)['Ascension'] == 'date'
        report = ExcelReport(out_file, font_name='Times New Roman', 
                            df=df,
                            emphasize=lambda x: x[1] > 8200,
                            debug=debug)
        report.rich_print()
        report.open()
    else:
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
        df = xldb.table(0)
        print(df_columns(df))
        assert df_columns(df)['Ascension'] == 'date'
        wks = Worksheet('Mountains', df, emphasize=lambda x: x[1] > 8500,
                        num_formats={'Feet': "#'##0"})
        print("Columns:", wks.columns)
        report.append(wks)
        print(report)

        title("Second worksheet")
        # filter where height > 8000
        MAX_ALTITUDE = 8000
        ic(MAX_ALTITUDE)
        df2 = df[df['Metres']>MAX_ALTITUDE]
        # assert df_columns(df2)['Ascension'] == 'date'
        wks = Worksheet('Higher than 8500', df2, 
                        header_color='#A1CAF1')
        report.append(wks)
        print("Number formats:")
        print(report.number_formats)


        title("Third worksheet")
        TABLE = "Cities"
        df = xldb.table(TABLE)
        wks = Worksheet(TABLE, df)
        report.append(wks)


        title("Fourth worksheet")
        # filter where height > 8000
        MAX_ALTITUDE = 7900
        TABLE = "Mountains Full"
        ic(MAX_ALTITUDE)
        myquery = f"""
        SELECT main.*, 
            city.Population as [City Population],
            city.Altitude as [City Altitude],
            city.Country as [City Country] 
        FROM Mountains as main 
        LEFT JOIN Cities as city
            ON main.[Closest City] = city.city
        WHERE Metres > :MAX_ALTITUDE
        """
        df = xldb.query(myquery)
        wks = Worksheet(TABLE, df)
        report.append(wks)

        title("Save")
        # no autosave by default:
        report.rich_print(1)
        report.save(open_file=True)
        print("Saved!")
        report.open()

if __name__ == '__main__':
    test()
