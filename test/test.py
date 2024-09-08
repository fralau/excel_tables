"Test the program"

import os


import click
import pandas as pd

from excel_tables import ExcelReport, Worksheet, df_columns 


TEST_FILENAME = "mountains.xlsx"
OUT_FILE = "output.xlsx"

def add_suffix(filename:str, suffix:str):
    "Add a suffix to the basename of a file"
    fn, extension = os.path.splitext(filename)
    return ''.join((fn, suffix, extension))


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
    xl = pd.ExcelFile(in_file)
    df = xl.parse(0)
    assert df_columns(df)['Ascension'] == 'date'

    # ------------------------
    # Prepare Excel
    # ------------------------
    if not extended:
        report = ExcelReport(out_file, font_name='Helvetica', 
                            df=df,
                            emphasize=lambda x: x[1] > 8200,
                            debug=debug)
        report.rich_print()
        report.open()
    else:
        print("First:")
        second_out_file = add_suffix(out_file, '_mult')
        print(f"  {second_out_file}")
        report = ExcelReport(second_out_file, 
                            font_name='Times New Roman', 
                            format_int="[>=1000]#'##0;[<1000]0",
                            format_float="[>=1000]#'##0.00;[<1000]0.00",
                            format_date="DD-MM-YYYY",
                            debug=debug)
        try:
            print(report)
        except KeyError:
            # No report available yet.
            pass
        wks = Worksheet('Mountains', df, emphasize=lambda x: x[1] > 8500,
                        num_formats={'Feet': "#'##0"})
        print("Columns:", wks.columns)
        report.append(wks)
        print(report)
        assert df_columns(df)['Ascension'] == 'date'

        print("Second:")
        # filter where height > 8000
        df2 = df[df['Metres']>8000]
        assert df_columns(df2)['Ascension'] == 'date'
        wks = Worksheet('Higher than 8500', df2, 
                        header_color='#A1CAF1')
        report.append(wks)
        print("Number formats:")
        print(report.number_formats)

        print("Save:")
        # no autosave by default:
        report.rich_print(1)
        report.save(open_file=True)
        report.open()

if __name__ == '__main__':
    test()
