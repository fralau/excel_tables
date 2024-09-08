<div align="center">

# Excel-Tables: An easy-to-use Python library to convert Pandas dataframes to pretty Excel files, for the business world

</div>

---

## Why Excel-Tables?

### The problem of Pandas in a business environment

Companies rely on databases to store their data, but
only a fraction of the staff know how to query one.  

Today, their only tool of work is Excel.

Pandas is great for querying relational databases very quickly,
or importing tables in a variety of formats. It also exports dataframes to Excel very well. It's a perfect tool for business use.


However, the Excel files are very bland: no colors, no number formatting, etc.
In a business environment, one cannot present an unformatted Excel
file to a colleague; this is considered poor work.

There are minimal standards for Excel spreadsheets.

So, every time one does any data extraction,
one has to spend _manually_ five minutes or up to a quarter of an hour, to reformat the file so that it can be sent.

And if a new extraction is required, even five minutes later?
Everything has to be redone.

This is a huge amount of boring, repetitive work.

### The solution

What if it was possible to export tables into an Excel workbook,
with the assurance that

- It would look good or almost good at the first attempt, with
   numbers or dates properly displayed, with nicely formatted header columns?

- One could adjust the format of some columns without jumping into a rabbit hole?

The solution? ** Excel-Tables**. It allows you to create Excel 
reports with one or more tables in it.

It is a higher standard than the plain Excel export from Pandas.

### What Excel-Tables is not

- It is not a tool that allows maximum flexibility in the presentation
of tables.

- It is not a diagram tool.

- It was not developed with scientific applications in mind. 
  It might become slow for large datasets.

### Cautions

** What we mean by _table_ is simple: a header row, and rows,
as in the extraction from a a relational database table.**

Excel-Tables is very standard in its presentation (opinionated), to keep
things maximally simple (it looks nice, but it has no bells and whistles).

In order to use it comfortably, you should first adhere to its
basic philosophy and choices.

** As a last resort, you could still use the openpyxl library
to rework the report.**
  



## Installation


```sh
pip install excel_tables
```


## Usage

### Simple example

```python
from excel_reports import ExcelReport
report = ExcelReport('Myfile.xlsx', df=df)
```

If you specify a dataframe at this stage, the report is immediately saved.

### More elaborate report

Specifies the font (Helvetica) and emphasizes (bold) the lines
where the second column (1) is higher than 1000.

```python
from excel_reports import ExcelReport
my_file = 'Myfile.xlsx'
report = ExcelReport(my_file, 
                    font_name='Helvetica', 
                    df=df,
                    num_formats={'Rate': "#'##0.0000"},
                    emphasize=lambda x: x[1] > 1000)
report.rich_print()
report.open()
```

- `num_formats` is used to specify additional formats for columns,
  if the default are not appropriate. It is a dictionary of columns, Excel formats.
- `rich_print()` prints a simplified view of the report on the console.
- `open()` opens the file in Excel.



## Advanced Usage

### Report with several worksheets

```python
from excel_reports import ExcelReport, Worksheet

report = ExcelReport(second_out_file, 
                    font_name='Times New Roman', 
                    format_int="[>=1000]#'##0;[<1000]0",
                    format_float="[>=1000]#'##0.00;[<1000]0.00")

wks = Worksheet('Income', df1, emphasize=lambda x: x[1] > 20000,
                        num_formats={'Feet': "#'##0"})
report.append(wks)

wks = Worksheet('Expenses', df2, num_formats={'Rates': "#'##0.0000"})
report.append(wks)
report.save(open_file=True)
```

You can can also specify arguments in this way:

```python
report = ExcelReport(second_out_file)
report.font_name='Times New Roman', 
report.format_int="[>=1000]#'##0;[<1000]0",
report.format_float="[>=1000]#'##0.00;[<1000]0.00"
```

Since no dataframe is provided when the report objected is created,
no auto_save is done; you must do it explicitly, after having
appended the worksheets.

- `format_int`: a specification for all integer columns, valid for the
   whole report.
- `format_float`: a specification for all float columns, valid for the
   whole report.

For worksheets:
- `num_formats` is used to specify additional formats for columns,
  if the default are not appropriate. It is a dictionary of columns, Excel formats.`num_formats` is a dictionary of column names and Excel numeric formats.


### Reworking the file in openpyxl

In the unlikely case you wish to rework a report using openpyxl.

An ExcelReport object has an attribute `workbook`.

** Keep in mind that the openpyxl `Worksheet` class is not the same as the one in ExcelReport. **

```python
wb = report.workbook
# Get the 'Main' worksheet (as in the tab)
first_sheet = wb.worksheets['Main']

...
wb.save(myfile)
```

## Basic rules

### General format decisions
1. Worksheets have no background and no grid for the cells
  (complete white).
2. A simple grid is applied to the cells of each table.
3. The font color for the data is black.
4. All tables have autofilter.

### General Number formats
For Excel, float, int or datetimes are all numbers.

excel_tables formats number according to its own logic



Internal Format  | Python Type | System Default | Your default
---------------- | ----------- | ------- | -------
integer | int |  No decimals, comma separator for thousands | `format_int`
float | float |  2 decimals, comma separator for thousands | `format_float`
percentage | float (between 0 and 1) | % symbol, 1 decimal | `format_perc`
date | datetime (no hours and minutes) | ISO (YYY-MM-DD) | `format_date`
datetime | datetime | ISO (YYYY-MM-DD HH:MM:SS) | N/A

The console printer adopts English conventions for numbers, and ISO for dates.

### Format for columns
You can specify exceptions to the above defaults, 
for each column of a Worksheet object.

This is done with the `numformats` attribute, which is a dictionary
of columns.

```pythons
myworkbook.num_formats={'Rates': "#,##0.0000", "Quality": "0.00%"}
```

This is field is also present in the ExcelReport; it is used
when a dataframe (df) is specified.

### Worksheet headers
1. Headers are all treated in the same way, with a default background color.
2. You can specify a default header background color for the whole report,
   or a specific header color for each workbook.
3. According to the luminance of the background color, the text of each
   header will be black or white.


