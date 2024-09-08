#!/usr/bin/env python3
"""
Utilities
(C) Laurent Franceschetti (2024)
"""
import pandas as pd
from datetime import datetime


# --------------------------------
# Data frame columns
# --------------------------------
def map_dtype(dtype) -> str:
    "Convert a dtype into Python type"
    if pd.api.types.is_string_dtype(dtype):
        return "str"
    elif pd.api.types.is_datetime64_any_dtype(dtype):
        return "datetime"
    elif pd.api.types.is_integer_dtype(dtype):
        return "int"
    elif pd.api.types.is_float_dtype(dtype):
        return "float"
    else:
        return "unknown"


def dates_no_time(df:pd.DataFrame, col:str) -> bool:
    """
    Check if a dataframe's column's data are pure dates
    (no hours, minutes, seconds...).
    It is assumed that the column is already dates.
    If 
    """
    # Normalize the dates to remove hours, minutes, and seconds
    normalized_dates = df[col].dt.normalize()

    # Check if the original dates are equal to the normalized dates
    return (df[col] == normalized_dates).all()




def df_columns(df) -> dict:
    "Returns a column description, as a list of name and type (Python)"
    cols = df.columns
    types = [map_dtype(item) for item in df.dtypes]
    ref = dict(zip(cols, types))

    # further checks on the values
    for col, col_type in ref.items():
        # print(col_type)
        if col_type == 'float':
            # check if all values are between 0 and 1
            column = df[col]
            is_between_0_and_1 = column.dropna().between(0, 1).all()
            if is_between_0_and_1:
                ref[col] = 'perc'
                # print("Is percentage")
        if col_type == 'datetime':
            if dates_no_time(df, col):
                ref[col] = 'date'
                # print("Is date")

    return ref

def apply_to_column(df:pd.DataFrame, col:str, func:callable):
    """
    Apply a function to all non-null values
    in a specified column of a DataFrame (in-place).

    Parameters:
    - df: The DataFrame containing the column.
    - col: The name of the column to apply the function to.
    - func: The function to apply to the column values.
    """
    df[col] = df[col].apply(lambda x: func(x) if pd.notnull(x) else x)


