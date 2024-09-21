#!/usr/bin/env python3
"""
Utilities
(C) Laurent Franceschetti (2024)
"""
import numpy as np
import re
from datetime import datetime

import pandas as pd

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


def is_column_str(df:pd.DataFrame, col: str) -> bool:
    "Check if a column is an str (regardless of its content)"
    col_type = df.dtypes[col]
    return pd.api.types.is_string_dtype(col_type)

def is_column_datetime(df:pd.DataFrame, col: str) -> bool:
    "Check if a column is a datetime (not str)"
    col_type = df.dtypes[col]
    return pd.api.types.is_datetime64_any_dtype(col_type)

def is_dates_no_time(df:pd.DataFrame, col:str) -> bool:
    """
    Check if a dataframe's column's data are pure dates
    (no hours, minutes, seconds...).
    It is assumed that the column is already dates.
    """
    if not is_column_datetime(df, col):
        return False
    # Normalize the dates to remove hours, minutes, and seconds
    normalized_dates = df[col].dt.normalize()

    # Check if the original dates are equal to the normalized dates
    return (df[col] == normalized_dates).all()

def is_iso_date(df:pd.DataFrame, col:str):
    """
    Detect if a string column is composed of ISO dates or datetimes
    (YYY-MM-DD or YYYY-MM-DD HH:MM:SS or YYYY-MM-DDTHH:MM:SS), 
    ignoring null values.

    It is assumed that the column is already str.
    
    Parameters:
    column (pd.Series): The column to check.
    
    Returns:
    bool: True if the column is composed of ISO dates or datetimes,
        False otherwise.
    """
    if not is_column_str(df, col):
        return False

    column = df[col]
    # Define ISO date and datetime patterns
    # YYYY-MM-DD:
    date_pattern = re.compile(r'^\d{4}-\d{2}-\d{2}$')  
    # YYYY-MM-DD HH:MM:SS:
    datetime_pattern = re.compile(r'^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?$')  
    
    # Drop null values and empty strings
    non_empty_values = column.dropna().replace('', np.nan).dropna()
    
    # Check each value and return False immediately if necessary:
    for value in non_empty_values:
        if not (date_pattern.match(value) or datetime_pattern.match(value)):
            # print(col, value, "not a a date")
            return False
    return True

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
        elif col_type == 'datetime':
            if is_dates_no_time(df, col):
                ref[col] = 'date'
                # print("Is date")
        elif col_type == 'str':
            if is_iso_date(df, col):
                ref[col] = 'date_ISO'

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

def convert_ISO_dates(df:pd.DataFrame):
    """
    Convert in place the ISO dates of a dataframe into dates
    """
    # print("Converting dates:", df_columns(df))
    for col, col_type in df_columns(df).items():
        if col_type == 'date_ISO':
            df[col]  = pd.to_datetime(df[col], errors='coerce')


