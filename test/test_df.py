"Test dataframe utilities"
import numpy as np
from datetime import datetime, timedelta


import pandas as pd
from excel_tables.df_util import (df_columns, convert_ISO_dates,
        is_dates_no_time, is_iso_date)

# --------------------------------
# Generate sample
# --------------------------------
dates = [datetime.now().replace(microsecond=0) - timedelta(days=i) for i in range(7)]
ints = np.random.randint(1, 100, size=7)
floats = np.random.uniform(1.0, 100.0, size=7)
floats_0_1 = np.random.uniform(0.0, 1.0, size=7)
iso_dates = [(datetime.now().replace(microsecond=0) - timedelta(days=i)).isoformat() for i in range(7)]
names = ["Alice", "Bob", "Charlie", "David", "Eve", "Frank", "Grace"]

DF = pd.DataFrame({
        "Participants": names,
        "Recorded": dates,
        "Weekly Hours": ints,
        "Minutes per day": floats,
        "Percentage": floats_0_1,
        "Recorded (ISO)": iso_dates
    })
    


def test_predicates():
    "Tests on predicate (boolean) functions"
    df = DF.copy(deep=True)
    # add a recording date (truncating hours, minutes, and seconds)
    df['Date'] = df['Recorded'].dt.normalize()
    assert not is_dates_no_time(df, "Participants") # str
    assert not is_dates_no_time(df, "Recorded")
    assert is_dates_no_time(df, "Date")

    assert not is_iso_date(df, 'Participants')
    assert not is_iso_date(df, 'Recorded')
    assert not is_iso_date(df, 'Date')
    assert is_iso_date(df, 'Recorded (ISO)')


def test_df():
    "Tests on dataframe functions"
    print("DataFrame Functions")
    # Generate the dataframe
    
    df = DF.copy(deep=True)
    print(df)

    # check the column types:
    column_types = df_columns(df)
    assert column_types == {
    'Participants': 'str',
    'Recorded': 'datetime',
    'Weekly Hours': 'int',
    'Minutes per day': 'float',
    'Percentage': 'perc',
    'Recorded (ISO)': 'date_ISO'
    }, "Failed to recognize column types"

    # convert ISO dates to date:
    convert_ISO_dates(df)
    ISO_column = 'Recorded (ISO)'
    column_types = df_columns(df)
    print(df[ISO_column])
    assert column_types[ISO_column] == 'datetime', 'Failed to convert to datetime'

