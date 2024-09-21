"Test the ExcelDB class"


import os
from excel_tables import ExcelDB 


CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_FILENAME = os.path.join(CURRENT_DIR, "mountains.xlsx")

def test_read_file():
    "Read the test file"

    xldb = ExcelDB(TEST_FILENAME)
    assert xldb.tables == ['Mountains', 'Cities']
    no_rows = len(xldb.table('Mountains'))
    assert no_rows == 25
    assert len(xldb.query("SELECT * FROM Mountains")) == no_rows
    assert len(xldb.get_values("SELECT Name FROM Mountains")) == no_rows
    assert xldb.get_value("SELECT count(*) FROM Mountains") == no_rows

    # Simple query, with a join
    myquery = f"""
    SELECT main.*, 
        city.Population as [City Population],
        city.Altitude as [City Altitude],
        city.Country as [City Country] 
    FROM Mountains as main 
    LEFT JOIN Cities as city
        ON main.[Closest City] = city.city
    """
    df = xldb.query(myquery)
    assert len(df) == no_rows

    # Parameterized query with a join
    MAX_ALTITUDE = 7900
    myquery = f"""
    SELECT main.*, 
        city.Population as [City Population],
        city.Altitude as [City Altitude],
        city.Country as [City Country] 
    FROM Mountains as main 
    LEFT JOIN Cities as city
        ON main.[Closest City] = city.city
    -- this part contains a parameter:
    WHERE Metres > :MAX_ALTITUDE
    """
    df_join = xldb.query(myquery)
    no_rows = len(df_join)

    count_query = """
    SELECT count(*) 
    FROM Mountains
    WHERE Metres > :MAX_ALTITUDE
    """
    assert no_rows == xldb.get_value(count_query)

    # Using f-strings (not recommended but you can):
    # You have to for tables
    mytable = 'Mountains'
    count_query = f"""
    SELECT count(*) 
    FROM {mytable}
    WHERE Metres > {MAX_ALTITUDE}
    """
    assert no_rows == xldb.get_value(count_query)