import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine.url import URL
import pypyodbc

# Path to Excel file
path = r'C:\python\export-excel-2-database\old_press_sentirings.xlsx'

# Connection string details
connection_string = (
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=;'
    'DATABASE=;'
    'UID=;'
    'PWD=;'
)

# Use SQLAlchemy's URL to create the engine
connection_url = URL.create('mssql+pyodbc', query={'odbc_connect': connection_string})
engine = create_engine(connection_url, module=pypyodbc)

# Load the Excel file
excel_data = pd.read_excel(path, sheet_name=None)

# Iterate over each sheet in the Excel file
for sheet_name, df in excel_data.items():
    table_name = sheet_name  
    
    # Dynamically create the table in the database based on the DataFrame's column names
    df.head(0).to_sql(table_name, engine, if_exists='replace', index=False)  
    
    # Insert the data into the table
    df.to_sql(table_name, engine, if_exists='append', index=False)

    print(f'Successfully loaded data from sheet {sheet_name} into table {table_name}')
