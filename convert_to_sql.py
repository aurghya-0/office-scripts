import pandas as pd
import mysql.connector
from mysql.connector import errorcode

# Function to create a connection to the MySQL database
def create_connection(host_name, user_name, user_password, db_name):
    connection = None
    try:
        connection = mysql.connector.connect(
            host=host_name,
            user=user_name,
            passwd=user_password,
            database=db_name
        )
        print("Connection to MySQL DB successful")
    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
            print("The credentials you provided are incorrect")
        elif err.errno == errorcode.ER_BAD_DB_ERROR:
            print("The database does not exist")
        else:
            print(err)
    return connection

# Function to execute a query
def execute_query(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        connection.commit()
        print("Query executed successfully")
    except mysql.connector.Error as err:
        print(f"Error: '{err}'")

# Function to read the Excel file and convert it to a MySQL table
def excel_to_mysql(file_path, table_name, connection):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Get the column names and types
    columns = df.columns
    columns_types = df.dtypes

    # Create table query
    create_table_query = f"CREATE TABLE IF NOT EXISTS {table_name} ("
    for column, dtype in zip(columns, columns_types):
        if dtype == "int64":
            create_table_query += f"{column} INT, "
        elif dtype == "float64":
            create_table_query += f"{column} FLOAT, "
        else:
            create_table_query += f"{column} TEXT, "
    create_table_query = create_table_query.rstrip(", ") + ");"

    # Execute the create table query
    execute_query(connection, create_table_query)

    # Insert data into the table
    for i, row in df.iterrows():
        insert_row_query = f"INSERT INTO {table_name} VALUES ("
        for value in row:
            if isinstance(value, str):
                insert_row_query += f"'{value}', "
            else:
                insert_row_query += f"{value}, "
        insert_row_query = insert_row_query.rstrip(", ") + ");"
        execute_query(connection, insert_row_query)

# Define your MySQL database connection details
host_name = "your_host"
user_name = "your_username"
user_password = "your_password"
db_name = "your_database"

# Define the path to your Excel file and the desired table name
file_path = "path_to_your_excel_file.xlsx"
table_name = "your_table_name"

# Create the connection
connection = create_connection(host_name, user_name, user_password, db_name)

# Convert the Excel file to a MySQL table
excel_to_mysql(file_path, table_name, connection)

# Close the connection
if connection.is_connected():
    connection.close()
    print("MySQL connection is closed")
