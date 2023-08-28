import tkinter as tk
import pandas as pd
import pyodbc
from tkinter import Toplevel, filedialog
from configparser import ConfigParser
# Import the required libraries
import tkinter as tk
from tkinter import Toplevel
from configparser import ConfigParser


config_window = None  # Declare config_window as a global variable

# Function to save SQL Server details to config.ini file
def save_config():
    global config_window  # Declare config_window as global
    config = ConfigParser()
    config['DATABASE'] = {
        'server': server_entry.get(),
        'database': database_entry.get(),
        'username': username_entry.get(),
        'password': password_entry.get()
    }
    with open('config.ini', 'w') as configfile:
        config.write(configfile)
        
    status_label.config(text='Configuration saved successfully!', fg='green')
    
    if config_window is not None:
        config_window.destroy() 


# Function to generate the report
def generate_report():
    try:
        # Read SQL Server details from config.ini file
        config = ConfigParser()
        config.read('config.ini')
        server = config.get('DATABASE', 'server')
        database = config.get('DATABASE', 'database')
        username = config.get('DATABASE', 'username')
        password = config.get('DATABASE', 'password')

        # Create a connection string based on the authentication method
        if username and password:
            connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        else:
            connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes"

        # Establish a connection to the SQL Server
        connection = pyodbc.connect(connection_string)

        # Execute the query and fetch the results
        query = """
            SELECT
    C.Code AS [Customer ID],
    C.FirstName AS [Customer First Name],
    C.LastName AS [Customer Last Name],
	LI.Branch AS [Branch],
    LI.LaybyNo AS [Layby ID],
    CAST(C.Phone1 AS VARCHAR(50)) AS [Mobile Number],
    LI.Logged AS [Date of Transaction],
    I.Description AS [Item Description],
    LI.TotalAfterTax AS [Price Sold],
    (SELECT SUM(LI2.TotalAfterTax) FROM [dbo].[LaybyInfo] LI2 WHERE LI2.Customer = C.Code AND LI2.LaybyNo = LI.LaybyNo AND LI2.TransType = 'L') AS [Total Laybuy],
    (SELECT ISNULL(SUM(LI3.TotalAfterTax), 0) FROM [dbo].[LaybyInfo] LI3 WHERE LI3.Customer = C.Code AND LI3.LaybyNo = LI.LaybyNo AND LI3.TransType = 'B') AS [Total Payments Made],
    (SELECT SUM(LI2.TotalAfterTax) FROM [dbo].[LaybyInfo] LI2 WHERE LI2.Customer = C.Code AND LI2.LaybyNo = LI.LaybyNo AND LI2.TransType = 'L') - 
        (SELECT ISNULL(SUM(LI3.TotalAfterTax), 0) FROM [dbo].[LaybyInfo] LI3 WHERE LI3.Customer = C.Code AND LI3.LaybyNo = LI.LaybyNo AND LI3.TransType = 'B') AS [Outstanding Amount],
    (SELECT COUNT(DISTINCT LI2.Logged) FROM [dbo].[LaybyInfo] LI2 WHERE LI2.Customer = C.Code AND LI2.LaybyNo = LI.LaybyNo AND LI2.TransType = 'B') AS [Number of Payments Made],
    (SELECT MAX(LI2.Logged) FROM [dbo].[LaybyInfo] LI2 WHERE LI2.Customer = C.Code AND LI2.LaybyNo = LI.LaybyNo AND LI2.TransType = 'B') AS [Last Payment Date]
FROM
    [dbo].[LaybyInfo] LI
    INNER JOIN [dbo].[Customers] C ON LI.Customer = C.Code
    INNER JOIN [dbo].[Items] I ON LI.UPC = I.UPC
GROUP BY
    C.Code,
    LI.Branch,
    C.FirstName,
    C.LastName,
    LI.LaybyNo,
    C.Phone1,
    LI.Logged,
    I.Description,
    LI.TotalAfterTax
HAVING
    (SELECT SUM(LI2.TotalAfterTax) FROM [dbo].[LaybyInfo] LI2 WHERE LI2.Customer = C.Code AND LI2.LaybyNo = LI.LaybyNo AND LI2.TransType = 'L') - 
        (SELECT ISNULL(SUM(LI3.TotalAfterTax), 0) FROM [dbo].[LaybyInfo] LI3 WHERE LI3.Customer = C.Code AND LI3.LaybyNo = LI.LaybyNo AND LI3.TransType = 'B') > 0
        """
        df = pd.read_sql_query(query, connection)

        df['Customer ID'] = "'" + df['Customer ID'] + "'"

        # df['Date of Transaction'] = "'" + df['Date of Transaction'].astype(str) + "'"
        # df['Last Payment Date'] = "'" + df['Last Payment Date'].astype(str) + "'"


        # Format 'Date of Transaction' and 'Last Payment Date' columns
        df['Date of Transaction'] = df['Date of Transaction'].dt.strftime('%d/%m/%Y %I:%M:%S %p')
        df['Last Payment Date'] = df['Last Payment Date'].dt.strftime('%d/%m/%Y %I:%M:%S %p')


        # Prepend a single quote to the numeric mobile numbers
        df['Mobile Number'] = df['Mobile Number'].apply(lambda x: f"'{x}'" if str(x).isnumeric() else x)





        # Open a file dialog to select the output file path
        file_path = filedialog.asksaveasfilename(defaultextension='.csv')
        if file_path:
            # Save the DataFrame to a CSV file
            df.to_csv(file_path, index=False)
            status_label.config(text='Report generated and saved successfully!', fg='green')
        else:
            status_label.config(text='Report generation cancelled.', fg='red')

        # Close the database connection
        connection.close()
    except Exception as e:
        status_label.config(text=f'Error: {str(e)}', fg='red')


def open_config_window():
    global server_entry, database_entry, username_entry, password_entry, config_window

    config_window = Toplevel(window)
    config_window.title("Update Configuration")
    
    title_label = tk.Label(config_window, text='Update Configuration', font=('Helvetica', 14, 'bold'))
    title_label.pack(pady=10)
    
    # Server Label and Entry
    server_label = tk.Label(config_window, text='Server:', font=('Helvetica', 12))
    server_label.pack()
    server_entry = tk.Entry(config_window, font=('Helvetica', 12))
    server_entry.pack(pady=5)
    
    # Database Label and Entry
    database_label = tk.Label(config_window, text='Database:', font=('Helvetica', 12))
    database_label.pack()
    database_entry = tk.Entry(config_window, font=('Helvetica', 12))
    database_entry.pack(pady=5)
    
    # Username Label and Entry
    username_label = tk.Label(config_window, text='Username (leave blank for Windows Authentication):', font=('Helvetica', 12))
    username_label.pack()
    username_entry = tk.Entry(config_window, font=('Helvetica', 12))
    username_entry.pack(pady=5)
    
    # Password Label and Entry
    password_label = tk.Label(config_window, text='Password (leave blank for Windows Authentication):', font=('Helvetica', 12))
    password_label.pack()
    password_entry = tk.Entry(config_window, show='*', font=('Helvetica', 12))
    password_entry.pack(pady=5)
    
    # Save Config Button
    save_config_button = tk.Button(config_window, text='Save Config', command=save_config, font=('Helvetica', 12))
    save_config_button.pack(pady=15)
    
    config_window.geometry('400x350')  # Adjust window size
    
    config_window.protocol("WM_DELETE_WINDOW", lambda: config_window.destroy())  # Handle window close
    
    config_window.transient(window)  # Set as transient to the main window
    config_window.grab_set()  # Prevent interaction with main window while config window is open

    
    


# Create the main application window
window = tk.Tk()
window.title('Layby Info Report')
window.geometry('400x300')

# Add a title label
title_label = tk.Label(window, text='Layby Info Report', font=('Helvetica', 18, 'bold'))
title_label.pack(pady=10)

# Generate Report Button
generate_report_button = tk.Button(window, text='Generate Report', command=generate_report, font=('Helvetica', 12))
generate_report_button.pack(pady=15)

# Update Configuration Button
update_config_button = tk.Button(window, text='Update Configuration', command=open_config_window, font=('Helvetica', 12))
update_config_button.pack(pady=5)

# Status Label
status_label = tk.Label(window, text='', font=('Helvetica', 10))
status_label.pack()

# # Footer Label
# footer_label = tk.Label(window, text='© 2023 Layby Corporation', font=('Helvetica', 8))
# footer_label.pack(pady=10)

# Start the GUI event loop
window.mainloop()
