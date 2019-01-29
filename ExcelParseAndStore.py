from tkinter import filedialog

import xlrd
import pymssql
import tkinter as tk

from Tools.scripts.treesync import raw_input

root = tk.Tk()
root.withdraw()

ps = 0
r_1 = 1
r_2 = 2
r_3 = 3
r_4 = 4

# print('We need to setup the DB connection')
# host = raw_input('Host: \t')
# username = raw_input('Username: ')
# passwrd = raw_input('Password: ')
# db_name = raw_input('Database: ')

# Get the file locations
# TODO: Modify to be able to select multiple excel files
# file_paths = filedialog.askopenfilenames()
file_paths = [('C:/Users/Nick/Downloads/Copy of Report For 01-02-2015.xlsx')
              , ('C:/Users/Nick/Downloads/Copy of Report For 01-02-2015.xlsx')
              , ('C:/Users/Nick/Downloads/Copy of Report For 01-02-2015.xlsx')
              , ('C:/Users/Nick/Downloads/Copy of Report For 01-02-2015.xlsx')]

# Iterate through each workbook and process the data
for file_path in file_paths:
    # Open the workbook we want to work with
    workbook = xlrd.open_workbook(file_path)

    # Add the five sheets we want (Production, Rack 1, Rack 2, Rack 3, Rack 4)
    sheets = []
    for i in range(5):
        sheets.append(workbook.sheet_by_index(i))

    # Assign the rows from Production sheet
    rows = [[], [], [], [], []]
    for sr in range(9, sheets[ps].nrows - 5):
        # According to how the excel sheet matches up to the db columns these are the correct cell references
        array = [sheets[ps].cell_value(sr, 18), sheets[ps].cell_value(sr, 17), sheets[ps].cell_value(sr, 5),
                 sheets[ps].cell_value(sr, 4), sheets[ps].cell_value(sr, 13), sheets[ps].cell_value(sr, 7),
                 sheets[ps].cell_value(sr, 19), sheets[ps].cell_value(sr, 14), sheets[ps].cell_value(sr, 12),
                 sheets[ps].cell_value(sr, 11), sheets[ps].cell_value(sr, 6), sheets[ps].cell_value(sr, 21),
                 sheets[ps].cell_value(sr, 16), sheets[ps].cell_value(sr, 3), sheets[ps].cell_value(sr, 20),
                 sheets[ps].cell_value(sr, 10), sheets[ps].cell_value(sr, 8), sheets[ps].cell_value(sr, 9),
                 sheets[ps].cell_value(sr, 2), sheets[ps].cell_value(sr, 15)]
        rows[ps].append(tuple(array))

    # Assign the rows from Racks 1-4 sheets
    for rr in range(1, 5):
        for sr in range(9, sheets[r_1].nrows - 5):
            # According to how the excel sheet matches up to the db columns these are the correct cell references
            array = [sheets[rr].cell_value(sr, 10), sheets[rr].cell_value(sr, 5), sheets[rr].cell_value(sr, 6),
                     sheets[rr].cell_value(sr, 12), sheets[rr].cell_value(sr, 9), sheets[rr].cell_value(sr, 7),
                     sheets[rr].cell_value(sr, 19), sheets[rr].cell_value(sr, 16), sheets[rr].cell_value(sr, 18),
                     sheets[rr].cell_value(sr, 3), sheets[rr].cell_value(sr, 15), sheets[rr].cell_value(sr, 17),
                     sheets[rr].cell_value(sr, 4), sheets[rr].cell_value(sr, 13), sheets[rr].cell_value(sr, 14),
                     sheets[rr].cell_value(sr, 2), sheets[rr].cell_value(sr, 8), sheets[rr].cell_value(sr, 11)]
            rows[rr].append(tuple(array))

    # Create a connection to the database
    connection = pymssql.connect("SKM-NICK", 'sa', 'c2k', 'Parker')
    cursor = connection.cursor()

    # Assemble the required query for each table
    qj_production_query = "INSERT INTO QJ_Production VALUES (%s, %s, %s, %s, %s, %s" \
                          ", %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
    r_1_query = "INSERT INTO rack1_Data VALUES (%s, %s, %s, %s, %s, %s" \
                ", %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
    r_2_query = "INSERT INTO rack2_Data VALUES (%s, %s, %s, %s, %s, %s" \
                ", %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
    r_3_query = "INSERT INTO rack3_Data VALUES (%s, %s, %s, %s, %s, %s" \
                ", %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
    r_4_query= "INSERT INTO rack4_Data VALUES (%s, %s, %s, %s, %s, %s" \
               ", %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"

    # Add the rows to be moved to the database
    cursor.executemany(qj_production_query, rows[ps])
    cursor.executemany(r_1_query, rows[r_1])
    cursor.executemany(r_2_query, rows[r_2])
    cursor.executemany(r_3_query, rows[r_3])
    cursor.executemany(r_4_query, rows[r_4])

    # Commit the changes
    connection.commit()
