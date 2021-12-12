""""------------------------------------------------------------------
Excel2GTFS v0.0.1
(c) Jeff Kessler, 2021-12-12-0745

0.0.1  Initial Commit
------------------------------------------------------------------"""

import openpyxl
import csv
import os
import time

wb = openpyxl.load_workbook(filename="ExcelGTFS.xlsx")

# Identify applicable sheets
config_sheets = {"Agency", "Routes", "Stops", "Fare Rules", "Fares", "Calendar", "Calendar Dates"}
services = set(wb.sheetnames) - config_sheets
config_sheets = config_sheets.intersection(wb.sheetnames)

# Create Output Directory
fp = "Excel2GTFS Output Created " + time.strftime("%Y-%m-%d-%H%M%S")
os.makedirs(fp)


# ------------------------------------------------------------------------------
# Process Configuration Sheets
# ------------------------------------------------------------------------------

for sheet_name in config_sheets:

    with open(f'{fp}/{sheet_name.lower().replace(" ", "_")}.csv', "w") as file:
        writer = csv.writer(file)
        writer.writerows([row for row in wb[sheet_name].values])


