""""------------------------------------------------------------------
Excel2GTFS v0.0.3
(c) Jeff Kessler, 2021-12-12-1050

0.0.1  Initial Commit
0.0.2  Schedule data processing
0.0.3  Support for calendar dates and overrides
------------------------------------------------------------------"""

import openpyxl
import csv
import os
import datetime

wb = openpyxl.load_workbook(filename="ExcelGTFS.xlsx", data_only=True)

# Identify applicable sheets
config_sheets = {"Agency", "Routes", "Stops", "Fare Rules", "Fares", "Shapes", "Calendar", "Calendar Overrides"}
services = set(wb.sheetnames) - config_sheets
config_sheets = config_sheets.intersection(wb.sheetnames)

# Create Output Directory
fp = "Excel2GTFS Output Created " + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")
os.makedirs(fp)


# ------------------------------------------------------------------------------
# Process and Write Configuration Sheets
# ------------------------------------------------------------------------------

for sheet_name in config_sheets:

    # Parse Dates in GTFS Format
    data = wb[sheet_name].values
    data = [[(val.strftime("%Y%m%d") if type(val)==datetime.datetime else val) for val in row] for row in data]

    # Process Calendar Overrides
    if sheet_name == "Calendar Overrides" and data[1:]:

        override_dates = {}
        calendar_dates = []

        # Process Override Entries and Extract Type 3s
        for row in data[1:]:

            # Convert to List of Dicts
            row = {data[0][index]: val for index, val in enumerate(row)}

            # Extract Type 3s or Add Generic Calendar Dates
            if str(row["exception_type"])=="3":
                override_dates[row["date"]].append(row["service_id"]) if row["date"] in override_dates else override_dates.update({row["date"]: [row["service_id"]]})
            else:
                calendar_dates.append(row)

        # Covert Override Entries to Type 1/2s
        for date, svcs in override_dates.items():
            [calendar_dates.append({"service_id": svc, "date": date, "exception_type": ("1" if svc in svcs else "2")}) for svc in services]

        # Covert Calendar Dates back to List vs Dict
        data = [list(calendar_dates[0])] + [[row[key] for key in list(calendar_dates[0])] for row in calendar_dates[1:]]


    # Save GTFS Configuration File
    with open(f'{fp}/{sheet_name.lower().replace(" ", "_")}.csv', "w") as file:
        writer = csv.writer(file)
        writer.writerows(data)


# ------------------------------------------------------------------------------
# Process Schedule Data
# ------------------------------------------------------------------------------

# Initialize Services
special_keys = ["route_id", "direction_id", "shape_id", "headsign", "wheelchair_accessible", "bikes_allowed", "Then Every", "Until"]
gtfs_entries = {"trips": [], "stop_times": [], "frequencies": []}

# Process Schedule Sheets
for service in services:

    svc_trips = list(wb[service].values)
    svc_trip_dicts = [{svc_trips[1][index]: val for index, val in enumerate(row)} for row in svc_trips[2:] if any(row)]

    for trip in svc_trip_dicts:

        # Identify Stops and Define trip_id by Origin and Departure Time
        trip_stop_times = sorted([(key, val.strftime("%H:%M:%S")) for key, val in trip.items() if key not in special_keys and val], key=lambda x: x[-1])
        trip_id = "-".join(str(item) for item in [service, *trip_stop_times[0]])

        # Append trips.txt Entries
        gtfs_entries["trips"].append({
            "trip_id": trip_id,
            "route_id": trip.get("route_id"),
            "direction_id": trip.get("direction_id"),
            "shape_id": trip.get("shape_id", ""),
            "trip_headsign": trip.get("headsign", ""),
            "wheelchair_accessible": trip.get("wheelchair_accessible", ""),
            "bikes_allowed": trip.get("bikes_allowed", ""),
        })

        # Append stop_time.txt Entries
        [gtfs_entries["stop_times"].append({
            "trip_id": trip_id,
            "arrival_time": val[1],
            "departure_time": val[1],
            "stop_id": str(val[0]),
            "stop_sequence": index,
            "pickup_type": "1" if index==len(trip_stop_times)-1 else "0",
            "drop_off_type": "1" if index==0 else "0"
        }) for index, val in enumerate(trip_stop_times)]

        # Append frequencies.txt Entries
        if trip.get("Then Every") and trip.get("Until"):
            gtfs_entries["frequencies"].append({
                "trip_id": trip_id,
                "start_time": trip_stop_times[0][-1],
                "end_time": trip["Until"].strftime("%H:%M:%S"),
                "headway_secs": (trip["Then Every"].hour*60*60 + trip["Then Every"].minute*60 + trip["Then Every"].second),
                "exact_times": "1"
            })

# ------------------------------------------------------------------------------
# Write Schedule Data
# ------------------------------------------------------------------------------

for key in gtfs_entries:
    if gtfs_entries[key]:
        with open(f'{fp}/{key}.csv', "w") as file:
            writer = csv.DictWriter(file, fieldnames=list(gtfs_entries[key][0]))
            writer.writeheader()
            writer.writerows(gtfs_entries[key])