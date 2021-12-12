## What is `excel2gtfs`?

The industry-standard for public transit information is that of GTFS (the "General Transit Feed Specification").

While many large and medium-sized operations use fancy planning systems with integrated GTFS exports, there are countless smaller operations that still rely on Microsoft Excel as the basis for their planning.

With GTFS having been designed for easy export from relational databases, the multitude of fields and requirements can often make it difficult for smaller operations to manage GTFS in-house.

This tool aims to address this challenge by providing smaller transportation agencies with the ability to control their own GTFS, thereby reducing the latency in updating schedules and increasing the feasibility of updating GTFS data for regular schedule changes (e.g. trackwork, planned detours, etc.).

## How does this work?

The primary means of interacting with the tool is via a standardized template file in Microsoft Excel. The provided `excel2gtfsTemplate.xlsm` serves as a template, using the Delaware River Port Authority's PATCO operation as an example.

Once you've finalized your excel spreadsheet, use the `excel2gtfs.py` python script to convert the spreadsheet data to a zipped GTFS package.

### Installation

[Download](https://python.org/) and install Python3 on your computer. This software was built with version 3.9.5.

After installation, use `pip` to install the `openpyxl` library by executing the following command on the command line: `py -m pip install openpyxl`. If your computer has a different Python keyword (e.g. typically `python3` on MacOS), use that in place of the preceding `py` e.g. `python3 -m pip install openpyxl`.

You will also need Microsoft Excel or a compatible editor to modify the `.xlsm` file.

### Usage

Command Line Usage: `python3 excel2gtfs.py <input_filepath>`

Example: `python3 excel2gtfs.py excel2gtfsTemplate.xlsm`


## What should my workflow be?

The first time you use this template, be sure to configure the above configuration tabs, specifically the Agency, Feed Info, Routes, and Stops files. If you don't have GPS coordinates for your system, feel free to discard the shape data, although you should work to include this as a best practice. Similarly, if your fare structure is too cumbersome to model, you may similarly omit this data. However, it is worth investing the time to accurately model your service, particularly as this only needs to be modified when the applicable elements of your operation change.

Once you've established your configuration, your focus will be the schedule tabs. Be sure to update the schedule tabs to appropriately reflect your service, then update the calendar to identify what's in effect and when.

## Components of the Spreadsheet Template

### Configuration Tabs

Within the excel spreadsheet, there are a number of sheets/tabs used to configure data *about* the transit network, without describing the schedule itself. All but one corresponds directly to files in the [GTFS standard](https://github.com/google/transit/blob/master/gtfs/spec/en/reference.md). These include:

- **Settings & Checks**: A summary view of tabs in the spreadsheet and various checks.
- **Agency**: Details about the agency operating each route.
- **Feed Info**: Information about the GTFS feed and the link from which updates can be downloaded.
- **Routes**: A list of all routes and the agency operating each route.
- **Stops**: A list of all stops in the transit system and details about each stop.
- **Shapes**: A sequential list of GPS coordinates to create a pathway for a trip on a map.
- **Fare Rules**: A list of fare zone combinations and the fare applied to each zone.
- **Fare Attributes**: A list of specific fares available on the system.
- **Calendar**: See [Configuring the Calendar](#calendar) below after defining schedules.
- **Calendar Dates**: See [Configuring the Calendar](#calendar) below after defining schedules.

All other tabs in the spreadsheet are assumed to be those of schedules, and are expected to be formatted as such (detailed below).

### Schedule Tabs

Schedule spreadsheets consist of a standardized format that should be duplicated from an existing sheet in the file. The name of the tab represents the given `service_id`, which is needed for defining trips in the calendar.

The first two rows in the spreadsheet represent the schedule's header. **Do NOT modify the values in row 1 (displayed diagonally); if needed, only modify the values in row 2 with the orange background.**

There are two types of columns in the spreadsheet: trip-information columns and stop-time values. If a field is not detected as a trip-information column, it is considered to be a stop-time column.

The following values represent the supported trip-information columns:

- `route_id`: The route being operated by the row's trip.
- `direction_id`: The direction indicator for the trip. Typically `1` represents "inbound" trips, and `0` represents "outbound" trips.
- `shape_id`: An identifier in the Shapes tab referencing the GPS path of the given trip. Leave blank if none defined.
- `headsign`: The "headsign" of the trip, effectively what would be displayed on the front of the trip's vehicle.
- `wheelchair_accessible`: `1` for a trip operating with an accessible vehicle, and `2` for an inaccessible vehicle. (NOTE: This relates only to the vehicle; wheelchair-accessible *stops* are indicated as such in the stops tab.)
- `bikes_allowed`: `1` for a trip that permits bikes, and `2` for one that does not.

All other columns represent a stop, with the row values indicating the time at which the vehicle serves that stop.

**IMPORTANT**: The orange editable values in Row 2 reflect the stop_id values defined in the stops tab. The diagonal value above indicates the stop_name associated with that stop_id where a match is found. If a match isn't found, the export will NOT be successful!

#### Headway-Service

There are also two additional optional values, which make defining headway service a breeze. Use "Then Every" to define the headway on which trips will operate, then "Until" to list when that headway ends. Note that trips will operate inclusive of the "Until" value, meaning a 10-minute headway for a trip that began at 8:00 will operate on the same schedule with initial departures at 8:00, 8:10, 8:20, and 8:30 if the "Until" value were set to 8:30, 8:35, or 8:39:59; setting the "Until" value to 8:40 would include an 8:40 departure.

**Warning:** Be careful to avoid include duplicate trips in your feed! If trips operate every 30-minutes until 2pm and then every 15-minutes thereafter, be sure to set the first headway's Until value to something before 2pm ***OR*** begin your 15-minute headway at 2:15pm. Otherwise, you'll have two simultaneous departures in the system.



### Configuring the Calendar {#calendar}

Once the trips have been defined in the applicable schedule tabs, the Calendar and Calendar Dates tabs are where the services can be scheduled.

The calendar tab forms the basis for *typical* service. In the calendar, list the sheet name of a given schedule as the schedule's `service_id`, then indicate if that schedule typically operates on each day of the week using `1` for yes and `0` for no, then indicate the start and end dates of the given schedule. The start/end dates are inclusive (e.g. service WILL operate on the end date listed, provided it matches the given day of the week).

The calendar dates tab forms the basis for *overriding* normal service. This tab contains a table with three very important values. Each row contains a given service_id (the tab name of the applicable schedule), a given date, and an `exception_type` command. There are three types of `exception_type` commands supported in excel2gtfs:

- `1`: Operate the given service_id's schedule on the date listed, in addition to all other regularly-scheduled service.
- `2`: Do NOT operate the given service_id's schedule on the date listed, regardless of whether or not it would otherwise be regularly scheduled to operate.
- `3`: Operate the given service_id's schedule on the date listed in addition to any other schedules defined with `exception_type=3` on the given date, but cancel all other regularly-scheduled service.

Note: If you inspect the output GTFS files, you won't see any `exception_type=3` values. The excel2gtfs converter automatically adjusts these exceptions to the appropriate set of type `1` and `2` exceptions for full GTFS support.


# Contributions

This tool was developed by [Jeff Kessler](http://jeffkess.com) of KASDAT Consulting in the hopes smaller agencies can better manage their own GTFS data and stay up-to-date with schedule changes and planned service adjustments.

# License

This tool is provided on an as-is basis with neither an explicit nor implied warranty of fitness for any purpose. Users accept any and all risks from use hereof, and agree to disclaim and indemnify all authors and contributors from any and all claims arising from use of this tool and/or any of its components. Public Transit Agencies under FRA or FTA jurisdiction are permitted to use this software in the furtherance of their operations, as are academics and hobbyists who may use it for noncommercial purposes. Individuals wishing to contribute to the development of this tool are invited to do so, provided the derivatives are similarly shared with equivalent license terms. The author reserves the right to modify and/or revoke this license.