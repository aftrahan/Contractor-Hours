This spreadsheet template was designed by me in several stages to automate and streamline a fully manual data cleaning process I inherited with my Clerk III position. It was something only I would have to use as the only part relevant to other departments and colleagues was the end result (the cleaned data) and therefore no documentation was written, beyond basic commenting of the code. There is also minimal error correction built in, for the same reason.

This readme will provide a basic tour through the spreadsheet by following the standard workflow in broad strokes and elaborating on the purpose of its different code and functions as they come up.

# The Data Cleaning Task

Our maintenance tracking software (TRAX, an SQL-based CMMS) allows for tracking of labour hours using a basic sign-in/sign-out system. However are were numerous opportunities for user error in this system, eg:

- workers who forgot to log out at the end of a shift
- workers who worked on multiple task cards but only logged into one
- workers who failed to log in at all
- workers who remained logged in to task cards for half- or full-hour lunch breaks

In addition, there were certain system errors that the database was prone to depending on its version status, such as workers who logged into multiple task cards at once having their total time attributed to each card in full, rather than being split evenly between all cards.

In order to allow for more accurate accounting, my predecessor had begun an entirely manual, very basic data cleaning process in which she would run labour reports in TRAX, copy & paste them into an Excel file, and visually check for bad data before correcting it in TRAX. When I inherited this task, I quickly became dissatisfied with the inefficiency of this method, especially due to a few issues specific to our data:

- There was no good way to quickly compare a worker's total hours with their worked shift;
- Compounding this, all airline operations are recorded in UTC, so that not just night shifts but afternoon shifts would frequently cross the date-line and make it unclear how to group a worker's total hours for comparison.

It was at this point that I began learning advanced Excel formulas and VBA in order to gradually improve this process and automate what parts could be automated.

## Limitations of the spreadsheet

I was constantly tweaking this sheet throughout my tenure in Maintenance Administration, and development ceased rather abruptly on my return to Technical Records. As such there are a number of functions that I considered adding but did not get to, or would have liked to implement but did not have the option given the data environment I was working in. As I was considered a clerk, not an analyst, and these improvements had all been on my own initiative, I was limited in what I could actually accomplish.

**Automatic Evaluation**. The spreadsheet contains several tools for highlighting and locating bad data but because of the way it was set up (as will be covered below) still required an employee to scroll through and pick out the items that needed fixing in TRAX. The next stage of development would have been setting up methods for automatically evaluating the bulk of the data and labelling for the user the items or days to be addressed.

**Automatic Updating**. Ideally the next step of any data cleaning system is to be able to automatically update the database instead of having to manually edit it. I have no doubt that a Python script would have been able to accomplish this; however, this would have required a level of access and permissions that I did not have and would probably have had a difficult time obtaining.

**Automatic DST**. Converting from UTC to EST is done via an offset value in the time calc columns (str/fin). To account for daylight savings time I would manually change this value. The closest to an automatic detection I ever got was for pay periods that covered the time shift, where I added an extra column that determined whether the offset should be 5 or 4 based on a hard-coded date range, but I had not found an efficient way to "detect" when daylight savings time started or ended by the time I left the department.

# Importing Data
The spreadsheet is set up to import data from TRAX labour reports of a specific format. These are produced by a custom SQL query I built and tweaked as I implemented new functions. Importing this data in an easily usable fashion requires the use of the MASTER LISTS sheet which must be located in the same folder as the TEMPLATE file for the spreadsheet to reference it and run.

A separate spreadsheet is used to reconcile each pay period. The start of the applicable pay period (always a Saturday) must be entered on the first sheet, labeled "HOUR TRACKING". November 21, 2020 has been entered into this example sheet already; any Saturday will work but the sheet and the data provided is set up to calculate for this specific pay period.

On pressing "Import WOs", the sheet will ask for a folder in which to find the labour reports ("WO Hours" in the provided files) and proceed to run through the files in that folder one-by-one.

The MASTER LISTS spreadsheet is used to load additional information about contractors/employees, as well as details about the work orders being referenced. The WO information is relevant to hour processing; the contractor/employee information is not. Therefore the spreadsheet will simply add any new contractors/employees it finds to MASTER LISTS and leave the extra fields blank, but prompt the user to provide the missing information on new WOs. This will not be encountered in a test run with the provided files as MASTER LISTS has already been updated on the example WOs.

>**WO Types**
>
>LINE - routine work done on active aircraft between flights
>SHOP - work done on airplane components, not the aircraft itself
>HMV - work done on aircraft during dedicated maintenance stays
>
>Only HMV work needed to be accounted for and updated; LINE and SHOP work is included to clarify gaps in the data and is automatically highlighted as not needing to be updated in TRAX.

The import code (comprised of several different functions - see ContractHours module or ContractHours.vba) checks that rows are within the correct date range, sorts rows between the Contractor and Employee sheets (employees have numerical IDs, contractors have alphanumeric IDs), and adds info to the Details sheet. It also checks if a file has been imported before, and only reads lines that were added after the last import; this is to account for later manual entry of labour times.

## Date and time columns
The start and end times for a given labour entry (columns str & fin) are calculated using a formula with an offset to represent the difference between UTC and EST/DST. The date column then refers to the converted times to determine if a given entry "actually" occurred on the previous day, and corrects itself accordingly. Any date entry that is different from the hidden transaction_date column (the raw data from the report itself) will automatically highlight in yellow to clarify that an edit occurred.

# Shift Entry
Obvious errors (task cards logged into for multiple days, hours calculated incorrectly, etc) can be identified just through these steps, but most errors are harder to spot without directly comparing the TRAX labour data to the shifts each worker was actually working. The Hours sheet allows for quick review of this information and how it lines up with the data in the Contractor and Employee sheets.

## Contractor shifts
Contractor hours are recorded using a time-stamp system on a physical time card, so unfortunately the only option is to manually enter their shifts. This system is also the reason for the "rough time" and "approved time" columns - though it only came up occasionally, sometimes there would be a mis-match between the time a contractor punched in for and the time signed off on by their supervisor. "Rough time" calculates based on the start and end times transcribed off the cart (columns str & fn), while approved time is manually entered based on the hours total provided by the supervisor.

In order to more easily showcase the functions of this sheet, it has been pre-filled with the contractor info for the pay period starting November 21, 2020.

## Employee shifts
Employees clock in and out using a digital system, so their shifts can be imported using reports from our HR software. The button "Input employee hours" on the first sheet (HOUR TRACKING) will automatically import into the Hours sheet. (To test, use file "Employee Hours Nov 21.xlsx" in the main folder.)

# Processing Times
Much of reconciling the data involves processes outside the spreadsheet itself - checking what was logged on contractors' time sheets, reviewing labour reports from TRAX to find additional WOs needing import, re-loading WOs to incorporate manual labour entry after the fact, and other forms of investigation. What follows is a brief overview of the tools within the spreadsheet that assist with all of this.

## Contractor & Employee sheets
When data is imported, it is in semi-random order based on the order of file import. I wrote two custom sort functions with keyboard shortcuts to help with the two main stages of data processing - chronologically by contractor (Ctrl+Shift+N) for editing contractor data, and by work order and task card (Ctrl+Shift+W) for updating TRAX data.

Labour hours are imported twice; once (hr & mn) to represent the original TRAX data and one (hr2 & mn2) to allow for user edits. The second set is what the majority of the spreadsheet functions run off of. By default, since the two data sets match on import, the "real time" column is highlighted in green via conditional formatting; if the second set changes, the "real time" column will un-highlight until a checkmark ('a') is placed in the TRAX column, indicating that the time has been updated in TRAX. (This does not occur in entries from LINE or SHOP WOs; their times must still be edited to reconcile with the Hours sheet - see below - but they do not need to be updated in TRAX so they will stay green no matter what. This is accomplished using conditional formatting with a VLOOKUP function that references the Details sheet to check WO status.)

### Code column
Certain processing tasks are accomplished using the "code" column. Entering a letter in here will do one of four things:
- **a**: highlights a row in yellow to signify that the work order or task card details are a(n educated) guess (only used with added lines, not imported data)
- **e**: highlights a row in blue to signify "extra time", TRAX hours that cannot be reconciled with the worker's official hours (usually occurring due to manual labour entry)
- **m**: highlights a row in red to indicate "missing time", time that a worker was supposed to be working but no labour info can be found (only used with added lines, not imported data)
- **x**: used when a work shift has crossed the date line, to sort the labour entry into the "previous" day

## Hours sheet
The Hours sheet is the main tool for reconciling data, as it will immediately indicate whether a worker's labour data matches the labour hours they reported. The "TRAX time" column uses an INDIRECT function to sum up all hours attributed to a given worker over a given day, using the second set of data in order to respond to user edits. If the times match their "approved time", the cell is clear; if not, it is highlighted as an error. Days with no attributed hours are simply blank.

The three following columns allow for finer clarification of hour entries. "Assigned time" is time associated with a specific work order and task card. "Missing time" is time unaccounted for; if a worker's hours have been fully reviewed and still fall under their official time count, an extra line can be added in the appropriate sheet for the appropriate date, with the time value of the missing hours and an 'm' in the Code column (see above). Similarly, "extra time" is time that goes over their accounted time; this happens when a worker manually enters labour time over and above their logged-in time, which can be marked with an 'x' in the code column. 

The "TRAX time" column's conditional formatting is responsive to these extra columns. If a day is reconciled except for an extra time entry, the cell will be highlighted in blue rather than in error colours. A day that has been reconciled using missing time entries will appear as normal, as it has undergone a full review and the missing time column itself serves to indicate the gap.
