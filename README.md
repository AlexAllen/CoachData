# CoachData

### Introduction

This script produces a spreadsheet with a summary of information about paddlesport coaches for clubs. It takes info from a range of spreadsheet reports that British Canoeing makes available to clubs, and compiles them into a single spreadsheet for ease of use.

The spreadsheet produced has space for clubs to add in any site specific training and assessment that their coaches have done, and lets them assign remits for each of the coaches. It has a sheet for each of the coaches with detailed information in it, and also produces a few summary sheets for ease of use. 

The summary sheets it currently produces are:
- Currency
  - This is based off the information BC provides, apart from the DBS column which is something that must be manually input to each coach's page
- Remits
  - This is based off the remits manually put input to each coach's page, and lists the craft each coach is allowed to operate from when working with each combination of craft and environment
- Providerships
  - This picks out a few key providerships (based on what my club has) and lists all coaches who can provide them.


### Instructions for Use 

CoachData is run from the command line, and was written for python 3.9.7 and pandas version 1.3.4. It should be placed in a folder with all the files necessary, and run from that location. It can either update an existing coaches summary it has produced, or it can generate a new one. 

It takes a single optional argument, which specifies which summary to update. If no argument is specified it will read in "Coaches Summary.xlsx" if it is present. Output is always to "Coaches Summary.xlsx".

In addition, the following files are used when updating:
- *Required:* "active_coaches.txt"
  - This is a text file containing the names of each coach to be included in the final file, one per line. These names must be in exactly the same format as used by the other files. E.g. it's no good putting "Alexander Allen" in this file if all the other files refer to "Alex Allen".
- Optional: "My Club Members with Coach Validation.xlsx"
  - This is the file produced by British Canoeing's "My Club Members with Coach Validation" report.
- Optional: "Safety Report.xlsx"
  - This is the file produced by British Canoeing's "Safety Report"
- Optional: "First Aid Training.xlsx"
  - This is the file produced by British Canoeing's "First Aid Training"
- Optional: "Safeguarding Report.xlsx"
  - This is the file produced by British Canoeing's "Safeguarding Report"
- Optional: "All Member Credentials.xlsx"
  - This is the file produced by Britsh Canoeing's "All Member Credentials"
