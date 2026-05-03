# LinkedIn India Job Market — Data Cleaning Project (Excel)

This project demonstrates a complete end‑to‑end data cleaning workflow on a LinkedIn job dataset containing 7,927 rows and 15 columns.
The goal was to clean, standardize, and structure the dataset for further analysis of the Indian job market.

## Dataset Overview
Original rows: 7,927

Columns: 15

Final cleaned rows: 7,816

Final columns: 14+ (after splitting and transformations)

“79 duplicate values found and removed… 33 empty rows removed… now we have 7816 rows.”

## Data Cleaning Steps

### 1. Remove Duplicates & Empty Rows
Removed 79 duplicate rows using Excel’s Remove Duplicates tool.

Deleted 33 rows where only job_id was filled and all other columns were blank.

“79 duplicates values found and removed… 33 empty rows are found… now we have 7816 rows.”

### 2. Clean Job Titles
Created a separate sheet clean_job_col and applied multi‑step cleaning:

Step 1 — Remove text after commas
excel
=TRIM(LEFT(A2, FIND(",", A2&",")-1))
Step 2 — Remove hyphens (-WFH, -Remote)
excel
=TRIM(LEFT(B2, FIND("-", B2&"-")-1))
Step 3 — Remove brackets
excel
=TRIM(LEFT(C2, FIND("(", C2&"(")-1))
Step 4 — Job Mapping
Extracted unique job titles using UNIQUE()

Created a mapping table (Raw_job → Standard_job)

Used filters + Flash Fill to categorize roles

Applied final mapping using XLOOKUP:

excel
=XLOOKUP(clean_job_col!D2, job_mapping!A:A, job_mapping!B:B, "Other")

### 3. Clean Location Column
Split location into city, state, country.

City
excel
=IFERROR(TRIM(LEFT(C4, FIND(",", C4)-1)), "")
State
(Using SUBSTITUTE + FIND logic)

Country
excel
=IFERROR(
TRIM(RIGHT(C4, LEN(C4) - FIND("@", SUBSTITUTE(C4, ",", "@", LEN(C4) - LEN(SUBSTITUTE(C4, ",", ""))))))
,"India")

### 4. Clean full_time_remote Column
Used Text to Columns with delimiter "·" to separate values.

### 5. Clean no_of_employ Column
Split employee count and industry using Text to Columns.

### 6. Clean no_of_application Column
Converted text to numbers

Removed inconsistent values containing “hours”, “minutes”, etc.

### 7. Clean posted_day_ago Column
Converted all time units into hours.

excel
=IF(N5="minutes", M5/60,
IF(N5="hour", M5,
IF(N5="hours", M5,
IF(N5="day", M5*24,
IF(N5="days", M5*24,
IF(N5="week", M5*168,
IF(N5="weeks", M5*168, "")))))))

### 8. Clean alumni Column
Removed “company alumni” using Ctrl + H and converted to numeric.

### 9. Clean linkedin_followers Column
Used Text to Columns → removed inconsistent values → validated remaining numbers.

## 📁 Project Structure

Code
linkedin-job-cleaning/
│
├── raw_data/
│   └── linkdin_Job_data.csv
│
├── cleaned_data/
│   └── linkedin_cleaned.xlsx
│
├── job_mapping.xlsx
│
└── README.md

## 🛠 Tools Used
Microsoft Excel

Excel formulas (TRIM, LEFT, MID, RIGHT, SUBSTITUTE, FIND, IFERROR, XLOOKUP)

Text-to-Columns

Flash Fill

Filters & Data Validation

## ⭐ Key Skills Demonstrated

Real‑world data cleaning

Handling inconsistent text formats

Extracting structured information

Creating mapping tables

Converting messy time units

Splitting multi‑value columns

Preparing data for analysis

## 👩‍💻 Author
Arpita  
Data Analyst | Excel | SQL | Data Cleaning




A short video script explaining your cleaning process

Just tell me what you want next.
