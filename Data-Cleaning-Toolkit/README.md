# Data Cleaning Toolkit (VBA)

## Problem
Messy datasets slow down analysis and require manual effort.

## Solution
This VBA toolkit automates cleaning tasks:
- Splits jumbled text into structured columns
- Trims spaces
- Standardizes dates
- Removes duplicates
- Validates fields

## Demo
- Input: [Sample_Messy_Data.xlsx](demo/Sample_Messy_Data.xlsx)
- Output: [Sample_Clean_Data.xlsx](demo/Sample_Clean_Data.xlsx)

## Code
See [mod_JumbledDataClean.bas](src/mod_JumbledDataClean.bas) for VBA implementation.

## Usage
1. Import `mod_JumbledDataClean.bas.bas` into Excel (ALT+F11 → File → Import).
2. Run `SplitJumbledDataToNewSheet`.
3. Cleaned data will appear in a new sheet with headers.

## Next Steps
Future enhancements: regex cleaning, SQL integration, Power BI export.
