# Badly Structured Sales Data 1 (VBA Cleaning)

## Problem
The dataset contains:
- Merged header rows
- Blank rows between sections
- Totals mixed with detail rows
- Inconsistent date formats

Such structure makes analysis impossible without manual cleanup.

## Solution
A VBA macro (`CleanSalesData1`) automates:
- Removing merged headers
- Deleting blank rows
- Removing totals rows
- Standardizing dates to `YYYY-MM-DD`
- Logging each step in a `CleaningLog` sheet

## Demo
- Input: [Messy Data.xls](src/Messy_Data.xlsx)
- Output: [Clean Data.xlsx](src/Clean_Data.xlsx)
- Output: [Log_File.xlsx](src/Log_File.xlsx)

## Code
See [Module_SalesData1DataCleaning.bas](demo/Module_SalesData1DataCleaning.bas) for implementation.

## Usage
1. Import `Module_SalesData1DataCleaning.bas` into Excel (`ALT+F11 → File → Import`).
2. Open `SalesData1_Raw.xlsx`.
3. Run `CleanSalesData1`.
4. Review `Log File` for a step-by-step audit trail.
5. Save the cleaned file as `Clean Data.xlsx`.
