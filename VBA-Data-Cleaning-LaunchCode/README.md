# Data Cleaning Toolkit (VBA)

## Problem
Messy datasets slow down analysis and require manual effort.

## Solution
This VBA toolkit automates cleaning tasks:

-Missing Data
    -VBA macro to scan each column and highlight blanks.

Irregular Data

Focus on transaction_total.

VBA can calculate quartiles 

Add conditional formatting via VBA.

Unnecessary Data

Detect duplicate columns (e.g., two “email” columns).

VBA can compare values and delete redundant columns.

Inconsistent Data

## Demo
- Input: [Sample_DirtyData.xlsx](demo/Sample_DirtyData.xlsx)
- Output: [Sample_CleanData.xlsx.xlsx](demo/Sample_CleanData.xlsx.xlsx)

## Code
See [Module_DataCleaning.bas](src/Module_DataCleaning.bas) for VBA implementation.

## Usage
1. Import `Module_DataCleaning.bas` into Excel (ALT+F11 → File → Import).
2. Run `CleanLaunchCodeData`.
3. Cleaned data will appear in a new sheet with headers.

Validate emails (must contain @ and not start with @).

Clean transaction_total by stripping $ signs and converting to numeric.
