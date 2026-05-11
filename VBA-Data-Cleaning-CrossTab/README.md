📄 Dataset Description
Dirty sheet
Layout: Crosstab with merged headers.

Row 1 → Ship Mode (First Class, Same Day, Second Class, Standard Class).
Row 2 → Segment (Consumer, Corporate, Home Office).
Column A → Order Date.
Columns B onward → Sales values under each Ship Mode/Segment combination.

Problem: This structure is not suitable for analysis — dates are in one column, but sales are spread across multiple columns under merged headers.

Clean sheet
Layout: Normalized flat table.

Columns: Ship Mode | Segment | Order Date | Sales.
Each row represents one transaction: a combination of Ship Mode, Segment, Order Date, and Sales.
This format is analysis‑ready and can be used directly in BI tools or pivot tables.



# Badly Structured Sales Data (Unpivot Transformation)

## Problem
The dataset in **Dirty** is stored in a cross-tab format:
- Ship Mode values are merged across the top row.
- Segment values are merged in the second row.
- Order Dates are listed in column A.
- Sales values are spread across multiple columns under each Ship Mode/Segment.

This structure makes it impossible to analyze directly with pivot tables or BI tools.

## Solution
A VBA macro (`UnpivotSalesData2`) automates the transformation:
- Reads Ship Mode from row 1.
- Reads Segment from row 2.
- Loops through each Order Date in column A.
- Extracts the corresponding Sales value from each Ship Mode/Segment column.
- Writes out rows in a normalized table with four fields:
  - Ship Mode
  - Segment
  - Order Date
  - Sales

## Demo
- Code: [Module_StructureMultiHeader.bas](demo/Module_StructureMultiHeader.bas)
- Input: [Unstructured_Data.xlsx](src/Unstructured_Data.xlsx) → **Dirty**
- Output: [Structured Data.xlsx](src/Structured Data.xlsx) → **Clean**

## Usage
1. Import `Module_StructureMultiHeader.bas` into Excel (`ALT+F11 → File → Import`).
2. Open the workbook containing **Dirty**.
3. Run `UnpivotSalesData2`.
4. The macro creates/overwrites the **Clean** sheet with normalized data.
5. Save the workbook to preserve the clean dataset.
