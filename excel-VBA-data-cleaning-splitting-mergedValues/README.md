# Invoices With Merged Categories and Merged Amounts

## Problem
In the **Merged** sheet, each order combines multiple items and their amounts into single cells:
- The *Category* column lists several items separated by “|”.
- The *Amount* column lists corresponding values separated by “|”.
- This makes it impossible to analyze item‑level sales or perform category‑wise aggregation.

Example (Dirty):
- Order ID: CA‑2011‑167199
- Category: Binders | Art | Phones | Fasteners | Paper
- Amount: 609.98 | 5.48 | 391.98 | 755.96 | 31.12

## Solution
A VBA macro (`SplitInvoicesMergedData`) automates the transformation:
- Reads each Order ID.
- Splits the *Category* and *Amount* fields by the “|” delimiter.
- Writes each item and its corresponding amount on a new row.
- Repeats the Order ID for every item.

Example (Clean):
- CA‑2011‑167199 | Binders   | 609.98
- CA‑2011‑167199 | Art       | 5.48
- CA‑2011‑167199 | Phones    | 391.98
- CA‑2011‑167199 | Fasteners | 755.96
- CA‑2011‑167199 | Paper     | 31.12

## Demo
- Code: [Module_UnMerged.bas](demo/Module_UnMerged.bas)
- Input: [Merged.xlsx](src/Merged.xlsx) → **Dirty**
- Output: [UnMerged.xlsx](src/UnMerged.xlsx) → **Clean**

## Usage
1. Import `Module_UnMerged.bas` into Excel (`ALT+F11 → File → Import`).
2. Open the workbook containing the **Merged** sheet.
3. Run `SplitInvoicesMergedData`.
4. The macro creates/overwrites the **Clean** sheet with normalized data.
5. Save the workbook to preserve the clean dataset.
