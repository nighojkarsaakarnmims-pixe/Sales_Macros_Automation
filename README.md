# Sales GST Automation System (Excel VBA)

## Mandatory Setup Before Running Macros

Before running any macro, the header names in Row 1 of the Sales sheet
MUST be replaced with the following numbers:

  Header Name         Replace With Number
  ------------------- ---------------------
  Date                1
  Particular          2
  Buyer               3
  Buyer Address       4
  Consignee           5
  Consignee Address   6
  Voucher Type        7
  Voucher No          8
  Voucher Ref No      9
  GSTIN/UIN           10
  Gross Total         11
  HSN Code            12
  Quantity            13
  Measurement         14

### Tax Column Structure:

-   Columns 30--45 → Taxable Value
-   Columns 46--55 → IGST
-   Columns 56--65 → CGST
-   Columns 66--75 → SGST / UTGST

All numeric columns must be in Number format.

## How to Import VBA Code

1.  Open Excel
2.  Press ALT + F11
3.  File → Import File
4.  Select SalesAutomation.bas
5.  Run the required macro
