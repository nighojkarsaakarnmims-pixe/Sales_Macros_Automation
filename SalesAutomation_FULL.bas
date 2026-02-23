Replacement of name:
“Date”=1
“Particular”=2
“Buyer”=3
“Buyer address”=4
“Consignee “=5
“Consignee address”=6
“Voucher type”=7
“Voucher no”=8
“Voucher ref no”=9
“GSTIN/UIN”=10
“Gross total”= 11
“Round of bills”=17
“Sale to sez unit”=18
HSN code=12(for sheet1)
Quantity=13
Measurement=14
30-45=taxable value
46-55=IGST
56-65=CGST
66-75=SGST/UTGST
Sequence for macros to run:
HSN CODE FROM SHEET 1 TO SALE REGISTER 
MACRO FOR SAME INVOICE NUMBER 
SUM OF SAME INOVICE NUMBER
SEGREGATE TALLY DATA CODE 
CREDIT AND CREDIT UNREGISTER
HSN SUMMARY CODE
TALLY B2B TO B2B
TALLY EXPORT TO EXPORT
CREDIT AND CREDIT UNREGISTER
TALLY B2C TO B2C
TALLY B2C 2.5 LAC TO B2C 2.5 LAC 


Macro for HSN code(to copy from “Sheet1” to “Sale Register”)
Sub CompareAndCopyByHeaders()

    Dim wsSales As Worksheet, wsSheet1 As Worksheet
    Dim col8_Sales As Long, col8_Sheet1 As Long
    Dim col12_Sheet1 As Long, col13_Sheet1 As Long, col14_Sheet1 As Long
    Dim col12_Sales As Long, col13_Sales As Long, col14_Sales As Long
    Dim lastRowSales As Long, lastRowSheet1 As Long
    Dim lastColSales As Long
    Dim i As Long, j As Long
    Dim matchValue As String
    Dim copiedVal12 As Variant, copiedVal13 As Variant, copiedVal14 As Variant

    Set wsSales = ThisWorkbook.Sheets("Sales")
    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")

    ' Find the column number of header "8" in both sheets
    col8_Sales = FindHeaderColumn(wsSales, "8")
    col8_Sheet1 = FindHeaderColumn(wsSheet1, "8")
    col12_Sheet1 = FindHeaderColumn(wsSheet1, "12")
    col13_Sheet1 = FindHeaderColumn(wsSheet1, "13")
    col14_Sheet1 = FindHeaderColumn(wsSheet1, "14")

    If col8_Sales * col8_Sheet1 * col12_Sheet1 * col13_Sheet1 * col14_Sheet1 = 0 Then Exit Sub

    ' Insert headers "12", "13", "14" in Sales Register if not found
    lastColSales = wsSales.Cells(1, wsSales.Columns.Count).End(xlToLeft).Column
    col12_Sales = FindOrInsertHeader(wsSales, "12", lastColSales + 1)
    col13_Sales = FindOrInsertHeader(wsSales, "13", col12_Sales + 1)
    col14_Sales = FindOrInsertHeader(wsSales, "14", col13_Sales + 1)

    ' Get last row based on column 8
    lastRowSales = wsSales.Cells(wsSales.Rows.Count, col8_Sales).End(xlUp).Row
    lastRowSheet1 = wsSheet1.Cells(wsSheet1.Rows.Count, col8_Sheet1).End(xlUp).Row

    ' Main loop to compare and copy
    For i = 2 To lastRowSales
        matchValue = wsSales.Cells(i, col8_Sales).Value
        copiedVal12 = ""
        copiedVal13 = ""
        copiedVal14 = ""

        For j = 2 To lastRowSheet1
            If wsSheet1.Cells(j, col8_Sheet1).Value = matchValue Then
                copiedVal12 = wsSheet1.Cells(j, col12_Sheet1).Value
                copiedVal13 = wsSheet1.Cells(j, col13_Sheet1).Value
                copiedVal14 = wsSheet1.Cells(j, col14_Sheet1).Value
                Exit For
            End If
        Next j

        wsSales.Cells(i, col12_Sales).Value = copiedVal12
        wsSales.Cells(i, col13_Sales).Value = copiedVal13
        wsSales.Cells(i, col14_Sales).Value = copiedVal14
    Next i

    MsgBox "Data copied successfully into columns 12, 13, and 14 in Sales Register."

End Sub

Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    Dim i As Long
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If Trim(ws.Cells(1, i).Value) = headerName Then
            FindHeaderColumn = i
            Exit Function
        End If
    Next i
    MsgBox "Header '" & headerName & "' not found in sheet '" & ws.Name & "'"
    FindHeaderColumn = 0
End Function

Function FindOrInsertHeader(ws As Worksheet, headerName As String, insertAtCol As Long) As Long
    Dim i As Long
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If Trim(ws.Cells(1, i).Value) = headerName Then
            FindOrInsertHeader = i
            Exit Function
        End If
    Next i
    ws.Columns(insertAtCol).Insert Shift:=xlToRight
    ws.Cells(1, insertAtCol).Value = headerName
    FindOrInsertHeader = insertAtCol
End Function

Macro for same invoice number :
Sub HighlightGSTRateAndDuplicateInvoices()
    Dim wsSrc As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim gstRateCol As Long
    Dim i As Long, j As Long
    Dim taxStartHeader As Long, taxEndHeader As Long
    Dim taxableStartHeader As Long, fullEndHeader As Long
    Dim taxCols() As Long, taxableCols() As Long, fullCols() As Long
    Dim taxCount As Long, taxableCount As Long, fullCount As Long
    Dim colInvoiceNo As Long, colGrossTotal As Long
    Dim col As Long
    Dim hdrVal As Variant
    Dim invoiceNo As String
    Dim invoiceDict As Object
    Dim sumTax As Double, sumTaxable As Double, sumFull As Double
    Dim gstRate As Double
    Dim key As Variant
    Dim coll As Collection
    Dim idx As Long
    Dim foundCol11 As Boolean

    Set wsSrc = ThisWorkbook.Sheets("Sales")  ' Your source sheet

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    ' Step 1: Find key column positions by header number
    colInvoiceNo = 0
    colGrossTotal = 0
    foundCol11 = False

    For col = 1 To lastCol
        hdrVal = wsSrc.Cells(1, col).Value
        If hdrVal = 8 Then colInvoiceNo = col
        If hdrVal = 11 Then
            colGrossTotal = col
            foundCol11 = True
        End If
    Next col

    ' If header 11 is missing, add it at next available column
    If Not foundCol11 Then
        colGrossTotal = lastCol + 1
        wsSrc.Cells(1, colGrossTotal).Value = 11
        lastCol = colGrossTotal
    End If

    ' GST RATE column (new column at the end)
    gstRateCol = lastCol + 1
    wsSrc.Cells(1, gstRateCol).Value = "GST RATE"

    ' Header ranges
    taxableStartHeader = 30
    taxStartHeader = 46
    taxEndHeader = 75
    fullEndHeader = 75

    ' Find column numbers corresponding to header values in row 1
    taxableCount = 0: taxCount = 0: fullCount = 0
    For col = 1 To lastCol
        hdrVal = wsSrc.Cells(1, col).Value
        If IsNumeric(hdrVal) Then
            If hdrVal >= 30 And hdrVal <= 45 Then
                taxableCount = taxableCount + 1
                ReDim Preserve taxableCols(1 To taxableCount)
                taxableCols(taxableCount) = col
            End If
            If hdrVal >= 46 And hdrVal <= 75 Then
                taxCount = taxCount + 1
                ReDim Preserve taxCols(1 To taxCount)
                taxCols(taxCount) = col
            End If
            If hdrVal >= 30 And hdrVal <= 75 Then
                fullCount = fullCount + 1
                ReDim Preserve fullCols(1 To fullCount)
                fullCols(fullCount) = col
            End If
        End If
    Next col

    ' Clear existing formatting
    wsSrc.Range(wsSrc.Rows(2), wsSrc.Rows(lastRow)).Interior.ColorIndex = xlNone

    ' Loop over data rows
    For i = 2 To lastRow
        sumTaxable = 0
        sumTax = 0
        sumFull = 0

        For j = 1 To taxableCount
            sumTaxable = sumTaxable + Nz(wsSrc.Cells(i, taxableCols(j)).Value)
        Next j

        For j = 1 To taxCount
            sumTax = sumTax + Nz(wsSrc.Cells(i, taxCols(j)).Value)
        Next j

        For j = 1 To fullCount
            sumFull = sumFull + Nz(wsSrc.Cells(i, fullCols(j)).Value)
        Next j

        ' Write Total (30-75) to column with header 11
        wsSrc.Cells(i, colGrossTotal).Value = Round(sumFull, 2)

        ' Calculate and write GST Rate
        If sumTaxable <> 0 Then
            gstRate = (sumTax / sumTaxable) * 100
        Else
            gstRate = 0
        End If
        wsSrc.Cells(i, gstRateCol).Value = Round(gstRate, 2)

    ' Highlight yellow if GST Rate is not 5, 12, 18, or 28
If Abs(gstRate - 5) > 0.1 And Abs(gstRate - 12) > 0.1 And Abs(gstRate - 18) > 0.1 And Abs(gstRate - 28) > 0.1 Then
    wsSrc.Rows(i).Interior.Color = vbYellow
End If
    Next i

    ' Highlight red for duplicate Invoice Numbers (header = 8)
    Set invoiceDict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        invoiceNo = Trim(CStr(wsSrc.Cells(i, colInvoiceNo).Value))
        If invoiceNo <> "" Then
            If invoiceDict.exists(invoiceNo) Then
                Set coll = invoiceDict(invoiceNo)
                coll.Add i
            Else
                Set coll = New Collection
                coll.Add i
                invoiceDict.Add invoiceNo, coll
            End If
        End If
    Next i

    For Each key In invoiceDict.Keys
        Set coll = invoiceDict(key)
        If coll.Count > 1 Then
            For idx = 1 To coll.Count
                wsSrc.Rows(coll(idx)).Interior.Color = vbRed
            Next idx
        End If
    Next key

    MsgBox "Done: GST Rate, Gross Total (col 11), and highlights updated.", vbInformation
End Sub

Function Nz(val As Variant) As Double
    If IsNumeric(val) Then
        Nz = CDbl(val)
    Else
        Nz = 0
    End If
End Function
Macro for the sum:
Sub SumGrossTotal_ByHeaderNumbers()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sales")
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim invoiceCol As Long, grossTotalCol As Long
    Dim i As Long
    
    ' Find columns by header value (Header 8 and Header 11)
    For i = 1 To lastCol
        If Trim(ws.Cells(1, i).Value) = "8" Then invoiceCol = i
        If Trim(ws.Cells(1, i).Value) = "11" Then grossTotalCol = i
    Next i

    If invoiceCol = 0 Or grossTotalCol = 0 Then
        MsgBox "Header '8' or '11' not found in row 1.", vbExclamation
        Exit Sub
    End If

    ' Build sum for each Invoice Number
    Dim invoiceDict As Object
    Set invoiceDict = CreateObject("Scripting.Dictionary")
    
    Dim invoiceNo As String
    Dim amount As Variant

    For i = 2 To lastRow
        invoiceNo = Trim(ws.Cells(i, invoiceCol).Value)
        amount = ws.Cells(i, grossTotalCol).Value

        If invoiceNo <> "" And IsNumeric(amount) Then
            If invoiceDict.exists(invoiceNo) Then
                invoiceDict(invoiceNo) = invoiceDict(invoiceNo) + CDbl(amount)
            Else
                invoiceDict.Add invoiceNo, CDbl(amount)
            End If
        End If
    Next i

    ' Write back the totals to same rows under Gross Total column
    For i = 2 To lastRow
        invoiceNo = Trim(ws.Cells(i, invoiceCol).Value)
        If invoiceNo <> "" Then
            If invoiceDict.exists(invoiceNo) Then
                ws.Cells(i, grossTotalCol).Value = invoiceDict(invoiceNo)
            End If
        End If
    Next i

    MsgBox "Gross Total summed per Invoice Number (Header 8), and written in Header 11.", vbInformation

End Sub

Macro to segregate the data into tally sheet :
Sub ClassifyAndSplitSalesRegister()

    Dim wsSrc As Worksheet
    Dim wsB2B As Worksheet, wsB2C25 As Worksheet, wsB2C As Worksheet
    Dim wsExport As Worksheet, wsCheck As Worksheet, wsCancelled As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, col As Long
    Dim invoiceNo As String, gstin As String, particulars As String
    Dim buyer As String, consignee As String
    Dim grossTotal As Double
    Dim invoiceType As String
    Dim invoiceTypeCol As Long

    Dim headerCol As Long
    Dim taxHeaderStart As Long, taxHeaderEnd As Long
    Dim taxCols() As Long
    Dim taxCount As Long
    Dim h As Long

    Dim colInvoiceNo As Long, colGSTIN As Long

    Set wsSrc = ThisWorkbook.Sheets("Sales Register")
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    invoiceTypeCol = lastCol + 1
    wsSrc.Cells(1, invoiceTypeCol).Value = "Invoice Type"

    ' Find headers
    colInvoiceNo = 0
    colGSTIN = 0
    For headerCol = 1 To lastCol
        Dim hdrVal As Variant
        hdrVal = wsSrc.Cells(1, headerCol).Value
        If hdrVal = 8 Then colInvoiceNo = headerCol
        If hdrVal = 10 Then colGSTIN = headerCol
    Next headerCol

    If colInvoiceNo = 0 Or colGSTIN = 0 Then
        MsgBox "Invoice No (8) or GSTIN (10) column not found!", vbCritical
        Exit Sub
    End If

    ' Tax headers
    taxHeaderStart = 45
    taxHeaderEnd = 75
    taxCount = 0
    For headerCol = 1 To lastCol
        hdrVal = wsSrc.Cells(1, headerCol).Value
        If IsNumeric(hdrVal) Then
            If hdrVal >= taxHeaderStart And hdrVal <= taxHeaderEnd Then
                taxCount = taxCount + 1
                ReDim Preserve taxCols(1 To taxCount)
                taxCols(taxCount) = headerCol
            End If
        End If
    Next headerCol

    ' Classification
    For i = 2 To lastRow
        invoiceNo = Trim(wsSrc.Cells(i, colInvoiceNo).Value)
        gstin = Trim(wsSrc.Cells(i, colGSTIN).Value)
        particulars = LCase(Trim(wsSrc.Cells(i, 2).Value))
        buyer = Trim(wsSrc.Cells(i, 3).Value)
        consignee = Trim(wsSrc.Cells(i, 5).Value)
        grossTotal = val(wsSrc.Cells(i, 11).Value)

        If Len(invoiceNo) > 16 Or (gstin <> "" And Len(gstin) <> 15) Then
            wsSrc.Cells(i, invoiceTypeCol).Value = "Cross Check"
            GoTo NextRowClassification
        End If

        Dim totalTax As Double
        totalTax = 0
        For h = 1 To taxCount
            totalTax = totalTax + Nz(wsSrc.Cells(i, taxCols(h)).Value)
        Next h

        ' === IGST, CGST, SGST values ===
        Dim igstVal As Double, cgstVal As Double, sgstVal As Double
        igstVal = 0: cgstVal = 0: sgstVal = 0
        For h = 1 To taxCount
            hdrVal = wsSrc.Cells(1, taxCols(h)).Value
            Select Case hdrVal
                Case 46 To 55: igstVal = igstVal + Nz(wsSrc.Cells(i, taxCols(h)).Value)
                Case 56 To 65: cgstVal = cgstVal + Nz(wsSrc.Cells(i, taxCols(h)).Value)
                Case 66 To 75: sgstVal = sgstVal + Nz(wsSrc.Cells(i, taxCols(h)).Value)
            End Select
        Next h

        ' === Classification ===
        If InStr(UCase(invoiceNo), "SEZ") > 0 Or _
           Left(UCase(invoiceNo), 3) = "EXP" Or _
           Left(UCase(invoiceNo), 3) = "SEZ" Then
            invoiceType = "Export"
        ElseIf gstin <> "" Then
            invoiceType = "B2B"
        Else
            If totalTax = 0 Then
                invoiceType = "Export"
            ElseIf grossTotal > 100000 And igstVal > 0 And cgstVal = 0 And sgstVal = 0 Then
                invoiceType = "B2C 2.5 Lac"
            Else
                invoiceType = "B2C"
            End If
        End If

        wsSrc.Cells(i, invoiceTypeCol).Value = invoiceType
NextRowClassification:
    Next i

    ' Prepare target sheets
    Set wsB2B = PrepareSheet("Tally B2B")
    Set wsB2C25 = PrepareSheet("Tally B2C 2.5 Lac")
    Set wsB2C = PrepareSheet("Tally B2C")
    Set wsExport = PrepareSheet("Tally Export")
    Set wsCheck = PrepareSheet("Cross Check")
    Set wsCancelled = PrepareSheet("Cancelled Entries")

    wsSrc.Rows(1).Copy wsB2B.Rows(1)
    wsSrc.Rows(1).Copy wsB2C25.Rows(1)
    wsSrc.Rows(1).Copy wsB2C.Rows(1)
    wsSrc.Rows(1).Copy wsExport.Rows(1)
    wsSrc.Rows(1).Copy wsCheck.Rows(1)
    wsSrc.Rows(1).Copy wsCancelled.Rows(1)

    Dim rowB2B As Long: rowB2B = 2
    Dim rowB2C25 As Long: rowB2C25 = 2
    Dim rowB2C As Long: rowB2C = 2
    Dim rowExport As Long: rowExport = 2
    Dim rowCheck As Long: rowCheck = 2
    Dim rowCancelled As Long: rowCancelled = 2

    ' Distribute rows
    For i = 2 To lastRow
        invoiceType = wsSrc.Cells(i, invoiceTypeCol).Value
        gstin = Trim(wsSrc.Cells(i, colGSTIN).Value)
        particulars = LCase(Trim(wsSrc.Cells(i, 2).Value))
        buyer = Trim(wsSrc.Cells(i, 3).Value)
        consignee = Trim(wsSrc.Cells(i, 5).Value)

        If (InStr(particulars, "cancelled") > 0 Or (gstin = "" And InStr(particulars, "cancelled") > 0)) _
           And (Len(particulars) = 0 Or Len(buyer) = 0 Or Len(consignee) = 0) Then
            wsSrc.Rows(i).Copy wsCancelled.Rows(rowCancelled)
            rowCancelled = rowCancelled + 1
        Else
            Select Case invoiceType
                Case "B2B"
                    wsSrc.Rows(i).Copy wsB2B.Rows(rowB2B): rowB2B = rowB2B + 1
                Case "B2C 2.5 Lac"
                    wsSrc.Rows(i).Copy wsB2C25.Rows(rowB2C25): rowB2C25 = rowB2C25 + 1
                Case "B2C"
                    wsSrc.Rows(i).Copy wsB2C.Rows(rowB2C): rowB2C = rowB2C + 1
                Case "Export"
                    wsSrc.Rows(i).Copy wsExport.Rows(rowExport): rowExport = rowExport + 1
                Case "Cross Check"
                    wsSrc.Rows(i).Copy wsCheck.Rows(rowCheck): rowCheck = rowCheck + 1
                Case Else
                    wsSrc.Rows(i).Copy wsCheck.Rows(rowCheck): rowCheck = rowCheck + 1
            End Select
        End If
    Next i

    ' Delete column with header 10 from Tally B2C
    Dim colToDelete As Long
    With wsB2C
        For colToDelete = .Cells(1, .Columns.Count).End(xlToLeft).Column To 1 Step -1
            If .Cells(1, colToDelete).Value = 10 Then
                .Columns(colToDelete).Delete
                Exit For
            End If
        Next colToDelete
    End With

    MsgBox "Classification complete. Cancelled and blank GSTIN, Buyer/Consignee handled.", vbInformation

End Sub

Function PrepareSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set PrepareSheet = ThisWorkbook.Sheets.Add
    PrepareSheet.Name = sheetName
End Function

Function Nz(val As Variant) As Double
    If IsNumeric(val) Then
        Nz = CDbl(val)
    Else
        Nz = 0
    End If
End Function


Flaws:
numbers should be number format and not any other formats 



HSN summary sheet:
Sub CreateHSNSummaryWithHSNAndTaxableAndTaxes()
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, invoiceCol As Long, hsnCol As Long, grossTotalCol As Long
    Dim col13 As Long, col14 As Long
    Dim destRow As Long, i As Long, col As Long
    Dim categories As Variant, cat As Variant
    Dim taxableCols As Collection, igstCols As Collection, cgstCols As Collection, sgstCols As Collection
    Dim taxableCol As Variant, igstCol As Variant, cgstCol As Variant, sgstCol As Variant
    Dim sumVal As Double, igstSum As Double, cgstSum As Double, sgstSum As Double
    Dim destColRate As Long, destColGross As Long, destCol13 As Long, destCol14 As Long

    Set wsSrc = ThisWorkbook.Sheets("Sales")
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    invoiceCol = 0: hsnCol = 0: col13 = 0: col14 = 0: grossTotalCol = 0
    Set taxableCols = New Collection
    Set igstCols = New Collection
    Set cgstCols = New Collection
    Set sgstCols = New Collection

    ' Detect headers
    For col = 1 To wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
        Select Case Trim(LCase(wsSrc.Cells(1, col).Value))
            Case "invoice type": invoiceCol = col
        End Select

        If Trim(wsSrc.Cells(1, col).Value) = "12" Then hsnCol = col
        If Trim(wsSrc.Cells(1, col).Value) = "11" Then grossTotalCol = col
        If Trim(wsSrc.Cells(1, col).Value) = "13" Then col13 = col
        If Trim(wsSrc.Cells(1, col).Value) = "14" Then col14 = col

        If IsNumeric(wsSrc.Cells(1, col).Value) Then
            Dim headerVal As Long
            headerVal = CLng(wsSrc.Cells(1, col).Value)
            If headerVal >= 30 And headerVal <= 45 Then taxableCols.Add col
            If headerVal >= 46 And headerVal <= 55 Then igstCols.Add col
            If headerVal >= 56 And headerVal <= 65 Then cgstCols.Add col
            If headerVal >= 66 And headerVal <= 75 Then sgstCols.Add col
        End If
    Next col

    ' Validate required headers
    If invoiceCol = 0 Then MsgBox "Invoice Type column not found.": Exit Sub
    If hsnCol = 0 Then MsgBox "Header '12' (HSN Code) not found.": Exit Sub
    If grossTotalCol = 0 Then MsgBox "Header '11' (Gross Total) not found.": Exit Sub
    If taxableCols.Count = 0 Then MsgBox "No headers from 30 to 45 found.": Exit Sub
    If igstCols.Count = 0 Then MsgBox "No headers from 46 to 55 found.": Exit Sub
    If cgstCols.Count = 0 Then MsgBox "No headers from 56 to 65 found.": Exit Sub
    If sgstCols.Count = 0 Then MsgBox "No headers from 66 to 75 found.": Exit Sub

    If col13 = 0 Then MsgBox "Header '13' not found. It will be skipped.", vbExclamation
    If col14 = 0 Then MsgBox "Header '14' not found. It will be skipped.", vbExclamation

    ' Create or clear HSN Summary sheet
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("HSN Summary")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSrc)
        wsDest.Name = "HSN Summary"
    Else
        wsDest.Cells.Clear
    End If
    On Error GoTo 0

    ' Write headers
    With wsDest
        .Cells(1, 1).Value = "Invoice Type"
        .Cells(1, 2).Value = "HSN Code"
        .Cells(1, 3).Value = "Taxable Value"
        .Cells(1, 4).Value = "IGST"
        .Cells(1, 5).Value = "CGST"
        .Cells(1, 6).Value = "SGST"
        .Cells(1, 7).Value = "Rate"
        .Cells(1, 8).Value = "Gross Total"
        destColRate = 7
        destColGross = 8
        destCol13 = 0: destCol14 = 0

        If col13 <> 0 Then
            destCol13 = 9
            .Cells(1, destCol13).Value = "13"
        End If
        If col14 <> 0 Then
            destCol14 = IIf(destCol13 > 0, destCol13 + 1, 9)
            .Cells(1, destCol14).Value = "14"
        End If
    End With
    destRow = 2

    ' Loop through categories
    categories = Array("B2B", "Export", "B2C", "B2C 2.5 Lac")

    For Each cat In categories
        For i = 2 To lastRow
            If Trim(UCase(wsSrc.Cells(i, invoiceCol).Value)) = UCase(cat) Then
                With wsDest
                    .Cells(destRow, 1).Value = wsSrc.Cells(i, invoiceCol).Value
                    .Cells(destRow, 2).Value = wsSrc.Cells(i, hsnCol).Value

                    sumVal = 0
                    For Each taxableCol In taxableCols
                        If IsNumeric(wsSrc.Cells(i, taxableCol).Value) Then
                            sumVal = sumVal + CDbl(wsSrc.Cells(i, taxableCol).Value)
                        End If
                    Next taxableCol
                    .Cells(destRow, 3).Value = sumVal

                    igstSum = 0
                    For Each igstCol In igstCols
                        If IsNumeric(wsSrc.Cells(i, igstCol).Value) Then
                            igstSum = igstSum + CDbl(wsSrc.Cells(i, igstCol).Value)
                        End If
                    Next igstCol
                    .Cells(destRow, 4).Value = igstSum

                    cgstSum = 0
                    For Each cgstCol In cgstCols
                        If IsNumeric(wsSrc.Cells(i, cgstCol).Value) Then
                            cgstSum = cgstSum + CDbl(wsSrc.Cells(i, cgstCol).Value)
                        End If
                    Next cgstCol
                    .Cells(destRow, 5).Value = cgstSum

                    sgstSum = 0
                    For Each sgstCol In sgstCols
                        If IsNumeric(wsSrc.Cells(i, sgstCol).Value) Then
                            sgstSum = sgstSum + CDbl(wsSrc.Cells(i, sgstCol).Value)
                        End If
                    Next sgstCol
                    .Cells(destRow, 6).Value = sgstSum

                    ' Rate column
                    .Cells(destRow, destColRate).Value = IIf(sumVal > 0, Round(((igstSum + cgstSum + sgstSum) / sumVal) * 100, 2), 0)

                    ' Gross Total column
                    .Cells(destRow, destColGross).Value = wsSrc.Cells(i, grossTotalCol).Value

                    ' Optional columns 13 and 14
                    If col13 <> 0 And destCol13 <> 0 Then .Cells(destRow, destCol13).Value = wsSrc.Cells(i, col13).Value
                    If col14 <> 0 And destCol14 <> 0 Then .Cells(destRow, destCol14).Value = wsSrc.Cells(i, col14).Value
                End With
                destRow = destRow + 1
            End If
        Next i
    Next cat

    MsgBox "HSN Summary with Gross Total, Rate, 13, and 14 created successfully!", vbInformation
End Sub

Macro tally b2b to b2b(2)( change the date format )
Sub CompareAndCopy_TallyB2B_to_B2B()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, destRow As Long
    Dim i As Long, j As Long
    Dim headerMap As Object
    Dim colHeader As Range
    Dim taxableSum As Double
    Dim igstSum As Double
    Dim cgstSum As Double
    Dim sgstSum As Double
    Dim gstRate As Double
    Dim stateMap As Object
    Dim gstin As String
    Dim stateCode As String

    Dim part1 As String, part2 As String, part3 As String
    Dim grossFallbackSum As Double

    Set wsSource = ThisWorkbook.Sheets("Tally B2B")
    Set wsDest = ThisWorkbook.Sheets("B2B")
    destRow = 9 ' Starting destination row

    ' Create header map (Header text ? column number)
    Set headerMap = CreateObject("Scripting.Dictionary")
    For Each colHeader In wsSource.Rows(1).Cells
        If Not IsEmpty(colHeader.Value) Then
            headerMap(CStr(colHeader.Value)) = colHeader.Column
        End If
    Next colHeader

    ' State code to state name map
    Set stateMap = CreateObject("Scripting.Dictionary")
    With stateMap
        .Add "01", "01-Jammu & Kashmir"
        .Add "02", "02-Himachal Pradesh"
        .Add "03", "03-Punjab"
        .Add "04", "04-Chandigarh"
        .Add "05", "05-Uttarakhand"
        .Add "06", "06-Haryana"
        .Add "07", "07-Delhi"
        .Add "08", "08-Rajasthan"
        .Add "09", "09-Uttar Pradesh"
        .Add "10", "10-Bihar"
        .Add "11", "11-Sikkim"
        .Add "12", "12-Arunachal Pradesh"
        .Add "13", "13-Nagaland"
        .Add "14", "14-Manipur"
        .Add "15", "15-Mizoram"
        .Add "16", "16-Tripura"
        .Add "17", "17-Meghalaya"
        .Add "18", "18-Assam"
        .Add "19", "19-West Bengal"
        .Add "20", "20-Jharkhand"
        .Add "21", "21-Odisha"
        .Add "22", "22-Chhattisgarh"
        .Add "23", "23-Madhya Pradesh"
        .Add "24", "24-Gujarat"
        .Add "25", "25-Daman & Diu"
        .Add "26", "26-Dadra & Nagar Haveli"
        .Add "27", "27-Maharashtra"
        .Add "28", "28-Andhra Pradesh (Old)"
        .Add "29", "29-Karnataka"
        .Add "30", "30-Goa"
        .Add "31", "31-Lakshadweep"
        .Add "32", "32-Kerala"
        .Add "33", "33-Tamil Nadu"
        .Add "34", "34-Puducherry"
        .Add "35", "35-Andaman & Nicobar Islands"
        .Add "36", "36-Telangana"
        .Add "37", "37-Andhra Pradesh"
        .Add "38", "38-Ladakh"
        .Add "97", "97-Other Territory"
    End With

    ' Get last row in source sheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    ' Loop through each row in source data
    For i = 2 To lastRow
        ' Copy column "1" if exists and format date as dd-mm-yyyy
        If headerMap.Exists("1") Then
            Dim dtValue As Variant
            dtValue = wsSource.Cells(i, headerMap("1")).Value
            If IsDate(dtValue) Then
                wsDest.Cells(destRow, 4).Value = CDate(dtValue)
                wsDest.Cells(destRow, 4).NumberFormat = "dd-mm-yyyy"
            Else
                wsDest.Cells(destRow, 4).Value = dtValue ' If not date, copy as is
            End If
        Else
            wsDest.Cells(destRow, 4).Value = ""
        End If

        ' Build concatenated string for columns 2,3,5
        If headerMap.Exists("2") Then part1 = wsSource.Cells(i, headerMap("2")).Value Else part1 = ""
        If headerMap.Exists("3") Then part2 = wsSource.Cells(i, headerMap("3")).Value Else part2 = ""
        If headerMap.Exists("5") Then part3 = wsSource.Cells(i, headerMap("5")).Value Else part3 = ""

        wsDest.Cells(destRow, 2).Value = Trim(part1 & " " & part2 & " " & part3)

        ' Copy GSTIN - header "10"
        If headerMap.Exists("10") Then
            wsDest.Cells(destRow, 1).Value = wsSource.Cells(i, headerMap("10")).Value
        Else
            wsDest.Cells(destRow, 1).Value = ""
        End If

        ' Copy Voucher Number - header "8"
        If headerMap.Exists("8") Then
            wsDest.Cells(destRow, 3).Value = wsSource.Cells(i, headerMap("8")).Value
        Else
            wsDest.Cells(destRow, 3).Value = ""
        End If

        ' Copy Gross Total - header "11" primary, fallback to sum 30–75
        If headerMap.Exists("11") Then
            wsDest.Cells(destRow, 5).Value = wsSource.Cells(i, headerMap("11")).Value
        Else
            grossFallbackSum = 0
            For j = 30 To 75
                If headerMap.Exists(CStr(j)) Then
                    grossFallbackSum = grossFallbackSum + val(wsSource.Cells(i, headerMap(CStr(j))).Value)
                End If
            Next j
            wsDest.Cells(destRow, 5).Value = grossFallbackSum
        End If

        ' Sum Taxable Value from headers 30 to 45
        taxableSum = 0
        For j = 30 To 45
            If headerMap.Exists(CStr(j)) Then
                taxableSum = taxableSum + val(wsSource.Cells(i, headerMap(CStr(j))).Value)
            End If
        Next j
        wsDest.Cells(destRow, 6).Value = taxableSum

        ' Sum IGST from headers 46 to 55
        igstSum = 0
        For j = 46 To 55
            If headerMap.Exists(CStr(j)) Then
                igstSum = igstSum + val(wsSource.Cells(i, headerMap(CStr(j))).Value)
            End If
        Next j
        wsDest.Cells(destRow, 9).Value = igstSum

        ' Sum CGST from headers 56 to 65
        cgstSum = 0
        For j = 56 To 65
            If headerMap.Exists(CStr(j)) Then
                cgstSum = cgstSum + val(wsSource.Cells(i, headerMap(CStr(j))).Value)
            End If
        Next j
        wsDest.Cells(destRow, 10).Value = cgstSum

        ' Sum SGST from headers 66 to 75
        sgstSum = 0
        For j = 66 To 75
            If headerMap.Exists(CStr(j)) Then
                sgstSum = sgstSum + val(wsSource.Cells(i, headerMap(CStr(j))).Value)
            End If
        Next j
        wsDest.Cells(destRow, 11).Value = sgstSum

        ' Calculate GST Rate
        If taxableSum <> 0 Then
            gstRate = ((igstSum + cgstSum + sgstSum) / taxableSum) * 100
        Else
            gstRate = 0
        End If
        wsDest.Cells(destRow, 8).Value = Round(gstRate, 2)

        ' Extract State from GSTIN (first two digits)
        gstin = wsDest.Cells(destRow, 1).Value
        If Len(gstin) >= 2 Then
            stateCode = Left(gstin, 2)
            If stateMap.Exists(stateCode) Then
                wsDest.Cells(destRow, 13).Value = stateMap(stateCode)
            Else
                wsDest.Cells(destRow, 13).Value = "Unknown State Code"
            End If
        Else
            wsDest.Cells(destRow, 13).Value = "Invalid GSTIN"
        End If

        destRow = destRow + 1
    Next i

    MsgBox "Data transfer completed successfully.", vbInformation
End Sub

`

Macro to copy data from tally b2b to b2b :
Sub CompareAndCopy_TallyB2B_to_B2B()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, destRow As Long
    Dim i As Long, j As Long
    Dim headerMap As Object
    Dim colHeader As Range
    Dim taxableSum As Double
    Dim igstSum As Double
    Dim cgstSum As Double
    Dim sgstSum As Double
    Dim gstRate As Double
    Dim stateMap As Object
    Dim gstin As String
    Dim stateCode As String

    Set wsSource = ThisWorkbook.Sheets("Tally B2B")
    Set wsDest = ThisWorkbook.Sheets("B2B")
    destRow = 9 ' Start writing from this row in B2B

    ' Build header map from Tally B2B sheet
    Set headerMap = CreateObject("Scripting.Dictionary")
    For Each colHeader In wsSource.Rows(1).Cells
        If Not IsEmpty(colHeader.Value) Then
            headerMap(colHeader.Value) = colHeader.Column
        End If
    Next colHeader

    ' Build state code to name mapping with your exact format
    Set stateMap = CreateObject("Scripting.Dictionary")
    With stateMap
        .Add "01", "01-Jammu & Kashmir"
        .Add "02", "02-Himachal Pradesh"
        .Add "03", "03-Punjab"
        .Add "04", "04-Chandigarh"
        .Add "05", "05-Uttarakhand"
        .Add "06", "06-Haryana"
        .Add "07", "07-Delhi"
        .Add "08", "08-Rajasthan"
        .Add "09", "09-Uttar Pradesh"
        .Add "10", "10-Bihar"
        .Add "11", "11-Sikkim"
        .Add "12", "12-Arunachal Pradesh"
        .Add "13", "13-Nagaland"
        .Add "14", "14-Manipur"
        .Add "15", "15-Mizoram"
        .Add "16", "16-Tripura"
        .Add "17", "17-Meghalaya"
        .Add "18", "18-Assam"
        .Add "19", "19-West Bengal"
        .Add "20", "20-Jharkhand"
        .Add "21", "21-Odisha"
        .Add "22", "22-Chhattisgarh"
        .Add "23", "23-Madhya Pradesh"
        .Add "24", "24-Gujarat"
        .Add "25", "25-Daman & Diu"
        .Add "26", "26-Dadra & Nagar Haveli"
        .Add "27", "27-Maharashtra"
        .Add "28", "28-Andhra Pradesh (Old)" ' Added 28 to keep full list consistent
        .Add "29", "29-Karnataka"
        .Add "30", "30-Goa"
        .Add "31", "31-Lakshadweep"
        .Add "32", "32-Kerala"
        .Add "33", "33-Tamil Nadu"
        .Add "34", "34-Puducherry"
        .Add "35", "35-Andaman & Nicobar Islands"
        .Add "36", "36-Telangana"
        .Add "37", "37-Andhra Pradesh"
        .Add "38", "38-Ladakh"
        .Add "97", "97-Other Territory"
    End With

    ' Find last row of data in Tally B2B
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    ' Loop through source data
    For i = 2 To lastRow
        ' === Your existing copying code ===
        wsDest.Cells(destRow, 4).Value = wsSource.Cells(i, headerMap("1")).Value
        wsDest.Cells(destRow, 2).Value = _
            wsSource.Cells(i, headerMap("2")).Value & " " & _
            wsSource.Cells(i, headerMap("3")).Value & " " & _
            wsSource.Cells(i, headerMap("5")).Value
        wsDest.Cells(destRow, 1).Value = wsSource.Cells(i, headerMap("10")).Value
        wsDest.Cells(destRow, 3).Value = wsSource.Cells(i, headerMap("8")).Value
        wsDest.Cells(destRow, 5).Value = wsSource.Cells(i, headerMap("11")).Value

        taxableSum = 0
        For j = 30 To 45
            If headerMap.exists(CStr(j)) Then
                taxableSum = taxableSum + val(wsSource.Cells(i, headerMap(CStr(j))).Value)
            End If
        Next j
        wsDest.Cells(destRow, 6).Value = taxableSum

        igstSum = 0
        For j = 46 To 55
            If headerMap.exists(CStr(j)) Then
                igstSum = igstSum + val(wsSource.Cells(i, headerMap(CStr(j))).Value)
            End If
        Next j
        wsDest.Cells(destRow, 9).Value = igstSum

        cgstSum = 0
        For j = 56 To 65
            If headerMap.exists(CStr(j)) Then
                cgstSum = cgstSum + val(wsSource.Cells(i, headerMap(CStr(j))).Value)
            End If
        Next j
        wsDest.Cells(destRow, 10).Value = cgstSum

        sgstSum = 0
        For j = 66 To 75
            If headerMap.exists(CStr(j)) Then
                sgstSum = sgstSum + val(wsSource.Cells(i, headerMap(CStr(j))).Value)
            End If
        Next j
        wsDest.Cells(destRow, 11).Value = sgstSum

        If taxableSum <> 0 Then
            gstRate = ((igstSum + cgstSum + sgstSum) / taxableSum) * 100
        Else
            gstRate = 0
        End If
        wsDest.Cells(destRow, 8).Value = Round(gstRate, 2)

        ' === Your new condition: Extract first 2 digits of GSTIN in column 1 and map state name ===
        gstin = wsDest.Cells(destRow, 1).Value
        If Len(gstin) >= 2 Then
            stateCode = Left(gstin, 2)
            If stateMap.exists(stateCode) Then
                wsDest.Cells(destRow, 13).Value = stateMap(stateCode)
            Else
                wsDest.Cells(destRow, 13).Value = "Unknown State Code"
            End If
        Else
            wsDest.Cells(destRow, 13).Value = "Invalid GSTIN"
        End If

        destRow = destRow + 1
    Next i

    MsgBox "Data transfer from Tally B2B to B2B completed successfully.", vbInformation
End Sub

	

Macro to copy from tally b2c 2.5 lac to b2c 2.5 lac:
Sub Transfer_Tally_B2C_2_5_Lac_With_GSTRate()
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim lastRowSrc As Long
    Dim srcRow As Long, destRow As Long
    Dim partName As String
    Dim sumTaxable As Double
    Dim sumIGST As Double
    Dim gstRate As Double
    Dim headerRow As Long
    Dim col As Long
    Dim headerVal As Long
    Dim taxableCols As Collection
    Dim igstCols As Collection
    Dim cellValue As Variant
    Dim c As Variant  ' For For Each loop

    ' Define worksheets
    Set wsSrc = ThisWorkbook.Sheets("Tally B2C 2.5 Lac")
    Set wsDest = ThisWorkbook.Sheets("B2C 2.5 Lac")

    headerRow = 1 ' Assuming headers are in row 1 in source sheet

    ' Collect all physical columns where header number is between 30 and 45 for Taxable Value
    Set taxableCols = New Collection
    ' Collect all physical columns where header number is between 46 and 55 for IGST
    Set igstCols = New Collection

    For col = 1 To wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column
        cellValue = wsSrc.Cells(headerRow, col).Value
        If IsNumeric(cellValue) Then
            headerVal = CLng(cellValue)
            If headerVal >= 30 And headerVal <= 45 Then
                taxableCols.Add col
            ElseIf headerVal >= 46 And headerVal <= 55 Then
                igstCols.Add col
            End If
        End If
    Next col

    ' Start copying from row 2 in source
    srcRow = 2

    ' Start pasting at row 9 in destination
    destRow = 9

    ' Find the last row in the source sheet from column 1
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    ' Loop through source rows from row 2 to last row
    Do While srcRow <= lastRowSrc
        ' --- Date (Header 1) to Column 3 ---
        wsDest.Cells(destRow, 3).Value = wsSrc.Cells(srcRow, 1).Value

        ' --- Party Name (Header 2 or 3 or 5) to Column 1 ---
        If wsSrc.Cells(srcRow, 2).Value <> "" Then
            partName = wsSrc.Cells(srcRow, 2).Value
        ElseIf wsSrc.Cells(srcRow, 3).Value <> "" Then
            partName = wsSrc.Cells(srcRow, 3).Value
        ElseIf wsSrc.Cells(srcRow, 5).Value <> "" Then
            partName = wsSrc.Cells(srcRow, 5).Value
        Else
            partName = ""
        End If
        wsDest.Cells(destRow, 1).Value = partName

        ' --- Number (Header 8) to Column 2 ---
        wsDest.Cells(destRow, 2).Value = wsSrc.Cells(srcRow, 8).Value

        ' --- Total (Header 11) to Column 4 ---
        wsDest.Cells(destRow, 4).Value = wsSrc.Cells(srcRow, 11).Value

        ' --- Sum of taxable columns (headers 30 to 45) ---
        sumTaxable = 0
        For Each c In taxableCols
            If IsNumeric(wsSrc.Cells(srcRow, c).Value) Then
                sumTaxable = sumTaxable + wsSrc.Cells(srcRow, c).Value
            End If
        Next c
        wsDest.Cells(destRow, 5).Value = sumTaxable

        ' --- Sum of IGST columns (headers 46 to 55) ---
        sumIGST = 0
        For Each c In igstCols
            If IsNumeric(wsSrc.Cells(srcRow, c).Value) Then
                sumIGST = sumIGST + wsSrc.Cells(srcRow, c).Value
            End If
        Next c
        wsDest.Cells(destRow, 8).Value = sumIGST

        ' --- Calculate GST Rate and paste in Column 7 ---
        If sumTaxable <> 0 Then
            gstRate = (sumIGST / sumTaxable) * 100
        Else
            gstRate = 0
        End If
        wsDest.Cells(destRow, 7).Value = gstRate

        ' Move to next rows
        srcRow = srcRow + 1
        destRow = destRow + 1
    Loop

    MsgBox "Data transferred with GST Rate calculated.", vbInformation
End Sub
Macro to copy tally export to  export :
Sub CopyTallyExportToExport()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRowSource As Long, i As Long, h As Long
    Dim sourceColDate As Long, sourceColParty1 As Long, sourceColParty2 As Long, sourceColParty3 As Long
    Dim sourceColNumber As Long, sourceColTotal As Long, sourceColGSTIN As Long
    Dim destRow As Long

    Set wsSource = ThisWorkbook.Sheets("Tally Export")
    Set wsDest = ThisWorkbook.Sheets("Export")

    destRow = 9 ' Starting row in Export sheet
    Dim headerRow As Long: headerRow = 1

    Dim colHeaders As Object
    Set colHeaders = CreateObject("Scripting.Dictionary")

    ' Build header mapping
    Dim c As Range
    For Each c In wsSource.Rows(headerRow).Cells
        If IsNumeric(c.Value) Then
            colHeaders(CLng(c.Value)) = c.Column
        End If
    Next c

    ' Assign headers
    sourceColDate = IIf(colHeaders.Exists(1), colHeaders(1), 0)
    sourceColParty1 = IIf(colHeaders.Exists(2), colHeaders(2), 0)
    sourceColParty2 = IIf(colHeaders.Exists(3), colHeaders(3), 0)
    sourceColParty3 = IIf(colHeaders.Exists(5), colHeaders(5), 0)
    sourceColNumber = IIf(colHeaders.Exists(10), colHeaders(10), 0)
    sourceColTotal = IIf(colHeaders.Exists(11), colHeaders(11), 0)
    sourceColGSTIN = IIf(colHeaders.Exists(10), colHeaders(10), 0) ' GSTIN also at header 10

    If sourceColDate = 0 Then
        MsgBox "Header 1 (Date) not found in Tally Export!", vbCritical
        Exit Sub
    End If

    lastRowSource = wsSource.Cells(wsSource.Rows.Count, sourceColDate).End(xlUp).Row
    If lastRowSource < headerRow + 1 Then
        MsgBox "No data found in Tally Export.", vbInformation
        Exit Sub
    End If

    ' Clear Export sheet (A to R)
    wsDest.Range(wsDest.Cells(destRow, 1), wsDest.Cells(wsDest.Rows.Count, 18)).ClearContents

    Dim destRowIndex As Long: destRowIndex = destRow
    Dim partyColToUse As Long

    If sourceColParty1 <> 0 Then
        partyColToUse = sourceColParty1
    ElseIf sourceColParty2 <> 0 Then
        partyColToUse = sourceColParty2
    ElseIf sourceColParty3 <> 0 Then
        partyColToUse = sourceColParty3
    Else
        partyColToUse = 0
    End If

    Dim taxableVal As Double, igstSum As Double, cgstSum As Double, sgstSum As Double, gstRate As Double

    ' Place of Supply Map
    Dim stateMap As Object: Set stateMap = CreateObject("Scripting.Dictionary")
    With stateMap
        .Add "01", "01-Jammu & Kashmir"
        .Add "02", "02-Himachal Pradesh"
        .Add "03", "03-Punjab"
        .Add "04", "04-Chandigarh"
        .Add "05", "05-Uttarakhand"
        .Add "06", "06-Haryana"
        .Add "07", "07-Delhi"
        .Add "08", "08-Rajasthan"
        .Add "09", "09-Uttar Pradesh"
        .Add "10", "10-Bihar"
        .Add "11", "11-Sikkim"
        .Add "12", "12-Arunachal Pradesh"
        .Add "13", "13-Nagaland"
        .Add "14", "14-Manipur"
        .Add "15", "15-Mizoram"
        .Add "16", "16-Tripura"
        .Add "17", "17-Meghalaya"
        .Add "18", "18-Assam"
        .Add "19", "19-West Bengal"
        .Add "20", "20-Jharkhand"
        .Add "21", "21-Odisha"
        .Add "22", "22-Chhattisgarh"
        .Add "23", "23-Madhya Pradesh"
        .Add "24", "24-Gujarat"
        .Add "25", "25-Daman & Diu"
        .Add "26", "26-Dadra & Nagar Haveli"
        .Add "27", "27-Maharashtra"
        .Add "28", "28-Andhra Pradesh (Old)"
        .Add "29", "29-Karnataka"
        .Add "30", "30-Goa"
        .Add "31", "31-Lakshadweep"
        .Add "32", "32-Kerala"
        .Add "33", "33-Tamil Nadu"
        .Add "34", "34-Puducherry"
        .Add "35", "35-Andaman & Nicobar Islands"
        .Add "36", "36-Telangana"
        .Add "37", "37-Andhra Pradesh"
        .Add "38", "38-Ladakh"
        .Add "97", "97-Other Territory"
    End With

    For i = headerRow + 1 To lastRowSource
        ' --- Basic Columns ---
        wsDest.Cells(destRowIndex, 5).Value = wsSource.Cells(i, sourceColDate).Value ' Date (E)
        wsDest.Cells(destRowIndex, 2).Value = wsSource.Cells(i, sourceColNumber).Value ' Invoice No (B)
        wsDest.Cells(destRowIndex, 3).Value = IIf(partyColToUse <> 0, wsSource.Cells(i, partyColToUse).Value, "")

        ' --- Gross Total in column 6 (F) ---
        wsDest.Cells(destRowIndex, 6).Value = wsSource.Cells(i, sourceColTotal).Value

        ' --- Taxable Value in column 7 (G) ---
        taxableVal = 0
        For h = 30 To 45
            If colHeaders.Exists(h) Then
                taxableVal = taxableVal + SafeVal(wsSource.Cells(i, colHeaders(h)).Value)
            End If
        Next h
        wsDest.Cells(destRowIndex, 7).Value = taxableVal

        ' --- IGST in column 14 (N) ---
        igstSum = 0
        For h = 46 To 55
            If colHeaders.Exists(h) Then
                igstSum = igstSum + SafeVal(wsSource.Cells(i, colHeaders(h)).Value)
            End If
        Next h
        wsDest.Cells(destRowIndex, 14).Value = igstSum

        ' --- GST Classification in column 1 (A) ---
        wsDest.Cells(destRowIndex, 1).Value = IIf(igstSum > 0, "With payment of GST", "Without payment of GST")

        ' --- CGST to column 15 (O) ---
        cgstSum = 0
        For h = 56 To 65
            If colHeaders.Exists(h) Then
                cgstSum = cgstSum + SafeVal(wsSource.Cells(i, colHeaders(h)).Value)
            End If
        Next h
        wsDest.Cells(destRowIndex, 15).Value = cgstSum

        ' --- SGST to column 16 (P) ---
        sgstSum = 0
        For h = 66 To 75
            If colHeaders.Exists(h) Then
                sgstSum = sgstSum + SafeVal(wsSource.Cells(i, colHeaders(h)).Value)
            End If
        Next h
        wsDest.Cells(destRowIndex, 16).Value = sgstSum

        ' --- GST Rate in column 13 (M) ---
        If taxableVal <> 0 Then
            gstRate = ((igstSum + cgstSum + sgstSum) / taxableVal) * 100
        Else
            gstRate = 0
        End If
        wsDest.Cells(destRowIndex, 13).Value = gstRate

        ' --- Place of Supply in column 18 (R) ---
        Dim gstinVal As String
        gstinVal = Trim(wsSource.Cells(i, sourceColGSTIN).Value)
        If Len(gstinVal) = 0 Then
            ' GSTIN is blank ? leave POS blank
            wsDest.Cells(destRowIndex, 18).Value = ""
        ElseIf Len(gstinVal) >= 2 Then
            Dim code2 As String: code2 = Left(gstinVal, 2)
            If stateMap.Exists(code2) Then
                wsDest.Cells(destRowIndex, 18).Value = stateMap(code2)
            Else
                wsDest.Cells(destRowIndex, 18).Value = "Unknown"
            End If
        Else
            wsDest.Cells(destRowIndex, 18).Value = ""
        End If

        destRowIndex = destRowIndex + 1
    Next i

    MsgBox "Data copied successfully with Place of Supply.", vbInformation
End Sub

Function SafeVal(v As Variant) As Double
    If IsNumeric(v) Then SafeVal = CDbl(v) Else SafeVal = 0
End Function

Macro to copy the data from tally b2c to b2c :
Sub SummarizePlaceCounts_FromTallyB2C_ToB2C()

    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim colAddr As Long, c As Long
    Dim lastRowSrc As Long, i As Long, j As Long
    Dim addr As String
    Dim places As Variant
    Dim groupKey As String, key As Variant
    Dim dictCount As Object, dictTaxable As Object, dictIGST As Object
    Dim dictCGST As Object, dictSGST As Object
    Dim taxableVal As Double, igstVal As Double, cgstVal As Double, sgstVal As Double
    Dim ratio As Double, closestVal As Double, minDiff As Double, diff As Variant
    Dim givenValues As Variant
    Dim stateName As String, stateCode As String

    places = Array( _
        "Jammu & Kashmir-01", "Himachal Pradesh-02", "Punjab-03", "Chandigarh-04", "Uttarakhand-05", _
        "Haryana-06", "Delhi-07", "Rajasthan-08", "Uttar Pradesh-09", "Bihar-10", "Sikkim-11", _
        "Arunachal Pradesh-12", "Nagaland-13", "Manipur-14", "Mizoram-15", "Tripura-16", _
        "Meghalaya-17", "Assam-18", "West Bengal-19", "Jharkhand-20", "Odisha-21", "Chhattisgarh-22", _
        "Madhya Pradesh-23", "Gujarat-24", "Daman and Diu and Dadra & Nagar Haveli-25", "Maharashtra-27", _
        "Andhra Pradesh (Old)-28", "Karnataka-29", "Goa-30", "Lakshadweep-31", "Kerala-32", _
        "Tamil Nadu-33", "Puducherry-34", "Andaman & Nicobar Islands-35", "Telangana-36", _
        "Andhra Pradesh (Newly Added)-37", "Ladakh (Newly Added)-38", "Other Territory-97", "Center Jurisdiction")

    givenValues = Array(0, 0.1, 0.25, 3, 5, 6, 12, 18, 28, 1, 1.5, 7.5)

    Set wsSrc = ThisWorkbook.Sheets("Tally B2C")
    Set wsDest = ThisWorkbook.Sheets("B2C")

    Set dictCount = CreateObject("Scripting.Dictionary")
    Set dictTaxable = CreateObject("Scripting.Dictionary")
    Set dictIGST = CreateObject("Scripting.Dictionary")
    Set dictCGST = CreateObject("Scripting.Dictionary")
    Set dictSGST = CreateObject("Scripting.Dictionary")

    ' find address column (header "4")
    For c = 1 To wsSrc.UsedRange.Columns.Count
        If Trim(wsSrc.Cells(1, c).Value) = "4" Then
            colAddr = c: Exit For
        End If
    Next c
    If colAddr = 0 Then MsgBox "Header '4' not found.", vbCritical: Exit Sub

    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, colAddr).End(xlUp).Row

    For i = 2 To lastRowSrc
        addr = LCase(Trim(wsSrc.Cells(i, colAddr).Value))
        stateName = "": stateCode = ""

        ' match address to state
        For j = LBound(places) To UBound(places)
            If InStr(addr, LCase(Split(places(j), "-")(0))) > 0 Then
                stateName = Split(places(j), "-")(0)
                stateCode = Split(places(j), "-")(1)
                Exit For
            End If
        Next j
        If stateName = "" Then GoTo NextRow

        ' taxable: 30–45
        taxableVal = 0
        For c = 1 To wsSrc.UsedRange.Columns.Count
            If val(wsSrc.Cells(1, c).Value) >= 30 And val(wsSrc.Cells(1, c).Value) <= 45 Then
                taxableVal = taxableVal + val(wsSrc.Cells(i, c).Value)
            End If
        Next c

        ' IGST: 46–55
        igstVal = 0
        For c = 1 To wsSrc.UsedRange.Columns.Count
            If val(wsSrc.Cells(1, c).Value) >= 46 And val(wsSrc.Cells(1, c).Value) <= 55 Then
                igstVal = igstVal + val(wsSrc.Cells(i, c).Value)
            End If
        Next c

        ' CGST: 56–65
        cgstVal = 0
        For c = 1 To wsSrc.UsedRange.Columns.Count
            If val(wsSrc.Cells(1, c).Value) >= 56 And val(wsSrc.Cells(1, c).Value) <= 65 Then
                cgstVal = cgstVal + val(wsSrc.Cells(i, c).Value)
            End If
        Next c

        ' SGST: 66–75
        sgstVal = 0
        For c = 1 To wsSrc.UsedRange.Columns.Count
            If val(wsSrc.Cells(1, c).Value) >= 66 And val(wsSrc.Cells(1, c).Value) <= 75 Then
                sgstVal = sgstVal + val(wsSrc.Cells(i, c).Value)
            End If
        Next c

        ' GST rate
        If taxableVal <> 0 Then
            ratio = ((igstVal + cgstVal + sgstVal) / taxableVal) * 100
        Else
            ratio = 0
        End If

        minDiff = 999999: closestVal = 0
        For Each diff In givenValues
            If Abs(ratio - diff) < minDiff Then
                minDiff = Abs(ratio - diff)
                closestVal = diff
            End If
        Next diff

        groupKey = stateName & "-" & stateCode & "|" & closestVal

        If Not dictCount.exists(groupKey) Then
            dictCount.Add groupKey, 1
            dictTaxable.Add groupKey, taxableVal
            dictIGST.Add groupKey, igstVal
            dictCGST.Add groupKey, cgstVal
            dictSGST.Add groupKey, sgstVal
        Else
            dictCount(groupKey) = dictCount(groupKey) + 1
            dictTaxable(groupKey) = dictTaxable(groupKey) + taxableVal
            dictIGST(groupKey) = dictIGST(groupKey) + igstVal
            dictCGST(groupKey) = dictCGST(groupKey) + cgstVal
            dictSGST(groupKey) = dictSGST(groupKey) + sgstVal
        End If

NextRow:
    Next i

    ' Output
    i = 9
    For Each key In dictCount.Keys
        stateName = Split(Split(key, "|")(0), "-")(0)
        stateCode = Split(Split(key, "|")(0), "-")(1)
        closestVal = Split(key, "|")(1)

        wsDest.Cells(i, 2).Value = stateCode & " - " & stateName
        wsDest.Cells(i, 4).Value = dictTaxable(key)
        wsDest.Cells(i, 6).Value = closestVal
        wsDest.Cells(i, 7).Value = dictIGST(key)
        wsDest.Cells(i, 8).Value = dictCGST(key)
        wsDest.Cells(i, 9).Value = dictSGST(key)

        If dictIGST(key) > 0 And dictCGST(key) = 0 And dictSGST(key) = 0 Then
            wsDest.Cells(i, 1).Value = "Inter-state"
        ElseIf dictIGST(key) = 0 And dictCGST(key) > 0 And dictSGST(key) > 0 Then
            wsDest.Cells(i, 1).Value = "Intra-state"
        End If

        i = i + 1
    Next key

    MsgBox "Summary with GST-rate + State grouping completed.", vbInformation

End Sub



Credit and Credit Unregister:
COMBINE OF BOTH AND TALLY B2C:
Sub InsertComputeCopyAndTallyB2C()
    Dim ws As Worksheet, wsCredit As Worksheet, wsCreditUnreg As Worksheet, wsTallyB2C As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim i As Long, j As Long
    Dim headerVal As Variant
    Dim col11 As Long, colTaxable As Long, colIGST As Long, colCGST As Long, colSGST As Long
    Dim headerExists As Object
    Dim newHeaders As Variant
    Dim destRowCredit As Long, destRowUnreg As Long, destRowTally As Long
    Dim valGSTIN As String, valInvoice As String, valGrossTotal As Double
    Dim taxableVal As Double, igstVal As Double, cgstVal As Double, sgstVal As Double
    Dim gstPercent As Double
    Dim stateDict As Object
    Dim srcStateName As String, prefixStateUnreg As String
    Dim prefixStateCredit As String

    Set ws = ThisWorkbook.Sheets("Credit Note Register")
    Set wsCredit = ThisWorkbook.Sheets("Credit")
    Set wsCreditUnreg = ThisWorkbook.Sheets("Credit Unregister")
    Set wsTallyB2C = ThisWorkbook.Sheets("Tally B2C")

    ' Find last column with data in row 1 (headers)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Create dictionary to store existing headers and their columns
    Set headerExists = CreateObject("Scripting.Dictionary")

    For i = 1 To lastCol
        headerVal = ws.Cells(1, i).Value
        If Not IsEmpty(headerVal) Then
            headerExists(headerVal) = i
        End If
    Next i

    ' Headers we want to ensure exist
    newHeaders = Array(11, "Taxable Value", "IGST", "CGST", "SGST")

    ' Insert columns for missing headers at the end, keep track of their column numbers
    For i = LBound(newHeaders) To UBound(newHeaders)
        If Not headerExists.exists(newHeaders(i)) Then
            ws.Cells(1, lastCol + 1).EntireColumn.Insert
            ws.Cells(1, lastCol + 1).Value = newHeaders(i)
            headerExists(newHeaders(i)) = lastCol + 1
            lastCol = lastCol + 1
        End If
    Next i

    ' Assign columns for each header
    col11 = headerExists(11)
    colTaxable = headerExists("Taxable Value")
    colIGST = headerExists("IGST")
    colCGST = headerExists("CGST")
    colSGST = headerExists("SGST")

    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Loop through each data row to compute sums and place in the respective columns
    For i = 2 To lastRow
        Dim sumTaxable As Double, sumIGST As Double, sumCGST As Double, sumSGST As Double
        Dim sumTotal As Double
        
        sumTaxable = 0
        sumIGST = 0
        sumCGST = 0
        sumSGST = 0
        sumTotal = 0
        
        For j = 1 To lastCol
            headerVal = ws.Cells(1, j).Value
            If IsNumeric(headerVal) Then
                If headerVal >= 30 And headerVal <= 45 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumTaxable = sumTaxable + ws.Cells(i, j).Value
                    End If
                ElseIf headerVal >= 46 And headerVal <= 55 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumIGST = sumIGST + ws.Cells(i, j).Value
                    End If
                ElseIf headerVal >= 56 And headerVal <= 65 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumCGST = sumCGST + ws.Cells(i, j).Value
                    End If
                ElseIf headerVal >= 66 And headerVal <= 75 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumSGST = sumSGST + ws.Cells(i, j).Value
                    End If
                End If
                
                ' Sum all in 30-75 range for header 11
                If headerVal >= 30 And headerVal <= 75 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumTotal = sumTotal + ws.Cells(i, j).Value
                    End If
                End If
            End If
        Next j
        
        ws.Cells(i, colTaxable).Value = sumTaxable
        ws.Cells(i, colIGST).Value = sumIGST
        ws.Cells(i, colCGST).Value = sumCGST
        ws.Cells(i, colSGST).Value = sumSGST
        ws.Cells(i, col11).Value = sumTotal
    Next i

    MsgBox "Headers 11, Taxable Value, IGST, CGST, SGST checked/added and sums computed."

    ' Refresh last row and last column after computations
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Verify required headers exist for processing
    If Not headerExists.exists(10) Or Not headerExists.exists(1) Or Not headerExists.exists(2) _
       Or Not headerExists.exists(4) Or Not headerExists.exists(8) Or Not headerExists.exists(11) Then
        MsgBox "One or more required headers (1,2,4,8,10,11) missing in source sheet.", vbCritical
        Exit Sub
    End If

    Dim colGSTIN As Long, colHeader1 As Long, colHeader2 As Long, colHeader4 As Long, colHeader8 As Long, colHeader11 As Long
    
    colGSTIN = headerExists(10)
    colHeader1 = headerExists(1)
    colHeader2 = headerExists(2)
    colHeader4 = headerExists(4) ' State Name column
    colHeader8 = headerExists(8)
    colHeader11 = headerExists(11)

    ' Create dictionary for state code and name prefix (code-name)
    Set stateDict = CreateObject("Scripting.Dictionary")
    With stateDict
        .Add "01", "01-Jammu & Kashmir"
        .Add "02", "02-Himachal Pradesh"
        .Add "03", "03-Punjab"
        .Add "04", "04-Chandigarh"
        .Add "05", "05-Uttarakhand"
        .Add "06", "06-Haryana"
        .Add "07", "07-Delhi"
        .Add "08", "08-Rajasthan"
        .Add "09", "09-Uttar Pradesh"
        .Add "10", "10-Bihar"
        .Add "11", "11-Sikkim"
        .Add "12", "12-Arunachal Pradesh"
        .Add "13", "13-Nagaland"
        .Add "14", "14-Manipur"
        .Add "15", "15-Mizoram"
        .Add "16", "16-Tripura"
        .Add "17", "17-Meghalaya"
        .Add "18", "18-Assam"
        .Add "19", "19-West Bengal"
        .Add "20", "20-Jharkhand"
        .Add "21", "21-Odisha"
        .Add "22", "22-Chhattisgarh"
        .Add "23", "23-Madhya Pradesh"
        .Add "24", "24-Gujarat"
        .Add "25", "25-Daman & Diu"
        .Add "26", "26-Dadra & Nagar Haveli"
        .Add "27", "27-Maharashtra"
        .Add "28", "28-Andhra Pradesh (Old)"
        .Add "29", "29-Karnataka"
        .Add "30", "30-Goa"
        .Add "31", "31-Lakshadweep"
        .Add "32", "32-Kerala"
        .Add "33", "33-Tamil Nadu"
        .Add "34", "34-Puducherry"
        .Add "35", "35-Andaman & Nicobar Islands"
        .Add "36", "36-Telangana"
        .Add "37", "37-Andhra Pradesh"
        .Add "38", "38-Ladakh"
        .Add "97", "97-Other Territory"
    End With

    destRowCredit = 9
    destRowUnreg = 9
    destRowTally = wsTallyB2C.Cells(wsTallyB2C.Rows.Count, 1).End(xlUp).Row
    If destRowTally < 9 Then destRowTally = 9 ' Ensure start at row 9 minimum

    For i = 2 To lastRow
        valGSTIN = Trim(ws.Cells(i, colGSTIN).Text)
        valInvoice = UCase(Trim(ws.Cells(i, colHeader8).Text))
        valGrossTotal = 0
        If IsNumeric(ws.Cells(i, colHeader11).Value) Then
            valGrossTotal = ws.Cells(i, colHeader11).Value
        End If

        ' Read GST components
        taxableVal = 0: igstVal = 0: cgstVal = 0: sgstVal = 0
        If colTaxable > 0 Then taxableVal = val(ws.Cells(i, colTaxable).Value)
        If colIGST > 0 Then igstVal = val(ws.Cells(i, colIGST).Value)
        If colCGST > 0 Then cgstVal = val(ws.Cells(i, colCGST).Value)
        If colSGST > 0 Then sgstVal = val(ws.Cells(i, colSGST).Value)

        ' Calculate GST %
        If taxableVal <> 0 Then
            gstPercent = Round((igstVal + cgstVal + sgstVal) / taxableVal * 100, 2)
            gstPercent = NearestGSTRate(gstPercent)
        Else
            gstPercent = ""
        End If

        If valGSTIN <> "" Then
            ' --- CREDIT Sheet copy ---
            wsCredit.Cells(destRowCredit, 1).Value = valGSTIN
            wsCredit.Cells(destRowCredit, 2).Value = ws.Cells(i, colHeader2).Value
            wsCredit.Cells(destRowCredit, 4).Value = ws.Cells(i, colHeader8).Value
            wsCredit.Cells(destRowCredit, 7).Value = ws.Cells(i, colHeader8).Value
            wsCredit.Cells(destRowCredit, 5).Value = ws.Cells(i, colHeader1).Value
            wsCredit.Cells(destRowCredit, 8).Value = ws.Cells(i, colHeader1).Value
            wsCredit.Cells(destRowCredit, 9).Value = ws.Cells(i, colHeader11).Value

            If colTaxable > 0 Then wsCredit.Cells(destRowCredit, 10).Value = taxableVal
            If colIGST > 0 Then wsCredit.Cells(destRowCredit, 14).Value = igstVal
            If colCGST > 0 Then wsCredit.Cells(destRowCredit, 15).Value = cgstVal
            If colSGST > 0 Then wsCredit.Cells(destRowCredit, 16).Value = sgstVal

            wsCredit.Cells(destRowCredit, 13).Value = gstPercent

            srcStateName = Trim(ws.Cells(i, colHeader4).Value)
            prefixStateCredit = ""
            Dim key As Variant
            For Each key In stateDict.Keys
                If InStr(1, stateDict(key), srcStateName, vbTextCompare) > 0 Then
                    prefixStateCredit = stateDict(key)
                    Exit For
                End If
            Next key

            wsCredit.Cells(destRowCredit, 6).Value = prefixStateCredit

            destRowCredit = destRowCredit + 1

        ElseIf (InStr(valInvoice, "EXP") > 0 Or InStr(valInvoice, "SEZ") > 0) Or valGrossTotal >= 100000 Then
            ' --- CREDIT UNREGISTER Sheet copy ---
            wsCreditUnreg.Cells(destRowUnreg, 1).Value = ws.Cells(i, colHeader2).Value
            wsCreditUnreg.Cells(destRowUnreg, 4).Value = ws.Cells(i, colHeader8).Value
            wsCreditUnreg.Cells(destRowUnreg, 8).Value = ws.Cells(i, colHeader8).Value
            wsCreditUnreg.Cells(destRowUnreg, 5).Value = ws.Cells(i, colHeader1).Value
            wsCreditUnreg.Cells(destRowUnreg, 9).Value = ws.Cells(i, colHeader1).Value
            wsCreditUnreg.Cells(destRowUnreg, 10).Value = ws.Cells(i, colHeader11).Value

            If colTaxable > 0 Then wsCreditUnreg.Cells(destRowUnreg, 11).Value = taxableVal
            If colIGST > 0 Then wsCreditUnreg.Cells(destRowUnreg, 14).Value = igstVal

            ' New condition for IGST and invoice number in Credit Unregister sheet
            If Left(valInvoice, 3) = "EXP" Or Left(valInvoice, 3) = "SEZ" Then
                If igstVal <> 0 Then
                    wsCreditUnreg.Cells(destRowUnreg, 2).Value = "Export With payment of GST"
                Else
                    wsCreditUnreg.Cells(destRowUnreg, 2).Value = "Export Without payment of GST"
                End If
            Else
                wsCreditUnreg.Cells(destRowUnreg, 2).Value = "B2CL"
            End If

            srcStateName = Trim(ws.Cells(i, colHeader4).Value)
            prefixStateUnreg = ""
            For Each key In stateDict.Keys
                If InStr(1, stateDict(key), srcStateName, vbTextCompare) > 0 Then
                    prefixStateUnreg = stateDict(key)
                    Exit For
                End If
            Next key

            wsCreditUnreg.Cells(destRowUnreg, 7).Value = prefixStateUnreg

            destRowUnreg = destRowUnreg + 1

        Else
            ' --- TALLY B2C Sheet copy for data not in Credit or Credit Unregister ---
            Dim tallyHeaders As Object
            Set tallyHeaders = CreateObject("Scripting.Dictionary")
            Dim c As Long, tallyLastCol As Long

            tallyLastCol = wsTallyB2C.Cells(1, wsTallyB2C.Columns.Count).End(xlToLeft).Column
            For c = 1 To tallyLastCol
                Dim hVal As Variant
                hVal = wsTallyB2C.Cells(1, c).Value
                If Not IsEmpty(hVal) Then tallyHeaders(hVal) = c
            Next c
            
            ' Ensure all needed headers exist, else skip this row in tally paste
            Dim requiredTallyHeaders As Variant
            requiredTallyHeaders = Array(1, 2, 4, 8, 11, 30, 46, 56, 66)
            Dim hasAllHeaders As Boolean
            hasAllHeaders = True
            Dim h As Variant
            For Each h In requiredTallyHeaders
                If Not tallyHeaders.exists(h) Then
                    hasAllHeaders = False
                    Exit For
                End If
            Next h
            If Not hasAllHeaders Then
                ' Skip this row or optionally notify user of missing headers
                GoTo NextIteration
            End If

            wsTallyB2C.Cells(destRowTally, tallyHeaders(8)).Value = ws.Cells(i, colHeader8).Value
            wsTallyB2C.Cells(destRowTally, tallyHeaders(1)).Value = ws.Cells(i, colHeader1).Value
            wsTallyB2C.Cells(destRowTally, tallyHeaders(2)).Value = ws.Cells(i, colHeader2).Value
            wsTallyB2C.Cells(destRowTally, tallyHeaders(4)).Value = ws.Cells(i, colHeader4).Value
            wsTallyB2C.Cells(destRowTally, tallyHeaders(11)).Value = -ws.Cells(i, colHeader11).Value
            
            If colTaxable > 0 Then
                wsTallyB2C.Cells(destRowTally, tallyHeaders(30)).Value = -ws.Cells(i, colTaxable).Value
            Else
                wsTallyB2C.Cells(destRowTally, tallyHeaders(30)).Value = 0
            End If
            
            If colIGST > 0 Then
                wsTallyB2C.Cells(destRowTally, tallyHeaders(46)).Value = -ws.Cells(i, colIGST).Value
            Else
                wsTallyB2C.Cells(destRowTally, tallyHeaders(46)).Value = 0
            End If
            
            If colSGST > 0 Then
                wsTallyB2C.Cells(destRowTally, tallyHeaders(56)).Value = -ws.Cells(i, colSGST).Value
            Else
                wsTallyB2C.Cells(destRowTally, tallyHeaders(56)).Value = 0
            End If
            
            If colCGST > 0 Then
                wsTallyB2C.Cells(destRowTally, tallyHeaders(66)).Value = -ws.Cells(i, colCGST).Value
            Else
                wsTallyB2C.Cells(destRowTally, tallyHeaders(66)).Value = 0
            End If

            destRowTally = destRowTally + 1

        End If
NextIteration:
    Next i

End Sub

Function NearestGSTRate(gstPercent As Double) As Double
    Dim rates As Variant
    Dim i As Integer
    Dim diff As Double
    Dim closestRate As Double
    Dim minDiff As Double
    
    rates = Array(0, 0.1, 0.25, 3, 5, 12, 18, 28)
    minDiff = 1000
    closestRate = 0
    
    For i = LBound(rates) To UBound(rates)
        diff = Abs(rates(i) - gstPercent)
        If diff < minDiff Then
            minDiff = diff
            closestRate = rates(i)
        End If
    Next i
    
    NearestGSTRate = closestRate
End Function

Macro from credit note register to hsn summary sheet:
Sub CopyToHSNSummary_B2B_B2C()
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim lastRowSrc As Long, lastRowDest As Long, i As Long
    Dim col As Long
    Dim taxable As Double, igst As Double, cgst As Double, sgst As Double, grossTotal As Double
    Dim rate As Double, invoiceType As String
    Dim gstin As String, invoiceNo As String, hsnCode As String
    Dim taxPresent As Boolean
    Dim headerRow As Long: headerRow = 1
    
    Set wsSrc = ThisWorkbook.Sheets("Credit Note Register")
    Set wsDest = ThisWorkbook.Sheets("HSN Summary")
    
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row + 1
    
    For i = 2 To lastRowSrc
        ' Read values from headers
        gstin = Trim(wsSrc.Cells(i, FindColumnByHeader(wsSrc, 10)).Value)
        invoiceNo = Trim(wsSrc.Cells(i, FindColumnByHeader(wsSrc, 8)).Value)
        hsnCode = Trim(wsSrc.Cells(i, FindColumnByHeader(wsSrc, 12)).Value)
        
        taxable = 0: igst = 0: cgst = 0: sgst = 0: grossTotal = 0
        taxPresent = False
        
        ' Sum taxable (headers 30–45)
        For col = 1 To wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
            If IsNumeric(wsSrc.Cells(headerRow, col).Value) Then
                Select Case CLng(wsSrc.Cells(headerRow, col).Value)
                    Case 30 To 45
                        taxable = taxable + val(wsSrc.Cells(i, col).Value)
                    Case 46 To 55
                        igst = igst + val(wsSrc.Cells(i, col).Value)
                    Case 56 To 65
                        cgst = cgst + val(wsSrc.Cells(i, col).Value)
                    Case 66 To 75
                        sgst = sgst + val(wsSrc.Cells(i, col).Value)
                End Select
                If CLng(wsSrc.Cells(headerRow, col).Value) >= 30 And CLng(wsSrc.Cells(headerRow, col).Value) <= 75 Then
                    grossTotal = grossTotal + val(wsSrc.Cells(i, col).Value)
                End If
            End If
        Next col
        
        If taxable <> 0 Or igst <> 0 Or cgst <> 0 Or sgst <> 0 Then taxPresent = True
        
        ' Decide type
        If gstin <> "" Then
            invoiceType = "B2B"
        ElseIf gstin = "" And Not (Left(invoiceNo, 3) = "EXP" Or Left(invoiceNo, 3) = "SEZ") And taxPresent And grossTotal < 100000 Then
            invoiceType = "B2C"
        Else
            invoiceType = ""
        End If
        
        ' Write to HSN Summary if valid
        If invoiceType <> "" And taxable <> 0 Then
            If taxable <> 0 Then
                rate = ((igst + cgst + sgst) / taxable) * 100
            Else
                rate = 0
            End If
            
            With wsDest
                .Cells(lastRowDest, 1).Value = invoiceType
                .Cells(lastRowDest, 2).Value = hsnCode
                .Cells(lastRowDest, 3).Value = -Abs(taxable)
                .Cells(lastRowDest, 4).Value = -Abs(igst)
                .Cells(lastRowDest, 5).Value = -Abs(cgst)
                .Cells(lastRowDest, 6).Value = -Abs(sgst)
                .Cells(lastRowDest, 7).Value = rate
                .Cells(lastRowDest, 8).Value = -Abs(grossTotal)
            End With
            
            lastRowDest = lastRowDest + 1
        End If
    Next i
End Sub

Function FindColumnByHeader(ws As Worksheet, headerNumber As Long) As Long
    Dim col As Long
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If Trim(ws.Cells(1, col).Value) = CStr(headerNumber) Then
            FindColumnByHeader = col
            Exit Function
        End If
    Next col
    FindColumnByHeader = 0
End Function

COMBINE MACRO OF BOTH:
Sub InsertComputeAndCopy()
    Dim ws As Worksheet, wsCredit As Worksheet, wsCreditUnreg As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim i As Long, j As Long
    Dim headerVal As Variant
    Dim sumVal As Double
    Dim col11 As Long, colTaxable As Long, colIGST As Long, colCGST As Long, colSGST As Long
    Dim headerExists As Object
    Dim newHeaders As Variant
    Dim destRowCredit As Long, destRowUnreg As Long
    Dim valGSTIN As String, valInvoice As String, valGrossTotal As Double
    Dim taxableVal As Double, igstVal As Double, cgstVal As Double, sgstVal As Double
    Dim gstPercent As Double
    Dim stateDict As Object
    Dim srcStateName As String, prefixStateUnreg As String
    Dim prefixStateCredit As String
    Dim gstinStateCode As String

    Set ws = ThisWorkbook.Sheets("Credit Note Register")
    Set wsCredit = ThisWorkbook.Sheets("Credit")
    Set wsCreditUnreg = ThisWorkbook.Sheets("Credit Unregister")

    ' Find last column with data in row 1 (headers)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Create dictionary to store existing headers and their columns
    Set headerExists = CreateObject("Scripting.Dictionary")

    For i = 1 To lastCol
        headerVal = ws.Cells(1, i).Value
        If Not IsEmpty(headerVal) Then
            headerExists(headerVal) = i
        End If
    Next i

    ' Headers we want to ensure exist
    newHeaders = Array(11, "Taxable Value", "IGST", "CGST", "SGST")

    ' Insert columns for missing headers at the end, keep track of their column numbers
    For i = LBound(newHeaders) To UBound(newHeaders)
        If Not headerExists.exists(newHeaders(i)) Then
            ws.Cells(1, lastCol + 1).EntireColumn.Insert
            ws.Cells(1, lastCol + 1).Value = newHeaders(i)
            headerExists(newHeaders(i)) = lastCol + 1
            lastCol = lastCol + 1
        End If
    Next i

    ' Assign columns for each header
    col11 = headerExists(11)
    colTaxable = headerExists("Taxable Value")
    colIGST = headerExists("IGST")
    colCGST = headerExists("CGST")
    colSGST = headerExists("SGST")

    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Loop through each data row to compute sums and place in the respective columns
    For i = 2 To lastRow
        Dim sumTaxable As Double, sumIGST As Double, sumCGST As Double, sumSGST As Double
        Dim sumTotal As Double
        
        sumTaxable = 0
        sumIGST = 0
        sumCGST = 0
        sumSGST = 0
        sumTotal = 0
        
        For j = 1 To lastCol
            headerVal = ws.Cells(1, j).Value
            If IsNumeric(headerVal) Then
                If headerVal >= 30 And headerVal <= 45 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumTaxable = sumTaxable + ws.Cells(i, j).Value
                    End If
                ElseIf headerVal >= 46 And headerVal <= 55 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumIGST = sumIGST + ws.Cells(i, j).Value
                    End If
                ElseIf headerVal >= 56 And headerVal <= 65 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumCGST = sumCGST + ws.Cells(i, j).Value
                    End If
                ElseIf headerVal >= 66 And headerVal <= 75 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumSGST = sumSGST + ws.Cells(i, j).Value
                    End If
                End If
                
                ' Sum all in 30-75 range for header 11
                If headerVal >= 30 And headerVal <= 75 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumTotal = sumTotal + ws.Cells(i, j).Value
                    End If
                End If
            End If
        Next j
        
        ws.Cells(i, colTaxable).Value = sumTaxable
        ws.Cells(i, colIGST).Value = sumIGST
        ws.Cells(i, colCGST).Value = sumCGST
        ws.Cells(i, colSGST).Value = sumSGST
        ws.Cells(i, col11).Value = sumTotal
    Next i

    MsgBox "Headers 11, Taxable Value, IGST, CGST, SGST checked/added and sums computed."

    ' Continue with CopyToCreditSheets logic
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Check required headers including column 4 (State) needed for prefix logic
    If Not headerExists.exists(10) Or Not headerExists.exists(1) Or Not headerExists.exists(2) _
       Or Not headerExists.exists(4) Or Not headerExists.exists(8) Or Not headerExists.exists(11) Then
        MsgBox "One or more required headers (1,2,4,8,10,11) missing in source sheet.", vbCritical
        Exit Sub
    End If

    colGSTIN = headerExists(10)
    colHeader1 = headerExists(1)
    colHeader2 = headerExists(2)
    colHeader4 = headerExists(4) ' State Name column
    colHeader8 = headerExists(8)
    colHeader11 = headerExists(11)

    ' Create dictionary for state code and name prefix (code-name)
    Set stateDict = CreateObject("Scripting.Dictionary")
    With stateDict
        .Add "01", "01-Jammu & Kashmir"
        .Add "02", "02-Himachal Pradesh"
        .Add "03", "03-Punjab"
        .Add "04", "04-Chandigarh"
        .Add "05", "05-Uttarakhand"
        .Add "06", "06-Haryana"
        .Add "07", "07-Delhi"
        .Add "08", "08-Rajasthan"
        .Add "09", "09-Uttar Pradesh"
        .Add "10", "10-Bihar"
        .Add "11", "11-Sikkim"
        .Add "12", "12-Arunachal Pradesh"
        .Add "13", "13-Nagaland"
        .Add "14", "14-Manipur"
        .Add "15", "15-Mizoram"
        .Add "16", "16-Tripura"
        .Add "17", "17-Meghalaya"
        .Add "18", "18-Assam"
        .Add "19", "19-West Bengal"
        .Add "20", "20-Jharkhand"
        .Add "21", "21-Odisha"
        .Add "22", "22-Chhattisgarh"
        .Add "23", "23-Madhya Pradesh"
        .Add "24", "24-Gujarat"
        .Add "25", "25-Daman & Diu"
        .Add "26", "26-Dadra & Nagar Haveli"
        .Add "27", "27-Maharashtra"
        .Add "28", "28-Andhra Pradesh (Old)"
        .Add "29", "29-Karnataka"
        .Add "30", "30-Goa"
        .Add "31", "31-Lakshadweep"
        .Add "32", "32-Kerala"
        .Add "33", "33-Tamil Nadu"
        .Add "34", "34-Puducherry"
        .Add "35", "35-Andaman & Nicobar Islands"
        .Add "36", "36-Telangana"
        .Add "37", "37-Andhra Pradesh"
        .Add "38", "38-Ladakh"
        .Add "97", "97-Other Territory"
    End With

    destRowCredit = 9
    destRowUnreg = 9

    For i = 2 To lastRow
        valGSTIN = Trim(ws.Cells(i, colGSTIN).Text)
        valInvoice = UCase(Trim(ws.Cells(i, colHeader8).Text))
        valGrossTotal = 0
        If IsNumeric(ws.Cells(i, colHeader11).Value) Then
            valGrossTotal = ws.Cells(i, colHeader11).Value
        End If

        ' Read GST components
        taxableVal = 0: igstVal = 0: cgstVal = 0: sgstVal = 0
        If colTaxable > 0 Then taxableVal = Val(ws.Cells(i, colTaxable).Value)
        If colIGST > 0 Then igstVal = Val(ws.Cells(i, colIGST).Value)
        If colCGST > 0 Then cgstVal = Val(ws.Cells(i, colCGST).Value)
        If colSGST > 0 Then sgstVal = Val(ws.Cells(i, colSGST).Value)

        ' Calculate GST %
        If taxableVal <> 0 Then
            gstPercent = Round((igstVal + cgstVal + sgstVal) / taxableVal * 100, 2)
            ' Round to nearest valid GST %
            gstPercent = NearestGSTRate(gstPercent)
        Else
            gstPercent = ""
        End If

        If valGSTIN <> "" Then
            ' --- CREDIT Sheet copy ---
            wsCredit.Cells(destRowCredit, 1).Value = valGSTIN
            wsCredit.Cells(destRowCredit, 2).Value = ws.Cells(i, colHeader2).Value
            wsCredit.Cells(destRowCredit, 4).Value = ws.Cells(i, colHeader8).Value
            wsCredit.Cells(destRowCredit, 7).Value = ws.Cells(i, colHeader8).Value
            wsCredit.Cells(destRowCredit, 5).Value = ws.Cells(i, colHeader1).Value
            wsCredit.Cells(destRowCredit, 8).Value = ws.Cells(i, colHeader1).Value
            wsCredit.Cells(destRowCredit, 9).Value = ws.Cells(i, colHeader11).Value

            If colTaxable > 0 Then wsCredit.Cells(destRowCredit, 10).Value = taxableVal
            If colIGST > 0 Then wsCredit.Cells(destRowCredit, 14).Value = igstVal
            If colCGST > 0 Then wsCredit.Cells(destRowCredit, 15).Value = cgstVal
            If colSGST > 0 Then wsCredit.Cells(destRowCredit, 16).Value = sgstVal

            wsCredit.Cells(destRowCredit, 13).Value = gstPercent

            ' Add state prefix for CREDIT sheet
            srcStateName = Trim(ws.Cells(i, colHeader4).Value)
            prefixStateCredit = ""
            ' Find matching state code by checking each entry for match
            Dim key As Variant
            For Each key In stateDict.keys
                ' Match state name case-insensitive with dictionary value after hyphen
                If InStr(1, stateDict(key), srcStateName, vbTextCompare) > 0 Then
                    prefixStateCredit = stateDict(key)
                    Exit For
                End If
            Next key

            wsCredit.Cells(destRowCredit, 6).Value = prefixStateCredit

            destRowCredit = destRowCredit + 1

        ElseIf (InStr(valInvoice, "EXP") > 0 Or InStr(valInvoice, "SEZ") > 0) Or valGrossTotal >= 100000 Then
            ' --- CREDIT UNREGISTER Sheet copy ---
            wsCreditUnreg.Cells(destRowUnreg, 1).Value = ws.Cells(i, colHeader2).Value
            wsCreditUnreg.Cells(destRowUnreg, 4).Value = ws.Cells(i, colHeader8).Value
            wsCreditUnreg.Cells(destRowUnreg, 8).Value = ws.Cells(i, colHeader8).Value
            wsCreditUnreg.Cells(destRowUnreg, 5).Value = ws.Cells(i, colHeader1).Value
            wsCreditUnreg.Cells(destRowUnreg, 9).Value = ws.Cells(i, colHeader1).Value
            wsCreditUnreg.Cells(destRowUnreg, 10).Value = ws.Cells(i, colHeader11).Value

            If colTaxable > 0 Then wsCreditUnreg.Cells(destRowUnreg, 11).Value = taxableVal
            If colIGST > 0 Then wsCreditUnreg.Cells(destRowUnreg, 14).Value = igstVal

            ' New condition for IGST and invoice number
            If Left(valInvoice, 3) = "EXP" Then
                If igstVal <> 0 Then
                    wsCreditUnreg.Cells(destRowUnreg, 2).Value = "Export With payment of GST"
                Else
                    wsCreditUnreg.Cells(destRowUnreg, 2).Value = "Export Without payment of GST"
                End If
            ElseIf Left(valInvoice, 3) = "SEZ" Then
                If igstVal <> 0 Then
                    wsCreditUnreg.Cells(destRowUnreg, 2).Value = "Export With payment of GST"
                Else
                    wsCreditUnreg.Cells(destRowUnreg, 2).Value = "Export Without payment of GST"
                End If
            Else
                wsCreditUnreg.Cells(destRowUnreg, 2).Value = "B2CL"
            End If

            wsCreditUnreg.Cells(destRowUnreg, 13).Value = gstPercent

            ' Add state prefix for CREDIT UNREGISTER sheet
            srcStateName = Trim(ws.Cells(i, colHeader4).Value)
            prefixStateUnreg = ""
            ' Find matching state code by checking each entry for match
            For Each key In stateDict.keys
                ' Match state name case-insensitive with dictionary value after hyphen
                If InStr(1, stateDict(key), srcStateName, vbTextCompare) > 0 Then
                    prefixStateUnreg = stateDict(key)
                    Exit For
                End If
            Next key

            wsCreditUnreg.Cells(destRowUnreg, 7).Value = prefixStateUnreg ' <-- Corrected column here

            destRowUnreg = destRowUnreg + 1

        End If
    Next i
End Sub

Function NearestGSTRate(gstPercent As Double) As Double
    Dim rates As Variant
    Dim i As Integer
    Dim diff As Double
    Dim closestRate As Double
    Dim minDiff As Double
    
    rates = Array(0, 0.1, 0.25, 3, 5, 12, 18, 28)
    minDiff = 1000
    closestRate = 0
    
    For i = LBound(rates) To UBound(rates)
        diff = Abs(rates(i) - gstPercent)
        If diff < minDiff Then
            minDiff = diff
            closestRate = rates(i)
        End If
    Next i
    
    NearestGSTRate = closestRate
End Function

1st code:
Sub InsertAndComputeHeadersUpdated()
    Dim ws As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim i As Long, j As Long
    Dim headerVal As Variant
    Dim sumVal As Double
    Dim col11 As Long, colTaxable As Long, colIGST As Long, colCGST As Long, colSGST As Long
    Dim headerExists As Object
    Dim newHeaders As Variant
    
    Set ws = ThisWorkbook.Sheets("Credit Note Register")
    
    ' Find last column with data in row 1 (headers)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Create dictionary to store existing headers and their columns
    Set headerExists = CreateObject("Scripting.Dictionary")
    
    For i = 1 To lastCol
        headerVal = ws.Cells(1, i).Value
        If Not IsEmpty(headerVal) Then
            headerExists(headerVal) = i
        End If
    Next i
    
    ' Headers we want to ensure exist
    newHeaders = Array(11, "Taxable Value", "IGST", "CGST", "SGST")
    
    ' Insert columns for missing headers at the end, keep track of their column numbers
    For i = LBound(newHeaders) To UBound(newHeaders)
        If Not headerExists.exists(newHeaders(i)) Then
            ws.Cells(1, lastCol + 1).EntireColumn.Insert
            ws.Cells(1, lastCol + 1).Value = newHeaders(i)
            headerExists(newHeaders(i)) = lastCol + 1
            lastCol = lastCol + 1
        End If
    Next i
    
    ' Assign columns for each header
    col11 = headerExists(11)
    colTaxable = headerExists("Taxable Value")
    colIGST = headerExists("IGST")
    colCGST = headerExists("CGST")
    colSGST = headerExists("SGST")
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each data row to compute sums and place in the respective columns
    For i = 2 To lastRow
        Dim sumTaxable As Double, sumIGST As Double, sumCGST As Double, sumSGST As Double
        Dim sumTotal As Double
        
        sumTaxable = 0
        sumIGST = 0
        sumCGST = 0
        sumSGST = 0
        sumTotal = 0
        
        For j = 1 To lastCol
            headerVal = ws.Cells(1, j).Value
            If IsNumeric(headerVal) Then
                If headerVal >= 30 And headerVal <= 45 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumTaxable = sumTaxable + ws.Cells(i, j).Value
                    End If
                ElseIf headerVal >= 46 And headerVal <= 55 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumIGST = sumIGST + ws.Cells(i, j).Value
                    End If
                ElseIf headerVal >= 56 And headerVal <= 65 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumCGST = sumCGST + ws.Cells(i, j).Value
                    End If
                ElseIf headerVal >= 66 And headerVal <= 75 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumSGST = sumSGST + ws.Cells(i, j).Value
                    End If
                End If
                
                ' Sum all in 30-75 range for header 11
                If headerVal >= 30 And headerVal <= 75 Then
                    If IsNumeric(ws.Cells(i, j).Value) Then
                        sumTotal = sumTotal + ws.Cells(i, j).Value
                    End If
                End If
            End If
        Next j
        
        ws.Cells(i, colTaxable).Value = sumTaxable
        ws.Cells(i, colIGST).Value = sumIGST
        ws.Cells(i, colCGST).Value = sumCGST
        ws.Cells(i, colSGST).Value = sumSGST
        ws.Cells(i, col11).Value = sumTotal
    Next i
    
    MsgBox "Headers 11, Taxable Value, IGST, CGST, SGST checked/added and sums computed."
End Sub
2nd code:
Sub CopyToCreditSheets()
    Dim wsSource As Worksheet, wsCredit As Worksheet, wsCreditUnreg As Worksheet
    Dim lastRow As Long, i As Long
    Dim colGSTIN As Long, colHeader1 As Long, colHeader2 As Long, colHeader4 As Long, colHeader8 As Long, colHeader11 As Long
    Dim colTaxable As Long, colIGST As Long, colCGST As Long, colSGST As Long
    Dim destRowCredit As Long, destRowUnreg As Long
    Dim valGSTIN As String, valInvoice As String, valGrossTotal As Double
    Dim headerExists As Object
    Dim lastCol As Long, j As Long, headerVal As Variant
    Dim taxableVal As Double, igstVal As Double, cgstVal As Double, sgstVal As Double
    Dim gstPercent As Double
    Dim stateDict As Object
    Dim srcStateName As String, prefixStateUnreg As String
    Dim prefixStateCredit As String
    Dim gstinStateCode As String

    Set wsSource = ThisWorkbook.Sheets("Credit Note Register")
    Set wsCredit = ThisWorkbook.Sheets("Credit")
    Set wsCreditUnreg = ThisWorkbook.Sheets("Credit Unregister")

    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    Set headerExists = CreateObject("Scripting.Dictionary")
    For j = 1 To lastCol
        headerVal = wsSource.Cells(1, j).Value
        If Not IsEmpty(headerVal) Then headerExists(headerVal) = j
    Next j

    ' Check required headers including column 4 (State) needed for prefix logic
    If Not headerExists.exists(10) Or Not headerExists.exists(1) Or Not headerExists.exists(2) _
       Or Not headerExists.exists(4) Or Not headerExists.exists(8) Or Not headerExists.exists(11) Then
        MsgBox "One or more required headers (1,2,4,8,10,11) missing in source sheet.", vbCritical
        Exit Sub
    End If

    colGSTIN = headerExists(10)
    colHeader1 = headerExists(1)
    colHeader2 = headerExists(2)
    colHeader4 = headerExists(4) ' State Name column
    colHeader8 = headerExists(8)
    colHeader11 = headerExists(11)

    colTaxable = 0: colIGST = 0: colCGST = 0: colSGST = 0
    If headerExists.exists("Taxable Value") Then colTaxable = headerExists("Taxable Value")
    If headerExists.exists("IGST") Then colIGST = headerExists("IGST")
    If headerExists.exists("CGST") Then colCGST = headerExists("CGST")
    If headerExists.exists("SGST") Then colSGST = headerExists("SGST")

    ' Create dictionary for state code and name prefix (code-name)
    Set stateDict = CreateObject("Scripting.Dictionary")
    With stateDict
        .Add "01", "01-Jammu & Kashmir"
        .Add "02", "02-Himachal Pradesh"
        .Add "03", "03-Punjab"
        .Add "04", "04-Chandigarh"
        .Add "05", "05-Uttarakhand"
        .Add "06", "06-Haryana"
        .Add "07", "07-Delhi"
        .Add "08", "08-Rajasthan"
        .Add "09", "09-Uttar Pradesh"
        .Add "10", "10-Bihar"
        .Add "11", "11-Sikkim"
        .Add "12", "12-Arunachal Pradesh"
        .Add "13", "13-Nagaland"
        .Add "14", "14-Manipur"
        .Add "15", "15-Mizoram"
        .Add "16", "16-Tripura"
        .Add "17", "17-Meghalaya"
        .Add "18", "18-Assam"
        .Add "19", "19-West Bengal"
        .Add "20", "20-Jharkhand"
        .Add "21", "21-Odisha"
        .Add "22", "22-Chhattisgarh"
        .Add "23", "23-Madhya Pradesh"
        .Add "24", "24-Gujarat"
        .Add "25", "25-Daman & Diu"
        .Add "26", "26-Dadra & Nagar Haveli"
        .Add "27", "27-Maharashtra"
        .Add "28", "28-Andhra Pradesh (Old)"
        .Add "29", "29-Karnataka"
        .Add "30", "30-Goa"
        .Add "31", "31-Lakshadweep"
        .Add "32", "32-Kerala"
        .Add "33", "33-Tamil Nadu"
        .Add "34", "34-Puducherry"
        .Add "35", "35-Andaman & Nicobar Islands"
        .Add "36", "36-Telangana"
        .Add "37", "37-Andhra Pradesh"
        .Add "38", "38-Ladakh"
        .Add "97", "97-Other Territory"
    End With

    destRowCredit = 9
    destRowUnreg = 9

    For i = 2 To lastRow
        valGSTIN = Trim(wsSource.Cells(i, colGSTIN).Text)
        valInvoice = UCase(Trim(wsSource.Cells(i, colHeader8).Text))
        valGrossTotal = 0
        If IsNumeric(wsSource.Cells(i, colHeader11).Value) Then
            valGrossTotal = wsSource.Cells(i, colHeader11).Value
        End If

        ' Read GST components
        taxableVal = 0: igstVal = 0: cgstVal = 0: sgstVal = 0
        If colTaxable > 0 Then taxableVal = val(wsSource.Cells(i, colTaxable).Value)
        If colIGST > 0 Then igstVal = val(wsSource.Cells(i, colIGST).Value)
        If colCGST > 0 Then cgstVal = val(wsSource.Cells(i, colCGST).Value)
        If colSGST > 0 Then sgstVal = val(wsSource.Cells(i, colSGST).Value)

        ' Calculate GST %
        If taxableVal <> 0 Then
            gstPercent = Round((igstVal + cgstVal + sgstVal) / taxableVal * 100, 2)
            ' Round to nearest valid GST %
            gstPercent = NearestGSTRate(gstPercent)
        Else
            gstPercent = ""
        End If

        If valGSTIN <> "" Then
            ' --- CREDIT Sheet copy ---
            wsCredit.Cells(destRowCredit, 1).Value = valGSTIN
            wsCredit.Cells(destRowCredit, 2).Value = wsSource.Cells(i, colHeader2).Value
            wsCredit.Cells(destRowCredit, 4).Value = wsSource.Cells(i, colHeader8).Value
            wsCredit.Cells(destRowCredit, 7).Value = wsSource.Cells(i, colHeader8).Value
            wsCredit.Cells(destRowCredit, 5).Value = wsSource.Cells(i, colHeader1).Value
            wsCredit.Cells(destRowCredit, 8).Value = wsSource.Cells(i, colHeader1).Value
            wsCredit.Cells(destRowCredit, 9).Value = wsSource.Cells(i, colHeader11).Value

            If colTaxable > 0 Then wsCredit.Cells(destRowCredit, 10).Value = taxableVal
            If colIGST > 0 Then wsCredit.Cells(destRowCredit, 14).Value = igstVal
            If colCGST > 0 Then wsCredit.Cells(destRowCredit, 15).Value = cgstVal
            If colSGST > 0 Then wsCredit.Cells(destRowCredit, 16).Value = sgstVal

            wsCredit.Cells(destRowCredit, 13).Value = gstPercent

            ' Add state prefix for CREDIT sheet
            srcStateName = Trim(wsSource.Cells(i, colHeader4).Value)
            prefixStateCredit = ""
            ' Find matching state code by checking each entry for match
            Dim key As Variant
            For Each key In stateDict.keys
                ' Match state name case-insensitive with dictionary value after hyphen
                If InStr(1, stateDict(key), srcStateName, vbTextCompare) > 0 Then
                    prefixStateCredit = stateDict(key)
                    Exit For
                End If
            Next key

            wsCredit.Cells(destRowCredit, 6).Value = prefixStateCredit

            destRowCredit = destRowCredit + 1

        ElseIf (InStr(valInvoice, "EXP") > 0 Or InStr(valInvoice, "SEZ") > 0) Or valGrossTotal >= 100000 Then
            ' --- CREDIT UNREGISTER Sheet copy ---
            wsCreditUnreg.Cells(destRowUnreg, 1).Value = wsSource.Cells(i, colHeader2).Value
            wsCreditUnreg.Cells(destRowUnreg, 4).Value = wsSource.Cells(i, colHeader8).Value
            wsCreditUnreg.Cells(destRowUnreg, 8).Value = wsSource.Cells(i, colHeader8).Value
            wsCreditUnreg.Cells(destRowUnreg, 5).Value = wsSource.Cells(i, colHeader1).Value
            wsCreditUnreg.Cells(destRowUnreg, 9).Value = wsSource.Cells(i, colHeader1).Value
            wsCreditUnreg.Cells(destRowUnreg, 10).Value = wsSource.Cells(i, colHeader11).Value

            If colTaxable > 0 Then wsCreditUnreg.Cells(destRowUnreg, 11).Value = taxableVal
            If colIGST > 0 Then wsCreditUnreg.Cells(destRowUnreg, 14).Value = igstVal

            wsCreditUnreg.Cells(destRowUnreg, 13).Value = gstPercent

            ' Add state prefix for CREDIT UNREGISTER sheet
            srcStateName = Trim(wsSource.Cells(i, colHeader4).Value)
            prefixStateUnreg = ""
            ' Find matching state code by checking each entry for match
            For Each key In stateDict.keys
                ' Match state name case-insensitive with dictionary value after hyphen
                If InStr(1, stateDict(key), srcStateName, vbTextCompare) > 0 Then
                    prefixStateUnreg = stateDict(key)
                    Exit For
                End If
            Next key

            wsCreditUnreg.Cells(destRowUnreg, 7).Value = prefixStateUnreg ' <-- Corrected column here

            destRowUnreg = destRowUnreg + 1

        End If
    Next i
End Sub

Function NearestGSTRate(gstPercent As Double) As Double
    Dim rates As Variant
    Dim i As Integer
    Dim diff As Double
    Dim closestRate As Double
    Dim minDiff As Double
    
    rates = Array(0, 0.1, 0.25, 3, 5, 12, 18, 28)
    minDiff = 1000
    closestRate = 0
    
    For i = LBound(rates) To UBound(rates)
        diff = Abs(rates(i) - gstPercent)
        If diff < minDiff Then
            minDiff = diff
            closestRate = rates(i)
        End If
    Next i
    
    NearestGSTRate = closestRate
End Function
Macro for HSN Summary SHEET1 TO SARAL HSN Summary :

Sub Copy_HSN_Summary_Mapped()
    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim lastSrcRow As Long, destRow As Long
    Dim i As Long
    
    Set wb = ThisWorkbook
    Set wsSrc = wb.Sheets("HSN Summary Sheet1")   ' Source sheet
    Set wsDest = wb.Sheets("HSN Summary")         ' Destination sheet
    
    ' Get the last row in the source sheet
    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    
    ' Set initial destination row (start from row 9)
    destRow = 9
    
    ' Loop through source data starting from row 2
    For i = 2 To lastSrcRow
        wsDest.Cells(destRow, 2).Value = wsSrc.Cells(i, 1).Value   ' 1st ? 2nd col
        wsDest.Cells(destRow, 3).Value = wsSrc.Cells(i, 2).Value   ' 2nd ? 3rd col
        wsDest.Cells(destRow, 9).Value = wsSrc.Cells(i, 3).Value   ' 3rd ? 9th col
        wsDest.Cells(destRow, 11).Value = wsSrc.Cells(i, 4).Value  ' 4th ? 11th col
        wsDest.Cells(destRow, 12).Value = wsSrc.Cells(i, 5).Value  ' 5th ? 12th col
        wsDest.Cells(destRow, 13).Value = wsSrc.Cells(i, 6).Value  ' 6th ? 13th col
        wsDest.Cells(destRow, 10).Value = wsSrc.Cells(i, 7).Value  ' 7th ? 10th col
        wsDest.Cells(destRow, 8).Value = wsSrc.Cells(i, 8).Value   ' 8th ? 8th col
        
        ' 9th column ? 7th col (optional; skip if blank)
        If Trim(wsSrc.Cells(i, 9).Value) <> "" Then
            wsDest.Cells(destRow, 7).Value = wsSrc.Cells(i, 9).Value
        End If
        
        destRow = destRow + 1
    Next i
    
    MsgBox "HSN Summary mapping completed.", vbInformation
End Sub
