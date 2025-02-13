Attribute VB_Name = "Module1"
Public Sub ProcessExclusionsAndTotals(exclusionFilePath As String)
    
    ' Run the exclusions process
    HandleExclusions exclusionFilePath

    ' Run the totals process
    CalculateTotals exclusionFilePath

    SortByColumns Array("Symbol") ' Update with the column names to sort by

    ' MsgBox "Exclusions processed and totals calculated.", vbInformation, "Task Complete"
End Sub


Sub HandleExclusions(exclusionsFilePath As String)
    Dim ws As Worksheet
    Dim exclusionsWorkbook As Workbook
    Dim exclusionsWs As Worksheet
    Dim exclusions() As Variant
    Dim symbolColIndex As Variant
    Dim cellValue As String
    Dim lastRow As Long
    Dim i As Long

    ' Use the active worksheet in the active workbook
    Set ws = ActiveWorkbook.ActiveSheet

    ' Open the exclusions workbook
    Set exclusionsWorkbook = Workbooks.Open(exclusionsFilePath)
    Set exclusionsWs = exclusionsWorkbook.Sheets("Exclusions") ' Update with the correct sheet name

    ' Read the exclusion symbols
    exclusions = exclusionsWs.Range("A1", exclusionsWs.Cells(exclusionsWs.Rows.Count, "A").End(xlUp)).Value
    exclusionsWorkbook.Close SaveChanges:=False

    ' Find the column named "symbol"
    symbolColIndex = Application.Match("symbol", ws.Rows(1), 0)
    If IsError(symbolColIndex) Then
        MsgBox "Column 'symbol' not found.", vbCritical
        Exit Sub
    End If

    ' Delete rows that match exclusions
    lastRow = ws.Cells(ws.Rows.Count, symbolColIndex).End(xlUp).Row
    For i = lastRow To 2 Step -1
        cellValue = ws.Cells(i, symbolColIndex).Value
        If Not IsError(Application.Match(cellValue, exclusions, 0)) Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Sub CalculateTotals(exclusionsFilePath As String)
    Dim ws As Worksheet
    Dim exclusionsWorkbook As Workbook
    Dim columnsWs As Worksheet
    Dim columnsToSum() As Variant
    Dim colIndex As Variant
    Dim total As Double
    Dim lastRow As Long
    Dim i As Long

    Set ws = ActiveWorkbook.ActiveSheet
    Set exclusionsWorkbook = Workbooks.Open(exclusionsFilePath)
    Set columnsWs = exclusionsWorkbook.Sheets("ColumnsToSum")

    ' Read column names to sum
    columnsToSum = columnsWs.Range("A1", columnsWs.Cells(columnsWs.Rows.Count, "A").End(xlUp)).Value
    exclusionsWorkbook.Close SaveChanges:=False

    ' Sum specified columns and place totals
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = LBound(columnsToSum, 1) To UBound(columnsToSum, 1)
        colIndex = Application.Match(columnsToSum(i, 1), ws.Rows(1), 0)
        If Not IsError(colIndex) Then
            total = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex)))
            ws.Cells(lastRow + 1, colIndex).Value = total
            ws.Cells(lastRow + 1, colIndex).Font.Bold = True
        End If
    Next i
End Sub

Sub SortByColumns(sortFields As Variant)
    Dim ws As Worksheet
    Dim sortRange As Range
    Dim colIndex As Variant
    Dim i As Long
    Dim lastRow As Long

    Set ws = ActiveWorkbook.ActiveSheet

    ' Determine the range to sort
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set sortRange = ws.Range("A1", ws.Cells(lastRow, ws.UsedRange.Columns.Count))

    ' Clear existing sort fields
    ws.Sort.sortFields.Clear

    ' Add sort fields based on column names
    For i = LBound(sortFields) To UBound(sortFields)
        colIndex = Application.Match(sortFields(i), ws.Rows(1), 0) ' Find column index by name
        If Not IsError(colIndex) Then
            ws.Sort.sortFields.Add Key:=ws.Cells(2, colIndex), Order:=xlAscending
        Else
            MsgBox "Column '" & sortFields(i) & "' not found.", vbExclamation
        End If
    Next i

    ' Perform the sort
    With ws.Sort
        .SetRange sortRange
        .Header = xlYes
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

