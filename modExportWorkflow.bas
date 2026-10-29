Option Explicit

Public Function BuildExportRows(ByVal results As Variant, Optional ByVal folderRoot As String = "") As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim arr() As Variant
    Dim itemCount As Long
    Dim rowValue As String

    If IsEmpty(results) Then
        BuildExportRows = Empty
        Exit Function
    End If

    If Not IsArray(results) Then
        BuildExportRows = Empty
        Exit Function
    End If

    itemCount = UBound(results, 1) - LBound(results, 1) + 1
    ReDim arr(0 To itemCount - 1, 0 To 7)

    For i = LBound(results, 1) To UBound(results, 1)
        arr(i - LBound(results, 1), 0) = NormalizeText(results(i, 0))
        arr(i - LBound(results, 1), 1) = NormalizeText(results(i, 1))
        arr(i - LBound(results, 1), 2) = SafeFormatDateValue(results(i, 2))
        arr(i - LBound(results, 1), 3) = SafeFormatTimeValue(results(i, 3))
        arr(i - LBound(results, 1), 4) = SafeFormatDateValue(results(i, 4))
        arr(i - LBound(results, 1), 5) = NormalizeText(results(i, 5))
        arr(i - LBound(results, 1), 6) = NormalizeText(results(i, 6))
        rowValue = NormalizeText(results(i, 8))
        If Len(rowValue) > 0 Then
            rowValue = NormalizeFolderLink(rowValue, folderRoot)
        End If
        arr(i - LBound(results, 1), 7) = rowValue
    Next i

    BuildExportRows = arr
End Function

Public Sub WriteExportRowsToSheet(ByVal ws As Worksheet, ByVal exportRows As Variant)
    Dim rowIndex As Long
    Dim i As Long
    Dim rowCount As Long

    If ws Is Nothing Then Exit Sub
    If IsEmpty(exportRows) Then Exit Sub

    rowCount = UBound(exportRows, 1) - LBound(exportRows, 1) + 1

    For i = LBound(exportRows, 1) To UBound(exportRows, 1)
        rowIndex = i + 2
        ws.Cells(rowIndex, 1).Value = exportRows(i, 0)
        ws.Cells(rowIndex, 2).Value = exportRows(i, 1)
        ws.Cells(rowIndex, 3).Value = exportRows(i, 2)
        ws.Cells(rowIndex, 4).Value = exportRows(i, 3)
        ws.Cells(rowIndex, 5).Value = exportRows(i, 4)
        ws.Cells(rowIndex, 6).Value = exportRows(i, 5)
        ws.Cells(rowIndex, 7).Value = exportRows(i, 6)
        ws.Cells(rowIndex, 8).Value = exportRows(i, 7)
    Next i

    ws.Columns.AutoFit
End Sub

Public Function PrepareExportSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    DeleteSheetIfExists sheetName
    Set ws = EnsureSheet(sheetName)
    ws.Cells.Clear

    ws.Range("A1").Value = "Table"
    ws.Range("B1").Value = "#"
    ws.Range("C1").Value = "Start Date"
    ws.Range("D1").Value = "Start Time"
    ws.Range("E1").Value = "End Date"
    ws.Range("F1").Value = "Description"
    ws.Range("G1").Value = "Notes"
    ws.Range("H1").Value = "Link"

    Set PrepareExportSheet = ws
End Function

Public Sub AddHyperlinksToExportSheet(ByVal ws As Worksheet)
    Dim rowIndex As Long
    Dim linkText As String

    If ws Is Nothing Then Exit Sub

    For rowIndex = 2 To ws.UsedRange.Rows.Count
        linkText = NormalizeText(ws.Cells(rowIndex, 8).Value)
        If Len(linkText) > 0 Then
            AddHyperlinkToCell ws.Cells(rowIndex, 8), linkText, "Link"
        End If
    Next rowIndex
End Sub

Public Function BuildOutputWorkbook(ByVal workbookName As String) As Workbook
    Dim wb As Workbook
    Dim fileNameValue As String

    fileNameValue = workbookName
    If InStr(1, fileNameValue, ".xlsx", vbTextCompare) = 0 Then
        fileNameValue = fileNameValue & ".xlsx"
    End If

    On Error Resume Next
    Set wb = Workbooks(fileNameValue)
    On Error GoTo 0

    If wb Is Nothing Then
        Set wb = Workbooks.Add
        wb.SaveAs Filename:=fileNameValue
    End If

    Set BuildOutputWorkbook = wb
End Function

Public Sub ExportResultsToWorkbook(ByVal results As Variant, Optional ByVal sheetName As String = "For Printing", Optional ByVal workbookName As String = "Export", Optional ByVal folderRoot As String = "")
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim exportRows As Variant

    If IsEmpty(results) Then Exit Sub

    exportRows = BuildExportRows(results, folderRoot)
    If IsEmpty(exportRows) Then Exit Sub

    Set ws = PrepareExportSheet(sheetName)
    WriteExportRowsToSheet ws, exportRows
    AddHyperlinksToExportSheet ws

    Set wb = BuildOutputWorkbook(workbookName)
    ws.Copy Before:=wb.Worksheets(1)
    SafeSaveWorkbook wb
End Sub

Public Sub PrepareAndExportWorksheet(ByVal results As Variant, Optional ByVal sheetName As String = "For Printing", Optional ByVal workbookName As String = "Export", Optional ByVal folderRoot As String = "")
    ExportResultsToWorkbook results, sheetName, workbookName, folderRoot
End Sub
