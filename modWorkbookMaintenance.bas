Option Explicit

Public Function EnsureSheet(ByVal sheetName As String, Optional ByVal afterSheetName As String = "") As Worksheet
    Dim ws As Worksheet
    Dim afterWs As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        If Len(Trim(afterSheetName)) > 0 Then
            On Error Resume Next
            Set afterWs = ThisWorkbook.Worksheets(afterSheetName)
            On Error GoTo 0
        End If

        If afterWs Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=afterWs)
        End If

        ws.Name = sheetName
    End If

    Set EnsureSheet = ws
End Function

Public Function DeleteSheetIfExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        DeleteSheetIfExists = False
        Exit Function
    End If

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    DeleteSheetIfExists = True
End Function

Public Function EnsureListObject(ByVal ws As Worksheet, ByVal listName As String, ByVal sourceRange As Range, Optional ByVal hasHeaders As XlYesNoGuess = xlYes) As ListObject
    Dim tbl As ListObject

    If ws Is Nothing Then
        Set EnsureListObject = Nothing
        Exit Function
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects(listName)
    On Error GoTo 0

    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=sourceRange, XlListObjectHasHeaders:=hasHeaders)
        tbl.Name = listName
    End If

    Set EnsureListObject = tbl
End Function

Public Function ClearRange(ByVal targetRange As Range) As Boolean
    If targetRange Is Nothing Then
        ClearRange = False
        Exit Function
    End If

    targetRange.Clear
    ClearRange = True
End Function

Public Function ResetPrintSheet(ByVal sheetName As String) As Worksheet
    Dim wsPrint As Worksheet

    DeleteSheetIfExists sheetName
    Set wsPrint = EnsureSheet(sheetName)
    wsPrint.Cells.Clear

    wsPrint.Range("A1").Value = "Table"
    wsPrint.Range("B1").Value = "#"
    wsPrint.Range("C1").Value = "Start Date"
    wsPrint.Range("D1").Value = "Start Time"
    wsPrint.Range("E1").Value = "End Date"
    wsPrint.Range("F1").Value = "Description"
    wsPrint.Range("G1").Value = "Notes"
    wsPrint.Range("H1").Value = "Link"

    Set ResetPrintSheet = wsPrint
End Function

Public Function EnsureOutputWorkbook(ByVal workbookName As String) As Workbook
    Dim wb As Workbook
    Dim isFound As Boolean

    isFound = False

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            Set EnsureOutputWorkbook = wb
            isFound = True
            Exit Function
        End If
    Next wb

    If Not isFound Then
        Set EnsureOutputWorkbook = Workbooks.Add
        EnsureOutputWorkbook.SaveAs Filename:=workbookName
    End If
End Function

Public Function IsSheetPresent(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    IsSheetPresent = Not (ws Is Nothing)
End Function

Public Function RemoveTempOutputSheet(ByVal sheetName As String) As Boolean
    RemoveTempOutputSheet = DeleteSheetIfExists(sheetName)
End Function

Public Function SafeSaveWorkbook(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then
        SafeSaveWorkbook = False
        Exit Function
    End If

    wb.Save
    SafeSaveWorkbook = True
End Function
