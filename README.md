Option Explicit

Public Function EnsureFolder(ByVal folderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(Trim(folderPath)) = 0 Then
        EnsureFolder = ""
        Exit Function
    End If

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

    EnsureFolder = folderPath
End Function

Public Function BuildFolderPath(ByVal basePath As String, ByVal folderName As String) As String
    Dim cleanBase As String
    Dim cleanFolderName As String

    cleanBase = Trim(basePath)
    cleanFolderName = Trim(folderName)

    If Len(cleanBase) = 0 Then
        BuildFolderPath = cleanFolderName
        Exit Function
    End If

    If Right$(cleanBase, 1) = "\" Then
        BuildFolderPath = cleanBase & cleanFolderName
    Else
        BuildFolderPath = cleanBase & "\" & cleanFolderName
    End If
End Function

Public Sub AddHyperlinksToLinkColumn(ByVal tbl As ListObject)
    Dim rowIndex As Long
    Dim linkValue As String
    Dim rowRange As Range

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To tbl.DataBodyRange.Rows.Count
        Set rowRange = tbl.DataBodyRange.Rows(rowIndex)
        linkValue = NormalizeText(rowRange.Cells(1, RecordColumn.ColLink).Value)

        If Len(linkValue) > 0 Then
            rowRange.Cells(1, RecordColumn.ColLink).Hyperlinks.Delete
            rowRange.Cells(1, RecordColumn.ColLink).Hyperlinks.Add Anchor:=rowRange.Cells(1, RecordColumn.ColLink), Address:=linkValue, TextToDisplay:="Link"
        End If
    Next rowIndex
End Sub

Public Sub ExportSelectedRowsToPrintSheet(ByVal results As Variant)
    Dim wsPrint As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim i As Long

    Set wsPrint = GetWorksheet(SHEET_PRINT)
    wsPrint.Cells.Clear

    wsPrint.Range("A1").Value = "Table"
    wsPrint.Range("B1").Value = "#"
    wsPrint.Range("C1").Value = "Start Date"
    wsPrint.Range("D1").Value = "Start Time"
    wsPrint.Range("E1").Value = "End Date"
    wsPrint.Range("F1").Value = "Description"
    wsPrint.Range("G1").Value = "Notes"
    wsPrint.Range("H1").Value = "Link"

    If IsEmpty(results) Then Exit Sub

    For i = LBound(results, 1) To UBound(results, 1)
        rowIndex = i + 2
        wsPrint.Cells(rowIndex, 1).Value = results(i, 0)
        wsPrint.Cells(rowIndex, 2).Value = results(i, 1)
        wsPrint.Cells(rowIndex, 3).Value = results(i, 2)
        wsPrint.Cells(rowIndex, 4).Value = results(i, 3)
        wsPrint.Cells(rowIndex, 5).Value = results(i, 4)
        wsPrint.Cells(rowIndex, 6).Value = results(i, 5)
        wsPrint.Cells(rowIndex, 7).Value = results(i, 6)
        wsPrint.Cells(rowIndex, 8).Value = results(i, 7)
    Next i

    wsPrint.Columns.AutoFit
End Sub
