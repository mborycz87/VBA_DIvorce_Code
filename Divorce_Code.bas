Sub ProtectSheets()
Dim wsheet As Worksheet
For Each wsheet In ActiveWorkbook.Worksheets
wsheet.Protect Password:="password"
Next wsheet
End Sub

Sub UnprotectSheets()
Dim wsheet As Worksheet
For Each wsheet In ActiveWorkbook.Worksheets
wsheet.Unprotect Password:="password"
Next wsheet
End Sub

Sub StartButton()
Dim i As Integer
Dim j As Integer
Dim x As Integer

Call UnprotectSheets

If Worksheets("Master").Range("O10") <> "" Then
    For i = 3 To ActiveWorkbook.Sheets.Count
        Sheets(i).Visible = True
        Sheets(i).Activate
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Select
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                If ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 9) <> "" Then
                    Str1 = Replace(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 9), Worksheets("Master").Range("O8"), Worksheets("Master").Range("O10"))
                    ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 9) = Str1
                End If
            Next x
        Next j
    Next i
    Worksheets("Master").Range("O8") = Worksheets("Master").Range("O10")
    Worksheets("Master").Range("O10") = ""
End If
With uf_Master
    .Height = 444
    .Width = 517
End With
'Call SortListBoxAlphabetically(uf_Master.cbx_ShCheck)
'Call SortListBoxAlphabetically(uf_Master.cbx_ShDelete)
'Call SortListBoxAlphabetically(uf_Master.cbx_RefineShCheck)
uf_Master.Show

Call ProtectSheets
End Sub

Sub BoxForDataEntry(ByRef listName As MSForms.TextBox, ByVal Str As String, ByVal Num As Integer)
Set table = ActiveWorkbook.ActiveSheet.ListObjects(uf_Master.cbx_TblCheck.Text)
indRow = 0
anSwer = MsgBox(Str & ":" & vbNewLine & listName.Text _
& vbNewLine & vbNewLine & "Would you like to change this?", vbYesNo, Str)

If anSwer = vbYes Then
    inptAns = InputBox("Type into field to change, or leave blank to delete", "Change or delete field")
    If StrPtr(inptAns) = 0 Then
        Exit Sub
    ElseIf StrPtr(inptAns) <> 0 Then
         anSwer = inptAns
    Do Until table.ListColumns(7).DataBodyRange(indRow, 1).Value = uf_Master.tb_Notes.Text
        indRow = indRow + 1
    Loop
    table.DataBodyRange(indRow, Num).Value = anSwer
    listName.Text = anSwer
    End If
    If IsNumeric(anSwer) Then
        anSwer = "'" & anSwer
    End If
ElseIf anSwer = vbNo Then
    Exit Sub
End If
 If anSwer = "" Then
    anSwer = "Enter " & Str & " when known"
    table.DataBodyRange(indRow, Num).Value = anSwer
    listName.Text = anSwer
 Else:
    table.DataBodyRange(indRow, Num).Value = anSwer
    listName.Text = anSwer
End If
End Sub

Sub GreenDataEntry(ByVal Str As String, ByVal Num As Integer, ByRef listName As MSForms.TextBox, ByRef table As ListObject)

If listName.Text <> "" Then
    If table.ListColumns(Num).DataBodyRange(table.ListRows.Count, 1) <> "" Then
        table.ListRows.Add
    End If
    'ActiveSheet.Unprotect Password:="mEEKAbUTTICH@!21"
    table.ListColumns(Num).DataBodyRange(table.ListRows.Count, 1).Select
    ActiveCell.Value = listName.Text
    If IsNumeric(ActiveCell.Value) Then
        ActiveCell.Value = "'" & ActiveCell.Value
    End If
ElseIf listName.Text = "" Then
    If table.ListColumns(Num).DataBodyRange(table.ListRows.Count, 1) <> "" Then
        table.ListRows.Add
    End If
    cellNum = 1
    indRow = 1
    table.ListColumns(Num).DataBodyRange(table.ListRows.Count, 1).Select
    ActiveCell.Value = "Enter " & Str & " when known"
   ' listName.Text = ActiveCell.Value
    For Each cell In table.ListColumns(Num).DataBodyRange
        If cell.Value Like "*" & "Enter " & Str & " when known" Then
             cell = "(" & cellNum & ") Enter " & Str & " when known"
             cellNum = cellNum + 1
             table.ListColumns(Num).DataBodyRange(indRow, 1) = cell
        End If
        indRow = indRow + 1
    Next cell
End If
listName.Text = ActiveCell.Value
End Sub

Sub sharepointButton()
Dim i As Integer 'count sheets
Dim x As Integer 'count sheets
Dim j As Integer 'count tables
Dim k As Integer 'count rows
Dim colCnt As Integer 'count columns
Dim newWB As Workbook 'New workbook
Dim filePath As String ' file path
Dim fileName As String 'file name
Dim cutPath As String 'cut the path for sharepoint
colCnt = -1
Num = 0

' Cuts the path for Sharepoint
cutPath = ActiveWorkbook.Sheets("Master").Range("O8").Text & "\All Files\"
' Define the desired file name and path
fileName = "Sharepoint Upload.xlsx"
' This example saves it to the user's Desktop
filePath = ActiveWorkbook.Sheets("Master").Range("O8").Text & "\" & fileName
For i = 1 To ActiveWorkbook.Sheets.Count
    If Sheets(i).Name = "Prepared for Sharepoint" Then
        anSwer = MsgBox(Sheets(i).Name & " alreadyexist do you want to create a new one?", vbYesNo, "Sheet Exist")
        If anSwer = vbYes Then
            Application.DisplayAlerts = False
            Sheets("Prepared for Sharepoint").Activate
            Sheets("Prepared for Sharepoint").Delete
            Num = ActiveWorkbook.Sheets.Count
            Application.DisplayAlerts = True
        ElseIf anSwer = vbNo Then
            Sheets(i).Visible = True
            Sheets(i).Activate
            MsgBox ("This sheet has been brought to the front")
            anSwer = MsgBox("Would you like to make this sheet Sharepoint ready?", vbYesNo, "Prepare for Sharepoint")
            If anSwer = vbYes Then
                Num = ActiveWorkbook.Sheets.Count
            ElseIf anSwer = vbNo Then
                Exit Sub
            End If
        End If
'Check to see if Sharepoint pages exists beginning----------------------------------------
Else
    Num = Num + 1
End If
Next i
'Check to see if Sharepoint pages exists ending----------------------------------------
If Num = ActiveWorkbook.Sheets.Count Then
    Dim arrayFullCombined 'All info with combined dates and times
    
    ActiveWorkbook.Sheets.Add after:=Sheets(Sheets.Count)
    Num = Num + 1
    ActiveWorkbook.Sheets(Num).Name = "Prepared for Sharepoint"
    'Count all the rows in entire workbook beginning------------------------------
        For x = 1 To ActiveWorkbook.Sheets.Count
            If Sheets(x).Name <> "Master" And _
            Sheets(x).Name <> "Key Words" And _
            Sheets(x).Name <> "Prepared for Sharepoint" And _
            Sheets(x).Name <> "For Printing" And _
            Sheets(x).Name <> "Before Divorce" Then
                For j = 1 To ActiveWorkbook.Sheets(x).ListObjects.Count
                    For k = 1 To ActiveWorkbook.Sheets(x).ListObjects(j).ListRows.Count
                        Numb = Numb + 1
                    Next k
                Next j
            End If
        Next x
            'Count all the rows in entire workbook ending------------------------------
    ReDim arrayFullCombined(1 To Numb, 1 To 8) 'All info with combined dates and times
    'Populate array beginning------------------------------------------------------
    Numr = 1 'start the count of number of open slots in array
    For i = 1 To ActiveWorkbook.Sheets.Count
        If Sheets(i).Name <> "Master" And _
        Sheets(i).Name <> "Key Words" And _
        Sheets(i).Name <> "Prepared for Sharepoint" And _
        Sheets(i).Name <> "For Printing" And _
        Sheets(i).Name <> "Before Divorce" Then
            Sheets(i).Visible = True
            Sheets(i).Activate
            For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
                For k = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                    ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Select
                    arrayFullCombined(Numr, 1) = "mborycz-" & Numr 'row number
                    arrayFullCombined(Numr, 2) = ActiveWorkbook.ActiveSheet.ListObjects(j).Name 'catagory
                    arrayFullCombined(Numr, 3) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 2) & " " & _
                    Format(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 3), "hh:mm") ' start date plus start time
                    arrayFullCombined(Numr, 4) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 4) & " " & _
                    Format(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 5), "hh:mm") 'end date plus end time
                    arrayFullCombined(Numr, 5) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 6) 'description
                    arrayFullCombined(Numr, 6) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 7) 'notes
                    arrayFullCombined(Numr, 7) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 8) 'keyword
                    arrayFullCombined(Numr, 8) = Replace(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 9), cutPath, "") 'link
                    Numr = Numr + 1 'counts number of open slots in array
                Next k
            Next j
        End If
    Next i
    'Populate array ending---------------------------------------------------------
    Sheets(Num).Visible = True
    Sheets(Num).Activate
    'Make Sharepoint table beginning----------------------------------------------------
                With ActiveWorkbook.ActiveSheet
                    .Cells(1, 1).Value = "#"
                    .Cells(1, 2).Value = "Catagory"
                    .Cells(1, 3).Value = "Start Event"
                    .Cells(1, 4).Value = "Finish Event"
                    .Cells(1, 5).Value = "Description"
                    .Cells(1, 6).Value = "Notes"
                    .Cells(1, 7).Value = "Key words"
                    Cells(1, 8).Value = "Links"
                    .ListObjects.Add(xlSrcRange, Cells(1, 1).CurrentRegion, , xlYes).Name = "PreparedforSharepoint"
                End With
                ActiveWorkbook.ActiveSheet.ListObjects("PreparedforSharepoint").ListRows.Add
                Cells(2, 1) = "blank table"
                Cells(2, 2) = "blank table"
                Cells(2, 3) = "blank table"
                Cells(2, 4) = "blank table"
                Cells(2, 5) = "blank table"
                Cells(2, 6) = "blank table"
                Cells(2, 7) = "blank table"
                Cells(2, 8) = "blank table"
    'Make Sharepoint table ending------------------------------------------------------
    'Populate Sharepoint table beginning-------------------------------------------------
    For j = 0 To Sheets(Num).ListObjects("PreparedforSharepoint").ListColumns.Count
        colCnt = colCnt + 1
    Next j
    For j = 2 To Numr
        'Cells(j, colCnt - 8) = arrayFullCombined(j - 1, colCnt - 8)
        Cells(j, colCnt - 7) = arrayFullCombined(j - 1, colCnt - 7)
        Cells(j, colCnt - 6) = arrayFullCombined(j - 1, colCnt - 6)
        Cells(j, colCnt - 5) = arrayFullCombined(j - 1, colCnt - 5)
        Cells(j, colCnt - 4) = arrayFullCombined(j - 1, colCnt - 4)
        Cells(j, colCnt - 3) = arrayFullCombined(j - 1, colCnt - 3)
        Cells(j, colCnt - 2) = arrayFullCombined(j - 1, colCnt - 2)
        Cells(j, colCnt - 1) = arrayFullCombined(j - 1, colCnt - 1)
        Cells(j, colCnt) = arrayFullCombined(j - 1, colCnt)
    Next j
    'Populate Sharepoint table ending---------------------------------------------------
    'Create new workbook and put table in it beginning------------------------------------

    ActiveWorkbook.Sheets("Prepared for Sharepoint").Cells(1, 1).CurrentRegion.Copy

    
    ' Create a new workbook and assign it to a variable
    Set newWB = Workbooks.Add
    
    ' Save the new workbook with the specified name and location
    ' Application.DisplayAlerts = False prevents prompts if the file already exists (optional)
    Application.DisplayAlerts = False
    newWB.Sheets(1).Name = ("For Sharepoint")
    newWB.Sheets("For Sharepoint").Cells(1, 1).PasteSpecial xlPasteAll
    newWB.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    ' Optional: Close the new workbook and provide confirmation message
    ' newWB.Close SaveChanges:=False
    MsgBox "New workbook created and named: " & fileName
    'Create new workbook and put table in it ending--------------------------------------
    
End If


End Sub

Private Sub ch_ChangeDataLocation_Click()
If Me.cbx_ChangeDataLocation.Text = "" Then
    MsgBox "Please add a table to Switch data to", , "Required Field Empty"
    Exit Sub
End If
Numb = 0
Dim ii As Integer
Dim jj As Integer
ReDim arrayListIndex(1 To 1, 1 To 9)
For i = 1 To ActiveWorkbook.Sheets.Count
    Worksheets(i).Visible = True
    Worksheets(i).Activate
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        If ActiveWorkbook.ActiveSheet.ListObjects(j).Name = uf_Master.cbx_TblCheck.Text Then
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                If uf_Master.lb_AllData.List(uf_Master.lb_AllData.ListIndex, 0) = _
                    ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 1) Then
                    arrayListIndex(1, 1) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 1)
                    arrayListIndex(1, 2) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 2)
                    arrayListIndex(1, 3) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 3)
                    arrayListIndex(1, 4) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 4)
                    arrayListIndex(1, 5) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 5)
                    arrayListIndex(1, 6) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 6)
                    arrayListIndex(1, 7) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 7)
                    arrayListIndex(1, 8) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 8)
                    arrayListIndex(1, 9) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 9)
                    ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 1).Delete
                    For Num = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                        ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(Num, 1) = Num
                    Next Num
                    For ii = 1 To ActiveWorkbook.Sheets.Count
                        Worksheets(ii).Visible = True
                        Worksheets(ii).Activate
                        For jj = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
                            If ActiveWorkbook.ActiveSheet.ListObjects(jj).Name = Me.cbx_ChangeDataLocation.Text Then
                                For Num = 1 To ActiveWorkbook.ActiveSheet.ListObjects(jj).ListRows.Count
                                    Numb = Numb + 1
                                Next Num
                                For Num = 1 To 9
                                    ActiveWorkbook.ActiveSheet.ListObjects(jj).DataBodyRange(Numb + 1, Num) = arrayListIndex(1, Num)
                                Next Num
                                For Num = 1 To ActiveWorkbook.ActiveSheet.ListObjects(jj).ListRows.Count
                                    ActiveWorkbook.ActiveSheet.ListObjects(jj).DataBodyRange(Num, 1) = Num
                                Next Num
                                    MsgBox "Data has been tranfered to Listindex " & _
                                    ActiveWorkbook.ActiveSheet.ListObjects(jj).DataBodyRange(Num - 1, 1) & _
                                    " of table " & Me.cbx_ChangeDataLocation.Text, , "Data Transfered Successfully"
                                    Me.cbx_ChangeDataLocation.Text = ""
                                    uf_Switch.Hide
                                    Call goToTable   ' Goes to a table
                                    Exit Sub
                            End If
                        Next jj
                    Next ii
                End If
            Next x
        End If
    Next j
Next i

End Sub

Sub Search_Array(ByVal xx As Integer, ByRef tbxNote1 As MSForms.TextBox)

Dim Num As Integer
Dim Numb As Integer
Dim Numr As Integer
Dim Numm As Integer
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim jj As Integer

Dim fullArray As Variant
If tbxNote1.Text = "" Then
    MsgBox "Please add something to search"
    Exit Sub
End If
' Counting sheets and rows in tables beinging----------------------------------------------
uf_Master.lb_Search.Clear
i = 0
For i = 1 To ActiveWorkbook.Sheets.Count
    If Sheets(i).Name = "For Printing" Then
        Sheets(i).Delete
    End If
Next i
j = 0
Numb = 0
For i = 1 To ActiveWorkbook.Sheets.Count
Worksheets(i).Visible = True
Worksheets(i).Activate
    If ActiveWorkbook.ActiveSheet.ListObjects.Count > 0 = True Then
    Num = 0
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
       Num = Num + 1
        x = 0
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(Num).ListRows.Count
            ActiveWorkbook.ActiveSheet.ListObjects(Num).Range.Activate
            ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, xx).Select
                If InStr(ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, xx), tbxNote1) <> 0 Then
                    Numb = Numb + 1
                End If
            Next x
    Next j
    End If
    If i > 1 = True Then
        Worksheets(i).Visible = False
    End If
Next i
If Numb = 0 Then
    MsgBox ("This data does not exist")
    Exit Sub
End If
' Counting sheets and rows in tables ending----------------------------------------------
ReDim fullArray(1 To Numb, 1 To 10)
' Work on arrays beginning-----------------------------------------------------------------
i = 0
j = 0
Numr = 0
jj = 1
For i = 1 To ActiveWorkbook.Sheets.Count
Worksheets(i).Visible = True
Worksheets(i).Activate
    If ActiveWorkbook.ActiveSheet.ListObjects.Count > 0 = True Then
    Num = 0
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
       Num = Num + 1
        x = 0
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(Num).ListRows.Count
            ActiveWorkbook.ActiveSheet.ListObjects(Num).Range.Activate
            ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, xx).Select
                If InStr(ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, xx), tbxNote1) <> 0 Then
                    fullArray(jj, 1) = ActiveWorkbook.ActiveSheet.ListObjects(Num)
                    fullArray(jj, 2) = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 1)
                    fullArray(jj, 3) = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 2)
                    fullArray(jj, 4) = Format(ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 3), "hh:mm")
                    fullArray(jj, 5) = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 4)
                    fullArray(jj, 6) = Format(ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 5), "hh:mm")
                    fullArray(jj, 7) = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 6)
                    fullArray(jj, 8) = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 7)
                    fullArray(jj, 9) = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 8)
                    fullArray(jj, 10) = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 9)
                    jj = jj + 1
                End If
            Next x
    Next j
    End If
    If i > 1 = True Then
        Worksheets(i).Visible = False
    End If
Next i
With uf_Master.lb_Search
    .List = fullArray
    .FontSize = 8
    .ColumnWidths = "75,15,50,30,50,30,112,113,0,0"
End With
End Sub

Sub goToTable() ' Goes to a table
Dim tblArray() As Variant, rowInd As Variant, colInd As Variant
Dim tblCreate As ComboBox
Dim Sh As Worksheet
'Dim table As ListObject
Dim j As Integer, i As Integer, rowNum As Integer, colNum As Integer, z As Integer

cbx_StartDateCheck = ""
cbx_StartTimeCheck = ""
cbx_EndDateCheck = ""
cbx_CheckAbbr = ""
cbx_CheckDesc = ""
cbx_CheckNotes = ""

tb_ColNum = ""
tb_StartDate = ""
tb_StartTime = ""
tb_EndDate = ""
tb_Abbreviate = ""
tb_Description = ""
tb_Notes = ""

Set tblCreate = uf_Master.cbx_TblCheck

If tblCreate.Text = "" Then
    MsgBox Prompt:="Enter a table to work on" _
    & vbNewLine & vbNewLine & "If there are no tables to choose from," _
    & vbNewLine & "create a table in 'Create Table' field", Title:="Go to Table Field Empty"
    Exit Sub
End If

Set table = ActiveWorkbook.ActiveSheet.ListObjects(tblCreate.Text)
Set Sh = ActiveWorkbook.ActiveSheet

j = 1 'For column count
z = 1 'For row count

Do Until colNum = table.Range.Columns.Count
    colNum = colNum + 1 'Count columns
Loop
Do Until rowNum = table.Range.Rows.Count
    rowNum = rowNum + 1 'Count rows
Loop

ReDim Preserve tblArray(1 To rowNum, 1 To colNum)
        Do Until z = rowNum + 1 'Populate rows
            For colInd = 1 To colNum
                tblArray(z, colInd) = table.Range(z + 1, colInd)
            Next colInd
            z = z + 1
        Loop

For i = 1 To Sh.ListObjects.Count
    If ActiveSheet.ListObjects(i) = tblCreate.Text Then
        ActiveWorkbook.ActiveSheet.ListObjects(i).Range.Select
    End If
Next i
With uf_Master.lb_AllData
    .Clear
    .ColumnHeads = False
    .ColumnCount = UBound(tblArray, 2)
    .List = tblArray()
End With

uf_Master.cbx_StartDateCheck.Clear
uf_Master.cbx_StartTimeCheck.Clear
uf_Master.cbx_EndDateCheck.Clear
uf_Master.cbx_CheckAbbr.Clear
uf_Master.cbx_CheckDesc.Clear
uf_Master.cbx_CheckNotes.Clear
For j = LBound(tblArray) To UBound(tblArray)
    uf_Master.cbx_StartDateCheck.AddItem (tblArray(j, 2)) 'Populate Start Date list to check
    uf_Master.cbx_StartTimeCheck.AddItem Format(tblArray(j, 3), "hh:mm") 'Populate Start Time list to check
    uf_Master.cbx_EndDateCheck.AddItem (tblArray(j, 4)) 'Populate End Date list to check
    uf_Master.cbx_CheckAbbr.AddItem Format(tblArray(j, 5), "hh:mm") 'Populate Abbreviation list to check
    uf_Master.cbx_CheckDesc.AddItem (tblArray(j, 6)) 'Populate Descriptions list to check
    uf_Master.cbx_CheckNotes.AddItem (tblArray(j, 7)) 'Populate Notes list to check
Next j
For cell = 1 To table.ListColumns(2).DataBodyRange.Count
    If table.ListColumns(2).DataBodyRange(table.ListRows.Count, 1) <> uf_Master.tb_Abbreviate.Text Then
        Num = 1
    Else:
        Num = Num + 1
    End If
Next cell

For cell = 0 To table.ListColumns(2).DataBodyRange.Count - 1
    uf_Master.lb_AllData.List(cell, 2) = Format(uf_Master.lb_AllData.List(cell, 2), "hh: mm")
    uf_Master.lb_AllData.List(cell, 4) = Format(uf_Master.lb_AllData.List(cell, 4), "hh: mm")
Next cell
End Sub

Sub workShtsIntoCombobox()
Dim shArray, i
Num = 1
Numb = 1

'Populating Worksheets into comboboxes beginning----------------------------------
For i = 1 To ActiveWorkbook.Sheets.Count - 1
    If Sheets(i).Name <> "Master" And Sheets(i).Name <> "Key Words" Then
        Num = Num + 1
    End If
Next i
ReDim shArray(1 To Num)

With uf_Master.lb_AllData
    .ColumnCount = 7
    .ColumnWidths = "15;50;50;50;50;130;130"
End With

For i = 1 To ActiveWorkbook.Sheets.Count
    If Sheets(i).Name <> "Master" And Sheets(i).Name <> "Key Words" Then
        shArray(Numb) = Sheets(i).Name
        Numb = Numb + 1
    End If
Next i

shArray = SortArrayAlphabetical(shArray)

'Populating Worksheets into comboboxes ending----------------------------------
uf_Master.cbx_ShCheck.List = shArray 'Adds sheet array to combobox seeing what sheets exist
uf_Master.cbx_ShDelete.List = shArray 'Adds sheet array to combobox for deletion options
uf_Master.cbx_RefineShCheck.List = shArray 'Adds sheets array to combobox on large search page
End Sub

Sub tblsIntoCombobox()
Numb = 0
Num = 1

Dim tblArray
Dim j As Integer
Dim i As Integer
Dim x As Integer

'Populating Tables into comboboxes beginning----------------------------------
For i = 1 To ActiveWorkbook.Sheets.Count
    If ActiveWorkbook.Worksheets(i).Name <> "Master" And ActiveWorkbook.Worksheets(i).Name <> "Key Words" Then
        Sheets(i).Visible = True
        Sheets(i).Activate
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            Numb = Numb + 1
        Next j
        Sheets(i).Visible = False
    End If
Next i

ReDim tblArray(1 To Numb)

For i = 1 To ActiveWorkbook.Sheets.Count
    If ActiveWorkbook.Worksheets(i).Name <> "Master" And ActiveWorkbook.Worksheets(i).Name <> "Key Words" Then
        Sheets(i).Visible = True
        Sheets(i).Activate
        For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            tblArray(Num) = Sheets(i).ListObjects(x).Name
                Num = Num + 1
            Next x
            Sheets(i).Visible = False
    End If
Next i

tblArray = SortArrayAlphabetical(tblArray)
'Populating Tables into comboboxes ending----------------------------------
uf_Master.cbx_RefineTblCheck.List = tblArray 'Adds tables array to combobox on large search page
uf_Switch.cbx_ChangeDataLocation.List = tblArray 'Adds tables array to combobox on large search page
End Sub

Sub keyWordsIntoCombobox()
Dim keyArray

Dim x As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer

Numr = 0

'Populating Keywords into comboboxes beginning----------------------------------

For i = 1 To ActiveWorkbook.Sheets.Count
    If ActiveWorkbook.Worksheets(i).Name = "Key Words" Then
        Sheets(i).Visible = True
        Sheets(i).Activate
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            If ActiveWorkbook.ActiveSheet.ListObjects(j) = "KeyWords" Then
                For k = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                    Numr = Numr + 1
                Next k
            End If
        Next j
        Sheets(i).Visible = False
    End If
Next i

ReDim keyArray(1 To Numr)
Num = 1
For i = 1 To ActiveWorkbook.Sheets.Count
    If ActiveWorkbook.Worksheets(i).Name = "Key Words" Then
        Sheets(i).Visible = True
        Sheets(i).Activate
        For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            If ActiveWorkbook.ActiveSheet.ListObjects(x).Name = "KeyWords" Then
                For k = 1 To ActiveWorkbook.ActiveSheet.ListObjects(x).ListRows.Count
                    keyArray(Num) = ActiveWorkbook.ActiveSheet.ListObjects(x).DataBodyRange(k, 1)
                    Num = Num + 1
                Next k
            End If
        Next x
        Sheets(i).Visible = False
    End If
Next i

keyArray = SortArrayAlphabetical(keyArray)
'Populating Keywords into comboboxes ending----------------------------------
uf_Master.cb_KeyWords.List = keyArray 'Adds keywords array to combobox on large search page
End Sub

Sub folderCreateAndSearch(ByRef tbxNote1 As MSForms.TextBox, ByRef tbxNote2 As MSForms.TextBox, ByRef tbxNote3 As MSForms.TextBox, ByRef tbxNote4 As MSForms.TextBox)
Dim pAth As String
Dim dateOne As String
Dim dateTwo As String
Dim finalDateOne As String
Dim finalDateTwo As String
Dim timeOne As String
Dim timeTwo As String
Dim anSwer As Integer

Dim fso As Object
Application.DisplayAlerts = False
Set table = ActiveWorkbook.ActiveSheet.ListObjects(uf_Master.cbx_TblCheck.Text)
Numb = 0
Set fso = CreateObject("Scripting.FileSystemObject")
Numb = 0
On Error Resume Next
If uf_Master.lb_AllData.List(uf_Master.lb_AllData.ListIndex, 0) = "blank table" Then
    MsgBox "Please as data to other fields before adding a folder"
    Exit Sub
End If

If tbxNote1 <> "" And tbxNote2 <> "" And tbxNote3 <> "" And tbxNote4 <> "" Then
    If Left(tbxNote3, 3) = ":" Then
        timeOne = Replace(tbxNote3, ": ", "")
    Else
        timeOne = Replace(tbxNote3, ": ", "")
    End If
    If Left(tbxNote4, 3) = ":" Then
        timeTwo = Replace(tbxNote4, ": ", "")
    Else
        timeTwo = Replace(tbxNote4, ": ", "")
    End If
    dateOne = tbxNote1
    dateTwo = tbxNote2
    If Mid(dateOne, 2, 1) = "/" And Mid(dateOne, 4, 1) = "/" Then
        finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 1) & "/" & Mid(dateOne, 3, 1)
    ElseIf Mid(dateOne, 4, 1) = "0" And Mid(dateOne, 5, 1) = "/" Then
        finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 1) & "/" & Mid(dateOne, 3, 2)
    ElseIf Mid(dateOne, 3, 1) = "/" And Mid(dateOne, 5, 1) = "/" And Left(dateOne, 2) <> "11" And Left(dateOne, 2) <> "12" Then
        finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 1) & "/" & Mid(dateOne, 4, 1)
    ElseIf Mid(dateOne, 3, 1) = "/" And Mid(dateOne, 5, 1) = "/" And Left(dateOne, 2) = "10" Then
        finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 2) & "/" & Mid(dateOne, 4, 1)
    ElseIf Mid(dateOne, 3, 1) = "/" And Mid(dateOne, 5, 1) = "/" And Left(dateOne, 2) = "11" Then
        finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 1) & "/" & Mid(dateOne, 4, 1)
    ElseIf Mid(dateOne, 3, 1) = "/" And Mid(dateOne, 6, 1) = "/" And Left(dateOne, 2) = "11" Then
        finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 2) & "/" & Mid(dateOne, 4, 2)
    ElseIf Mid(dateOne, 3, 1) = "/" And Mid(dateOne, 5, 1) = "/" And Left(dateOne, 2) <> "12" Then
        finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 1) & "/" & Mid(dateOne, 4, 1)
    ElseIf Mid(dateOne, 3, 1) = "/" And Mid(dateOne, 5, 1) = "/" And Left(dateOne, 2) = "12" Then
        finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 2) & "/" & Mid(dateOne, 4, 1)
    ElseIf Mid(dateOne, 3, 1) = "/" And Mid(dateOne, 6, 1) = "/" And Left(dateOne, 2) = "12" Then
        finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 2) & "/" & Mid(dateOne, 4, 2)
    ElseIf Mid(dateOne, 2, 1) = "/" And Mid(dateOne, 5, 1) = "/" Then
        finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 1) & "/" & Mid(dateOne, 3, 2)
    End If
    If Mid(dateTwo, 2, 1) = "/" And Mid(dateTwo, 4, 1) = "/" Then
        finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 1) & "/" & Mid(dateTwo, 3, 1)
    ElseIf Mid(dateTwo, 4, 1) = "0" And Mid(dateTwo, 5, 1) = "/" Then
        finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 1) & "/" & Mid(dateTwo, 3, 2)
    ElseIf Mid(dateTwo, 3, 1) = "/" And Mid(dateTwo, 5, 1) = "/" And Left(dateTwo, 2) <> "11" And Left(dateTwo, 2) <> "12" Then
        finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 1) & "/" & Mid(dateTwo, 4, 1)
    ElseIf Mid(dateTwo, 3, 1) = "/" And Mid(dateTwo, 5, 1) = "/" And Mid(dateTwo, 1, 2) = "10" Then
        finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 2) & "/" & Mid(dateTwo, 4, 1)
    ElseIf Mid(dateTwo, 3, 1) = "/" And Mid(dateTwo, 5, 1) = "/" And Left(dateTwo, 2) = "11" Then
        finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 1) & "/" & Mid(dateTwo, 4, 1)
    ElseIf Mid(dateTwo, 3, 1) = "/" And Mid(dateTwo, 6, 1) = "/" And Left(dateTwo, 2) = "11" Then
        finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 2) & "/" & Mid(dateTwo, 4, 2)
    ElseIf Mid(dateTwo, 3, 1) = "/" And Mid(dateTwo, 5, 1) = "/" And Left(dateTwo, 2) <> "12" Then
        finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 1) & "/" & Mid(dateTwo, 4, 1)
    ElseIf Mid(dateTwo, 3, 1) = "/" And Mid(dateTwo, 5, 1) = "/" And Left(dateTwo, 2) = "12" Then
        finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 2) & "/" & Mid(dateTwo, 4, 1)
    ElseIf Mid(dateTwo, 3, 1) = "/" And Mid(dateTwo, 6, 1) = "/" And Left(dateTwo, 2) = "12" Then
        finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 2) & "/" & Mid(dateTwo, 4, 2)
    ElseIf Mid(dateTwo, 2, 1) And Mid(dateTwo, 5, 1) = "/" Then
        finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 1) & "/" & Mid(dateTwo, 3, 2)
    End If
     pAth = Worksheets("Master").Range("O8") & "\All Files" & "\" & Replace(finalDateOne, "/", "-") & " at " & timeOne & " to " & Replace(finalDateTwo, "/", "-") & " at " & timeTwo
    '"C:\Users\mludwigborycz\OneDrive - Plastics Family Americas\Desktop\Dont Delete\Divorce\All Files" & "\" & Replace(finalDateOne, "/", "-") & " at " & timeOne & " to " & Replace(finalDateTwo, "/", "-") & " at " & timeTwo
    
    If Not fso.FolderExists(pAth) Then
        If table.ListRows.Count = 1 Then
            Numb = Numb + 1
        Else
            For x = 1 To uf_Master.tb_ColNum
                Numb = Numb + 1
            Next x
        End If
        MkDir pAth
        uf_Master.cb_FolderSearch.Caption = "Folder Exists"
        MsgBox "A new folder has been created, do NOT change the name"
        table.ListColumns(9).DataBodyRange(Numb, 1) = pAth
        ActiveWorkbook.FollowHyperlink Address:=pAth, NewWindow:=True
    Else
        uf_Master.cb_FolderSearch.Caption = "Folder Exists"
        anSwer = MsgBox("A folder for this date and time already exists, do NOT change the name. Do you want to open it?", vbQuestion + vbYesNo)
        If anSwer = vbYes Then
        If table.ListRows.Count = 1 Then
            Numb = Numb + 1
        Else
            For x = 1 To uf_Master.tb_ColNum
                Numb = Numb + 1
            Next
        End If
            table.ListColumns(9).DataBodyRange(Numb, 1) = pAth
            ActiveWorkbook.FollowHyperlink Address:=pAth, NewWindow:=True
        ElseIf anSwer = vbNo Then
        End If
    End If
    fso = Nothing
Else
    MsgBox "Please fill all fields before entering data"
    Exit Sub
End If
Application.DisplayAlerts = False
End Sub

Sub addingForChanges(Str As String)
Dim wordString As String
Dim wordArray As Variant
Dim smallArray As Variant
Dim wordList As Variant
Dim i As Integer
Dim x As Integer
Dim j As Integer


If uf_Master.cb_Switch.Caption = "Click for Single Select" Then
    MsgBox "Please select 'Click for Single Select', before adding and ending to a key word", , "Choose Single Select"
    Exit Sub
End If
If uf_Master.cb_KeyWords.Text = "" Then
    MsgBox "Put a key word into the text box to add " & Str & " at the end of the key word", , "Add Pattern Change Simble"
    Exit Sub
End If

wordArray = Split((uf_Master.lb_Search.List(uf_Master.lb_Search.ListIndex, 8)), ";")

For i = 0 To UBound(wordArray)
    If uf_Master.cb_KeyWords.Text & "-no" = wordArray(i) Or _
    uf_Master.cb_KeyWords.Text & "-maybe" = wordArray(i) Or _
    uf_Master.cb_KeyWords.Text & "-yes" = wordArray(i) Or _
    uf_Master.cb_KeyWords.Text = wordArray(i) Then
        If InStr(wordArray(i), "-") > 0 Then
            smallArray = Split(wordArray(i), "-")
        End If
        If uf_Master.cb_KeyWords.Text & Str = wordArray(i) Then
            MsgBox "This keyword already had a '" & Str & " extention", , "Extention exists"
        ElseIf wordArray(i) = uf_Master.cb_KeyWords.Text Then
            wordArray(i) = uf_Master.cb_KeyWords.Text & Str
        ElseIf InStr(wordArray(i), "-") <> Str Then
            wordArray(i) = smallArray(0) & Str
        End If
    End If
Next i

uf_Master.tb_KeyWordsUsed.Text = Join(wordArray, ";")
For i = 1 To ActiveWorkbook.Worksheets.Count
    Worksheets(i).Visible = True
    Worksheets(i).Activate
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        If ActiveWorkbook.ActiveSheet.ListObjects(j).Name = uf_Master.lb_Search.List(uf_Master.lb_Search.ListIndex, 0) Then
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                If uf_Master.lb_Search.List(uf_Master.lb_Search.ListIndex, 1) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 1) Then
                    ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 8) = Join(wordArray, ";")
                End If
            Next x
        End If
    Next j
    If Worksheets(i).Name <> "Master" Then
        Worksheets(i).Visible = False
    End If
Next i
uf_Master.lb_Search.List(uf_Master.lb_Search.ListIndex, 8) = Join(wordArray, ";")
End Sub

Sub Check_Date(ByRef tbxNote1 As MSForms.TextBox)

Dim strVal As String
Dim dDate As Date
If IsNumeric(tbxNote1) And Len(tbxNote1.Value) = 8 Then
    strVal = Format(tbxNote1, "00000000")
    If Application.International(xlDateOrder) = 1 Then 'mm/dd/yyyy
        dDate = DateValue(Left(strVal, 2) & "/" & Mid(strVal, 3, 2) & "/" & Right(strVal, 4))
     Else 'mm/dd/yyyy
        dDate = DateValue(Left(strVal, 2) & "/" & Mid(strVal, 3, 2) & "/" & Right(strVal, 4))
        tbxNote1.Text = dDate
    End If
ElseIf Len(tbxNote1.Value) < 8 Or Len(tbxNote1.Value) < 10 Then
    MsgBox "Please enter a date formated as mm/dd/yyyy"
    tbxNote1.Text = ""
    Exit Sub
ElseIf IsNumeric(tbxNote1.Value) = False Then
    MsgBox "Only numbers aloud in this box"
    tbxNote1.Text = ""
    Exit Sub
End If

If Not IsDate(dDate) Then
    MsgBox "Please enter a date formated as mm/dd/yyyy"
    tbxNote1.Text = ""
    Exit Sub
Else
    If Mid(tbxNote1, 2, 1) = "/" And Mid(tbxNote1, 4, 1) <> "/" Then
        tbxNote1.Text = Format(tbxNote1.Text, "m/dd/yyyy")
    ElseIf Mid(tbxNote1, 5, 1) = "/" Then
        tbxNote1.Text = Format(tbxNote1.Text, "mm/d/yyyy")
    ElseIf Mid(tbxNote1, 4, 1) = "/" Then
        tbxNote1.Text = Format(tbxNote1.Text, "m/d/yyyy")
    End If
End If
End Sub

Sub CheckTblInfo(ByVal indRow As Integer, Num As Integer, Numb As Integer, Numr As Integer, Numm As Integer, Numbb As Integer, Numrr As Integer, _
tbxNote1 As MSForms.TextBox, tbxNote2 As MSForms.TextBox, tbxNote3 As MSForms.TextBox, tbxNote4 As MSForms.TextBox, tbxNote5 As MSForms.TextBox, _
tbxNote6 As MSForms.TextBox, cbxNote As MSForms.ComboBox)
On Error Resume Next
indRow = 0
Set table = ActiveWorkbook.ActiveSheet.ListObjects(uf_Master.cbx_TblCheck.Text)
    tbxNote1.Text = cbxNote.Text
    If Num = 3 Or Num = 5 Then
        Do Until Format(table.DataBodyRange(indRow, Num).Value, "hh:mm") = cbxNote.Value
            indRow = indRow + 1
        Loop
    Else
        Do Until table.DataBodyRange(indRow, Num).Value = cbxNote.Value
            indRow = indRow + 1
        Loop
    End If
    tbxNote2.Text = table.DataBodyRange(indRow, Numb).Value
    tbxNote3.Text = table.DataBodyRange(indRow, Numr).Value
    tbxNote4.Text = table.DataBodyRange(indRow, Numm).Value
    tbxNote5.Text = table.DataBodyRange(indRow, Numbb).Value
    tbxNote6.Text = table.DataBodyRange(indRow, Numrr).Value
    tb_StartTime = Format(tb_StartTime, "hh:mm")
    tb_Abbreviate = Format(tb_Abbreviate, "hh:mm")
    cbxNote.Clear
End Sub

Sub SortListBoxAlphabetically(ByRef TargetListBox As MSForms.ListBox)
    Dim i As Long
    Dim j As Long
    Dim Temp As Variant
    Dim blnSorted As Boolean
    
    ' Loop until no swaps are needed (Bubble Sort logic)
    Do Until blnSorted
        blnSorted = True
        ' Use LCase to ensure case-insensitive sorting
        For i = 0 To TargetListBox.ListCount - 2
            If LCase(TargetListBox.List(i)) > LCase(TargetListBox.List(i + 1)) Then
                ' Swap items
                Temp = TargetListBox.List(i)
                TargetListBox.List(i) = TargetListBox.List(i + 1)
                TargetListBox.List(i + 1) = Temp
                blnSorted = False ' A swap occurred, so list is not fully sorted yet
            End If
        Next i
    Loop
End Sub

Public Function loopWorksheets()
For i = 1 To ActiveWorkbook.Worksheets.Count
    Worksheets(i).Visible = True
    Worksheets(i).Activate

Next i
End Function

Sub makeItColored()
Dim listInd As Integer
Dim x As Integer
Dim j As Integer

Dim keyWords As String
Dim noChange As String
Dim maybeChange As String
Dim yesChange As String

noChange = "-no"
maybeChange = "-maybe"
yesChange = "-yes"

For listInd = 0 To lb_Search.ListCount - 1
    keyWords = lb_Search.List(listInd, 8)
    Do While InStr(keyWords, noChange) > 0
        lb_Search.List(listInd).Interior.ColorIndex &HFF&
        Application.Wait (Now + TimeValue("0:00:01"))
        lb_Search.List(listInd).Interior.ColorIndex &H80000005
        Application.Wait (Now + TimeValue("0:00:01"))
        DoEvents
    Loop
Next listInd
End Sub

Function SortArrayAlphabetical(myArray As Variant) As Variant
Dim i As Long
Dim j As Long
Dim Temp As Variant
    
' Loop through the array
For i = LBound(myArray) To UBound(myArray) - 1
    For j = i + 1 To UBound(myArray)
        ' Compare items (UCase makes the sort case-insensitive)
        If UCase(myArray(i)) > UCase(myArray(j)) Then
            ' Swap items
            Temp = myArray(i)
            myArray(i) = myArray(j)
            myArray(j) = Temp
        End If
    Next j
Next i
'Return the sorted array
SortArrayAlphabetical = myArray
End Function

Function FolderExistsFSO(ByVal sFolderName As String) As Boolean
    Dim fso As Object
    ' Create the FileSystemObject late-bound
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Use the built-in FolderExists method
    FolderExistsFSO = fso.FolderExists(sFolderName)
    
    ' Clean up
    Set fso = Nothing
End Function

Function GetDateTaken(filePath As String) As Date
    Dim objShell As Object
    Dim objFolder As Object
    Dim objFolderItem As Object
    Dim fileName As String
    Dim folderPath As String
    
    ' Split the full path into folder path and file name
    folderPath = Left(filePath, InStrRev(filePath, "\"))
    fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
    
    ' Create a Shell object (using late binding)
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(folderPath)
    Set objFolderItem = objFolder.ParseName(fileName)
    
    If Not objFolderItem Is Nothing Then
        ' GetDetailsOf method uses an index to get properties
        ' Index 12 is typically "Date Taken" in Windows Explorer
        Dim dateTakenStr As String
        dateTakenStr = objFolder.GetDetailsOf(objFolderItem, 12)
        
        ' Check if a date was found and convert it
        If dateTakenStr <> "" Then
            ' The format returned by GetDetailsOf can vary by system settings
            ' A reliable way is to try converting it
            On Error Resume Next
            GetDateTaken = CDate(dateTakenStr)
            On Error GoTo 0
        End If
    End If
    
    ' Clean up
    Set objFolderItem = Nothing
    Set objFolder = Nothing
    Set objShell = Nothing
End Function

' Example Usage:
Sub TestGetDateTaken()
    Dim myDate As Date
    ' **Change this to your actual file path**
    myDate = GetDateTaken("C:\Users\mludwigborycz\OneDrive - Plastics Family Americas\Desktop\Dont Delete\Divorce\All Files\2026-2-8 at 1215 to 2026-2-8 at 2000\B-2026-2-8 at 1215 to 2026-2-8 at 2000.JPEG")
    
    If myDate > #1/1/1900# Then ' Check if a valid date was returned
        MsgBox "The picture was taken on: " & myDate
    Else
        MsgBox "Could not retrieve Date Taken or it does not exist."
    End If
End Sub

Sub changePeriods()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim update_Str As String
For i = 3 To ActiveWorkbook.Sheets.Count
    Sheets(i).Visible = True
    Sheets(i).Activate
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Activate
        For k = 0 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count - 1
             If InStr(1, ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 2), ".") <> 0 And InStr(1, ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 4), ".") <> 0 Then
                ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 2).Select
                updateD_Str = Replace(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 2), ".", "/", 1)
                ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 2) = updateD_Str
                ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 4).Select
                updateD_Str = Replace(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 4), ".", "/", 1)
                ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 4) = updateD_Str
            End If
        Next k
    Next j
Next i
End Sub

Sub changeTime()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim update_Str As String
For i = 3 To ActiveWorkbook.Sheets.Count
    Sheets(i).Visible = True
    Sheets(i).Activate
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Activate
        For k = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
            ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 3).Select
            ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 3) = Format(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 3), "hh:mm")
            ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 5).Select
            ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 5) = Format(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 5), "hh:mm")
        Next k
    Next j
Next i
End Sub

Sub changeLinks()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim fullStr As String
Dim addedStr As String
Dim fstPtStr As String
Dim lstPtStr As String
Dim update_Str As String
Call UnprotectSheets
addedStr = "OneDrive - Plastics Family Americas"
For i = 3 To ActiveWorkbook.Sheets.Count
    Sheets(i).Visible = True
    Sheets(i).Activate
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Activate
        For k = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
            If ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 9) <> "" Then
                ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 9).Select
                fullStr = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 9)
                fstPtStr = Left(fullStr, 23)
                lstPtStr = Mid(fullStr, 23)
                update_Str = fstPtStr & addedStr & lstPtStr
                ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 9) = update_Str
            End If
        Next k
    Next j
Next i
End Sub

Sub addAllLinks()
Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim pAth As String
Dim sDate As String
Dim sTime As String
Dim eDate As String
Dim eTime As String

For i = 1 To ActiveWorkbook.Worksheets.Count
    If Sheets(i).Name <> "Master" And Sheets(i).Name <> "Key Words" Then
        Sheets(i).Visible = True
        Sheets(i).Activate
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Activate
            For k = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                sDate = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 2)
                sTime = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 3)
                eDate = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 4)
                eTime = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 5)
            Next k
        Next j
    End If
    If Sheets(i).Name <> "Master" Then
        Sheets(i).Visible = False
    Else
        Sheets(i).Visible = True
    End If
Next i
pAth = ActiveWorkbook.Sheets("Master").Range("O8") ' Specify the folder path to check
    
If FolderExistsFSO(pAth) Then
    MsgBox "Folder " & strPath & " exists"
Else
    MsgBox "Folder " & strPath & " does NOT exist"
End If
End Sub

Sub ProtectSheets()
Dim wsheet As Worksheet
For Each wsheet In ActiveWorkbook.Worksheets
wsheet.Protect Password:="password"
Next wsheet
End Sub

Sub UnprotectSheets()
Dim wsheet As Worksheet
For Each wsheet In ActiveWorkbook.Worksheets
wsheet.Unprotect Password:="password"
Next wsheet
End Sub

Sub StartButton()
Dim i As Integer
Dim j As Integer
Dim x As Integer

Call UnprotectSheets

If Worksheets("Master").Range("O10") <> "" Then
    For i = 3 To ActiveWorkbook.Sheets.Count
        Sheets(i).Visible = True
        Sheets(i).Activate
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Select
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                If ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 9) <> "" Then
                    Str1 = Replace(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 9), Worksheets("Master").Range("O8"), Worksheets("Master").Range("O10"))
                    ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 9) = Str1
                End If
            Next x
        Next j
    Next i
    Worksheets("Master").Range("O8") = Worksheets("Master").Range("O10")
    Worksheets("Master").Range("O10") = ""
End If
With uf_Master
    .Height = 444
    .Width = 517
End With
'Call SortListBoxAlphabetically(uf_Master.cbx_ShCheck)
'Call SortListBoxAlphabetically(uf_Master.cbx_ShDelete)
'Call SortListBoxAlphabetically(uf_Master.cbx_RefineShCheck)
uf_Master.Show

Call ProtectSheets
End Sub

Sub BoxForDataEntry(ByRef listName As MSForms.TextBox, ByVal Str As String, ByVal Num As Integer)
Set table = ActiveWorkbook.ActiveSheet.ListObjects(uf_Master.cbx_TblCheck.Text)
indRow = 0
anSwer = MsgBox(Str & ":" & vbNewLine & listName.Text _
& vbNewLine & vbNewLine & "Would you like to change this?", vbYesNo, Str)

If anSwer = vbYes Then
    inptAns = InputBox("Type into field to change, or leave blank to delete", "Change or delete field")
    If StrPtr(inptAns) = 0 Then
        Exit Sub
    ElseIf StrPtr(inptAns) <> 0 Then
         anSwer = inptAns
    Do Until table.ListColumns(7).DataBodyRange(indRow, 1).Value = uf_Master.tb_Notes.Text
        indRow = indRow + 1
    Loop
    table.DataBodyRange(indRow, Num).Value = anSwer
    listName.Text = anSwer
    End If
    If IsNumeric(anSwer) Then
        anSwer = "'" & anSwer
    End If
ElseIf anSwer = vbNo Then
    Exit Sub
End If
 If anSwer = "" Then
    anSwer = "Enter " & Str & " when known"
    table.DataBodyRange(indRow, Num).Value = anSwer
    listName.Text = anSwer
 Else:
    table.DataBodyRange(indRow, Num).Value = anSwer
    listName.Text = anSwer
End If
End Sub

Sub GreenDataEntry(ByVal Str As String, ByVal Num As Integer, ByRef listName As MSForms.TextBox, ByRef table As ListObject)

If listName.Text <> "" Then
    If table.ListColumns(Num).DataBodyRange(table.ListRows.Count, 1) <> "" Then
        table.ListRows.Add
    End If
    'ActiveSheet.Unprotect Password:="mEEKAbUTTICH@!21"
    table.ListColumns(Num).DataBodyRange(table.ListRows.Count, 1).Select
    ActiveCell.Value = listName.Text
    If IsNumeric(ActiveCell.Value) Then
        ActiveCell.Value = "'" & ActiveCell.Value
    End If
ElseIf listName.Text = "" Then
    If table.ListColumns(Num).DataBodyRange(table.ListRows.Count, 1) <> "" Then
        table.ListRows.Add
    End If
    cellNum = 1
    indRow = 1
    table.ListColumns(Num).DataBodyRange(table.ListRows.Count, 1).Select
    ActiveCell.Value = "Enter " & Str & " when known"
   ' listName.Text = ActiveCell.Value
    For Each cell In table.ListColumns(Num).DataBodyRange
        If cell.Value Like "*" & "Enter " & Str & " when known" Then
             cell = "(" & cellNum & ") Enter " & Str & " when known"
             cellNum = cellNum + 1
             table.ListColumns(Num).DataBodyRange(indRow, 1) = cell
        End If
        indRow = indRow + 1
    Next cell
End If
listName.Text = ActiveCell.Value
End Sub

Sub sharepointButton()
Dim i As Integer 'count sheets
Dim x As Integer 'count sheets
Dim j As Integer 'count tables
Dim k As Integer 'count rows
Dim colCnt As Integer 'count columns
Dim newWB As Workbook 'New workbook
Dim filePath As String ' file path
Dim fileName As String 'file name
Dim cutPath As String 'cut the path for sharepoint
colCnt = -1
Num = 0

' Cuts the path for Sharepoint
cutPath = ActiveWorkbook.Sheets("Master").Range("O8").Text & "\All Files\"
' Define the desired file name and path
fileName = "Sharepoint Upload.xlsx"
' This example saves it to the user's Desktop
filePath = ActiveWorkbook.Sheets("Master").Range("O8").Text & "\" & fileName
For i = 1 To ActiveWorkbook.Sheets.Count
    If Sheets(i).Name = "Prepared for Sharepoint" Then
        anSwer = MsgBox(Sheets(i).Name & " alreadyexist do you want to create a new one?", vbYesNo, "Sheet Exist")
        If anSwer = vbYes Then
            Application.DisplayAlerts = False
            Sheets("Prepared for Sharepoint").Activate
            Sheets("Prepared for Sharepoint").Delete
            Num = ActiveWorkbook.Sheets.Count
            Application.DisplayAlerts = True
        ElseIf anSwer = vbNo Then
            Sheets(i).Visible = True
            Sheets(i).Activate
            MsgBox ("This sheet has been brought to the front")
            anSwer = MsgBox("Would you like to make this sheet Sharepoint ready?", vbYesNo, "Prepare for Sharepoint")
            If anSwer = vbYes Then
                Num = ActiveWorkbook.Sheets.Count
            ElseIf anSwer = vbNo Then
                Exit Sub
            End If
        End If
'Check to see if Sharepoint pages exists beginning----------------------------------------
Else
    Num = Num + 1
End If
Next i
'Check to see if Sharepoint pages exists ending----------------------------------------
If Num = ActiveWorkbook.Sheets.Count Then
    Dim arrayFullCombined 'All info with combined dates and times
    
    ActiveWorkbook.Sheets.Add after:=Sheets(Sheets.Count)
    Num = Num + 1
    ActiveWorkbook.Sheets(Num).Name = "Prepared for Sharepoint"
    'Count all the rows in entire workbook beginning------------------------------
        For x = 1 To ActiveWorkbook.Sheets.Count
            If Sheets(x).Name <> "Master" And _
            Sheets(x).Name <> "Key Words" And _
            Sheets(x).Name <> "Prepared for Sharepoint" And _
            Sheets(x).Name <> "For Printing" And _
            Sheets(x).Name <> "Before Divorce" Then
                For j = 1 To ActiveWorkbook.Sheets(x).ListObjects.Count
                    For k = 1 To ActiveWorkbook.Sheets(x).ListObjects(j).ListRows.Count
                        Numb = Numb + 1
                    Next k
                Next j
            End If
        Next x
            'Count all the rows in entire workbook ending------------------------------
    ReDim arrayFullCombined(1 To Numb, 1 To 8) 'All info with combined dates and times
    'Populate array beginning------------------------------------------------------
    Numr = 1 'start the count of number of open slots in array
    For i = 1 To ActiveWorkbook.Sheets.Count
        If Sheets(i).Name <> "Master" And _
        Sheets(i).Name <> "Key Words" And _
        Sheets(i).Name <> "Prepared for Sharepoint" And _
        Sheets(i).Name <> "For Printing" And _
        Sheets(i).Name <> "Before Divorce" Then
            Sheets(i).Visible = True
            Sheets(i).Activate
            For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
                For k = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                    ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Select
                    arrayFullCombined(Numr, 1) = "mborycz-" & Numr 'row number
                    arrayFullCombined(Numr, 2) = ActiveWorkbook.ActiveSheet.ListObjects(j).Name 'catagory
                    arrayFullCombined(Numr, 3) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 2) & " " & _
                    Format(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 3), "hh:mm") ' start date plus start time
                    arrayFullCombined(Numr, 4) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 4) & " " & _
                    Format(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 5), "hh:mm") 'end date plus end time
                    arrayFullCombined(Numr, 5) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 6) 'description
                    arrayFullCombined(Numr, 6) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 7) 'notes
                    arrayFullCombined(Numr, 7) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 8) 'keyword
                    arrayFullCombined(Numr, 8) = Replace(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(k, 9), cutPath, "") 'link
                    Numr = Numr + 1 'counts number of open slots in array
                Next k
            Next j
        End If
    Next i
    'Populate array ending---------------------------------------------------------
    Sheets(Num).Visible = True
    Sheets(Num).Activate
    'Make Sharepoint table beginning----------------------------------------------------
                With ActiveWorkbook.ActiveSheet
                    .Cells(1, 1).Value = "#"
                    .Cells(1, 2).Value = "Catagory"
                    .Cells(1, 3).Value = "Start Event"
                    .Cells(1, 4).Value = "Finish Event"
                    .Cells(1, 5).Value = "Description"
                    .Cells(1, 6).Value = "Notes"
                    .Cells(1, 7).Value = "Key words"
                    Cells(1, 8).Value = "Links"
                    .ListObjects.Add(xlSrcRange, Cells(1, 1).CurrentRegion, , xlYes).Name = "PreparedforSharepoint"
                End With
                ActiveWorkbook.ActiveSheet.ListObjects("PreparedforSharepoint").ListRows.Add
                Cells(2, 1) = "blank table"
                Cells(2, 2) = "blank table"
                Cells(2, 3) = "blank table"
                Cells(2, 4) = "blank table"
                Cells(2, 5) = "blank table"
                Cells(2, 6) = "blank table"
                Cells(2, 7) = "blank table"
                Cells(2, 8) = "blank table"
    'Make Sharepoint table ending------------------------------------------------------
    'Populate Sharepoint table beginning-------------------------------------------------
    For j = 0 To Sheets(Num).ListObjects("PreparedforSharepoint").ListColumns.Count
        colCnt = colCnt + 1
    Next j
    For j = 2 To Numr
        'Cells(j, colCnt - 8) = arrayFullCombined(j - 1, colCnt - 8)
        Cells(j, colCnt - 7) = arrayFullCombined(j - 1, colCnt - 7)
        Cells(j, colCnt - 6) = arrayFullCombined(j - 1, colCnt - 6)
        Cells(j, colCnt - 5) = arrayFullCombined(j - 1, colCnt - 5)
        Cells(j, colCnt - 4) = arrayFullCombined(j - 1, colCnt - 4)
        Cells(j, colCnt - 3) = arrayFullCombined(j - 1, colCnt - 3)
        Cells(j, colCnt - 2) = arrayFullCombined(j - 1, colCnt - 2)
        Cells(j, colCnt - 1) = arrayFullCombined(j - 1, colCnt - 1)
        Cells(j, colCnt) = arrayFullCombined(j - 1, colCnt)
    Next j
    'Populate Sharepoint table ending---------------------------------------------------
    'Create new workbook and put table in it beginning------------------------------------

    ActiveWorkbook.Sheets("Prepared for Sharepoint").Cells(1, 1).CurrentRegion.Copy

    
    ' Create a new workbook and assign it to a variable
    Set newWB = Workbooks.Add
    
    ' Save the new workbook with the specified name and location
    ' Application.DisplayAlerts = False prevents prompts if the file already exists (optional)
    Application.DisplayAlerts = False
    newWB.Sheets(1).Name = ("For Sharepoint")
    newWB.Sheets("For Sharepoint").Cells(1, 1).PasteSpecial xlPasteAll
    newWB.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    ' Optional: Close the new workbook and provide confirmation message
    ' newWB.Close SaveChanges:=False
    MsgBox "New workbook created and named: " & fileName
    'Create new workbook and put table in it ending--------------------------------------
    
End If


End Sub
