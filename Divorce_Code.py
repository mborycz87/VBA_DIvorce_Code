Dim table As ListObject
Dim Sh As Worksheet
'Dim listName As MSForms.TextBox
'Dim Num As Integer

Dim CheckAbbr As MSForms.ComboBox, CheckDesc As MSForms.ComboBox, CheckNotes As MSForms.ComboBox, _
cbxNote As MSForms.ComboBox, tblCheck As MSForms.ComboBox, CheckStartDate As MSForms.ComboBox, _
CheckStartTime As MSForms.ComboBox, CheckEndDate As MSForms.ComboBox

Dim EnterAbbr As MSForms.TextBox, EnterDesc As MSForms.TextBox, EnterNotes As MSForms.TextBox, _
startDate As MSForms.TextBox, startTime As MSForms.TextBox, EnterEndDate As MSForms.TextBox, endDate As MSForms.TextBox, _
Abbreviate As MSForms.TextBox, Description As MSForms.TextBox, Notes As MSForms.TextBox, _
listName As MSForms.TextBox, tbxNote1 As MSForms.TextBox, tbxNote2 As MSForms.TextBox, tbxNote3 As MSForms.TextBox, _
tbxNote4 As MSForms.TextBox, tbxNote5 As MSForms.TextBox, tbxNote6 As MSForms.TextBox

Dim lbAllData As MSForms.ListBox

Dim anSwer As String, inptAns As String, Str As String, Str1 As String
Dim indRow As Integer, indCol As Integer, cellNum As Integer, Num As Integer, Numb As Integer, Numr As Integer, Numm As Integer, _
Numbb As Integer, Numrr As Integer

Dim cell As Variant

Dim x As Integer

Option Explicit

Private IsArrow As Boolean

Private Sub cb_BigDelete_Click()
If Me.cb_BigDelete.Caption = "Delete Selected" Then
    With Me.cb_BigDelete
        .Caption = "Delete All"
        .ForeColor = &H0&
        .BackColor = &HFF&
    End With
    With cb_KeyWordsDelete
        .ForeColor = &H0&
        .BackColor = &HFF&
    End With
ElseIf Me.cb_BigDelete.Caption = "Delete All" Then
    With Me.cb_BigDelete
        .Caption = "Delete Selected"
        .ForeColor = &HFFFFFF
        .BackColor = &HC0C000
    End With
    With cb_KeyWordsDelete
        .ForeColor = &H40C0&
        .BackColor = &HE0E0E0
    End With
End If
End Sub

Private Sub cb_CheckOrDelete_Click()
If Me.cb_CheckOrDelete.Caption = "Check Mode" Then
    With Me.cb_CheckOrDelete
        .Caption = "Delete Mode"
        .ForeColor = &H0&
        .BackColor = &HFF&
    End With
ElseIf Me.cb_CheckOrDelete.Caption = "Delete Mode" Then
    With Me.cb_CheckOrDelete
        .Caption = "Switch Mode"
        .ForeColor = &HFFFF&
        .BackColor = &HC000&
    End With
ElseIf Me.cb_CheckOrDelete.Caption = "Switch Mode" Then
    With Me.cb_CheckOrDelete
        .Caption = "Check Mode"
        .ForeColor = &HFFFFFF
        .BackColor = &HC0C000
    End With
End If
End Sub

Private Sub cb_CopytoExcel_Click()
Dim j As Integer
Dim i As Integer
Dim ii As Integer
Dim jj As Integer
Dim xx As Integer
Dim rowCnt As Integer
Dim colCnt As Integer
Dim anSwer As String
Dim folderPath As String
Dim MyFSO As New FileSystemObject
Dim coPier
x = 1
j = 0
i = 0
ii = 0
jj = 0
xx = 0
rowCnt = 0
colCnt = -1


anSwer = InputBox("What should the name of the folder for files be named?", "Enter Folder Name")

folderPath = ThisWorkbook.Worksheets("Master").Range("O8") & "\" & anSwer
If MyFSO.FolderExists(folderPath) = False Then
    MyFSO.CreateFolder (folderPath)
Else
    MsgBox "Those folder already exists"
End If

Dim printArray() As Variant

For j = 0 To Me.lb_Search.ListCount - 1
    rowCnt = rowCnt + 1
Next j

For j = 0 To Me.lb_Search.ColumnCount
    colCnt = colCnt + 1
Next j


ReDim fullArray(1 To rowCnt, 1 To colCnt - 1)

For ii = 2 To ActiveWorkbook.Worksheets.Count
    If Sheets(ii).Name = "For Printing" Then
        Sheets(ii).Delete
        GoTo out
    End If
Next ii
out:
ii = 0

Sheets.Add.Name = "For Printing"
Sheets("For Printing").Visible = True

 If ActiveWorkbook.ActiveSheet.Cells(1, x) = "" Then
    With ActiveWorkbook.ActiveSheet
        .Cells(1, x).Value = "Table"
        .Cells(1, x + 1).Value = "#"
        .Cells(1, x + 2).Value = "Start Date"
        .Cells(1, x + 3).Value = "Start Time"
        .Cells(1, x + 4).Value = "End Date"
        .Cells(1, x + 5).Value = "End Time"
        .Cells(1, x + 6).Value = "Description"
        .Cells(1, x + 7).Value = "Notes"
        .Cells(1, x + 8).Value = "Link"
        .ListObjects.Add(xlSrcRange, Cells(1, x).CurrentRegion, , xlYes).Name = "For Printing"
    End With
        ActiveWorkbook.ActiveSheet.ListObjects("For Printing").ListRows.Add
        Cells(2, x) = "blank table"
        Cells(2, x + 1) = "blank table"
        Cells(2, x + 2) = "blank table"
        Cells(2, x + 3) = "blank table"
        Cells(2, x + 4) = "blank table"
        Cells(2, x + 5) = "blank table"
        Cells(2, x + 6) = "blank table"
        Cells(2, x + 6).ColumnWidth = 30
        Cells(2, x + 7) = "blank table"
        Cells(2, x + 7).ColumnWidth = 30
        Cells(2, x + 8) = "blank table"
    ElseIf ActiveWorkbook.ActiveSheet.Cells(1, x) <> "" Then
        x = x + 9
End If

Worksheets("For Printing").Activate
'For ii = 2 To ActiveWorkbook.Worksheets.Count
'    Worksheets(ii).Visible = True
'    Worksheets(ii).Activate
'    If ActiveWorkbook.ActiveSheet.ListObjects.Count > 0 Then
'        For jj = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
'            If Me.lb_Search.List(Me.lb_Search.ListIndex, 0) = ActiveWorkbook.ActiveSheet.ListObjects(jj).Name Then
'                ActiveWorkbook.ActiveSheet.ListObjects(jj).Range.Activate
'                Set table = ActiveWorkbook.ActiveSheet.ListObjects(jj)
                For i = 1 To rowCnt
'                    For xx = 1 To table.ListRows.Count
                        'If Me.lb_Search.List(Me.lb_Search.ListIndex, 1) = CStr(table.DataBodyRange(xx, 1)) Then 'And Me.lb_Search.List(Me.lb_Search.ListIndex, 2) = Str(table.DataBodyRange(xx, 2)) And Me.lb_Search.List(Me.lb_Search.ListIndex, 4) = table.DataBodyRange(xx, 4) And Me.lb_Search.List(Me.lb_Search.ListIndex, 6) = table.DataBodyRange(xx, 6) Then
                            fullArray(i, 1) = Me.lb_Search.List(i - 1, 0)
                            fullArray(i, 2) = Me.lb_Search.List(i - 1, 1)
                            fullArray(i, 3) = Me.lb_Search.List(i - 1, 2)
                            fullArray(i, 4) = Me.lb_Search.List(i - 1, 3)
                            fullArray(i, 5) = Me.lb_Search.List(i - 1, 4)
                            fullArray(i, 6) = Me.lb_Search.List(i - 1, 5)
                            fullArray(i, 7) = Me.lb_Search.List(i - 1, 6)
                            fullArray(i, 8) = Me.lb_Search.List(i - 1, 7)
                            fullArray(i, 9) = Me.lb_Search.List(i - 1, 9) 'table.DataBodyRange(i, 9)
                        'End If
'                    Next xx
                Next i
'            End If
'        Next jj
'    End If
'Next ii

Workbooks.Add
ActiveWorkbook.SaveAs fileName:=anSwer & ".xlsx"

ThisWorkbook.Worksheets("For Printing").Activate
Set coPier = CreateObject("Scripting.FileSystemObject")
For i = 1 To rowCnt
    If fullArray(i, 9) <> "" Then
        coPier.CopyFolder Source:=fullArray(i, 9), Destination:=folderPath & "\"
        fullArray(i, 9) = Replace(fullArray(i, 9), "All Files", anSwer) 'Change path
    End If
Next i
'For i = 1 To rowCnt
'    If fullArray(i, 9) <> "" Then
'        coPier.CopyFolder Source:=fullArray(i, 9), Destination:="\"
'        fullArray(i, 9) = Replace(fullArray(i, 9), folderPath, anSwer) 'Change path
'    End If
'Next i
For j = 2 To rowCnt + 1
    Cells(j, colCnt - 9) = fullArray(j - 1, colCnt - 9)
    Cells(j, colCnt - 8) = fullArray(j - 1, colCnt - 8)
    Cells(j, colCnt - 7) = fullArray(j - 1, colCnt - 7)
    Cells(j, colCnt - 6) = fullArray(j - 1, colCnt - 6)
    Cells(j, colCnt - 5) = fullArray(j - 1, colCnt - 5)
    Cells(j, colCnt - 4) = fullArray(j - 1, colCnt - 4)
    Cells(j, colCnt - 3) = fullArray(j - 1, colCnt - 3)
    Cells(j, colCnt - 2) = fullArray(j - 1, colCnt - 2)
    Cells(j, colCnt - 1) = fullArray(j - 1, colCnt - 1)
Next j
Set table = ActiveWorkbook.ActiveSheet.ListObjects("For Printing")
For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects("For Printing").ListRows.Count
    If table.ListColumns(9).DataBodyRange(x, 1) <> "" Then
        With ActiveWorkbook.ActiveSheet
            .Hyperlinks.Add anchor:=table.ListColumns(9).DataBodyRange(x, 1), _
            Address:=table.ListColumns(9).DataBodyRange(x, 1), _
            TextToDisplay:="Link"
        End With
    End If
Next x
Worksheets("For Printing").Cells.WrapText = True
Sheets("For Printing").Copy Before:=Workbooks(anSwer & ".xlsx").Sheets(1)
ThisWorkbook.Worksheets("For Printing").Delete
End Sub

Private Sub cb_CreateShSave_Click()
ActiveWorkbook.Save
MsgBox Prompt:="Saved Events Sheet", Title:="Work Saved"
End Sub

Private Sub cb_CreateTblSave_Click()
ActiveWorkbook.Save
MsgBox Prompt:="Saved Events", Title:="Work Saved"
End Sub

Private Sub cb_FolderEnter_Click()
Set tbxNote1 = Me.tb_StartDateEnter
Set tbxNote2 = Me.tb_EndDateEnter
Set tbxNote3 = Me.tb_StartTimeEnter
Set tbxNote4 = Me.tb_EnterAbbr
Call folderCreateAndSearch(tbxNote1, tbxNote2, tbxNote3, tbxNote4)
End Sub

Private Sub cb_FolderSearch_Click()
Set tbxNote1 = Me.tb_StartDate
Set tbxNote2 = Me.tb_EndDate
Set tbxNote3 = Me.tb_StartTime
Set tbxNote4 = Me.tb_Abbreviate
Dim anSwer As Integer
Call folderCreateAndSearch(tbxNote1, tbxNote2, tbxNote3, tbxNote4)
Dim timeOne As String
Dim timeTwo As String
Dim dateOne As String
Dim dateTwo As String
Dim finalDateOne As String
Dim finalDateTwo As String
Dim pAth As String
Numb = 0

If Me.cb_FolderSearch.Caption = "Nothing Selected" Then
    MsgBox "Select an item from the ListBox to detect a folder"
    Exit Sub
ElseIf Me.cb_FolderSearch.Caption = "No Folder" Then
    anSwer = MsgBox("Add a folder for this selection?", vbQuestion + vbYesNo)
    If anSwer = vbYes Then
        With ActiveWorkbook.ActiveSheet
            table.ListColumns(9).DataBodyRange(Numb, 1) = pAth
        End With
    ElseIf anSwer = vbNo Then
        Exit Sub
    End If
ElseIf Me.cb_FolderSearch.Caption = "Folder Exists" Then
    If table.ListRows.Count = 1 Then
        Numb = Numb + 1
    Else
    For x = 1 To Me.tb_ColNum
        Numb = Numb + 1
    Next x
    End If
    If table.ListColumns(9).DataBodyRange(Numb, 1).Text = "" Then
        With ActiveWorkbook.ActiveSheet
            ActiveWorkbook.ActiveSheet.ListObjects(table).ListColumns(9).DataBodyRange(Numb, 1) = pAth
        End With
    ElseIf table.ListColumns(9).DataBodyRange(Numb, 1).Text <> "" Or table.ListColumns(9).DataBodyRange(Numb, 1).Text <> "blank table" Then
        MsgBox ("A hyperlink already exist for this folder")
        Exit Sub
    End If
End If
End Sub

Private Sub CB_Instructions_Click()
MsgBox Prompt:="Welcome to your reference helper." _
& vbNewLine & "This program is for personal notes" _
& vbNewLine & vbNewLine & "On the previous page, 'Type of Job', there were instructions on how to" _
& vbNewLine & "create a job, or sheets in excel, if you didn't do that please go back and create a job. The 'Table' portion is to create tables in those sheets." _
& vbNewLine & "So in each sheet created (ex. TTH600) there would be tables (ex. single lite, four lite)" _
& vbNewLine & vbNewLine & "Steps:" _
& vbNewLine & "1)Create a table in the 'New Table' field." _
& vbNewLine & "2)Select that table from the 'Choose Table' field" _
& vbNewLine & vbNewLine & "To write notes in that table:" _
& vbNewLine & "3)In the 'Enter Job' box put in a job number, job description, and any notes you prefer, then press 'Enter Data'" _
& vbNewLine & "(Note: it is ok to leave a field empty, it can be filled later" _
& vbNewLine & "4)In the 'Check for Job' box find what was just entered, then press 'Search'" _
& vbNewLine & "5)Double click one of the 3 large boxes, see what happens" _
& vbNewLine & "6)To delete a line, double click from the list on the very bottom" _
& vbNewLine & "7)Press Save" _
& vbNewLine & "Note: 'Delete Table' will delete an entire table.", Title:="Instructions"
End Sub

Private Sub cb_KeyWords_Change()
Dim i  As Integer
Dim j As Integer
Num = 0

For i = 1 To ActiveWorkbook.Sheets.Count
    If ActiveWorkbook.Sheets(i).Name = "Key Words" Then
        Sheets(i).Visible = True
        Sheets(i).Activate
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            If ActiveWorkbook.ActiveSheet.ListObjects(j).Name = "KeyWords" Then
                For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                    If Me.cb_KeyWords.Text = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 1) Then
                        Me.tb_KeyWords.Text = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 2)
                        Exit Sub
                    Else
                        Me.tb_KeyWords.Text = ""
                    End If
                Next x
            End If
        Next j
        Worksheets("Master").Visible = True
        Sheets(i).Visible = False
    End If
Next i
End Sub

Private Sub cb_KeyWordsAdd_Click()
Dim ii As Integer
Dim j As Integer
Dim jj As Integer
Dim listInd As Integer
Dim numOfRows As Integer
listInd = 0

If Me.cb_KeyWords.Text = "" Then ' used to be cb_KeyWords
    MsgBox "Please put in a keyword to add"
    Exit Sub
End If

'Set table = ActiveWorkbook.ActiveSheet.ListObjects(lb_Search.Selected(ii))
For ii = 1 To ActiveWorkbook.Sheets.Count
    If ActiveWorkbook.Sheets(ii).Name = "Key Words" Then
        Sheets(ii).Visible = True
        Sheets(ii).Activate
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            If ActiveWorkbook.ActiveSheet.ListObjects(j).Name = "KeyWords" Then
                For jj = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                    If Me.cb_KeyWords.Text <> ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(jj, 1) Then
                        listInd = listInd + 1
                    End If
                Next jj
            End If
        Next j
    End If
 Next ii
If listInd = jj - 1 Then
    listInd = 1
    anSwer = InputBox("This is a new key word. Please insert a new definition for later reference.")
    If anSwer = "" Then
        MsgBox "Nothing has been added"
        Exit Sub
    ElseIf anSwer <> "" Then
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            If ActiveWorkbook.ActiveSheet.ListObjects(j).Name = "KeyWords" Then
                For jj = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                    listInd = listInd + 1
                Next jj
                ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(listInd, 1) = Me.cb_KeyWords.Text
                ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(listInd, 2) = anSwer
                Me.tb_KeyWords.Text = anSwer
                Call keyWordsIntoCombobox
                Exit Sub
            End If
        Next j
    End If
End If
If Me.cb_Switch.Caption = "Click for Multi Select" Then
    MsgBox "Please select the button that says Click for Multi Select"
    Exit Sub
ElseIf Me.cb_Switch.Caption = "Click for Single Select" Then
    For j = 2 To ActiveWorkbook.Worksheets.Count
        Worksheets(j).Activate
        For jj = 1 To Sheets(j).ListObjects.Count
            numOfRows = Sheets(j).ListObjects(jj).ListRows.Count
            listInd = 0
            For x = 1 To numOfRows
                For listInd = 0 To lb_Search.ListCount - 1
                    If lb_Search.List(listInd, 0) = Sheets(j).ListObjects(jj) And lb_Search.List(listInd, 1) = x And lb_Search.Selected(listInd) = True Then
                            If InStr(1, Sheets(j).ListObjects(jj).DataBodyRange(x, 8), cb_KeyWords.Text) Then
                                MsgBox "This keyword already exist for Table " & Sheets(j).ListObjects(jj).Name & " row " & Sheets(j).ListObjects(jj).DataBodyRange(x, 1)
                            ElseIf Sheets(j).ListObjects(jj).DataBodyRange(x, 8) = "" Then
                                Sheets(j).ListObjects(jj).DataBodyRange(x, 8) = cb_KeyWords.Text & ";"
                            Else
                                Sheets(j).ListObjects(jj).DataBodyRange(x, 8) = Sheets(j).ListObjects(jj).DataBodyRange(x, 8) & cb_KeyWords.Text & ";"
                            End If
                        End If
                    Next listInd
                Next x
            Next jj
        Next j
    End If
End Sub

Private Sub cb_KeyWordsDelete_Click()
Dim ii As Integer
Dim j As Integer
Dim jj As Integer
Dim listInd As Integer
Dim numOfRows As Integer
Dim Str As String

'Set table = ActiveWorkbook.ActiveSheet.ListObjects(lb_Search.Selected(ii))
If Me.cb_BigDelete.Caption = "Delete Selected" Then
    If Me.cb_KeyWords.Text = "" Then
        MsgBox "Please put in a keyword to delete"
        Exit Sub
    End If
    For j = 2 To ActiveWorkbook.Worksheets.Count
        Worksheets(j).Activate
        For jj = 1 To Sheets(j).ListObjects.Count
            numOfRows = Sheets(j).ListObjects(jj).ListRows.Count
            listInd = 0
            For x = 1 To numOfRows
                For listInd = 0 To lb_Search.ListCount - 1
                Sheets(j).Visible = True
                Sheets(j).Activate
                Sheets(j).ListObjects(jj).DataBodyRange(x, 8).Select
                    If lb_Search.List(listInd, 0) = Sheets(j).ListObjects(jj) And lb_Search.List(listInd, 1) = x And lb_Search.Selected(listInd) = True Then
                        If InStr(1, Sheets(j).ListObjects(jj).DataBodyRange(x, 8), cb_KeyWords.Text) <> 0 Then
                            Str = Sheets(j).ListObjects(jj).DataBodyRange(x, 8)
                            If InStr(Str, Me.cb_KeyWords.Text & "-yes;") > 0 Then
                                Str1 = Replace(Str, Me.cb_KeyWords.Text & "-yes;", "")
                            ElseIf InStr(Str, Me.cb_KeyWords.Text & "-maybe;") > 0 Then
                                Str1 = Replace(Str, Me.cb_KeyWords.Text & "-maybe;", "")
                            ElseIf InStr(Str, Me.cb_KeyWords.Text & "-no;") > 0 Then
                                Str1 = Replace(Str, Me.cb_KeyWords.Text & "-no;", "")
                            ElseIf InStr(Str, Me.cb_KeyWords.Text) > 0 Then
                                Str1 = Replace(Str, Me.cb_KeyWords.Text & ";", "")
                            End If
                            Me.tb_KeyWordsUsed.Text = Str1
                            Sheets(j).ListObjects(jj).DataBodyRange(x, 8) = Str1
                        Else
                            MsgBox "This keyword does not exist"
                        End If
                    End If
                Next listInd
            Next x
        Next jj
    Next j
ElseIf Me.cb_BigDelete.Caption = "Delete All" Then
    For j = 2 To ActiveWorkbook.Worksheets.Count
        Worksheets(j).Activate
        For jj = 1 To Sheets(j).ListObjects.Count
            For x = 1 To Sheets(j).ListObjects(jj).ListRows.Count
                If InStr(1, Sheets(j).ListObjects(jj).DataBodyRange(x, 8), cb_KeyWords.Text & ";") <> 0 Then
                    Str = Replace(Sheets(j).ListObjects(jj).DataBodyRange(x, 8), cb_KeyWords.Text & ";", "", 1)
                    Sheets(j).ListObjects(jj).DataBodyRange(x, 8) = Str
                End If
            Next x
        Next jj
        Worksheets(j).Visible = False
    Next j
End If
End Sub

Private Sub cb_KeyWordsSearch_Click()
Dim i As Integer
Dim x As Integer
Dim k As Integer
Dim ii As Integer
Dim j As Integer
Dim jj As Integer
Dim numOfRows As Integer
Dim rowsInListBox As Integer
Dim Str As String
Numb = 0
rowsInListBox = -1
Dim searchArray()

If Me.cb_KeyWords.Text = "" Then
    MsgBox "Please put in a keyword to search"
    Exit Sub
End If

For i = 1 To Sheets.Count
    Sheets(i).Visible = True
    Sheets(i).Select
    For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        For k = 1 To ActiveWorkbook.ActiveSheet.ListObjects(x).ListRows.Count
            If InStr(1, Sheets(i).ListObjects(x).DataBodyRange(k, 8), cb_KeyWords.Text) <> 0 Then
                rowsInListBox = rowsInListBox + 1
            End If
        Next k
    Next x
Next i
If rowsInListBox = 0 Then
    MsgBox "This data does not exist"
    Exit Sub
End If
ReDim searchArray(0 To rowsInListBox, 0 To 9)

If Me.cb_KeyWords.Text = "" Then
    MsgBox "Please enter data into the field"
Else
    For j = 2 To ActiveWorkbook.Worksheets.Count
        Worksheets(j).Visible = True
        Worksheets(j).Activate
        For jj = 1 To Sheets(j).ListObjects.Count
            numOfRows = Sheets(j).ListObjects(jj).ListRows.Count
            For x = 1 To numOfRows
                Sheets(j).ListObjects(jj).DataBodyRange(x, 8).Select
                If InStr(1, Sheets(j).ListObjects(jj).DataBodyRange(x, 8), cb_KeyWords.Text) <> 0 Then
                    searchArray(Numb, 0) = Sheets(j).ListObjects(jj).Name
                    searchArray(Numb, 1) = Sheets(j).ListObjects(jj).DataBodyRange(x, 1)
                    searchArray(Numb, 2) = Sheets(j).ListObjects(jj).DataBodyRange(x, 2)
                    searchArray(Numb, 3) = Format(Sheets(j).ListObjects(jj).DataBodyRange(x, 3), "hh:mm")
                    searchArray(Numb, 4) = Sheets(j).ListObjects(jj).DataBodyRange(x, 4)
                    searchArray(Numb, 5) = Format(Sheets(j).ListObjects(jj).DataBodyRange(x, 5), "hh:mm")
                    searchArray(Numb, 6) = Sheets(j).ListObjects(jj).DataBodyRange(x, 6)
                    searchArray(Numb, 7) = Sheets(j).ListObjects(jj).DataBodyRange(x, 7)
                    searchArray(Numb, 8) = Sheets(j).ListObjects(jj).DataBodyRange(x, 8)
                    searchArray(Numb, 9) = Sheets(j).ListObjects(jj).DataBodyRange(x, 9)
                    Numb = Numb + 1
                End If
            Next x
        Next jj
    Next j
End If
    With Me.lb_Search
        .List = searchArray
        .FontSize = 8
        .ColumnWidths = "75,15,50,30,50,30,112,113,0,0"
    End With
End Sub
Private Sub cb_YesChange_Click()
Str = "-yes"
Call addingForChanges(Str)
End Sub
Private Sub cb_MaybeChange_Click()
Str = "-maybe"
Call addingForChanges(Str)
End Sub

Private Sub cb_NoChange_Click()
Str = "-no"
Call addingForChanges(Str)
End Sub

Private Sub cb_RefindChange_Click()
Dim i As Integer
Dim j As Integer
If Me.lb_Search.ListIndex >= 0 And Me.cb_Switch.Caption = "Click for Multi Select" Then
    Me.MultiPage1.Value = 1
    Me.cbx_TblCheck.Text = Me.lb_Search.List(Me.lb_Search.ListIndex, 0)
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Worksheets(i).Visible = True
        Worksheets(i).Activate
        If ActiveWorkbook.ActiveSheet.ListObjects.Count > 0 Then
            For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
                If Me.cbx_TblCheck.Text = ActiveWorkbook.ActiveSheet.ListObjects(j).Name Then
                    ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Activate
                    Call goToTable
                End If
            Next j
        End If
    Next i
End If

For i = 1 To ActiveWorkbook.Sheets.Count
Worksheets(i).Visible = True
Worksheets(i).Activate
    If ActiveWorkbook.ActiveSheet.ListObjects.Count > 0 = True Then
    Num = 0
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        Num = Num + 1
        If ActiveWorkbook.ActiveSheet.ListObjects(Num).Name = cbx_TblCheck Then
            ActiveWorkbook.ActiveSheet.ListObjects(Num).Range.Activate
            Exit Sub
        End If
    Next j
    End If
Next i
End Sub

Private Sub cb_RefindEndtDate_Click()
Dim xx As Integer
Set tbxNote1 = tb_RefindEndtDate
xx = 4
Call Search_Array(xx, tbxNote1)
End Sub

Private Sub cb_RefindNotes_Click()
Dim xx As Integer
Set tbxNote1 = tb_RefindNotes
xx = 7
Call Search_Array(xx, tbxNote1)
End Sub

Private Sub cb_RefindShCheck_Click()
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim fullArray As Variant
Num = 0
Numb = 1
For i = 1 To ActiveWorkbook.Sheets.Count
    Sheets(i).Visible = True
    Sheets(i).Activate
    If ActiveWorkbook.Sheets(i).Name = Me.cbx_RefineShCheck.Text Then
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                Num = Num + 1
            Next x
        Next j
    End If
    Worksheets("Master").Visible = True
Next i

ReDim fullArray(1 To Num, 1 To 10)

For i = 1 To ActiveWorkbook.Sheets.Count
    Sheets(i).Visible = True
    Sheets(i).Activate
    If ActiveWorkbook.Sheets(i).Name = Me.cbx_RefineShCheck.Text Then
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Select
                fullArray(Numb, 1) = ActiveWorkbook.ActiveSheet.ListObjects(j).Name
                fullArray(Numb, 2) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 1)
                fullArray(Numb, 3) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 2)
                fullArray(Numb, 4) = Format(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 3), "hh:mm")
                fullArray(Numb, 5) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 4)
                fullArray(Numb, 6) = Format(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 5), "hh:mm")
                fullArray(Numb, 7) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 6)
                fullArray(Numb, 8) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 7)
                fullArray(Numb, 9) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 8)
                fullArray(Numb, 10) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 9)
                Numb = Numb + 1
            Next x
        Next j
    End If
Next i
With Me.lb_Search
    .List = fullArray
    .FontSize = 8
    .ColumnWidths = "75,15,50,30,50,30,112,113,0,0"
End With
End Sub

Private Sub cb_RefindStartDate_Click()
Dim xx As Integer
Set tbxNote1 = tb_RefindStartDate
xx = 2
Call Search_Array(xx, tbxNote1)
End Sub



Private Sub cb_RefindStartDescription_Click()
Dim xx As Integer
Set tbxNote1 = tb_RefindStartDescription
xx = 6
Call Search_Array(xx, tbxNote1)
End Sub

Private Sub cb_RefindTblCheck_Click()

Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim fullArray As Variant
Num = 0
Numb = 1
For i = 1 To ActiveWorkbook.Sheets.Count
    Sheets(i).Visible = True
    Sheets(i).Activate
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        If ActiveWorkbook.ActiveSheet.ListObjects(j).Name = Me.cbx_RefineTblCheck.Text Then
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                Num = Num + 1
            Next x
        End If
    Next j
    If Worksheets("Master").Visible = False Then
        Worksheets("Master").Visible = True
        Sheets(i).Visible = False
    End If
Next i

ReDim fullArray(1 To Num, 1 To 9)

For i = 1 To ActiveWorkbook.Sheets.Count
    Sheets(i).Visible = True
    Sheets(i).Activate
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        If ActiveWorkbook.ActiveSheet.ListObjects(j).Name = Me.cbx_RefineTblCheck.Text Then
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Select
                fullArray(Numb, 1) = ActiveWorkbook.ActiveSheet.ListObjects(j).Name
                fullArray(Numb, 2) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 1)
                fullArray(Numb, 3) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 2)
                fullArray(Numb, 4) = Format(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 3), "hh:mm")
                fullArray(Numb, 5) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 4)
                fullArray(Numb, 6) = Format(ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 5), "hh:mm")
                fullArray(Numb, 7) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 6)
                fullArray(Numb, 8) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 7)
                fullArray(Numb, 9) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 8)
                'fullArray(Numb, 10) = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 9)
                Numb = Numb + 1
            Next x
        End If
    Next j
Next i
With Me.lb_Search
    .List = fullArray
    .FontSize = 8
    .ColumnWidths = "75,15,50,30,50,30,112,113,0,0"
End With
End Sub

Private Sub cb_RefindToFromDate_Click()

Dim Num As Integer
Dim Numb As Integer
Dim Numr As Integer
Dim Numm As Integer
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim jj As Integer
Dim stDate As Date
Dim endDate As Date

Dim fullArray As Variant
If tb_RefindStartDate.Text & tb_RefindEndtDate.Text = "" Then
    MsgBox "Please enter dates in both the Start Date and End Date boxes"
    Exit Sub
ElseIf tb_RefindStartDate.Text = "" Then
    MsgBox "Please enter dates in both the Start Date box"
    Exit Sub
ElseIf tb_RefindEndtDate.Text = "" Then
    MsgBox "Please enter dates in both the End Date box"
    Exit Sub
End If
' Counting sheets and rows in tables beinging----------------------------------------------
Me.lb_Search.Clear
i = 0
j = 0
Numb = 0
For i = 3 To ActiveWorkbook.Sheets.Count
Worksheets(i).Visible = True
Worksheets(i).Activate
    If ActiveWorkbook.ActiveSheet.ListObjects.Count > 0 = True Then
    Num = 0
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
       Num = Num + 1
        x = 0
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(Num).ListRows.Count
            ActiveWorkbook.ActiveSheet.ListObjects(Num).Range.Activate
            stDate = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 2)
            ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 2).Select
            endDate = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 4)
            ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 4).Select
                If stDate >= CDate(tb_RefindStartDate.Text) And endDate <= CDate(tb_RefindEndtDate.Text) Then
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
For i = 3 To ActiveWorkbook.Sheets.Count
Worksheets(i).Visible = True
Worksheets(i).Activate
    If ActiveWorkbook.ActiveSheet.ListObjects.Count > 0 = True Then
    Num = 0
    For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
       Num = Num + 1
        x = 0
            For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(Num).ListRows.Count
            ActiveWorkbook.ActiveSheet.ListObjects(Num).Range.Activate
            stDate = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 2)
            ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 2).Select
            endDate = ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 4)
            ActiveWorkbook.ActiveSheet.ListObjects(Num).DataBodyRange(x, 4).Select
                If stDate >= CDate(tb_RefindStartDate.Text) And endDate <= CDate(tb_RefindEndtDate.Text) Then
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
With Me.lb_Search
    .List = fullArray
    .FontSize = 8
    .ColumnWidths = "75,15,50,30,50,30,112,113,0,0"
End With

End Sub

Private Sub cb_Sort_Click()
Dim i As Long
Dim j As Long
Dim c As Integer
Dim k As Variant
'Sort Date Beginning----------------------------------
With Me.lb_Search
    For i = 0 To .ListCount - 1
        For j = i + 1 To .ListCount - 1
            If DateValue(.List(i, 2)) > DateValue(.List(j, 2)) Then
                For c = 0 To Me.lb_Search.ColumnCount - 1
                    k = .List(i, c)
                    .List(i, c) = .List(j, c)
                    .List(j, c) = k
                Next c
            End If
        Next j
    Next i
End With
'Sort Date Ending----------------------------------
'Sort Time Beginning------------------------------
With Me.lb_Search
    For i = 0 To .ListCount - 1
        For j = i + 1 To .ListCount - 1
            If DateValue(.List(i, 2)) = DateValue(.List(j, 2)) And (.List(i, 3)) > (.List(j, 3)) Then
                For c = 0 To 9
                    k = .List(i, c)
                    .List(i, c) = .List(j, c)
                    .List(j, c) = k
                Next c
            End If
        Next j
    Next i
End With
'Sort Time Ending----------------------------------
End Sub

Private Sub cb_Switch_Click()
If Me.cb_Switch.Caption = "Click for Multi Select" Then
    Me.cb_Switch.Caption = "Click for Single Select"
    Me.lb_Search.MultiSelect = 1
ElseIf Me.cb_Switch.Caption = "Click for Single Select" Then
    Me.cb_Switch.Caption = "Click for Multi Select"
    Me.lb_Search.MultiSelect = fmMultiSelectSingle
End If
End Sub

Private Sub cbx_StartDateCheck_Change()
Set tblCheck = Me.cbx_TblCheck
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
'On Error Resume Next
Dim i As Long
'table.ListRows.Add
    If Not IsArrow Then
        With Me.cbx_StartDateCheck
            .List = table.ListColumns(2).DataBodyRange.Value
            .ListRows = Application.WorksheetFunction.Min(6, .ListCount)
            .DropDown
            If Len(.Text) Then
                For i = .ListCount - 1 To 0 Step -1
                    If InStr(1, .List(i), .Text, vbTextCompare) = 0 Then .RemoveItem i
                Next
                .DropDown
            End If
        End With
    End If

indRow = table.DataBodyRange.Rows.Count

If cbx_StartDateCheck = "" Then
    GoTo thatsIt
End If

For Each cell In table.ListColumns(2).DataBodyRange.Value
    If InStr(cell, cbx_CheckAbbr) = 0 Then
        x = x + 1
    End If
Next cell

If x = table.ListColumns(2).DataBodyRange.Count And cbx_StartDateCheck <> "" Then
MsgBox Prompt:="This data does not exist. Please try again", Title:=" Data not found!!!"
cbx_StartDateCheck = ""
End If

thatsIt:
For i = 1 To indRow
    If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then Rows(i).EntireRow.Delete
Next
tb_StartTime = Format(tb_StartTime, "hh:mm")
tb_Abbreviate = Format(tb_Abbreviate, "hh:mm")
x = 0
endThis:
x = 0
End Sub

Private Sub cbx_StartDateCheck_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 'Goes with checking enddate combobox
On Error Resume Next
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
    IsArrow = (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown)
    If KeyCode = vbKeyReturn Then Me.cbx_StartDateCheck.List = table.ListColumns(2).DataBodyRange.Value
End Sub

Private Sub cbx_StartTimeCheck_Change()
Set tblCheck = Me.cbx_TblCheck
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
'On Error Resume Next
Dim i As Long
'table.ListRows.Add
    If Not IsArrow Then
        With Me.cbx_StartTimeCheck
            .List = table.ListColumns(3).DataBodyRange.Value
            .ListRows = Application.WorksheetFunction.Min(6, .ListCount)
            .DropDown
            If Len(.Text) Then
                For i = .ListCount - 1 To 0 Step -1
                    If InStr(1, .List(i), .Text, vbTextCompare) = 0 Then .RemoveItem i
                Next
                .DropDown
            End If
        End With
    End If

indRow = table.DataBodyRange.Rows.Count

If cbx_StartTimeCheck = "" Then
    GoTo thatsIt
End If

For Each cell In table.ListColumns(3).DataBodyRange.Value
    If InStr(cell, cbx_CheckAbbr) = 0 Then
        x = x + 1
    End If
Next cell

If x = table.ListColumns(2).DataBodyRange.Count And cbx_StartTimeCheck <> "" Then
MsgBox Prompt:="This data does not exist. Please try again", Title:=" Data not found!!!"
cbx_StartTimeCheck = ""
End If

thatsIt:
For i = 1 To indRow
    If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then Rows(i).EntireRow.Delete
Next
tb_StartTime = Format(tb_StartTime, "hh:mm")
tb_Abbreviate = Format(tb_Abbreviate, "hh:mm")
x = 0
endThis:
x = 0
End Sub

Private Sub cbx_StartTimeCheck_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 'Goes with checking enddate combobox
On Error Resume Next
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
    IsArrow = (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown)
    If KeyCode = vbKeyReturn Then Me.cbx_StartTimeCheck.List = table.ListColumns(3).DataBodyRange.Value
End Sub

Private Sub cbx_EndDateCheck_Change()
Set tblCheck = Me.cbx_TblCheck
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
'On Error Resume Next
Dim i As Long
'table.ListRows.Add
    If Not IsArrow Then
        With Me.cbx_EndDateCheck
            .List = table.ListColumns(4).DataBodyRange.Value
            .ListRows = Application.WorksheetFunction.Min(6, .ListCount)
            .DropDown
            If Len(.Text) Then
                For i = .ListCount - 1 To 0 Step -1
                    If InStr(1, .List(i), .Text, vbTextCompare) = 0 Then .RemoveItem i
                Next
                .DropDown
            End If
        End With
    End If

indRow = table.DataBodyRange.Rows.Count

If cbx_EndDateCheck = "" Then
    GoTo thatsIt
End If

For Each cell In table.ListColumns(4).DataBodyRange.Value
    If InStr(cell, cbx_CheckAbbr) = 0 Then
        x = x + 1
    End If
Next cell

If x = table.ListColumns(2).DataBodyRange.Count And cbx_EndDateCheck <> "" Then
MsgBox Prompt:="This data does not exist. Please try again", Title:=" Data not found!!!"
cbx_EndDateCheck = ""
End If

thatsIt:
For i = 1 To indRow
    If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then Rows(i).EntireRow.Delete
Next
x = 0
endThis:
x = 0
End Sub

Private Sub cbx_EndDateCheck_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 'Goes with checking enddate combobox
On Error Resume Next
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
    IsArrow = (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown)
    If KeyCode = vbKeyReturn Then Me.cbx_EndDateCheck.List = table.ListColumns(4).DataBodyRange.Value
End Sub

Private Sub cbx_CheckAbbr_Change() 'Cycle threw checking abbreviation list box
Set tblCheck = Me.cbx_TblCheck
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
'On Error Resume Next
Dim i As Long

'table.ListRows.Add
    If Not IsArrow Then
        With Me.cbx_CheckAbbr
            .List = table.ListColumns(5).DataBodyRange.Value
            debugprint = table.ListColumns(5).DataBodyRange.Value
            .ListRows = Application.WorksheetFunction.Min(6, .ListCount)
            .DropDown
            If Len(.Text) Then
                For i = .ListCount - 1 To 0 Step -1
                    If InStr(1, .List(i), .Text, vbTextCompare) = 0 Then .RemoveItem i
                Next
                .DropDown
            End If
        End With
    End If

indRow = table.DataBodyRange.Rows.Count

If cbx_CheckAbbr = "" Then
    GoTo thatsIt
End If

For Each cell In table.ListColumns(5).DataBodyRange.Value
    If InStr(cell, cbx_CheckAbbr) = 0 Then
        x = x + 1
    End If
Next cell

If x = table.ListColumns(2).DataBodyRange.Count And cbx_CheckAbbr <> "" Then
MsgBox Prompt:="This data does not exist. Please try again", Title:=" Data not found!!!"
cbx_CheckAbbr = ""
End If

thatsIt:
For i = 1 To indRow
    If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then Rows(i).EntireRow.Delete
Next
x = 0
endThis:
x = 0
End Sub

Private Sub cbx_CheckAbbr_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 'Goes with checking abbrevation combobox
On Error Resume Next
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
    IsArrow = (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown)
    If KeyCode = vbKeyReturn Then Me.cbx_CheckAbbr.List = table.ListColumns(5).DataBodyRange.Value
End Sub

Private Sub cbx_CheckDesc_Change() 'Cycles through checking description combobox
Set tblCheck = Me.cbx_TblCheck
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
Dim i As Long

    If Not IsArrow Then
        With Me.cbx_CheckDesc
            .List = table.ListColumns(6).DataBodyRange.Value
            .ListRows = Application.WorksheetFunction.Min(6, .ListCount)
            .DropDown
            If Len(.Text) Then
                For i = .ListCount - 1 To 0 Step -1
                    If InStr(1, .List(i), .Text, vbTextCompare) = 0 Then .RemoveItem i
                Next
                .DropDown
            End If
        End With
    End If
    
indRow = table.DataBodyRange.Rows.Count

If cbx_CheckDesc = "" Then
    GoTo thatsIt
End If

For Each cell In table.ListColumns(6).DataBodyRange.Value
    If InStr(cell, cbx_CheckDesc) = 0 Then
        x = x + 1
    End If
Next cell

If x = table.ListColumns(2).DataBodyRange.Count And cbx_CheckDesc <> "" Then
MsgBox Prompt:="This data does not exist. Please try again", Title:=" Data not found!!!"
cbx_CheckDesc = ""
End If

thatsIt:
For i = 1 To indRow
    If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then Rows(i).EntireRow.Delete
Next
x = 0
endThis:
x = 0
End Sub


Private Sub cbx_CheckDesc_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 'Goes with checking description combobox
On Error Resume Next
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
    IsArrow = (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown)
    If KeyCode = vbKeyReturn Then Me.cbx_CheckDesc.List = table.ListColumns(6).DataBodyRange.Value
End Sub

Private Sub cbx_CheckNotes_Change() 'Cycles through checking notes combobox

On Error Resume Next
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
Dim i As Long

    If Not IsArrow Then
        With Me.cbx_CheckNotes
            .List = table.ListColumns(7).DataBodyRange.Value
            .ListRows = Application.WorksheetFunction.Min(6, .ListCount)
            .DropDown
            If Len(.Text) Then
                For i = .ListCount - 1 To 0 Step -1
                    If InStr(1, .List(i), .Text, vbTextCompare) = 0 Then .RemoveItem i
                Next
                .DropDown
            End If
        End With
    End If
    
indRow = table.DataBodyRange.Rows.Count

If cbx_CheckNotes = "" Then
    GoTo thatsIt
End If

For Each cell In table.ListColumns(7).DataBodyRange.Value
    If InStr(cell, cbx_CheckNotes) = 0 Then
        x = x + 1
    End If
Next cell

If x = table.ListColumns(2).DataBodyRange.Count And cbx_CheckNotes <> "" Then
MsgBox Prompt:="This data does not exist. Please try again", Title:=" Data not found!!!"
cbx_CheckNotes = ""
End If

thatsIt:
For i = 1 To indRow
    If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then Rows(i).EntireRow.Delete
Next
x = 0
endThis:
x = 0
    
End Sub

Private Sub cbx_CheckNotes_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 'Goes with checking notes combobox
On Error Resume Next
Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
    IsArrow = (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown)
    If KeyCode = vbKeyReturn Then Me.cbx_CheckNotes.List = table.ListColumns(7).DataBodyRange.Value
End Sub

Private Sub cb_DeleteSh_Click() 'For deleting sheets
Dim delText As String
Dim cbxDel As ComboBox
Dim cbxShChk As ComboBox

Set cbxDel = uf_Master.cbx_ShDelete
Set cbxShChk = uf_Master.cbx_ShCheck

delText = Me.cbx_ShDelete.Text
If Me.cbx_ShDelete.Text <> "" Then
    If delText <> "Master" Or delText <> "Key Words" Then
        ActiveWorkbook.Worksheets(delText).Delete
    ElseIf delText = "Master" Or delText = "Key Words" Then
        MsgBox Prompt:="This worksheet cannot be deleted", Title:="Action unavailable"
    End If
ElseIf Me.cbx_ShDelete.Text = "" Then
    MsgBox Prompt:="Enter a Manufacturer name to delete", Title:="Manufacturer Required"
End If


cbxShChk.Clear
cbxDel.Clear

Call UserForm_Initialize
Call MultiPage1_Change
End Sub

Private Sub cb_CreateSh_Click() 'For creating sheets
Dim i As Integer
Dim Exists As Boolean
Dim cbxDel As ComboBox, cbxShChk As ComboBox

Set cbxDel = uf_Master.cbx_ShDelete
Set cbxShChk = uf_Master.cbx_ShCheck

If Me.tb_NewShName.Text <> "" Then
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets(i).Name = tb_NewShName.Value Then
            Exists = True
            MsgBox Prompt:="This worksheet already exist.", Title:="Action Unavailable"
            tb_NewShName.Text = ""
        End If
    Next i
        If Not Exists Then
            Sheets.Add(after:=Sheets(Sheets.Count)).Name = tb_NewShName.Value
            ActiveWorkbook.Worksheets(tb_NewShName.Text).Visible = False
            tb_NewShName.Text = ""
        End If
ElseIf Me.tb_NewShName.Text = "" Then
    MsgBox Prompt:="For a new manufacturer sheet" & vbNewLine & "type a name in the field.", Title:="Manufacturer Required"
    GoTo stopPro
End If
cbxDel.Clear
cbxShChk.Clear
    
Call UserForm_Initialize
Call MultiPage1_Change
stopPro:
End Sub

Private Sub cb_DeleteTbl_Click() 'Deletes a table
Dim tblArray, j
Dim i As Long
Dim tblDel As ComboBox, tblCreate As ComboBox
Dim Exist As Boolean

Set tblCreate = Me.cbx_TblCheck
Set tblDel = Me.cbx_TblDelete

If tblDel.Text = "" Then
    MsgBox Prompt:="Enter a table in 'Go to Table' field to delete it", Title:=" Delete Field Empty"
    Exit Sub
End If

If ActiveWorkbook.ActiveSheet.ListObjects.Count > 0 Then
    For i = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
        If ActiveWorkbook.ActiveSheet.ListObjects(i).Name = tblDel.Text Then
            ActiveSheet.ListObjects(tblDel.Text).Range.Select
            Selection.Clear
            tblDel.Text = ""
            GoTo out
        End If
    Next i
End If
out:
tblCreate.Clear
tblDel.Clear
startDate.Text = ""
startTime.Text = ""
EnterEndDate.Text = ""
EnterAbbr.Text = ""
EnterDesc.Text = ""
EnterNotes.Text = ""
lb_AllData.Clear

Call UserForm_Initialize
Call MultiPage1_Change
End Sub

Private Sub cb_CreateTbl_Click() 'Creates a table
Dim NewTblName As MSForms.TextBox
Dim i As Integer, x As Integer, y As Integer
Dim tblName As String
Dim Exists As Boolean

x = 1
y = 1
Set NewTblName = uf_Master.tb_NewTblName

If NewTblName.Text = "" Then
    MsgBox Prompt:="Enter a name in the 'New Table' field" _
    & vbNewLine & "to start working on it." & vbNewLine & vbNewLine & _
    "Or select from the list of existing tables" & vbNewLine & _
    " in the 'Go to Table' field to work on an exisiting table", Title:="Create Table Field Empty"
    Exit Sub
ElseIf ActiveWorkbook.ActiveSheet.Name = "Master" Then
    MsgBox Prompt:="A table cannot be put onto the master page," _
    & vbNewLine & "To create a table choose or create a manufacturer sheet", _
    Title:="Master Page Uneditable"
    NewTblName.Text = ""
    Exit Sub
End If

For i = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
    Exists = True
    If ActiveSheet.ListObjects(i) = NewTblName Then
        MsgBox Prompt:="This table already exists", Title:="Action Unavailable"
        NewTblName.Text = ""
    End If
Next i

    If Not Exists Or NewTblName <> "" Then
        Do Until y = ActiveWorkbook.ActiveSheet.ListObjects.Count + 2
            If ActiveWorkbook.ActiveSheet.Cells(1, x) = "" Then
                With ActiveWorkbook.ActiveSheet
                    .Cells(1, x).Value = "#"
                    .Cells(1, x + 1).Value = "Start Date"
                    .Cells(1, x + 2).Value = "Start Time"
                    .Cells(1, x + 3).Value = "End Date"
                    .Cells(1, x + 4).Value = "End Time"
                    .Cells(1, x + 5).Value = "Description"
                    .Cells(1, x + 6).Value = "Notes"
                    Cells(1, x + 7).Value = "Key Words"
                    Cells(1, x + 8).Value = "Link"
                    .ListObjects.Add(xlSrcRange, Cells(1, x).CurrentRegion, , xlYes).Name = NewTblName.Text
                End With
                ActiveWorkbook.ActiveSheet.ListObjects(NewTblName.Text).ListRows.Add
                Cells(2, x) = "blank table"
                Cells(2, x + 1) = "blank table"
                Cells(2, x + 2) = "blank table"
                Cells(2, x + 3) = "blank table"
                Cells(2, x + 4) = "blank table"
                Cells(2, x + 5) = "blank table"
                Cells(2, x + 6) = "blank table"
                Cells(2, x + 7) = "blank table"
                Cells(2, x + 8) = "blank table"
            ElseIf ActiveWorkbook.ActiveSheet.Cells(1, x) <> "" Then
                x = x + 10
            End If
            y = y + 1
        Loop
    Else
        MsgBox Prompt:="Error table not named", Title:="Action Unavailable"
    End If

Call UserForm_Initialize
Call MultiPage1_Change
Call ch_GoToSh_Click
NewTblName.Text = ""
'uf_Master.cbx_TblCheck.Text = ""
End Sub

Private Sub cb_EnterData_Click() 'Enter data into a table
uf_Master.tb_StartTimeEnter = Left(uf_Master.tb_StartTimeEnter, 2) & ":" & Right(uf_Master.tb_StartTimeEnter, 2)
uf_Master.tb_EnterAbbr = Left(uf_Master.tb_EnterAbbr, 2) & ":" & Right(uf_Master.tb_EnterAbbr, 2)

Dim pAth As String
Dim timeOne As String
Dim timeTwo As String
Dim dateOne As String
Dim dateTwo As String
Dim finalDateOne As String
Dim finalDateTwo As String
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set startDate = uf_Master.tb_StartDateEnter
Set startTime = uf_Master.tb_StartTimeEnter
Set EnterEndDate = uf_Master.tb_EndDateEnter
Set EnterAbbr = uf_Master.tb_EnterAbbr
Set EnterDesc = uf_Master.tb_EnterDesc
Set EnterNotes = uf_Master.tb_EnterNotes
Set tblCheck = uf_Master.cbx_TblCheck

If tblCheck = "" Then
    MsgBox Prompt:="Enter a table to add data to", Title:="Table required"
    startDate.Text = ""
    startTime.Text = ""
    EnterEndDate.Text = ""
    EnterAbbr.Text = ""
    EnterDesc.Text = ""
    EnterNotes.Text = ""
    Exit Sub
End If

Set table = ActiveWorkbook.ActiveSheet.ListObjects(tblCheck.Text)

Num = 2
Str = "Start Date"
Set listName = uf_Master.tb_StartDateEnter
Call GreenDataEntry(Str, Num, listName, table)

Num = 3
Str = "Start Time"
Set listName = uf_Master.tb_StartTimeEnter
Call GreenDataEntry(Str, Num, listName, table)

Num = 4
Str = "End Date"
Set listName = uf_Master.tb_EndDateEnter
Call GreenDataEntry(Str, Num, listName, table)

Num = 5
Str = "End Time"
Set listName = uf_Master.tb_EnterAbbr
Call GreenDataEntry(Str, Num, listName, table)

Num = 6
Str = "Description"
Set listName = uf_Master.tb_EnterDesc
Call GreenDataEntry(Str, Num, listName, table)

Num = 7
Str = "Notes"
Set listName = uf_Master.tb_EnterNotes
Call GreenDataEntry(Str, Num, listName, table)

If table.ListColumns(2).DataBodyRange.Cells(1, 1) = "blank table" Then
    table.ListRows(1).Delete
End If


For cell = 1 To table.ListColumns(1).DataBodyRange.Count
    table.ListColumns(1).DataBodyRange(cell, 1) = ""
Next cell
Num = 1
For cell = 1 To table.ListColumns(2).DataBodyRange.Count
    table.ListColumns(1).DataBodyRange(cell, 1) = Num
    Num = Num + 1
Next cell

Numb = 1
Do Until table.ListColumns(1).DataBodyRange(Numb, 1) = table.ListColumns(1).DataBodyRange.Count
    Numb = Numb + 1
Loop

table.ListColumns(3).DataBodyRange(Numb, 1) = Format(table.ListColumns(3).DataBodyRange(Numb, 1), "hh:mm")
table.ListColumns(5).DataBodyRange(Numb, 1) = Format(table.ListColumns(5).DataBodyRange(Numb, 1), "hh:mm")

Me.tb_ColNum.Text = table.ListColumns(1).DataBodyRange(Numb, 1)
Me.tb_StartDate.Text = table.ListColumns(2).DataBodyRange(Numb, 1)
Me.tb_StartTime.Text = Format(table.ListColumns(3).DataBodyRange(Numb, 1), "hh:mm")
Me.tb_EndDate.Text = table.ListColumns(4).DataBodyRange(Numb, 1)
Me.tb_Abbreviate.Text = Format(table.ListColumns(5).DataBodyRange(Numb, 1), "hh:mm")
Me.tb_Description.Text = table.ListColumns(6).DataBodyRange(Numb, 1)
Me.tb_Notes.Text = table.ListColumns(7).DataBodyRange(Numb, 1)

'pAth = "C:\Users\mludwigborycz\Desktop\Dont Delete\Divorce\All Files" & "\" & Replace(finalDateOne, "/", "-") & " at " & timeOne & " to " & Replace(finalDateTwo, "/", "-") & " at " & timeTwo
If Left(Me.tb_StartTime, 3) = ":" Then
    timeOne = Replace(Me.tb_StartTime, ":", "")
Else
    timeOne = Replace(Me.tb_StartTime, ":", "")
End If
If Left(Me.tb_Abbreviate, 3) = ":" Then
    timeTwo = Replace(Me.tb_Abbreviate, ":", "")
Else
    timeTwo = Replace(Me.tb_Abbreviate, ":", "")
End If
dateOne = startDate
dateTwo = EnterEndDate
finalDateOne = Right(dateOne, 4) & "/" & Left(dateOne, 2) & "/" & Mid(dateOne, 4, 2)
finalDateTwo = Right(dateTwo, 4) & "/" & Left(dateTwo, 2) & "/" & Mid(dateTwo, 4, 2)
pAth = Worksheets("Master").Range("O8") & "\All Files" & "\" & Replace(finalDateOne, "/", "-") & " at " & timeOne & " to " & Replace(finalDateTwo, "/", "-") & " at " & timeTwo
   'C:\Users\mludwigborycz\OneDrive - Plastics Family Americas\Desktop\Dont Delete\Divorce\All Files" & "\" & Replace(finalDateOne, "/", "-") & " at " & timeOne & " to " & Replace(finalDateTwo, "/", "-") & " at " & timeTwo
If Not fso.FolderExists(pAth) Then


Else
    With ActiveWorkbook.ActiveSheet
        .Hyperlinks.Add anchor:=table.ListColumns(9).DataBodyRange(Numb, 1), _
        Address:=pAth
    End With
End If

Call goToTable
startDate.Text = ""
startTime.Text = ""
EnterEndDate.Text = ""
EnterAbbr.Text = ""
EnterDesc.Text = ""
EnterNotes.Text = ""

Me.tb_StartTime = Format(Me.tb_StartTime, "hh:mm")
Me.tb_Abbreviate = Format(Me.tb_Abbreviate, "hh:mm")
Me.cb_FolderSearch.Caption = "No Folder"
Me.lb_AllData.Selected(CInt(Me.tb_ColNum.Text) - 1) = True
End Sub

Private Sub ch_GoToSh_Click() 'Go to a sheet
Dim shCheck As ComboBox
Dim i As Integer
Dim x As Integer
Set shCheck = Me.cbx_ShCheck
On Error GoTo 0



If shCheck.Text <> "Master" And shCheck.Text <> "" Then
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets(i).Name <> shCheck.Text And ActiveWorkbook.Worksheets(i).Name <> "Master" Then
            ActiveWorkbook.Worksheets(i).Visible = False
        End If
    Next i
   
ElseIf shCheck.Text = "" Then
    MsgBox Prompt:="Enter name of a Manufacturer sheet to open", Title:="Manufacturer name required"
ElseIf shCheck = "Master" Then
    ActiveWorkbook.Worksheets(shCheck.Text).Select
    shCheck.Text = ""
End If

x = 0
For i = 1 To ActiveWorkbook.Worksheets.Count
    If ActiveWorkbook.Worksheets(i).Name = shCheck.Text Then
    x = x + 1
    Exit For
    End If
Next i

If x > 0 Then
    ActiveWorkbook.Worksheets(shCheck.Text).Visible = True
    ActiveWorkbook.Worksheets(shCheck.Text).Select
Else
    MsgBox Prompt:="Please select from the exisiting list for this field", Title:="Sheet not found"
    shCheck.Clear
End If
'Call UserForm_Initialize
'shCheck.Text = ""
End Sub



Private Sub ch_GoToTbl_Click()
Call goToTable
End Sub

Private Sub lb_AllData_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'Deleting row using combobox
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim arrayListIndex

Set tblCheck = Me.cbx_TblCheck

If Me.cb_CheckOrDelete.Caption = "Delete Mode" Then
    If tblCheck.Text = "" Or Me.lb_AllData.Text = "" Then
        MsgBox Prompt:="Enter a table  in 'Go to Table' field to delete table", Title:="Go to Table Field Empty"
        Exit Sub
    End If
    Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)
    For indRow = 1 To table.Range.Rows.Count
        Do Until table.ListColumns(1).Range(indRow, 1) = uf_Master.lb_AllData.List(lb_AllData.ListIndex)
            indRow = indRow + 1
        Loop
        anSwer = MsgBox(Prompt:="Are you sure you want to delete this row", Buttons:=vbYesNo, Title:="DELETE ENTIRE ROW!?")
            If anSwer = vbYes Then
                table.ListRows(indRow - 1).Range.Select
                Selection.Delete
                GoTo jumpOut1
            ElseIf anSwer = vbNo Then
                GoTo jumpOut2
            End If
    Next indRow
jumpOut1:
jumpOut2:
    If table.ListRows.Count = 0 Then
        table.ListRows.Add
        table.Range.Rows(2) = "blank table"
        GoTo Populate
    End If

    For cell = 1 To table.ListColumns(1).DataBodyRange.Count
        table.ListColumns(1).DataBodyRange(cell, 1) = ""
    Next cell
    Num = 1
    For cell = 1 To table.ListColumns(1).DataBodyRange.Count
        table.ListColumns(1).DataBodyRange(cell, 1) = Num
        Num = Num + 1
    Next cell

Populate:
    Me.tb_ColNum.Text = ""
    Me.tb_StartDate.Text = ""
    Me.tb_StartTime.Text = ""
    Me.tb_EndDate.Text = ""
    Me.tb_Abbreviate.Text = ""
    Me.tb_Description.Text = ""
    Me.tb_Notes.Text = ""
    Call goToTable 'Call table madule to reset all lists
    'cb_EnterData_Click 'To renumber all unknown fields
    Exit Sub
ElseIf Me.cb_CheckOrDelete.Caption = "Check Mode" Then
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Worksheets(i).Visible = True
        Worksheets(i).Activate
        If ActiveWorkbook.ActiveSheet.ListObjects.Count > 0 Then
            For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
                If Me.cbx_TblCheck.Text = ActiveWorkbook.ActiveSheet.ListObjects(j).Name Then
                    ActiveWorkbook.ActiveSheet.ListObjects(j).Range.Activate
                    Set table = ActiveWorkbook.ActiveSheet.ListObjects(j)
                    For x = 1 To table.ListRows.Count
                        If Me.lb_AllData.List(Me.lb_AllData.ListIndex, 0) = table.DataBodyRange(x, 1) And Me.lb_AllData.List(Me.lb_AllData.ListIndex, 1) = table.DataBodyRange(x, 2) And Me.lb_AllData.List(Me.lb_AllData.ListIndex, 3) = table.DataBodyRange(x, 4) And Me.lb_AllData.List(Me.lb_AllData.ListIndex, 5) = table.DataBodyRange(x, 6) Then
                            Me.tb_ColNum = Me.lb_AllData.List(Me.lb_AllData.ListIndex, 0)
                            tb_StartDate = Me.lb_AllData.List(Me.lb_AllData.ListIndex, 1)
                            tb_StartTime = Me.lb_AllData.List(Me.lb_AllData.ListIndex, 2)
                            tb_EndDate = Me.lb_AllData.List(Me.lb_AllData.ListIndex, 3)
                            tb_Abbreviate = Me.lb_AllData.List(Me.lb_AllData.ListIndex, 4)
                            tb_Description = Me.lb_AllData.List(Me.lb_AllData.ListIndex, 5)
                            tb_Notes = Me.lb_AllData.List(Me.lb_AllData.ListIndex, 6)
                            If table.DataBodyRange(x, 9) = "" Or table.DataBodyRange(x, 9) = "blank table" Then
                                cb_FolderSearch.Caption = "No Folder"
                            ElseIf table.DataBodyRange(x, 9) <> "" Then
                                cb_FolderSearch.Caption = "Folder Exists"
                            End If
                            Exit Sub
                        End If
                    Next x
                End If
            Next j
        End If
    Next i
ElseIf Me.cb_CheckOrDelete.Caption = "Switch Mode" Then
    uf_Switch.Show
End If
End Sub

Private Sub lb_Search_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'Deleting row using combobox
Dim i As Integer
Dim j As Integer
Dim x As Integer
Me.tb_KeyWordsUsed.Text = ""
Me.tb_DisplaySelectedDescription.Text = ""
Me.tb_DisplaySelectedNotes.Text = ""
If Me.lb_Search.ListIndex >= 0 Then
    Me.tb_DisplaySelectedDescription.Text = Me.lb_Search.List(Me.lb_Search.ListIndex, 6)
    Me.tb_DisplaySelectedNotes.Text = Me.lb_Search.List(Me.lb_Search.ListIndex, 7)
    If Me.cb_Switch.Caption = "Click for Multi Select" Then
        For i = 1 To ActiveWorkbook.Sheets.Count
            Sheets(i).Visible = True
            Sheets(i).Activate
            For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
                If ActiveWorkbook.ActiveSheet.ListObjects(j).Name = Me.lb_Search.List(Me.lb_Search.ListIndex, 0) Then
                    For x = 1 To ActiveWorkbook.ActiveSheet.ListObjects(j).ListRows.Count
                        If ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 1) = Me.lb_Search.List(Me.lb_Search.ListIndex, 1) Then
                            Me.tb_KeyWordsUsed.Text = ActiveWorkbook.ActiveSheet.ListObjects(j).DataBodyRange(x, 8)
                            Exit Sub
                        End If
                    Next x
                End If
            Next j
            If Worksheets(i).Name = "Master" Then
                Sheets(i).Visible = True
            Else
                Sheets(i).Visible = False
            End If
        Next i
    End If
End If
End Sub

Private Sub MultiPage1_Change() 'Adds tables to combo boxes
Dim tblArray, j

Me.cbx_TblCheck.Clear

If MultiPage1.Value = 1 And ActiveWorkbook.ActiveSheet.Name <> "Master" And ActiveWorkbook.ActiveSheet.ListObjects.Count > 0 Then
        ReDim tblArray(1 To ActiveSheet.ListObjects.Count)
        For j = 1 To ActiveWorkbook.ActiveSheet.ListObjects.Count
            tblArray(j) = ActiveWorkbook.ActiveSheet.ListObjects(j).Name
        Next j
    tblArray = SortArrayAlphabetical(tblArray)
    Me.cbx_TblCheck.List = tblArray 'Adds table array to combobox seeing what tables in a sheet exist
    Me.cbx_TblDelete.List = tblArray 'Adds table array to combobox for deletion option
End If
End Sub

Private Sub tb_EndDateEnter_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
Set tbxNote1 = uf_Master.tb_EndDateEnter
Call Check_Date(tbxNote1)
End Sub

Private Sub tb_RefindEndtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
Set tbxNote1 = uf_Master.tb_RefindEndtDate
Call Check_Date(tbxNote1)
End Sub

Private Sub tb_RefindStartDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
Set tbxNote1 = uf_Master.tb_RefindStartDate
Call Check_Date(tbxNote1)
End Sub

Private Sub tb_StartDateEnter_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
Set tbxNote1 = uf_Master.tb_StartDateEnter
Call Check_Date(tbxNote1)
End Sub

Private Sub tb_StartTimeEnter__BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
Dim strVal As String
Dim dTime As Date

If IsNumeric(tb_StartTimeEnter) And Len(tb_StartTimeEnter.Value) = 4 Then
    strVal = Format(tb_StartTimeEnter, "0000")
    dTime = Left(strVal, 2) & ":" & Right(strVal, 2)
    tb_StartTimeEnter.Text = dTime
ElseIf Len(tb_StartTimeEnter.Value) > 4 Or Len(tb_StartTimeEnter.Value) < 4 Then
    MsgBox "Please enter a time formated in a 24 hour time"
    tb_StartTimeEnter.Text = ""
    Exit Sub
ElseIf IsNumeric(tb_StartTimeEnter.Value) = False Then
    MsgBox "Only numbers aloud in this box"
    tb_StartTimeEnter.Text = ""
    Exit Sub
End If

End Sub

Private Sub UserForm_Initialize()
Call workShtsIntoCombobox
Call tblsIntoCombobox
Call keyWordsIntoCombobox

If Me.cbx_ShCheck.Text <> "" Then
    Sheets(cbx_ShCheck.Text).Visible = True
    Sheets(cbx_ShCheck.Text).Activate
End If

End Sub

Private Sub cb_TblSearch_Click() ' Searching table data
Set CheckStartDate = Me.cbx_StartDateCheck
Set CheckStartTime = Me.cbx_StartTimeCheck
Set CheckEndDate = Me.cbx_EndDateCheck

Set CheckAbbr = Me.cbx_CheckAbbr
Set CheckDesc = Me.cbx_CheckDesc
Set CheckNotes = Me.cbx_CheckNotes

Set startDate = Me.tb_StartDate
Set startTime = Me.tb_StartTime
Set endDate = Me.tb_EndDate

Set Abbreviate = Me.tb_Abbreviate
Set Description = Me.tb_Description
Set Notes = Me.tb_Notes
Set tblCheck = Me.cbx_TblCheck

If tblCheck = "" Then
    MsgBox Prompt:="Enter a table to add data to", Title:="Table required"
    CheckStartDate.Text = ""
    CheckStartTime.Text = ""
    CheckEndDate.Text = ""
    CheckAbbr.Text = ""
    CheckDesc.Text = ""
    CheckNotes.Text = ""
    Exit Sub
End If

Set table = ActiveWorkbook.ActiveSheet.ListObjects(Me.cbx_TblCheck.Text)

If CheckStartDate.Text <> "" And CheckStartTime.Text = "" And CheckEndDate.Text = "" And CheckAbbr.Text = "" And CheckDesc.Text = "" And CheckNotes.Text = "" Then
    indRow = 0
    Num = 2
    Numb = 3
    Numr = 4
    Numm = 5
    Numbb = 6
    Numrr = 7
    Set tbxNote1 = startDate
    Set tbxNote2 = startTime
    Set tbxNote3 = endDate
    Set tbxNote4 = Abbreviate
    Set tbxNote5 = Description
    Set tbxNote6 = Notes
    Set cbxNote = CheckStartDate
    Call CheckTblInfo(indRow, Num, Numb, Numr, Numm, Numbb, Numrr, tbxNote1, tbxNote2, tbxNote3, tbxNote4, tbxNote5, tbxNote6, cbxNote)
ElseIf CheckStartDate.Text = "" And CheckStartTime.Text <> "" And CheckEndDate.Text = "" And CheckAbbr.Text = "" And CheckDesc.Text = "" And CheckNotes.Text = "" Then
    indRow = 0
    Num = 3
    Numb = 2
    Numr = 4
    Numm = 5
    Numbb = 6
    Numrr = 7
    Set tbxNote1 = startTime
    Set tbxNote2 = startDate
    Set tbxNote3 = endDate
    Set tbxNote4 = Abbreviate
    Set tbxNote5 = Description
    Set tbxNote6 = Notes
    Set cbxNote = CheckStartTime
    Call CheckTblInfo(indRow, Num, Numb, Numr, Numm, Numbb, Numrr, tbxNote1, tbxNote2, tbxNote3, tbxNote4, tbxNote5, tbxNote6, cbxNote)
 ElseIf CheckStartDate.Text = "" And CheckStartTime.Text = "" And CheckEndDate.Text <> "" And CheckAbbr.Text = "" And CheckDesc.Text = "" And CheckNotes.Text = "" Then
    indRow = 0
    Num = 4
    Numb = 2
    Numr = 3
    Numm = 5
    Numbb = 6
    Numrr = 7
    Set tbxNote1 = endDate
    Set tbxNote2 = startDate
    Set tbxNote3 = startTime
    Set tbxNote4 = Abbreviate
    Set tbxNote5 = Description
    Set tbxNote6 = Notes
    Set cbxNote = CheckEndDate
    Call CheckTblInfo(indRow, Num, Numb, Numr, Numm, Numbb, Numrr, tbxNote1, tbxNote2, tbxNote3, tbxNote4, tbxNote5, tbxNote6, cbxNote)
ElseIf CheckStartDate.Text = "" And CheckStartTime.Text = "" And CheckEndDate.Text = "" And CheckAbbr.Text <> "" And CheckDesc.Text = "" And CheckNotes.Text = "" Then
    indRow = 0
    Num = 5
    Numb = 2
    Numr = 3
    Numm = 4
    Numbb = 6
    Numrr = 7
    Set tbxNote1 = Abbreviate
    Set tbxNote2 = startDate
    Set tbxNote3 = startTime
    Set tbxNote4 = endDate
    Set tbxNote5 = Description
    Set tbxNote6 = Notes
    Set cbxNote = CheckAbbr
    Call CheckTblInfo(indRow, Num, Numb, Numr, Numm, Numbb, Numrr, tbxNote1, tbxNote2, tbxNote3, tbxNote4, tbxNote5, tbxNote6, cbxNote)
ElseIf CheckStartDate.Text = "" And CheckStartTime.Text = "" And CheckEndDate.Text = "" And CheckAbbr.Text = "" And CheckDesc.Text <> "" And CheckNotes.Text = "" Then
    indRow = 0
    Num = 6
    Numb = 2
    Numr = 3
    Numm = 4
    Numbb = 5
    Numrr = 7
    Set tbxNote1 = Description
    Set tbxNote2 = startDate
    Set tbxNote3 = startTime
    Set tbxNote4 = endDate
    Set tbxNote5 = Abbreviate
    Set tbxNote6 = Notes
    Set cbxNote = CheckDesc
    Call CheckTblInfo(indRow, Num, Numb, Numr, Numm, Numbb, Numrr, tbxNote1, tbxNote2, tbxNote3, tbxNote4, tbxNote5, tbxNote6, cbxNote)
ElseIf CheckStartDate.Text = "" And CheckStartTime.Text = "" And CheckEndDate.Text = "" And CheckAbbr.Text = "" And CheckDesc.Text = "" And CheckNotes.Text <> "" Then
    indRow = 0
    Num = 7
    Numb = 2
    Numr = 3
    Numm = 4
    Numbb = 5
    Numrr = 6
    Set tbxNote1 = Notes
    Set tbxNote2 = startDate
    Set tbxNote3 = startTime
    Set tbxNote4 = endDate
    Set tbxNote5 = Abbreviate
    Set tbxNote6 = Description
    Set cbxNote = CheckNotes
    Call CheckTblInfo(indRow, Num, Numb, Numr, Numm, Numbb, Numrr, tbxNote1, tbxNote2, tbxNote3, tbxNote4, tbxNote5, tbxNote6, cbxNote)
   Else:
    MsgBox Prompt:="To search for an item only populate one of the three fields", Title:="Populate One Field"
    CheckStartDate.Text = ""
    CheckStartTime.Text = ""
    CheckEndDate.Text = ""
    CheckAbbr.Text = ""
    CheckDesc.Text = ""
    CheckNotes.Text = ""
    indRow = 0
    Exit Sub
End If

Num = 1
Do Until table.ListColumns(7).DataBodyRange(Num, 1) = Me.tb_Notes
    Num = Num + 1
Loop
Me.tb_StartTime = Format(Me.tb_StartTime, "hh:mm")
Me.tb_Abbreviate = Format(Me.tb_Abbreviate, "hh:mm")
Me.tb_ColNum = table.ListColumns(1).DataBodyRange(Num, 1)
End Sub



Private Sub tb_StartDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'Double click for the Start Date

Set startDate = Me.tb_StartDate
Set tblCheck = Me.cbx_TblCheck
If tblCheck = "" Then
    MsgBox Prompt:="Enter a table in 'Go to Table' field" _
    & vbNewLine & "to add data", Title:="'Go to Table' Field Empty"
     startDate = ""
    Exit Sub
End If
Set listName = Me.tb_StartDate
Str = "Start Date"
Num = 2
Call BoxForDataEntry(listName, Str, Num)
Call goToTable
If Me.tb_EnterAbbr.Text <> "" Or Me.tb_EnterDesc.Text <> "" Or Me.tb_EnterNotes.Text <> "" Then
    Call cb_EnterData_Click
End If
End Sub

Private Sub tb_StartTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'Double click for the Start Time

Set startTime = Me.tb_StartTime
Set tblCheck = Me.cbx_TblCheck
If tblCheck = "" Then
    MsgBox Prompt:="Enter a table in 'Go to Table' field" _
    & vbNewLine & "to add data", Title:="'Go to Table' Field Empty"
     startTime = ""
    Exit Sub
End If
Set listName = Me.tb_StartTime
Str = "Start Time"
Num = 3
Call BoxForDataEntry(listName, Str, Num)
Call goToTable
If Me.tb_EnterAbbr.Text <> "" Or Me.tb_EnterDesc.Text <> "" Or Me.tb_EnterNotes.Text <> "" Then
    Call cb_EnterData_Click
End If
End Sub

Private Sub tb_EndDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'Double click for the End Date
Set endDate = Me.tb_EndDate
Set tblCheck = Me.cbx_TblCheck
If tblCheck = "" Then
    MsgBox Prompt:="Enter a table in 'Go to Table' field" _
    & vbNewLine & "to add data", Title:="'Go to Table' Field Empty"
     endDate = ""
    Exit Sub
End If
Set listName = Me.tb_EndDate
Str = "End Date"
Num = 4
Call BoxForDataEntry(listName, Str, Num)
Call goToTable
If Me.tb_EnterAbbr.Text <> "" Or Me.tb_EnterDesc.Text <> "" Or Me.tb_EnterNotes.Text <> "" Then
    Call cb_EnterData_Click
End If
End Sub

Private Sub tb_Abbreviate_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'Double click for the Abbreviation
Set Abbreviate = Me.tb_Abbreviate
Set tblCheck = Me.cbx_TblCheck
If tblCheck = "" Then
    MsgBox Prompt:="Enter a table in 'Go to Table' field" _
    & vbNewLine & "to add data", Title:="'Go to Table' Field Empty"
     Abbreviate = ""
    Exit Sub
End If
Set listName = Me.tb_Abbreviate
Str = "Job Number"
Num = 5
Call BoxForDataEntry(listName, Str, Num)
Call goToTable
If Me.tb_EnterAbbr.Text <> "" Or Me.tb_EnterDesc.Text <> "" Or Me.tb_EnterNotes.Text <> "" Then
    Call cb_EnterData_Click
End If

End Sub
Private Sub tb_Description_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'Double click for the Description
Set Description = Me.tb_Description
Set tblCheck = Me.cbx_TblCheck
If tblCheck = "" Then
    MsgBox Prompt:="Enter a table in 'Go to Table' field" _
    & vbNewLine & "to add data", Title:="'Go to Table' Field Empty"
    Description.Text = ""
    Exit Sub
End If
Set listName = Me.tb_Description
Str = "Description"
Num = 6
Call BoxForDataEntry(listName, Str, Num)
Call goToTable
If Me.tb_EnterAbbr.Text <> "" Or Me.tb_EnterDesc.Text <> "" Or Me.tb_EnterNotes.Text <> "" Then
    Call cb_EnterData_Click
End If

End Sub
Private Sub tb_Notes_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 'Double click for the Notes
Set Notes = Me.tb_Notes
Set tblCheck = Me.cbx_TblCheck
If tblCheck = "" Then
    MsgBox Prompt:="Enter a table in 'Go to Table' field" _
    & vbNewLine & "to add data", Title:="'Go to Table' Field Empty"
     Notes = ""
    Exit Sub
End If
Set listName = Me.tb_Notes
Str = "Notes"
Num = 7
Call BoxForDataEntry(listName, Str, Num)
Call goToTable
If Me.tb_EnterAbbr.Text <> "" Or Me.tb_EnterDesc.Text <> "" Or Me.tb_EnterNotes.Text <> "" Then
    Call cb_EnterData_Click
End If

End Sub

Private Sub cb_CreateShCancel_Click() 'Cancel userform
For Each Sh In ActiveWorkbook.Worksheets
    If Sh.Name <> "Master" Then
        Sh.Visible = False
    End If
Next Sh
Application.Quit
End Sub
Private Sub cb_CreateTblCancel_Click() 'Cancel userform
For Each Sh In ActiveWorkbook.Worksheets
    If Sh.Name <> "Master" Then
        Sh.Visible = False
    End If
Next Sh
Application.Quit
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'Hides excel except for userform
Application.Visible = True
For Each Sh In ActiveWorkbook.Worksheets
    If Sh.Name <> "Master" And Sh.Name <> "For Printing" Then
        Sh.Visible = False
    End If
Next Sh
Unload Me
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

