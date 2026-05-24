"These are all the button in the code"

"This will make a user go to a sheet in the divorce log"
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

"This one will create a sheet in the divorce log"
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

"This will delete a sheet in the divorce log"
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