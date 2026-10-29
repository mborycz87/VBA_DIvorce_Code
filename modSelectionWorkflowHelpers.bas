Option Explicit

Public Function ValidateSelectedResult(ByVal lbTarget As MSForms.ListBox) As Boolean
    If lbTarget Is Nothing Then
        ValidateSelectedResult = False
        Exit Function
    End If

    If lbTarget.ListIndex < 0 Then
        ValidateSelectedResult = False
        Exit Function
    End If

    If Len(Trim(NormalizeText(lbTarget.List(lbTarget.ListIndex, 0)))) = 0 Then
        ValidateSelectedResult = False
        Exit Function
    End If

    ValidateSelectedResult = True
End Function

Public Function ResolveSelectedResult(ByVal lbTarget As MSForms.ListBox) As Variant
    Dim tableName As String
    Dim rowIndex As Long
    Dim jobNumber As String
    Dim tbl As ListObject
    Dim result(0 To 2) As Variant

    If Not ValidateSelectedResult(lbTarget) Then
        ResolveSelectedResult = Empty
        Exit Function
    End If

    tableName = NormalizeText(lbTarget.List(lbTarget.ListIndex, 0))
    jobNumber = NormalizeText(lbTarget.List(lbTarget.ListIndex, 1))
    rowIndex = FindMatchingTableRow(FindTableByName(tableName), jobNumber)

    result(0) = tableName
    result(1) = rowIndex
    result(2) = jobNumber

    ResolveSelectedResult = result
End Function

Public Function BulkApplyKeywordState(ByVal lbTarget As MSForms.ListBox, ByVal keyword As String, ByVal stateSuffix As String) As Boolean
    Dim resolved As Variant
    Dim tableName As String
    Dim rowIndex As Long
    Dim tbl As ListObject

    If Not ValidateSelectedResult(lbTarget) Then
        BulkApplyKeywordState = False
        Exit Function
    End If

    resolved = ResolveSelectedResult(lbTarget)
    If IsEmpty(resolved) Then
        BulkApplyKeywordState = False
        Exit Function
    End If

    tableName = CStr(resolved(0))
    rowIndex = CLng(resolved(1))
    Set tbl = FindTableByName(tableName)

    If tbl Is Nothing Then
        BulkApplyKeywordState = False
        Exit Function
    End If

    If rowIndex <= 0 Then
        BulkApplyKeywordState = False
        Exit Function
    End If

    BulkApplyKeywordState = AddKeywordStateToRow(tbl, rowIndex, keyword, stateSuffix)
End Function

Public Function ApplySelectedKeywordState(ByVal lbTarget As MSForms.ListBox, ByVal keyword As String, ByVal stateSuffix As String) As Boolean
    If Not ValidateKeyword(keyword) Then
        ApplySelectedKeywordState = False
        Exit Function
    End If

    ApplySelectedKeywordState = BulkApplyKeywordState(lbTarget, keyword, stateSuffix)
End Function

Public Function RemoveSelectedKeywordState(ByVal lbTarget As MSForms.ListBox, ByVal keyword As String) As Boolean
    Dim resolved As Variant
    Dim tableName As String
    Dim rowIndex As Long
    Dim tbl As ListObject

    If Not ValidateSelectedResult(lbTarget) Then
        RemoveSelectedKeywordState = False
        Exit Function
    End If

    resolved = ResolveSelectedResult(lbTarget)
    If IsEmpty(resolved) Then
        RemoveSelectedKeywordState = False
        Exit Function
    End If

    tableName = CStr(resolved(0))
    rowIndex = CLng(resolved(1))
    Set tbl = FindTableByName(tableName)

    If tbl Is Nothing Then
        RemoveSelectedKeywordState = False
        Exit Function
    End If

    If rowIndex <= 0 Then
        RemoveSelectedKeywordState = False
        Exit Function
    End If

    RemoveSelectedKeywordState = RemoveKeywordStateFromRow(tbl, rowIndex, keyword)
End Function

Public Function PopulateListFromVariant(ByVal results As Variant, ByVal lbTarget As MSForms.ListBox) As Boolean
    If lbTarget Is Nothing Then
        PopulateListFromVariant = False
        Exit Function
    End If

    If IsEmpty(results) Then
        PopulateListFromVariant = False
        Exit Function
    End If

    lbTarget.Clear
    lbTarget.List = results
    lbTarget.FontSize = 8
    lbTarget.ColumnWidths = "75,15,50,30,50,30,112,113,0,0"
    PopulateListFromVariant = True
End Function

Public Function GetListCellText(ByVal lbTarget As MSForms.ListBox, ByVal columnIndex As Long) As String
    If lbTarget Is Nothing Then
        GetListCellText = ""
        Exit Function
    End If

    If lbTarget.ListIndex < 0 Then
        GetListCellText = ""
        Exit Function
    End If

    If columnIndex < 0 Or columnIndex > lbTarget.ColumnCount - 1 Then
        GetListCellText = ""
        Exit Function
    End If

    GetListCellText = NormalizeText(lbTarget.List(lbTarget.ListIndex, columnIndex))
End Function

Public Sub BubbleKeywordRefresh(ByVal form As Object, ByVal keyword As String, ByVal stateSuffix As String)
    If form Is Nothing Then Exit Sub
    If Not ApplySelectedKeywordState(form.lb_Search, keyword, stateSuffix) Then
        MsgBox "Unable to apply keyword state to the selected result.", vbExclamation
    End If
End Sub

Public Sub BubbleSelectionCheck(ByVal form As Object)
    If form Is Nothing Then Exit Sub
    If Not ValidateSelectedResult(form.lb_Search) Then
        MsgBox "Please select an item from the list before continuing.", vbExclamation
    End If
End Sub
