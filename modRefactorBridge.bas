Option Explicit

Public Sub PopulateResultsList(ByVal targetList As MSForms.ListBox, ByVal results As Variant)
    If targetList Is Nothing Then Exit Sub

    targetList.Clear
    If IsEmpty(results) Then Exit Sub

    targetList.List = results
    targetList.FontSize = 8
    targetList.ColumnWidths = "75,15,50,30,50,30,112,113,0,0"
End Sub

Public Function TryKeywordSearch(ByVal keyword As String, ByVal targetList As MSForms.ListBox) As Boolean
    Dim results As Variant

    If targetList Is Nothing Then
        TryKeywordSearch = False
        Exit Function
    End If

    If Not ValidateKeyword(keyword) Then
        TryKeywordSearch = False
        Exit Function
    End If

    results = SearchByKeyword(keyword)
    If IsEmpty(results) Then
        MsgBox MSG_NO_DATA, vbInformation
        TryKeywordSearch = False
        Exit Function
    End If

    PopulateResultsList targetList, results
    TryKeywordSearch = True
End Function

Public Function TryDateRangeSearch(ByVal startText As String, ByVal endText As String, ByVal targetList As MSForms.ListBox) As Boolean
    Dim results As Variant
    Dim startDateValue As Date
    Dim endDateValue As Date

    If targetList Is Nothing Then
        TryDateRangeSearch = False
        Exit Function
    End If

    If Not ValidateDateRange(startText, endText) Then
        TryDateRangeSearch = False
        Exit Function
    End If

    startDateValue = CDate(startText)
    endDateValue = CDate(endText)
    results = SearchByDateRange(startDateValue, endDateValue)

    If IsEmpty(results) Then
        MsgBox MSG_NO_DATA, vbInformation
        TryDateRangeSearch = False
        Exit Function
    End If

    PopulateResultsList targetList, results
    TryDateRangeSearch = True
End Function

Public Function TryTableSearch(ByVal tableName As String, ByVal targetList As MSForms.ListBox) As Boolean
    Dim results As Variant

    If targetList Is Nothing Then
        TryTableSearch = False
        Exit Function
    End If

    If Len(Trim(tableName)) = 0 Then
        MsgBox "Please select a table name.", vbExclamation
        TryTableSearch = False
        Exit Function
    End If

    results = SearchByTable(tableName)
    If IsEmpty(results) Then
        MsgBox MSG_NO_DATA, vbInformation
        TryTableSearch = False
        Exit Function
    End If

    PopulateResultsList targetList, results
    TryTableSearch = True
End Function

Public Sub ApplyKeywordStateToSelection(ByVal keyword As String, ByVal stateSuffix As String, ByVal targetList As MSForms.ListBox, ByVal targetTable As ListObject)
    Dim selectedIndex As Long
    Dim rowIndex As Long

    If targetTable Is Nothing Then Exit Sub
    If targetList Is Nothing Then Exit Sub

    selectedIndex = targetList.ListIndex
    If selectedIndex < 0 Then Exit Sub

    rowIndex = selectedIndex + 1
    If rowIndex < 1 Then Exit Sub

    SetKeywordStateOnRow targetTable, rowIndex, keyword, stateSuffix
End Sub

Public Sub ExampleKeywordSearchHandler(ByVal form As Object)
    Dim ok As Boolean

    ok = TryKeywordSearch(form.cb_KeyWords.Text, form.lb_Search)
    If Not ok Then Exit Sub
End Sub

Public Sub ExampleDateSearchHandler(ByVal form As Object)
    Dim ok As Boolean

    ok = TryDateRangeSearch(form.tb_RefindStartDate.Text, form.tb_RefindEndtDate.Text, form.lb_Search)
    If Not ok Then Exit Sub
End Sub
