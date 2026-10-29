Option Explicit

Public Function SafeSetButtonState(ByVal buttonObj As Object, ByVal captionText As String, ByVal foreColor As Long, ByVal backColor As Long) As Boolean
    If buttonObj Is Nothing Then
        SafeSetButtonState = False
        Exit Function
    End If

    buttonObj.Caption = captionText
    buttonObj.ForeColor = foreColor
    buttonObj.BackColor = backColor
    SafeSetButtonState = True
End Function

Public Function SafeSetListboxFormat(ByVal listBoxObj As MSForms.ListBox, ByVal columnWidths As String, Optional ByVal fontSize As Long = 8) As Boolean
    If listBoxObj Is Nothing Then
        SafeSetListboxFormat = False
        Exit Function
    End If

    listBoxObj.FontSize = fontSize
    listBoxObj.ColumnWidths = columnWidths
    SafeSetListboxFormat = True
End Function

Public Function SafeClearListbox(ByVal listBoxObj As MSForms.ListBox) As Boolean
    If listBoxObj Is Nothing Then
        SafeClearListbox = False
        Exit Function
    End If

    listBoxObj.Clear
    SafeClearListbox = True
End Function

Public Function SafePopulateListbox(ByVal listBoxObj As MSForms.ListBox, ByVal results As Variant, Optional ByVal columnWidths As String = "75,15,50,30,50,30,112,113,0,0", Optional ByVal fontSize As Long = 8) As Boolean
    If listBoxObj Is Nothing Then
        SafePopulateListbox = False
        Exit Function
    End If

    If IsEmpty(results) Then
        call SafeClearListbox(listBoxObj)
        SafePopulateListbox = False
        Exit Function
    End If

    SafeClearListbox listBoxObj
    listBoxObj.List = results
    listBoxObj.FontSize = fontSize
    listBoxObj.ColumnWidths = columnWidths
    SafePopulateListbox = True
End Function

Public Function SafeRefreshListbox(ByVal listBoxObj As MSForms.ListBox, ByVal results As Variant, Optional ByVal columnWidths As String = "75,15,50,30,50,30,112,113,0,0", Optional ByVal fontSize As Long = 8) As Boolean
    Return SafePopulateListbox(listBoxObj, results, columnWidths, fontSize)
End Function

Public Function SafeShowMessage(ByVal messageText As String, Optional ByVal titleText As String = "Information", Optional ByVal styleValue As VbMsgBoxStyle = vbInformation) As Boolean
    If Len(Trim(messageText)) = 0 Then
        SafeShowMessage = False
        Exit Function
    End If

    MsgBox messageText, styleValue, titleText
    SafeShowMessage = True
End Function

Public Function SafeToggleSelectionMode(ByVal listBoxObj As MSForms.ListBox, ByVal isMultiSelect As Boolean) As Boolean
    If listBoxObj Is Nothing Then
        SafeToggleSelectionMode = False
        Exit Function
    End If

    If isMultiSelect Then
        listBoxObj.MultiSelect = 1
    Else
        listBoxObj.MultiSelect = fmMultiSelectSingle
    End If

    SafeToggleSelectionMode = True
End Function
