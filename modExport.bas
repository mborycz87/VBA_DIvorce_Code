Option Explicit

Public Function RequireText(ByVal value As Variant, ByVal fieldName As String) As Boolean
    If ValueIsBlank(value) Then
        MsgBox fieldName & " is required.", vbExclamation
        RequireText = False
    Else
        RequireText = True
    End If
End Function

Public Function ValidateDateRange(ByVal startDate As Variant, ByVal endDate As Variant) As Boolean
    If ValueIsBlank(startDate) Or ValueIsBlank(endDate) Then
        MsgBox "Please enter both dates.", vbExclamation
        ValidateDateRange = False
        Exit Function
    End If

    If IsDate(startDate) = False Or IsDate(endDate) = False Then
        MsgBox "Both values must be valid dates.", vbExclamation
        ValidateDateRange = False
        Exit Function
    End If

    If CDate(startDate) > CDate(endDate) Then
        MsgBox "Start date must be earlier than or equal to the end date.", vbExclamation
        ValidateDateRange = False
        Exit Function
    End If

    ValidateDateRange = True
End Function

Public Function ValidateKeyword(ByVal keyword As String) As Boolean
    If Len(Trim(keyword)) = 0 Then
        MsgBox "Please enter a keyword to search.", vbExclamation
        ValidateKeyword = False
    Else
        ValidateKeyword = True
    End If
End Function

Public Function HasTableRowData(ByVal tbl As ListObject) As Boolean
    If tbl Is Nothing Then
        HasTableRowData = False
        Exit Function
    End If

    HasTableRowData = Not (tbl.DataBodyRange Is Nothing)
End Function
