Option Explicit

Public Const SHEET_MASTER As String = "Master"
Public Const SHEET_PRINT As String = "For Printing"
Public Const SHEET_KEYWORDS As String = "Key Words"

Public Enum RecordColumn
    ColTableName = 1
    ColJobNumber = 2
    ColStartDate = 3
    ColStartTime = 4
    ColEndDate = 5
    ColDescription = 6
    ColNotes = 7
    ColKeywords = 8
    ColLink = 9
End Enum

Public Type JobRecord
    TableName As String
    JobNumber As String
    StartDate As Variant
    StartTime As String
    EndDate As Variant
    Description As String
    Notes As String
    Keywords As String
    Link As String
End Type

Public Const MSG_SELECT_ITEM As String = "Select an item from the ListBox to detect a folder"
Public Const MSG_NO_DATA As String = "This data does not exist"
Public Const MSG_ENTER_KEYWORD As String = "Please put in a keyword to search"
