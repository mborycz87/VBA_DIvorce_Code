# VBA Divorce Code - cleaned architecture proposal

This repository contains a large Excel/VBA workbook macro file. The original implementation is functional, but it mixes UI handling, search logic, data access, validation, and export logic in the same form module.

This folder adds a cleaner structure that keeps the same features while making the code easier to maintain.

## Goals

- keep the same business behavior
- split logic by responsibility
- remove repeated `Activate` / `Select` patterns
- centralize workbook and table access
- make search and export functions reusable
- create a pattern that can be migrated gradually from the current form-based logic

## Proposed modules

- `modConstants.bas` - shared constants and record definition
- `modWorkbookHelpers.bas` - workbook and worksheet helper functions
- `modTableHelpers.bas` - row reading/writing and table helpers
- `modSearch.bas` - keyword, date range, and table-based search logic
- `modValidation.bas` - validation and required field checks
- `modExport.bas` - folder creation and export helpers

## Recommended next step

Import these modules into the VBA project and then modify each existing form event like this:

```vb
Private Sub cb_KeyWordsSearch_Click()
    Dim results As Variant

    If Not ValidateKeyword(Me.cb_KeyWords.Text) Then Exit Sub

    results = SearchByKeyword(Me.cb_KeyWords.Text)
    If IsEmpty(results) Then
        MsgBox MSG_NO_DATA, vbInformation
        Exit Sub
    End If

    Me.lb_Search.List = results
End Sub
```

This keeps the UI thin and moves the real logic into the shared modules.

## Why this is cleaner

- UI code only handles user actions and display
- search logic is reusable and testable
- table access is standardized through `ReadRow` and `WriteRow`
- export logic is separated from form events
- validation rules are centralized instead of repeated in every button click

## Migration approach

1. Keep the original `Divorce_Code.bas` file as a safety copy.
2. Import the new refactor modules.
3. Replace the largest button handlers one by one.
4. Remove repeated `Activate`, `Select`, and hard-coded sheet traversal from the UI code.
5. Keep the same user-facing behavior while gradually replacing old logic.

## Example refactor pattern

Instead of repeatedly doing this in every button:

- activate a sheet
- select a table
- loop across all sheets manually
- hard-code column indexes
- duplicate validation logic

Use this pattern instead:

- `GetWorksheet("Master")`
- `GetListObject(ws, "MyTable")`
- `ReadRow(tbl, rowIndex)`
- `SearchByKeyword(keyword)`
- `ValidateDateRange(startDate, endDate)`

This keeps the project close to the original functionality while making the code much easier to maintain.

## Notes

This repository currently contains a single large VBA form file. A true end-to-end refactor would replace the event-heavy code with a modular structure over time. These files provide the foundation for that work without changing the original workbook behavior.

The next practical step is to import the modules and then refactor the biggest event functions in `Divorce_Code.bas` to call these shared helpers.

