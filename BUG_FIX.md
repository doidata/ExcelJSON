# Bug Fix: Undeclared Variable in CommandButton2_Click()

## Issue
In the file `ExcelJSON.xlsm`, module `Arkusz1.cls`, there is an undeclared variable `id` used in the `CommandButton2_Click()` procedure.

### Location
Line in the second loop of the procedure:
```vba
id = Worksheets(1).Cells(i, 1).Value
```

### Problem
The variable `id` is used without being declared with a `Dim` statement. This causes VBA to implicitly create a Variant variable, which can lead to:
- Type confusion and unexpected behavior
- Harder debugging
- Reduced performance
- Potential runtime errors in strict compilation mode

### Fix
Add the declaration at the beginning of the procedure with the other variable declarations.

## Before (Buggy Code)
```vba
Private Sub CommandButton2_Click()
    Dim i As Integer
    i = 2
    Dim nodes As New Collection

    ' ... code ...

    i = 2
    While Worksheets(1).Cells(i, 1).Value <> ""
        Dim parent As String
        parent = Worksheets(1).Cells(i, 2).Value
        id = Worksheets(1).Cells(i, 1).Value  ' <-- BUG: 'id' not declared
        If parent <> "root" Then
            On Error Resume Next
            nodes(parent).children.Add nodes(id)
        End If

        i = i + 1
    Wend

    ' ... rest of code ...
End Sub
```

## After (Fixed Code)
```vba
Private Sub CommandButton2_Click()
    Dim i As Integer
    i = 2
    Dim nodes As New Collection

    ' ... code ...

    i = 2
    While Worksheets(1).Cells(i, 1).Value <> ""
        Dim parent As String
        Dim id As String  ' <-- FIX: Properly declare 'id' variable
        parent = Worksheets(1).Cells(i, 2).Value
        id = Worksheets(1).Cells(i, 1).Value
        If parent <> "root" Then
            On Error Resume Next
            nodes(parent).children.Add nodes(id)
        End If

        i = i + 1
    Wend

    ' ... rest of code ...
End Sub
```

## How to Apply This Fix

1. Open `ExcelJSON.xlsm` in Microsoft Excel
2. Press `Alt + F11` to open the VBA Editor
3. In the Project Explorer, find `Arkusz1` (or Sheet1)
4. Double-click to open the code
5. Find the `CommandButton2_Click()` procedure
6. Locate the line `id = Worksheets(1).Cells(i, 1).Value`
7. Add `Dim id As String` on the line before it (or group it with other Dim statements)
8. Save the file
9. Test the macro to ensure it works correctly

## Additional Bugs Found (Not Fixed in This Iteration)

### 1. Incorrect Collection.Count Usage
VBA Collections use `Count` as a property, not a method. Remove parentheses:
- Change `nodes.Count()` to `nodes.Count`
- Change `nodes(nodes.Count())` to `nodes(nodes.Count)`
- This appears multiple times throughout the code

### 2. Poor Error Handling
Multiple `On Error Resume Next` statements suppress all errors without proper handling. Consider:
- Using proper error handling with `On Error GoTo ErrorHandler`
- Checking for specific error conditions
- Logging or displaying meaningful error messages
