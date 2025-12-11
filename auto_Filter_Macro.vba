' ========================================================================================
' Macro Name: AutoFilter
' Description: Splits a dataset into multiple worksheets based on unique values in a user-selected column.
'              Features:
'              - Prompts for Data Range and Filter Column
'              - Handles User Cancellation safely
'              - Sanitizes sheet names (removes invalid chars, handles duplicates)
'              - Improving performance by disabling Screen Updating
'              - Optionally generates serial numbers in the new sheets
' Author: [Your Name/GitHub User]
' License: MIT (See LICENSE file)
' ========================================================================================

Sub AutoFilter()

    ' --- Variable Declarations ---
    Dim wsSource As Worksheet       ' The sheet containing the original data
    Dim wsTarget As Worksheet       ' Validated new sheet for filtered data
    Dim rngData As Range            ' The full range of data selected by the user
    Dim rngCopy As Range            ' The visible data to copy after filtering
    Dim lastRow As Long             ' (Unused) Could be used for finding last row dynamically
    Dim filteredValue As Variant    ' Iterator for unique values found in the column
    Dim uniqueValues As Collection  ' Collection to store unique keys from the filter column
    Dim cell As Range               ' Iterator context for loop
    Dim filterColumn As Range       ' The single cell defining the column to filter by
    Dim generateSerial As String    ' User response (Yes/No) for serial number generation
    Dim headerRow As Long           ' (Unused) Logic handled currently via Offset
    Dim i As Long                   ' Loop counter for serial numbers
    Dim cleanName As String         ' Temporary string for sanitizing sheet names

    
    ' --- Initialization ---
    
    ' Set source worksheet to the currently active sheet
    Set wsSource = ActiveSheet

    ' --- User Input: Select Data Range ---
    ' Note: Type:=8 ensures the input box returns a Range object.
    On Error Resume Next
    Set rngData = Application.InputBox("Select the data range (including headers)", "Select Data", Type:=8)
    On Error GoTo 0

    ' VALIDATION: Check if user pressed Cancel or selected nothing
    If rngData Is Nothing Then
        MsgBox "Operation cancelled. No range selected.", vbExclamation
        GoTo Cleanup
    End If

    ' --- User Input: Select Filter Column ---
    ' User should pick one cell in the column they want to group by (e.g., "Category")
    On Error Resume Next
    Set filterColumn = Application.InputBox("Select a single cell in the column you want to filter by", "Select Filter Column", Type:=8)
    On Error GoTo 0

    ' VALIDATION: Check if user pressed Cancel
    If filterColumn Is Nothing Then
        MsgBox "Operation cancelled. No filter column selected.", vbExclamation
        GoTo Cleanup
    End If

    ' --- Performance Optimization ---
    ' Disable screen updating to speed up the loop and prevent flickering.
    Application.ScreenUpdating = False

    ' --- User Input: Serial Numbers ---
    generateSerial = MsgBox("Do you want to generate new serial numbers (1, 2, 3...) in the first column of the new sheets?", vbYesNo, "Generate Serial Numbers")

    ' --- Step 1: Identify Unique Values ---
    Set uniqueValues = New Collection

    ' Loop through the Data Body (skipping the header row) to find unique values.
    ' We use a Collection with a Key to automatically filter duplicates (Error 457 is ignored).
    On Error Resume Next
    ' Logic: 
    ' 1. Calculate the column offset within the range.
    ' 2. Use Offset(1, 0) and Resize to exclude the first row (header).
    For Each cell In rngData.Columns(filterColumn.Column - rngData.Column + 1).Offset(1, 0).Resize(rngData.Rows.Count - 1).Cells
        ' Add item to collection: Value, Key (Key must be string)
        uniqueValues.Add cell.Value, CStr(cell.Value)
    Next cell
    On Error GoTo 0

    ' --- Step 2: Iterate and Process Each Unique Value ---
    For Each filteredValue In uniqueValues

        ' Apply AutoFilter to the original data
        ' Field is calculated relative to the start of the rngData
        rngData.AutoFilter Field:=filterColumn.Column - rngData.Column + 1, Criteria1:=filteredValue

        ' Check if any visible data exists (Subtotal 103 counts visible non-empty cells)
        If Application.WorksheetFunction.Subtotal(103, rngData.Offset(1, 0).Resize(rngData.Rows.Count - 1).Columns(1)) < 1 Then
            ' Corner case: Value exists in list but is hidden or empty in filter context
            MsgBox "There is no filtered data to copy for value " & filteredValue & ".", vbInformation, "Copy Filtered Data"
        Else
            ' Create a new worksheet for this category
            Set wsTarget = Worksheets.Add(After:=wsSource)

            ' --- Sheet Naming Logic ---
            ' We must ensure the sheet name is valid and unique in Excel.
            cleanName = CStr(filteredValue)
            
            ' Strip invalid characters: : \ / ? * [ ]
            cleanName = Replace(cleanName, ":", "")
            cleanName = Replace(cleanName, "\", "")
            cleanName = Replace(cleanName, "/", "")
            cleanName = Replace(cleanName, "?", "")
            cleanName = Replace(cleanName, "*", "")
            cleanName = Replace(cleanName, "[", "")
            cleanName = Replace(cleanName, "]", "")
            
            ' Excel sheet names max length is 31 chars
            cleanName = Left(cleanName, 31)
            
            ' Fallback for empty strings
            If cleanName = "" Then cleanName = "FilteredData"
            
            ' Handle Duplicates: If sheet exists, append timestamp
            On Error Resume Next
            wsTarget.Name = cleanName
            If Err.Number <> 0 Then
                ' Append current seconds to make it likely unique
                wsTarget.Name = Left(cleanName, 20) & "_" & Format(Now, "mmss")
            End If
            On Error GoTo 0

            ' --- Copy Data ---
            
            ' 1. Copy Header Row
            ' We determine the header row from the source range top row
            wsSource.Range("A1:" & wsSource.AutoFilter.Range.Rows(1).Address).Copy
            wsTarget.Range("A1").PasteSpecial (xlPasteColumnWidths)
            wsTarget.Range("A1").PasteSpecial (xlPasteAll)

            ' 2. Copy Visible Body Data
            On Error Resume Next
            Set rngCopy = rngData.Offset(1, 0).Resize(rngData.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
            On Error GoTo 0

            If Not rngCopy Is Nothing Then
                rngCopy.Copy
                ' Paste data right below the header (Header is row 1, so Paste at Row 2... or dynamic)
                ' Note: Implementation assumes headers are 1 row tall.
                wsTarget.Range("A" & rngData.Rows(1).Row + 1).PasteSpecial (xlPasteColumnWidths)
                wsTarget.Range("A" & rngData.Rows(1).Row + 1).PasteSpecial (xlPasteAll)

                ' --- Optional: Generate Serial Numbers ---
                ' Overwrites Column A with 1, 2, 3...
                If generateSerial = vbYes Then
                    For i = rngData.Rows(1).Row + 1 To wsTarget.Cells(wsTarget.Rows.Count, 2).End(xlUp).Row
                        wsTarget.Cells(i, 1).Value = i - rngData.Rows(1).Row
                    Next i
                End If

                ' Move the completed new sheet to the end of the workbook
                wsTarget.Move After:=Worksheets(Worksheets.Count)

            Else
                ' Fallback if SpecialCells failed
                MsgBox "There is no selected filtered data to copy for value " & filteredValue & ".", vbInformation, "Copy Filtered Data"
                Application.DisplayAlerts = False
                wsTarget.Delete
                Application.DisplayAlerts = True
            End If
        End If

        ' Clear the filter for the next iteration
        rngData.AutoFilter Field:=filterColumn.Column - rngData.Column + 1

    Next filteredValue

Cleanup:
    ' Restore environment settings
    wsSource.Activate
    Application.ScreenUpdating = True
End Sub



