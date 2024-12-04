# Excel VBA Useful Custom Functions

(https://github.com/danvuk567/Excel_VBA-Useful-Custom-Functions/blob/main/images/Excel_VBA.jpg)

Here are some useful functions that I have defined when working with common Excel VBA code requirements. 

## Importing Excel Files

The procedure *Import_excel_file* can be used when importing an external excel file (*.csv, *.xlsx, *.xlsm) into the current workbook sheet. It requires the file path that includes the file name, the source sheet no, the cell range of the source file, the current workbook sheet no, and the current sheet starting range.

    ' f_path: The source file path including the file name
    ' f_sheet: The source sheet (can be a name or sheet number which is always 1 if it is a csv file)
    ' f_start_row: The source file starting row
    ' f_end_row: The source file last row
    ' f_start_col: The source file start column as an index or letter (ex: 1 or A)
    ' f_end_col: The source file end column as an index or letter (ex: 2 or B)
    ' curr_sheet: The current sheet (can be a name or sheet number)
    ' curr_sheet_no: The current sheet number we are working with
    ' curr_row: The current sheet starting row
    ' curr_col: The current sheet start column as an index or letter (ex: 1 or A)

    Sub Import_excel_file(f_path As String, f_sheet As Variant, f_start_row As Integer, f_end_row As Integer, f_start_col As Variant, f_end_col As Variant, curr_sheet As Variant, curr_row As Integer, curr_col As Variant)
        Dim wb1 As Workbook, wb2 As Workbook
        Dim ws1 As Worksheet, ws2 As Worksheet
        Dim sourceRange As Range, targetRange As Range
    
        Set wb1 = ThisWorkbook
        ' Define this workbook sheet which can be the sheet number or name
        Set ws1 = wb1.Sheets(curr_sheet)
    
        ' Open the Excel file
        Workbooks.Open Filename:=f_path, UpdateLinks:=0
        Set wb2 = ActiveWorkbook
        ' Define this sheet which can be the sheet number or name
        Set ws2 = wb2.Sheets(f_sheet)
    
        ' Define the source range based on the column type
        If IsNumeric(f_start_col) Then
            ' If the source excel column is passed as an index
            Set sourceRange = ws2.Range(ws2.Cells(f_start_row, f_start_col), ws2.Cells(f_end_row, f_end_col))
        Else
            ' If the source column is passed as a letter
            Set sourceRange = ws2.Range(f_start_col & f_start_row & ":" & f_end_col & f_end_row)
        End If

        ' Define the target range in the current workbook
        If IsNumeric(curr_col) Then
            ' If the current column is passed as an index (numeric)
            Set targetRange = ws1.Cells(curr_row, curr_col)
        Else
            ' If the currentcolumn is passed as a letter (string)
            Set targetRange = ws1.Range(curr_col & curr_row)
        End If

        ' Copy the source range and paste values to the target range
        sourceRange.Copy
        targetRange.PasteSpecial Paste:=xlPasteValues

        ' Close the external workbook without saving
        wb2.Close SaveChanges:=False

End Sub
