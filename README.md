# Excel VBA Useful Custom Functions

![Excel_VBA.jpg](https://github.com/danvuk567/Excel_VBA-Useful-Custom-Functions/blob/main/images/Excel_VBA.jpg?raw=true)

Here are some useful functions that I have defined when working with common Excel VBA code requirements. 


## Retrieving any higher Level Parent Folder Path

This function can get a higher level folder path after specifying how many folders to go back from the folder specified. This can be useful in cases where the absolute path is not known such as when searching for the same folder in different locations.

        ' curr_path: The path specified
        ' num_subfolders_back: The number of prior subfolders specified

        Function Get_relative_path_start(curr_path As String, num_subfolders_back As Integer) As String
            Dim i As Integer
            Dim slash_pos As Integer
    
            ' Ensure the path ends with a backslash
            If Right(curr_path, 1) <> "\" Then
                curr_path = curr_path + "\"
            End If

            slash_pos = InStrRev(curr_path, "\")
    
            ' Look for the path ending position for num_subfolders_back and store it in slash_pos
            If num_subfolders_back <> 0 Then
                For i = 1 To num_subfolders_back
                    slash_pos = InStrRev(curr_path, "\", slash_pos - 1)
                    If slash_pos = 0 Then
                        ' If there are no more slashes, return an empty string
                        Get_relative_path_start = ""
                        Exit Function
                    End If
                Next i
            End If
    
            ' Return the function as the left portion of curr_path before the postion slash_pos
            Get_relative_path_start = Left(curr_path, slash_pos - 1)
    
        End Function

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
