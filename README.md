# Excel VBA Useful Custom Functions

![Excel_VBA.jpg](https://github.com/danvuk567/Excel_VBA-Useful-Custom-Functions/blob/main/images/Excel_VBA.jpg?raw=true)

Here are some useful functions that I have defined when working with common Excel VBA code requirements. 


## 1. Retrieving any Higher Level Parent Folder Path

The function *Get_relative_path_start* can retrieve a higher level parent folder path after specifying how many folders to go back from the folder specified. This can be useful in cases where the absolute path is not known such as when searching for the same folder in different locations.

        ' curr_path: The path specified
        ' num_folders_back: The number of prior folders specified. Retrieves the current folder if num_folders_back = 0

        Function Get_relative_path_start(curr_path As String, num_folders_back As Integer) As String
            Dim i As Integer
            Dim slash_pos As Integer
    
            ' Ensure the path ends with a backslash
            If Right(curr_path, 1) <> "\" Then
                curr_path = curr_path + "\"
            End If

            slash_pos = InStrRev(curr_path, "\")
    
            ' Look for the path ending position for num_folders_back and store it in slash_pos
            If num_folders_back <> 0 Then
                For i = 1 To num_folders_back
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

## 2. Retrieve File Type List from a Folder Path

The function *Get_directory_files* can retrieve a list of files from a folder path based on file name or file type. It stores the file names and timestamps in a 2-dimensional array.

        ' file_name_type: The name and/or file type (ex: *.pdf)
        ' file_path: The path to the folder containing the files

        Function Get_directory_files(file_name_type As String, f_path As String) As Variant
                Dim f_names() As Variant
                Dim f_dates() As Variant
                Dim file_names As String
                Dim f_cnt As Integer
                Dim f_names_dates() As Variant
    
                ' Ensure the path ends with a backslash
                If Right(f_path, 1) <> "\" Then
                        f_path = f_path + "\"
                End If
    
                ' Get the first file in the folder
                file_names = Dir(f_path & file_name_type)
    
                ' Get all files in the directory of file_names and store the names in the f_names array and the dates in the f_dates array
                f_cnt = 0
                Do While file_names <> ""
                        f_cnt = f_cnt + 1
                        ReDim Preserve f_names(0 To f_cnt - 1)
                        ReDim Preserve f_dates(0 To f_cnt - 1)
                        f_names(f_cnt - 1) = file_names
                        f_dates(f_cnt - 1) = FileDateTime(f_path & "\" & file_names)
                        file_names = Dir
                Loop
    
                ' If no files found, return an empty array
                If f_cnt = 0 Then
                        Get_directory_files = Array()
                        Exit Function
                End If
    
                ReDim f_names_dates(0 To f_cnt - 1, 0 To 1)
        
                ' Store both names and dates in the f_names_dates 2-dimensional array
                For i = 0 To f_cnt - 1
                        f_names_dates(i, 0) = f_names(i)
                        f_names_dates(i, 1) = f_dates(i)
                Next i
    
                ' Return the the array f_names_dates
                Get_directory_files = f_names_dates
    
        End Function

## 3. Sort List of files based on Timestamp

The function *Sort_files_by_date* sorts the file names in ascending or descending order within a 2-dimensional array based on timestamp. 

        ' f_names: A 2-dimensional array of filenames and their timestamps
        ' sort_order: Sort order defined as ascending when True, and False when descending

        Function Sort_files_by_date(f_names As Variant, sort_order As Boolean) As Variant
                Dim i As Integer
                Dim j As Integer
                Dim start_ind As Integer
                Dim end_ind As Integer
                Dim f_name_tmp As String
                Dim f_date_tmp As String
                Dim f_cnt As Integer
    
                f_cnt = UBound(f_names, 1)
    
                ' Check if the array has more than one element to sort
                If f_cnt > 0 Then
                        For i = 0 To f_cnt - 1
                            For j = i + 1 To f_cnt
                                If IsDate(f_names(i, 1)) And IsDate(f_names(j, 1)) Then
                                    f_date1 = CDate(f_names(i, 1))  ' Convert once to Date
                                    f_date2 = CDate(f_names(j, 1))  ' Convert once to Date
                
                                    If sort_order Then
                                        If f_date1 > f_date2 Then
                                            ' Swap elements based on dates
                                            f_name_tmp = f_names(i, 0)
                                            f_date_tmp = f_names(i, 1)
                                            f_names(i, 0) = f_names(j, 0)
                                            f_names(i, 1) = f_names(j, 1)
                                            f_names(j, 0) = f_name_tmp
                                            f_names(j, 1) = f_date_tmp
                                        End If
                                    Else
                                        If f_date1 < f_date2 Then
                                            ' Swap elements based on dates
                                            f_name_tmp = f_names(i, 0)
                                            f_date_tmp = f_names(i, 1)
                                            f_names(i, 0) = f_names(j, 0)
                                            f_names(i, 1) = f_names(j, 1)
                                            f_names(j, 0) = f_name_tmp
                                            f_names(j, 1) = f_date_tmp
                                        End If
                                    End If
                                End If
                            Next j
                        Next i
                End If
    
                ' Return the sorted array f_names
                Sort_files_by_date = f_names
    
        End Function

## 4. Importing Excel Files

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

## 5. Search for Value in Excel File

The function *Search_for_value* will search for a value such as text, number or blank in the column of an excel sheet. It returns the row number if it is found, otherwise it returns 0.

        ' start_row: The starting row to start search
        ' curr_sheet: The current sheet (can be a name or sheet number)
        ' curr_col: The current sheet search column as an index or letter (ex: 1 or A)
        ' search_val: The value to search for such as text, number or blank

        Function Search_for_value(start_row As Integer, curr_sheet As Variant, curr_col As Variant, search_val As Variant) As Variant
            Dim wb1 As Workbook
            Dim ws1 As Worksheet
            Dim row_cnt As Integer
            Dim NotFound As Boolean
            Dim col_type As String
            Dim cell_value As Variant

            Set wb1 = ActiveWorkbook
            Set ws1 = wb1.Sheets(curr_sheet)
    
            ' Check if the column label being passed is a string type as column letter or number type as column index
            If VarType(curr_col) = vbString Then
                    col_type = "string"
            Else
                    col_type = "number"
            End If

            ws1.Activate
            NotFound = True
            row_cnt = start_row - 1

            ' Loop until we find the value
            Do While NotFound
                row_cnt = row_cnt + 1

                ' Use letter and row as cell reference in Range if col_type is string
                If col_type = "string" Then
                    cell_value = ws1.Range(curr_col & row_cnt).Value
                Else
                    cell_value = ws1.Cells(row_cnt, curr_col).Value
                End If

                ' Exit the loop if the current cell matches the search value (including blank match)
                If cell_value = search_val Then
                    NotFound = False
                End If
        
                ' Exit the loop if the current cell is blank and search value was not a blank
                If (search_val <> "") And (cell_value = "") Then
                    NotFound = False
                End If
        
            Loop

            ' Return 0 if no match found for non-blank search
            If (search_val <> "") And (cell_value = "") Then
                Search_for_value = 0
            Else
                Search_for_value = row_cnt
            End If

        End Function

## 6. Clear Section of Excel File

The procedure *Clear_Section* will clear a section of an excel sheet based on the range of rows and columns specified.

        ' start_row: The starting row within range to clear
        ' end_row: The last row within range to clear
        ' start_col: The starting column index or column letter within range to clear
        ' end_col: The ending column index or column letter within range to clear

        Sub Clear_Section(start_row As Integer, end_row As Integer, start_col As Variant, end_col As Variant)
            Dim wb1 As Workbook
            Dim ws1 As Worksheet
    
            Set wb1 = ActiveWorkbook
            Set ws1 = wb1.ActiveSheet
    
            ' Check if the column lable being passed is a string type as column letter or number type as column index
            If VarType(start_col) = vbString Then
                    col_type = "string"
            Else
                    col_type = "number"
            End If

            ' Clear rows only if they are not already cleared
            If end_row > start_row Then
                ' Use letter and row as cell reference in Range if col_type is string to clear contents
                If col_type = "string" Then
                    ws1.Range(start_col & start_row & ":" & end_col & end_row).ClearContents
                ' Use Cell index in Range if col_type is number to clear contents
                Else
                    ws1.Range(ws1.Cells(start_row, start_col), ws1.Cells(end_row, end_col)).ClearContents
                End If
            End If
            
        End Sub

## Practical Example of using functions #1 to #6

Here is an example of code using the functions oultined in #1 to #6 to search for the latest file in a higher parent folder, and importing that file in a sheet within the working excel file.

        Dim f_name As String
        Dim f_path As String
        Dim get_files As Variant
        Dim f_sheet As Variant
        Dim f_start_row As Integer
        Dim f_start_col As Variant
        Dim f_end_col As Variant
        Dim curr_sheet As Variant
        Dim curr_col As Variant
        Dim f_end_row As Integer
        Dim curr_row As Integer
        Dim end_row As Integer
        Dim end_col As Variant

        ' Get the path to the Excel folder
        f_path = Get_relative_path_start(ThisWorkbook.Path, 1) & "\Excel\"
    
        ' Retrieve the excel files of type *.xlsx
        get_files = Get_directory_files("*xlsx", f_path)
    
        ' Sort the files in descending order
        get_files = Sort_files_by_date(get_files, False)
    
        ' Retrieve the first file which is the latest file
        f_path = f_path & get_files(0, 0)
    
        ' Set copy paremeters as first sheet, 2nd to 11th row of column A to B
        f_sheet = 1
        f_start_row = 2
        f_end_row = 11
        f_start_col = "A"
        f_end_col = "B"
    
        ' Set paste parameters as Sheet1, 1st row of column A
        curr_sheet = "Sheet1"
        curr_row = 1
        curr_col = "A"
        end_col = "B"
    
        ' Get the 1st row that is blank
        end_row = Search_for_value(curr_row, curr_sheet, curr_col, "")

        ' Clear the section where data will be imported including any rows that might not be overwritten
        Clear_Section curr_row, end_row - 1, curr_col, end_col
    
        ' Call Import_excel_file to copy/paste data into Sheet1 of this workbook
        Import_excel_file f_path, f_sheet, f_start_row, f_end_row, f_start_col, f_end_col, curr_sheet, curr_row, curr_col



