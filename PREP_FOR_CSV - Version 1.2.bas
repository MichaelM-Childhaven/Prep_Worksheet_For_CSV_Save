Attribute VB_Name = "PREP_FOR_CSV"
Sub PREP_WORKSHEET_FOR_CSV_SAVE()
Attribute PREP_WORKSHEET_FOR_CSV_SAVE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Prep_CSV_File Macro
'
    Dim lLastUsedColumn As Long
    Dim sLastUsedColumn As String
    Dim lLastUsedRow As Long
    
    Dim lFirstUnusedColumn As Long
    Dim sFirstUnusedColumn As String
    Dim lFirstUnusedRow As Long
    
    Dim lFirstUsedColumn As Long
    Dim sFirstUsedColumn As String
    Dim lFirstUsedRow As Long

    Dim lLastUnusedColumn As Long
    Dim sLastUnusedColumn As String
    Dim lLastUnusedRow As Long

    'Get the last actually USED (not merely formatted) column and row of the worksheet
    lLastUsedColumn = Get_Last_Column()
    sLastUsedColumn = GetColumnLetterFromColumnNumber(lLastUsedColumn)
    lLastUsedRow = Get_Last_Row()
    
    ' Now get the FIRST UNUSED column and row of the worksheet
    sFirstUnusedColumn = GetColumnLetterFromColumnNumber(lLastUsedColumn + 1)
    lFirstUnusedRow = lLastUsedRow + 1
    
    ' Move to start of COLUMNS to delete, select the first cell in that column,
    ' Then select that cell's column, then extend the selection of columns as far to the
    ' right as possible.
    Range(sFirstUnusedColumn & "1").Select
    Selection.EntireColumn.Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
 
    ' Move to start of ROWS to delete, select the first cell in that row,
    ' Then select that cell's row, then extend the selection of rows as far to the
    ' bottom as possible.
    Range("A" & lFirstUnusedRow).Select
    Selection.EntireRow.Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    ' Some people might have blank rows and columns at the beginning of their worksheet.
    ' Let's remove those now.
    ' Start by getting the FIRST actually USED (not merely formatted) column and row of the worksheet
    lFirstUsedColumn = Get_First_Column()
    sFirstUsedColumn = GetColumnLetterFromColumnNumber(lLastUsedColumn)
    lFirstUsedRow = Get_First_Row()
    
    ' Now subtract one to get the LAST UNUSED column and row at the start of the data
    lLastUnusedColumn = lFirstUsedColumn - 1
    lLastUnusedRow = lFirstUsedRow - 1
    
    'There will only be beginning unused columns to delete if lLastUnusedColumn is greater than zero
    If lLastUnusedColumn > 0 Then
    
        ' Get the column letter of the last unused column
        sLastUnusedColumn = GetColumnLetterFromColumnNumber(lLastUnusedColumn)
        
        ' Move to start of unused COLUMNS to delete (always column A), select that column,
        ' and then extend the selection of columns all the way to the last unused column.
        Columns("A:" & sLastUnusedColumn).Select
        Selection.Delete Shift:=xlToLeft
        
    End If
    
    'There will only be beginning unused rows to delete if lLastUnusedRow is greater than zero
    If lLastUnusedRow > 0 Then
    
        ' Move to start of unused ROWS to delete (always row 1), select that row,
        ' and then extend the selection of rows all the way to the last unused row.
        Rows("1:" & lLastUnusedRow).Select
        Selection.Delete Shift:=xlUp
        
    End If
    
    ' Leave macro with A1 as the active cell
    Range("A1").Select
    
End Sub
'
' From **Jon Acampora** at Excel Campus
' https://www.excelcampus.com/vba/find-last-row-column-cell/
'
' Another way to compute the last row used in a worksheet.
'
Private Function Get_Last_Row() As Long
'Finds the last non-blank row on a sheet/range.
On Error GoTo Get_Last_Row_NoData

    Dim lRow As Long
    
    lRow = Cells.Find(what:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    
    Get_Last_Row = lRow
    GoTo Get_Last_Row_End
    
Get_Last_Row_NoData:

    Get_Last_Row = 1

Get_Last_Row_End:
End Function
'
' From **Jon Acampora** at Excel Campus
' https://www.excelcampus.com/vba/find-last-row-column-cell/
'
' Another way to compute the last row used in a worksheet.
'
Private Function Get_Last_Column() As Long
'Finds the last non-blank column on a sheet/range.
On Error GoTo Get_Last_Column_NoData

    Dim lCol As Long
    
    lCol = Cells.Find(what:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Column
    
    Get_Last_Column = lCol
    GoTo Get_Last_Column_End
    
Get_Last_Column_NoData:

    Get_Last_Column = 1

Get_Last_Column_End:
End Function
'
' From **Jon Acampora** at Excel Campus
' https://www.excelcampus.com/vba/find-last-row-column-cell/
'
' I adapted his last row method to find the first used row.
'
Private Function Get_First_Row() As Long
'Finds the first non-blank row on a sheet/range.
On Error GoTo Get_First_Row_NoData

    Dim lRow As Long
    
    lRow = Cells.Find(what:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False).Row
    
    Get_First_Row = lRow
    GoTo Get_First_Row_End
    
Get_First_Row_NoData:

    Get_First_Row = 1

Get_First_Row_End:
End Function
'
' From **Jon Acampora** at Excel Campus
' https://www.excelcampus.com/vba/find-last-row-column-cell/
'
' I adapted his last column method to find the first used column.
'
Private Function Get_First_Column() As Long
'Finds the first non-blank column on a sheet/range.
On Error GoTo Get_First_Column_NoData

    Dim lCol As Long
    
    lCol = Cells.Find(what:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False).Column
    
    Get_First_Column = lCol
    GoTo Get_First_Column_End
    
Get_First_Column_NoData:

    Get_First_Column = 1

Get_First_Column_End:
End Function
'
' From "brettdj" on StackOverflow
' https://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
'
' See also TheSpreadsheetGuru's take on this for an equivalent (and nearly identical) solution:
' https://www.thespreadsheetguru.com/the-code-vault/vba-code-to-convert-column-number-to-letter-or-letter-to-number
'
'This function returns the column letter for a given column number.
'
Private Function GetColumnLetterFromColumnNumber(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    GetColumnLetterFromColumnNumber = vArr(0)
End Function
