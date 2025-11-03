Sub TblEventLogAddRow()
'
' TblEventLogAddRow Macro
'

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim pwd As String
    Dim targetColumn As ListColumn
    Dim lastRow As ListRow
    Dim targetCell As Range
    Dim maxVal As Double
    Dim rowCount As Long
    Dim Target As Range
    Dim tsCol As Long
    Dim eventDateCol As Long

    ' Set your password here
    pwd = GetPassword()

    ' Set your worksheet and table name
    Set ws = ThisWorkbook.Sheets("Event Log")
    Set tbl = ActiveSheet.ListObjects("TBLEventLog") ' The table"
    targetColumnName = "Event ID"                 ' The Column
    Set targetColumn = tbl.ListColumns(targetColumnName)

    ' Unprotect the sheet
    ws.Unprotect Password:=pwd
    
    ' Add a new row to the bottom of the table
    Application.EnableEvents = False

    Set newRow = tbl.ListRows.Add
    Set newRowRange = newRow.Range
    newRowRange.Font.Bold = False

    
    ' Find max numeric value in the target column
    maxVal = 0
    For Each cell In targetColumn.DataBodyRange
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            If cell.Value > maxVal Then maxVal = cell.Value
        End If
    Next cell
   
     Application.EnableEvents = False
     
    ' Ensure the table is not empty
    If tbl.ListRows.Count > 0 Then
        Set lastRow = tbl.ListRows(tbl.ListRows.Count)
        Set targetCell = lastRow.Range.Cells(1, tbl.ListColumns(targetColumnName).Index)
        targetCell.Value = maxVal + 1
    Else
        MsgBox "Table has no rows."
        GoTo CleanExit
    End If
    
    'Default event date to today
    eventDateCol = tbl.ListColumns("Event Date").Index
    newRowRange.Cells(1, eventDateCol).Value = Date
    newRowRange.Cells(1, eventDateCol).NumberFormat = "dd/mm/yyyy"

    
    'Call ApplyChangeLogic fo audit fields update
    Set Target = newRowRange
    Call ApplyChangeLogic(Target, ws)


    ' Reset the focus to Event Source
    targetColumnName = "Event Type"                 ' The Column
    Set targetCell = lastRow.Range.Cells(1, tbl.ListColumns(targetColumnName).Index)
    targetCell.Select
    
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ws.Protect Password:=pwd, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True

End Sub

Sub ReapplyTableStyle()
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("TBLEventLog")

    With tbl
        .TableStyle = "TableStyleMedium2" ' Or any built-in style
        .ShowTableStyleRowStripes = True
    End With
End Sub

Sub ClearManualShading()
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("TBLEventLog")

    tbl.DataBodyRange.Interior.ColorIndex = xlColorIndexNone
End Sub


