Sub TblAssetConfigAddRow()
'
' TblAssetConfig Macro
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
    Dim DateRecordedCol As Long

    ' Set your password here
    pwd = GetPassword()

    ' Set your worksheet and table name
    Set ws = ThisWorkbook.Sheets("Asset Config.")
    Set tbl = ActiveSheet.ListObjects("TBLAssetConfig") ' The table"
    targetColumnName = "Asset ID"                 ' The Column
    Set targetColumn = tbl.ListColumns(targetColumnName)

    ' Unprotect the sheet
    ws.Unprotect Password:=pwd
    
    ' Add a new row to the bottom of the table
    Application.EnableEvents = False

    Set newRow = tbl.ListRows.Add
    Set newRowRange = newRow.Range
    
   
    'Default event date to today
    DateRecordedCol = tbl.ListColumns("Date Recorded").Index
    newRowRange.Cells(1, DateRecordedCol).Value = Date
    newRowRange.Cells(1, DateRecordedCol).NumberFormat = "dd/mm/yyyy"

    
    'Call ApplyChangeLogic fo audit fields update
    Set Target = newRowRange
    Call ApplyChangeLogic(Target, ws)


    ' Reset the focus to Asset ID
    Set targetCell = newRowRange.Cells(1, targetColumn.Index)
    targetCell.Select

    
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ws.Protect Password:=pwd, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True

End Sub


