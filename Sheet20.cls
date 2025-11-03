Sub TblMaintenanceEventsLinkageAddRow()
    Dim pwd As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As ListRow
    Dim targetCell As Range
    Dim targetColumnName As String

    pwd = GetPassword()
    Set ws = ThisWorkbook.Sheets("Maintenance - Events Linkage")

    ' Unprotect the sheet
    ws.Unprotect Password:=pwd

    ' Reference the table
    Set tbl = ws.ListObjects("TblMaintenanceEventLinks")

    ' Add a new row to the table
    Set lastRow = tbl.ListRows.Add(AlwaysInsert:=True)

    ' Define the column name you want to focus on
    targetColumnName = "Maintenance ID"

    ' Set the target cell in the newly added row
    Set targetCell = lastRow.Range.Cells(1, tbl.ListColumns(targetColumnName).Index)
    targetCell.Select

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Reprotect the sheet
    ws.Protect Password:=pwd, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
End Sub

