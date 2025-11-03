'Sub InsertBoMItemsIntoTimeAndMaterials()
'    Dim wsBoM As Worksheet, wsTAM As Worksheet
'    Dim tblBoM As ListObject, tblTAM As ListObject
'    Dim bomID As String
'    Dim bomRow As ListRow, newRow As ListRow
'    Dim bomIDCol As Long, bomInvIDCol As Long, bomDescCol As Long
'    Dim tamInvIDCol As Long, tamDescCol As Long
'
'    ' Set worksheets
'    Set wsBoM = ThisWorkbook.Sheets("BoM")
'    Set wsTAM = ThisWorkbook.Sheets("Time & Materials")
'
'    ' Set tables
'    Set tblBoM = wsBoM.ListObjects("TblBoM")
'    Set tblTAM = wsTAM.ListObjects("TblTimeAndMaterials")
'
'    ' Prompt for BoM ID
'    bomID = InputBox("Enter the BoM ID to import:", "Select BoM")
'
'    If bomID = "" Then
'        MsgBox "Operation cancelled.", vbExclamation
'        Exit Sub
'    End If
'
'    ' Identify relevant columns in TblBoM
'    bomIDCol = tblBoM.ListColumns("BoM ID").Index
'    bomInvIDCol = tblBoM.ListColumns("Inventory ID & Description").Index
'    'bomDescCol = tblBoM.ListColumns("Description").Index
'
'    ' Identify target columns in TblTimeAndMaterials
'    tamInvIDCol = tblTAM.ListColumns("Inventory Item").Index
'    'tamDescCol = tblTAM.ListColumns("Description").Index
'
'    ' Loop through TblBoM and copy matching rows
'    For Each bomRow In tblBoM.ListRows
'        If bomRow.Range.Cells(1, bomIDCol).Value = bomID Then
'            Set newRow = tblTAM.ListRows.Add
'            newRow.Range.Cells(1, tamInvIDCol).Value = bomRow.Range.Cells(1, bomInvIDCol).Value
'            'newRow.Range.Cells(1, tamDescCol).Value = bomRow.Range.Cells(1, bomDescCol).Value
'        End If
'    Next bomRow
'
'    MsgBox "BoM items inserted into Time & Materials table.", vbInformation
'End Sub

Sub TblTimeAndMaterialsAddRow()
    Dim pwd As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As ListRow
    Dim targetCell As Range
    Dim targetColumnName As String

    pwd = GetPassword()
    Set ws = ThisWorkbook.Sheets("Time & Materials")

    ' Unprotect the sheet
    ws.Unprotect Password:=pwd

    ' Reference the table
    Set tbl = ws.ListObjects("TblTimeAndMaterials")

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

Sub RefreshPivotTablesOnTimeAndMaterials()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pwd As String

    pwd = GetPassword() ' Update with your actual password

    Set ws = ThisWorkbook.Sheets("Time & Materials") ' Update with your sheet name

    With ws
        .Unprotect Password:=pwd

        For Each pt In .PivotTables
            pt.RefreshTable
        Next pt

        .Protect Password:=pwd, _
                 AllowUsingPivotTables:=True, _
                 AllowFiltering:=True, _
                 AllowSorting:=True, _
                 UserInterfaceOnly:=True
    End With

    MsgBox "PivotTables refreshed and sheet reprotected.", vbInformation
End Sub


