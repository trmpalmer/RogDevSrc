Attribute VB_Name = "ModBoMFrm"


'Sub ShowFrmBoMSelection()
'    FrmBoMSelection.Show
'End Sub

Public Sub PopulateBoMIDCombo(cbo As ComboBox)
    Dim wsBoM As Worksheet, tblBoM As ListObject
    Dim cell As Range
    Dim dictBoMIDs As Object ' Late-bound Dictionary

    Set wsBoM = ThisWorkbook.Sheets("BoM")
    Set tblBoM = wsBoM.ListObjects("TblBoM")
    Set dictBoMIDs = CreateObject("Scripting.Dictionary")

    ' Build dictionary of unique BoM IDs
    For Each cell In tblBoM.ListColumns("BoM ID").DataBodyRange
        If Not dictBoMIDs.exists(cell.Value) And Len(cell.Value) > 0 Then
            dictBoMIDs.Add cell.Value, True
        End If
    Next cell

    ' Populate ComboBox
    cbo.Clear
    Dim key As Variant
    For Each key In dictBoMIDs.Keys
        cbo.AddItem key
    Next key
End Sub

Public Sub PopulateCMBMaintenanceID(cbo As ComboBox)
    Dim wsMnt As Worksheet, TblMaintenanceRecord As ListObject
    Dim cell As Range
    Dim dictMntIDs As Object ' Late-bound Dictionary

    Set wsMnt = ThisWorkbook.Sheets("Maintenance Visit")
    Set TblMaintenanceRecord = wsMnt.ListObjects("TblMaintenanceRecord")
    Set dictMntIDs = CreateObject("Scripting.Dictionary")

    ' Build dictionary of unique BoM IDs
    For Each cell In TblMaintenanceRecord.ListColumns("Maintenance ID").DataBodyRange
        If Not dictMntIDs.exists(cell.Value) And Len(cell.Value) > 0 Then
            dictMntIDs.Add cell.Value, True
        End If
    Next cell

    ' Populate ComboBox
    cbo.Clear
    Dim key As Variant
    For Each key In dictMntIDs.Keys
        cbo.AddItem key
    Next key
End Sub

Public Sub InsertBoMItemsIntoTimeAndMaterials()
    Dim frm As FrmBoMSelection
    Dim selectedBoMID As String, selectedMaintID As String
    Dim parts() As String

    Set frm = New FrmBoMSelection
    frm.Show
    parts = Split(frm.Tag, "|")
    Unload frm

    If UBound(parts) < 1 Then
        MsgBox "Operation cancelled or incomplete.", vbExclamation
        Exit Sub
    End If

    selectedBoMID = parts(0)
    selectedMaintID = parts(1)

    Call CopyBoMItems(selectedBoMID, selectedMaintID)

End Sub

Private Sub CopyBoMItems(bomID As String, maintID As String)
    Dim wsBoM As Worksheet, wsTAM As Worksheet
    Dim tblBoM As ListObject, tblTAM As ListObject
    Dim bomRow As ListRow, newRow As ListRow
    Dim bomIDCol As Long, bomInvIDCol As Long, bomQtyCol As Long
    Dim tamInvIDCol As Long, tamQtyCol As Long, tamMaintIDCol As Long, tamBomIDCol As Long
    
    Dim Target As Range
    Dim Sh As Worksheet

    Dim auditTargets As Collection
    Set auditTargets = New Collection

    ' Set worksheets and tables
    Set wsBoM = ThisWorkbook.Sheets("BoM")
    Set wsTAM = ThisWorkbook.Sheets("Time & Materials")
    Set tblBoM = wsBoM.ListObjects("TblBoM")
    Set tblTAM = wsTAM.ListObjects("TblTimeAndMaterials")

    Dim tsCol As Long, AppUserCol As Long
    tsCol = tblTAM.ListColumns("Modified Date").Index
    AppUserCol = tblTAM.ListColumns("Modified User").Index
    pwd = GetPassword()

    On Error GoTo CleanExit

    ' Identify columns in TblBoM
    bomIDCol = tblBoM.ListColumns("BoM ID").Index
    bomInvIDCol = tblBoM.ListColumns("Inventory ID & Description").Index
    bomQtyCol = tblBoM.ListColumns("QTY").Index

    ' Identify columns in TblTimeAndMaterials
    tamInvIDCol = tblTAM.ListColumns("Inventory Item").Index
    tamQtyCol = tblTAM.ListColumns("QTY").Index
    tamMaintIDCol = tblTAM.ListColumns("Maintenance ID").Index
    tamBomIDCol = tblTAM.ListColumns("BoM ID").Index

    ' Unprotect once before the loop
    wsTAM.Unprotect Password:=pwd
   
    ' Disable events and screen updating for performance
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through BoM rows and insert into Time & Materials
    For Each bomRow In tblBoM.ListRows
        If bomRow.Range.Cells(1, bomIDCol).Value = bomID Then
            ' Add row and assign values cell-by-cell
            Set newRow = tblTAM.ListRows.Add
            With newRow.Range
                .Cells(1, tamInvIDCol).Value = bomRow.Range.Cells(1, bomInvIDCol).Value
                .Cells(1, tamQtyCol).Value = bomRow.Range.Cells(1, bomQtyCol).Value
                .Cells(1, tamMaintIDCol).Value = maintID
                .Cells(1, tamBomIDCol).Value = bomID
            
                ' Inline audit logic
                .Cells(1, tsCol).Value = Now
                .Cells(1, tsCol).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                .Cells(1, AppUserCol).Value = Application.userName
            End With
        End If
    Next bomRow

    MsgBox "BoM items inserted with Maintenance ID '" & maintID & "' and QTY values.", vbInformation

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Reprotect the sheet
    wsTAM.Protect Password:=pwd, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
End Sub








