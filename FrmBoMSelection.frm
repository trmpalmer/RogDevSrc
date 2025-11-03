VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmBoMSelection 
   Caption         =   "FrmBoMSelection"
   ClientHeight    =   3040
   ClientLeft      =   50
   ClientTop       =   150
   ClientWidth     =   3540
   OleObjectBlob   =   "FrmBoMSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmBoMSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnCancel_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub BtnOK_Click()
    If Me.CmbBomID.Value = "" Or Me.CMBMaintenanceID.Value = "" Then
        MsgBox "Please select both BoM ID and Maintenance ID.", vbExclamation
        Exit Sub
    End If
    Me.Tag = Me.CmbBomID.Value & "|" & Me.CMBMaintenanceID.Value
    Me.Hide

End Sub



Private Sub UserForm_Initialize()
    Call PopulateBoMIDCombo(Me.CmbBomID)
    Call PopulateCMBMaintenanceID(Me.CMBMaintenanceID)
End Sub

'Sub InsertBoMItemsViaForm()
'    Dim frm As frmSelectBoM
'    Dim selectedBoMID As String
'    Dim wsBoM As Worksheet, wsTAM As Worksheet
'    Dim tblBoM As ListObject, tblTAM As ListObject
'    Dim bomRow As ListRow, newRow As ListRow
'    Dim bomIDCol As Long, bomInvIDCol As Long, bomDescCol As Long
'    Dim tamInvIDCol As Long, tamDescCol As Long
'
'    ' Show form
'    Set frm = New frmSelectBoM
'    frm.Show
'
'    selectedBoMID = frm.Tag
'    Unload frm
'
'    If selectedBoMID = "" Then
'        MsgBox "Operation cancelled.", vbExclamation
'        Exit Sub
'    End If
'
'    ' Set worksheets and tables
'    Set wsBoM = ThisWorkbook.Sheets("BoM")
'    Set wsTAM = ThisWorkbook.Sheets("Time & Materials")
'    Set tblBoM = wsBoM.ListObjects("TblBoM")
'    Set tblTAM = wsTAM.ListObjects("TblTimeAndMaterials")
'
'    ' Identify columns
'    bomIDCol = tblBoM.ListColumns("BoMID").Index
'    bomInvIDCol = tblBoM.ListColumns("InventoryID").Index
'    bomDescCol = tblBoM.ListColumns("Description").Index
'    tamInvIDCol = tblTAM.ListColumns("InventoryID").Index
'    tamDescCol = tblTAM.ListColumns("Description").Index
'
'    ' Loop and insert
'    For Each bomRow In tblBoM.ListRows
'        If bomRow.Range.Cells(1, bomIDCol).Value = selectedBoMID Then
'            Set newRow = tblTAM.ListRows.Add
'            newRow.Range.Cells(1, tamInvIDCol).Value = bomRow.Range.Cells(1, bomInvIDCol).Value
'            newRow.Range.Cells(1, tamDescCol).Value = bomRow.Range.Cells(1, bomDescCol).Value
'        End If
'    Next bomRow
'
'    MsgBox "BoM items inserted into Time & Materials table.", vbInformation
'End Sub
