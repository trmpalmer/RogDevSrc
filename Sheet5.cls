Sub TblAssetDocumentAddRow()
'
' TblAssetDocumentAddRow Macro
'
    Dim pwd As String
    Dim ws As Worksheet
   
    
    
    ' Set your password here
    pwd = GetPassword()
    Set ws = ThisWorkbook.Sheets("Asset Documents")
    
    ' Unprotect the sheet
    ws.Unprotect Password:=pwd

    Range("TblAssetDocuments[AssetID]").Select
    Selection.ListObject.ListRows.Add AlwaysInsert:=True
    Range("C5").Select
    
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ws.Protect Password:=pwd, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
End Sub

