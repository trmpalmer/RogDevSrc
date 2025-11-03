'Private Sub Worksheet_Change(ByVal Target As Range)
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim tblRange As Range
'    Dim cell As Range
'    Dim tsCol As Long
'    Dim AppUser As Long
'    Dim pwd As String
'
'    Set ws = ThisWorkbook.Sheets("Asset Header")
'
'    ' Set your password here
'    pwd = ""
'
'    ' Unprotect the sheet
'    ws.Unprotect Password:=pwd
'
'    Set tbl = Me.ListObjects("TblAssetHeader")
'    Set tblRange = Intersect(tbl.DataBodyRange, Target)
'
'    If tblRange Is Nothing Then Exit Sub
'
'    ' Find the column number of [Modified Date]
'    On Error Resume Next
'    tsCol = tbl.ListColumns("Modified Date").Range.Column
'    AppUser = tbl.ListColumns("Modified User").Range.Column '****
'    On Error GoTo 0
'
'    If tsCol = 0 Then Exit Sub ' Column not found
'
'    Application.EnableEvents = False
'    For Each cell In tblRange
'    'If Not IsEmpty(cell.Value) And cell.Column <> tsCol And cell.Column <> User Then
'    If cell.Column <> tsCol And cell.Column <> User Then
'        Me.Cells(cell.Row, tsCol).Value = Now
'        Me.Cells(cell.Row, tsCol).NumberFormat = "dd-mm-yyyy hh:mm:ss"
'
'        Me.Cells(cell.Row, AppUser).Value = Application.UserName
'    End If
'
'    Next cell
'    Application.EnableEvents = True
'    ' Reprotect the sheet
'    ws.Protect Password:=pwd, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
'
'End Sub

'Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'    On Error Resume Next
'    Call ApplyChangeLogic(Target, Sh)
'    On Error GoTo 0
'End Sub


