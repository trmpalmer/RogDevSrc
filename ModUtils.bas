Attribute VB_Name = "ModUtils"
sub tims ()
end sub

Sub ShowLoggedInUsername()
    Dim windowsUser As String
    windowsUser = CreateObject("WScript.Network").userName
    MsgBox "Logged in as windows user: " & windowsUser
    MsgBox "Logged in as application user: " & Application.userName
End Sub


Sub AuditAndCleanNamedItems()
    Dim nm As Name
    Dim ws As Worksheet
    Dim reportWS As Worksheet
    Dim i As Long

    ' Create a new sheet for the audit report
    Set reportWS = ThisWorkbook.Sheets.Add
    reportWS.Name = "Named Items Audit"
    reportWS.Range("A1:E1").Value = Array("Name", "Refers To", "Visible", "Sheet Scope", "Status")
    
    i = 2
    For Each nm In ThisWorkbook.Names
        On Error Resume Next
        Dim status As String
        status = "OK"
        
        ' Check for broken references
        If InStr(nm.RefersTo, "#REF!") > 0 Then status = "Broken (#REF!)"
        If nm.RefersTo = "" Then status = "Empty"
        
        ' Log details
        reportWS.Cells(i, 1).Value = nm.Name
        reportWS.Cells(i, 2).Value = "'" & nm.RefersTo
        reportWS.Cells(i, 3).Value = nm.Visible
        reportWS.Cells(i, 4).Value = IIf(nm.Parent.Type = xlWorksheet, nm.Parent.Name, "Workbook")
        reportWS.Cells(i, 5).Value = status
        i = i + 1
        On Error GoTo 0
    Next nm

    ' Autofit for readability
    reportWS.Columns("A:E").AutoFit

    MsgBox "Audit complete. See 'Named Items Audit' sheet.", vbInformation
End Sub

Sub DiagnosePivotTableRefresh()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim logMsg As String
    Dim pwd As String

    pwd = GetPassword()
    Set ws = ActiveSheet
    logMsg = "?? PivotTable Refresh Diagnostics for '" & ws.Name & "':" & vbCrLf & vbCrLf

    On Error Resume Next
    ws.Unprotect Password:=pwd

    For Each pt In ws.PivotTables
        Err.Clear
        pt.PivotCache.Refresh

        If Err.Number <> 0 Then
            logMsg = logMsg & "? '" & pt.Name & "' failed to refresh." & vbCrLf & _
                     "    ? Error: " & Err.Description & vbCrLf & _
                     "    ? SourceData: " & pt.PivotCache.SourceData & vbCrLf & vbCrLf
            GoTo CleanExit
        Else
            logMsg = logMsg & "? '" & pt.Name & "' refreshed successfully." & vbCrLf
        End If
    Next pt

    MsgBox logMsg, vbInformation, "PivotTable Refresh Report"

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ws.Protect Password:=pwd, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True

'    ws.Protect Password:=pwd, _
'        AllowUsingPivotTables:=True, _
'        UserInterfaceOnly:=True
'    On Error GoTo 0


End Sub


Sub ListAllPivotTables()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim outputSheet As Worksheet
    Dim row As Long

    ' Create a new sheet for the output
    Set outputSheet = ThisWorkbook.Worksheets.Add
    outputSheet.Name = "PivotTableList2"

    ' Set headers
    With outputSheet
        .Cells(1, 1).Value = "Worksheet Name"
        .Cells(1, 2).Value = "PivotTable Name"
        .Cells(1, 3).Value = "Source Data"
        .Cells(1, 4).Value = "Table Address"
        .Cells(1, 5).Value = "Last Refreshed"
        .Cells(1, 6).Value = "Refreshed By"
        .Rows(1).Font.Bold = True
    End With

    row = 2

    ' Loop through all sheets and PivotTables
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            With outputSheet
                .Cells(row, 1).Value = ws.Name
                .Cells(row, 2).Value = pt.Name
                .Cells(row, 3).Value = pt.SourceData
                .Cells(row, 4).Value = pt.TableRange2.Address
                .Cells(row, 5).Value = pt.RefreshDate
                .Cells(row, 6).Value = pt.RefreshName
            End With
            row = row + 1
        Next pt
    Next ws

    outputSheet.Columns("A:F").AutoFit
    MsgBox "PivotTable list created on sheet: " & outputSheet.Name, vbInformation
End Sub

Function WorksheetExists(sheetName As String) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Integer

    On Error Resume Next
    For i = 1 To Application.Workbooks.Count
        Set wb = Application.Workbooks(i)
        Set ws = wb.Worksheets(sheetName)
        If Not ws Is Nothing Then
            WorksheetExists = True
            Exit Function
        End If
    Next i
    WorksheetExists = False
    On Error GoTo 0
End Function


Sub TestContext()
    MsgBox "Workbook count: " & Application.Workbooks.Count
End Sub
Sub tim()
  MsgBox "Hello from tim"
End Sub





