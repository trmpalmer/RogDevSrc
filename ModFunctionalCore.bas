Attribute VB_Name = "ModFunctionalCore"
Function GetPassword() As String
    GetPassword = "magyar"
End Function


Sub DisableEvents()
    Application.EnableEvents = False
End Sub


Sub ReenableEvents()
    Application.EnableEvents = True
End Sub

Sub ProtectAllSheets()
    Dim ws As Worksheet
    Dim pwd As String

    ' Define your password
    pwd = GetPassword()

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        With ws
            ' Unprotect first in case it's already protected
            .Unprotect Password:=pwd

            ' Apply protection with desired criteria
            .Protect Password:=pwd, _
                     AllowFormattingCells:=True, _
                     AllowSorting:=True, _
                     AllowFiltering:=True, _
                     AllowUsingPivotTables:=True, _
                     AllowInsertingRows:=False, _
                     AllowDeletingRows:=False, _
                     DrawingObjects:=True, _
                     Scenarios:=True, _
                     UserInterfaceOnly:=True, _
                     AllowFormattingColumns:=True, _
                     AllowFormattingRows:=True

            .EnableOutlining = True

            ' Control what users can select
            .EnableSelection = xlUnlockedCells ' Options: xlNoRestrictions, xlNoSelection
        End With
    Next ws

    ' Protect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=True

    ' Activate the workbook and the active sheet
    ThisWorkbook.Activate
    ThisWorkbook.ActiveSheet.Activate

    Application.EnableEvents = True

    'MsgBox "All sheets and workbook structure protected with consistent criteria.", vbInformation
End Sub

Sub UnprotectAllSheets()
    Dim ws As Worksheet
    Dim pwd As String

    ' Define your password
    pwd = GetPassword()

    ' Loop through each worksheet and unprotect
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next ' In case sheet is already unprotected or wrong password
        ws.Unprotect Password:=pwd
        On Error GoTo 0
    Next ws

    ' UnProtect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=False
    
    ' Activate the workbook and the active sheet
    ThisWorkbook.Activate
    ThisWorkbook.ActiveSheet.Activate

    Application.EnableEvents = False
    
    'MsgBox "All sheets have been unprotected.", vbInformation
End Sub

Sub SetDevOrProdColours()
'?? Important Note:
'Macros must be enabled when you open the file for this to run. If macros are disabled, the Workbook_Open() event wont execute..

    
    Dim fileName As String
    Dim sheet As Worksheet
    Dim environment As String
    Dim pwd As String

    ' Define your password
    pwd = GetPassword()
    
    ' UnProtect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=False
    
    Call UnprotectAllSheets

    ' Get the file name without the extension
    fileName = ThisWorkbook.Name
    fileName = Left(fileName, InStrRev(fileName, ".") - 1)

    ' Determine if the file is for Development or Production based on the file name
    If InStr(1, fileName, "Dev", vbTextCompare) > 0 Or InStr(1, fileName, "Development", vbTextCompare) > 0 Then
        environment = "DevVersion"
    ElseIf InStr(1, fileName, "Live", vbTextCompare) > 0 Or InStr(1, fileName, "Live", vbTextCompare) > 0 Then
        environment = "LiveVersion"
    Else
        environment = "Unknown"
    End If

    ' Loop through each worksheet in the workbook
    For Each sheet In ThisWorkbook.Sheets
        ' Set the text in row 1 based on the environment
        sheet.Rows(1).Value = environment

        ' Set colors based on environment
        Select Case environment
            Case "LiveVersion"
                sheet.Rows(1).Interior.Color = RGB(144, 238, 144) ' Light green
                sheet.Rows(1).Font.Color = RGB(0, 100, 0)         ' Dark green
            Case "DevVersion"
                sheet.Rows(1).Interior.Color = RGB(255, 0, 0)   ' Green
                sheet.Rows(1).Font.Color = RGB(0, 0, 0)         ' Black
            Case Else
                sheet.Rows(1).Interior.Color = RGB(211, 211, 211) ' Light gray
                sheet.Rows(1).Font.Color = RGB(169, 169, 169)     ' Dark gray
        End Select
    Next sheet
    
    ' Protect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=True
    Call ProtectAllSheets
   ' MsgBox "All sheets have been set to either Dev or Production based on file name.", vbInformation
End Sub

Sub ShowAdminTab()
    Dim pwd As String

    ' Define your password
    pwd = GetPassword()
    
        ' UnProtect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=False
    
    Worksheets("Admin").Visible = xlSheetVisible
    
            ' Protect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=True
    
End Sub
Sub HideAdminTab()
    Dim pwd As String

    ' Define your password
    pwd = GetPassword()

    ' UnProtect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=False
    
    Worksheets("Admin").Visible = xlSheetVeryHidden
    
    ' Protect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=True
End Sub
Sub ShowAllTabs()
    ' Hide all sheets first (except one to avoid error)
    Dim ws As Worksheet
    Dim pwd As String

    ' Define your password
    pwd = GetPassword()

    ' UnProtect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=False
    For Each ws In ThisWorkbook.Worksheets
        'ws.Visible = xlSheetVeryHidden
        ws.Visible = xlSheetVisible
    Next ws
    ' Protect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=True
End Sub
Sub ShowOnlyAssetAndMntTabs()
    Dim userName As String
    Dim pwd As String

    ' Define your password
    pwd = GetPassword()
    ' UnProtect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=False
    
    userName = Environ("Username") ' Gets Windows login name

    ' Hide all sheets except "Menu" first (except one to avoid error)
    Dim ws As Worksheet
    Dim keepVisible As String
    keepVisible = "Menu" ' Change to the sheet you want to keep visible

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> keepVisible Then
            ws.Visible = xlSheetVeryHidden
        Else
            ws.Visible = xlSheetVisible
        End If
    Next ws


    ' Show desired Tabs
    Worksheets("Menu").Visible = xlSheetVisible
    Worksheets("Asset Header").Visible = xlSheetVisible
    Worksheets("Asset Config.").Visible = xlSheetVisible
    Worksheets("Asset Documents").Visible = xlSheetVisible
    Worksheets("Insurance Records").Visible = xlSheetVisible
    Worksheets("Event Log").Visible = xlSheetVisible
    Worksheets("Maintenance Visit").Visible = xlSheetVisible
    Worksheets("Maintenance - Events Linkage").Visible = xlSheetVisible
    Worksheets("Time & Materials").Visible = xlSheetVisible
    Worksheets("PVTEvents For Mnt. Review").Visible = xlSheetVisible
    Worksheets("PVTEventLog").Visible = xlSheetVisible
    Worksheets("Admin").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Event Log").Activate

    ' Show sheets based on user
'    Select Case userName
'        Case "Tim.Jones"
'            Worksheets("Finance").Visible = xlSheetVisible
'            Worksheets("Dashboard").Visible = xlSheetVisible
'        Case "Alice.Smith"
'            Worksheets("HR").Visible = xlSheetVisible
'        Case Else
'            MsgBox "You are not authorized to view this workbook.", vbExclamation
'            ThisWorkbook.Close SaveChanges:=False
'    End Select
    ' Protect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=True
End Sub

Sub ShowInventoryTabs()
    Dim userName As String
    Dim pwd As String

    ' Define your password
    pwd = GetPassword()
    ' UnProtect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=False
    
    userName = Environ("Username") ' Gets Windows login name

    ' Hide all sheets except "Menu" first (except one to avoid error)
    Dim ws As Worksheet
    Dim keepVisible As String
    keepVisible = "Menu" ' Change to the sheet you want to keep visible

'    For Each ws In ThisWorkbook.Worksheets
'        If ws.Name <> keepVisible Then
'            ws.Visible = xlSheetVeryHidden
'        Else
'            ws.Visible = xlSheetVisible
'        End If
'    Next ws


    ' Show desired Tabs
    Worksheets("Menu").Visible = xlSheetVisible
    Worksheets("Inventory").Visible = xlSheetVisible
    Worksheets("BoM").Visible = xlSheetVisible
    Worksheets("Suppliers & Products").Visible = xlSheetVisible
    Worksheets("Admin").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Inventory").Activate

    ' Show sheets based on user
'    Select Case userName
'        Case "Tim.Jones"
'            Worksheets("Finance").Visible = xlSheetVisible
'            Worksheets("Dashboard").Visible = xlSheetVisible
'        Case "Alice.Smith"
'            Worksheets("HR").Visible = xlSheetVisible
'        Case Else
'            MsgBox "You are not authorized to view this workbook.", vbExclamation
'            ThisWorkbook.Close SaveChanges:=False
'    End Select
    ' Protect workbook structure to prevent sheet deletion, renaming, or moving
    ThisWorkbook.Protect Password:=pwd, Structure:=True
End Sub


Public Sub ApplyChangeLogic(ByVal Target As Range, ByVal ws As Worksheet)
    Dim tbl As ListObject
    Dim tblRange As Range
    Dim cell As Range
    Dim tsCol As Long
    Dim AppUser As Long
    Dim pwd As String
    Dim tableMap As Collection
    Dim mapItem As Variant
    Dim mapKey As String
    Dim foundTableName As String
    Dim i As Long
    Dim lastRowIndex As Long
    Dim lastRow As Range
    Dim newRow As Range
    Dim colIndex As Long

    If ws Is Nothing Then Exit Sub
    
    pwd = GetPassword()
    On Error Resume Next
    ws.Unprotect Password:=pwd
    On Error GoTo 0
    
    ' Create mapping of worksheet names to table names
    Set tableMap = New Collection
    tableMap.Add Array("Asset Header", "TblAssetHeader")
    tableMap.Add Array("Asset Config.", "TblAssetConfig")
    tableMap.Add Array("Asset Documents", "TblAssetDocuments")
    tableMap.Add Array("Insurance Records", "TblInsuranceRecords")
    tableMap.Add Array("Event Log", "TblEventLog")
    tableMap.Add Array("Maintenance Visit", "TblMaintenanceRecord")
    tableMap.Add Array("Maintenance - Events Linkage", "TblMaintenanceEventLinks")
    tableMap.Add Array("Time & Materials", "TblTimeAndMaterials")
    tableMap.Add Array("Inventory", "TblInventory")
    tableMap.Add Array("BoM", "TblBoM")
    ' Add more mappings as needed...
    
    ' Lookup table name based on worksheet name
    foundTableName = ""
    For i = 1 To tableMap.Count
        mapItem = tableMap(i)
        If mapItem(0) = ws.Name Then
            foundTableName = mapItem(1)
            Exit For
        End If
    Next i
    
    If foundTableName = "" Then GoTo CleanExit
    Set tbl = ws.ListObjects(foundTableName)
    
    Set tblRange = Intersect(tbl.DataBodyRange, Target)
    If tblRange Is Nothing Then GoTo CleanExit
    
    'Check if a new row was added and copy formatting from previous last row
    lastRowIndex = tbl.ListRows.Count - 1
    If lastRowIndex > 0 Then
        Set lastRow = tbl.ListRows(lastRowIndex).Range
        Set newRow = tbl.ListRows(tbl.ListRows.Count).Range
    
        ' Only copy if newRow is different from lastRow
        If Not newRow.Address = lastRow.Address Then
            For colIndex = 1 To lastRow.Columns.Count
                With newRow.Cells(1, colIndex)
                    '.Interior.Color = lastRow.Cells(1, colIndex).Interior.Color
                    .Font.Name = lastRow.Cells(1, colIndex).Font.Name
                    .Font.Size = lastRow.Cells(1, colIndex).Font.Size
                    .Font.Bold = lastRow.Cells(1, colIndex).Font.Bold
                    .NumberFormat = lastRow.Cells(1, colIndex).NumberFormat
                    ' Add more formatting as needed
                End With
            Next colIndex
        End If
    End If
    
    
    On Error Resume Next
    tsCol = tbl.ListColumns("Modified Date").Range.Column
    AppUser = tbl.ListColumns("Modified User").Range.Column
    On Error GoTo 0
    
    If tsCol = 0 Or AppUser = 0 Then GoTo CleanExit
    
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    For Each cell In tblRange
        If cell.Column <> tsCol And cell.Column <> AppUser Then
            ws.Cells(cell.row, tsCol).Value = Now
            ws.Cells(cell.row, tsCol).NumberFormat = "dd/mm/yyyy hh:mm:ss"
            ws.Cells(cell.row, AppUser).Value = Application.userName
        End If
    Next cell
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    ws.Protect Password:=pwd, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
    On Error GoTo 0
End Sub


Sub ShowAdminButtonForAdmins()
    Dim userName As String
    Dim userList As Range
    Dim matchFound As Boolean

    userName = Application.userName ' Get Windows login name
    ' Ensure named range is scoped to the workbook
    matchFound = Not IsError(Application.Match(userName, ThisWorkbook.Names("VarSysAdmin").RefersToRange, 0))

    With Sheets("Menu")
        .Shapes("BTNShowAdminTab").Visible = matchFound
    End With
End Sub


Sub RefreshPivotTablesSafely()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pwd As String
    Dim sheetsWithPivots As Collection
    Dim sheetName As Variant
    Dim callingSheet As Worksheet
    
    Set callingSheet = ActiveSheet


    On Error GoTo CleanExit

    pwd = GetPassword()
    Set sheetsWithPivots = New Collection

    ' Unprotect workbook structure
    ThisWorkbook.Unprotect Password:=pwd

    ' Phase 1: Unprotect sheets with PivotTables
'    For Each ws In ThisWorkbook.Worksheets
'        If ws.PivotTables.Count > 0 Then
'            ws.Unprotect Password:=pwd
'            sheetsWithPivots.Add ws.Name
'        End If
'    Next ws

Call UnprotectAllSheets

    ' Phase 2: Refresh all PivotTables
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            On Error Resume Next
            'Debug.Print Worksheets("PVTEvent").Visible

            pt.RefreshTable
            If Err.Number <> 0 Then
                Debug.Print "? Error refreshing '" & pt.Name & "' on '" & ws.Name & "': " & Err.Description
                Err.Clear
            End If
            On Error GoTo CleanExit
        Next pt
    Next ws

    ' Phase 3: Reprotect sheets that had PivotTables
    For Each sheetName In sheetsWithPivots
        Set ws = ThisWorkbook.Sheets(sheetName)
        ws.Protect Password:=pwd, _
            AllowUsingPivotTables:=True, _
            AllowFiltering:=True, _
            AllowSorting:=True, _
            UserInterfaceOnly:=True
    Next sheetName
    
    Call ProtectAllSheets

    ' Reprotect workbook structure
    ThisWorkbook.Protect Password:=pwd, Structure:=True
    
    ' Reset focus to sheet where refresh button was clicked
    If Not callingSheet Is Nothing Then callingSheet.Activate

    MsgBox "All report data has refreshed successfully.", vbInformation

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Sub InsertHyperlinkViaPrompt()
    Dim ws As Worksheet
    Dim pwd As String
    Dim targetCell As Range
    Dim linkAddress As String
    Dim linkText As String
    Dim cellref As String

    Set ws = ActiveSheet
    pwd = GetPassword() ' Replace with your password logic if needed

    ' Use currently selected cell as default
    cellref = InputBox("Enter the cell address to insert the hyperlink:", Default:=ActiveCell.Address(False, False))

    If cellref = "" Then Exit Sub

    On Error Resume Next
    Set targetCell = ws.Range(cellref)
    On Error GoTo 0

    If targetCell Is Nothing Then
        MsgBox "Invalid cell address. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    ' Prompt for hyperlink address
    linkAddress = InputBox("Enter the hyperlink address (e.g., https://example.com):")
    If linkAddress = "" Then
        MsgBox "No address entered. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    ' Prompt for display text
    linkText = InputBox("Enter the display text for the hyperlink:")
    If linkText = "" Then
        MsgBox "No display text entered. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    ' Unprotect, insert hyperlink, shift focus, re-protect
    With ws
        .Unprotect Password:=pwd
        .Hyperlinks.Add Anchor:=targetCell, Address:=linkAddress, TextToDisplay:=linkText
        On Error Resume Next
        targetCell.Offset(0, 1).Select ' Safely shift focus
        On Error GoTo 0
        .Protect Password:=pwd, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
    End With

    MsgBox "Hyperlink inserted successfully.", vbInformation
End Sub
