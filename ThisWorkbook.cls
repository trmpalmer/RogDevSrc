'
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    Call ApplyChangeLogic(Target, Sh)
    On Error GoTo 0
End Sub
'Sub can be used to only show sheets by user
'Private Sub Workbook_Open()
'    Dim userName As String
'    userName = Environ("Username") ' Gets Windows login name
'
'    ' Hide all sheets first (except one to avoid error)
'    Dim ws As Worksheet
'    For Each ws In ThisWorkbook.Worksheets
'        ws.Visible = xlSheetVeryHidden
'        'ws.Visible = xlSheetVisible
'    Next ws
'
'    ' Always keep one sheet visible to avoid runtime error
'    Worksheets("Menu").Visible = xlSheetVisible
'
'    ' Show sheets based on user
'    Select Case userName
'        Case "Tim.Jones"
'            Worksheets("Finance").Visible = xlSheetVisible
'            Worksheets("Dashboard").Visible = xlSheetVisible
'        Case "Alice.Smith"
'            Worksheets("HR").Visible = xlSheetVisible
'        Case Else
'            MsgBox "You are not authorized to view this workbook.", vbExclamation
'            ThisWorkbook.Close SaveChanges:=False
''    End Select
'End Sub



Private Sub Workbook_Open()
    Dim userName As String
    userName = Application.userName ' Gets application login name
    Dim userList As Range
    Dim matchFound As Boolean
  
'  MsgBox "Workbook_Open fired"
'  On Error GoTo 0
'  Stop
  

 
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo CatchError
 
' If Application.userName = "Tim Palmer" Then Stop

    'Show admin button if user is...
    'userName = Application.userName ' Get Windows login name
    ' Ensure named range is scoped to the workbook
    matchFound = Not IsError(Application.Match(userName, ThisWorkbook.Names("VarSysAdmin").RefersToRange, 0))
    ' Only show button if user is in sys admin list
    With Sheets("Menu")
        .Shapes("BTNShowAdminTab").Visible = matchFound
    End With
    
    Exit Sub

CatchError:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Debugger Trigger"
    Stop  ' Break into debugger here

'    With Sheets("Menu")
'        Select Case userName
'            Case "Tim Palmer"
'                .Shapes("BTNShowAdminTab").Visible = True
'            Case Else
'                .Shapes("BTNShowAdminTab").Visible = False
'        End Select
'    End With

'    ' Hide all sheets first (except one to avoid error)
'    Dim ws As Worksheet
'    For Each ws In ThisWorkbook.Worksheets
'        ws.Visible = xlSheetVeryHidden
'        'ws.Visible = xlSheetVisible
'    Next ws
'
'    ' Always keep one sheet visible to avoid runtime error
'    Worksheets("Menu").Visible = xlSheetVisible
'
'    ' Show sheets based on user
'    Select Case userName
'        Case "Tim.Jones"
'            Worksheets("Finance").Visible = xlSheetVisible
'            Worksheets("Dashboard").Visible = xlSheetVisible
'        Case "Alice.Smith"
'            Worksheets("HR").Visible = xlSheetVisible
'        Case Else
'            MsgBox "You are not authorized to view this workbook.", vbExclamation
'            ThisWorkbook.Close SaveChanges:=False
''    End Select
End Sub
