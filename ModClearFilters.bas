Attribute VB_Name = "ModClearFilters"
' This is the parameterized subroutine (kept modular)
Sub ClearAllTableFilters(ByVal TableName As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim sc As SlicerCache
    Dim found As Boolean

    On Error GoTo SafeExit
    Application.ScreenUpdating = False

    ' Locate the table by name
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = TableName Then
                found = True
                If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
                Exit For
            End If
        Next lo
        If found Then Exit For
    Next ws

    If Not found Then
        MsgBox "Table '" & TableName & "' not found.", vbExclamation
        GoTo SafeExit
    End If

    ' Clear slicers connected to the table
    For Each sc In ThisWorkbook.SlicerCaches
        If Not sc.ListObject Is Nothing Then
            If sc.ListObject.Name = TableName Then
                sc.ClearManualFilter
            End If
        End If
    Next sc

SafeExit:
    Application.ScreenUpdating = True
End Sub

' This is the macro-safe wrapper you can assign to a button
Sub ClearFilters_MyTable()
    Call ClearAllTableFilters("MyTableName") ' Replace with your actual table name
End Sub

Sub ClearFilters_ActiveSheetTable()
    Dim btn As Button
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim i As Long
    Dim isConnectedOnlyToThisTable As Boolean

    On Error GoTo SafeExit
    Application.ScreenUpdating = False

    ' Identify the button and its parent sheet
    Set btn = ActiveSheet.Buttons(Application.Caller)
    Set ws = btn.Parent

    ' Validate there's exactly one table
    If ws.ListObjects.Count <> 1 Then
        MsgBox "Expected exactly one table on this sheet. Found " & ws.ListObjects.Count & ".", vbExclamation
        GoTo SafeExit
    End If

    Set lo = ws.ListObjects(1)

    ' Clear AutoFilter
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData

    ' Clear slicers connected ONLY to this table
    For Each sc In ThisWorkbook.SlicerCaches
        isConnectedOnlyToThisTable = True

        ' Check if slicer cache is connected to anything other than this table
        For i = 1 To sc.ListObjects.Count
            If Not sc.ListObjects(i) Is lo Then
                isConnectedOnlyToThisTable = False
                Exit For
            End If
        Next i
        
        If isConnectedOnlyToThisTable Then
            sc.ClearManualFilter
        End If
    Next sc

SafeExit:
    Application.ScreenUpdating = True
End Sub




