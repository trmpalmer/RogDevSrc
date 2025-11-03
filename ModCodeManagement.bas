Attribute VB_Name = "ModCodeManagement"
Sub ExportVBAModules()
    Dim vbComp As VBComponent
    Dim exportPath As String
    Dim countExported As Long
    Dim fileName As String

    exportPath = "C:\Users\trmpa\OneDrive - Rail Operation Group\Class 93 Project\Dev and Prod\DevSrcVba\"
    countExported = 0

    MsgBox "Starting export of VBA modules to Git folder:" & vbCrLf & exportPath, vbInformation, "VBA Export"

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        fileName = exportPath & vbComp.Name & "." & GetExtension(vbComp.Type)
        vbComp.Export fileName
        countExported = countExported + 1
        Debug.Print "Exported: " & fileName
    Next vbComp

    MsgBox "Export complete." & vbCrLf & _
           "Modules exported: " & countExported & vbCrLf & _
           "Location: " & exportPath, vbInformation, "VBA Export Summary"
End Sub

Function GetExtension(compType As vbext_ComponentType) As String
    Select Case compType
        Case vbext_ct_StdModule: GetExtension = "bas"
        Case vbext_ct_ClassModule: GetExtension = "cls"
        Case vbext_ct_MSForm: GetExtension = "frm"
        Case Else: GetExtension = "txt"
    End Select
End Function




