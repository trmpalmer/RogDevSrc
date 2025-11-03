VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalender 
   Caption         =   "Pick a Date"
   ClientHeight    =   3940
   ClientLeft      =   70
   ClientTop       =   280
   ClientWidth     =   4110
   OleObjectBlob   =   "frmCalender.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
    Dim cell As Object
    For Each cell In Selection.Cells
       cell.Value = DateClicked
    Next cell
    Unload Me
End Sub


Private Sub UserForm_Click()

End Sub
