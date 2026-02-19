VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tableSelectorForm 
   Caption         =   "TableSelectorForm"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "tableSelectorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tableSelectorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public SelectedTables As Collection
Public Cancelled As Boolean

Private Sub btnImport_Click()
    Set SelectedTables = New Collection
    Dim i As Integer
    For i = 0 To lstTables.ListCount - 1
        If lstTables.Selected(i) Then SelectedTables.Add lstTables.List(i)
    Next i
    
    If SelectedTables.Count = 0 Then
        MsgBox "Please select at least one table.", vbExclamation
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub UserForm_Click()

End Sub
