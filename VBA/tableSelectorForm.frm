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