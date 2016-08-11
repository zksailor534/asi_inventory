Option Compare Database


Private Sub Form_Load()
    screenSize
    'UserRoleProperties
End Sub


Private Sub Form_Resize()
    screenSize
End Sub


Private Sub screenSize()
    On Error Resume Next
    Me.Form.Width = Round(Me.InsideWidth * 0.9)
End Sub


Private Sub UserRoleProperties()
    If (EmployeeRole = SalesLevel) Then
    ElseIf (EmployeeRole = ProdLevel) Then
    ElseIf (EmployeeRole = AdminLevel) Then
    ElseIf (EmployeeRole = DevelLevel) Then
    End If
End Sub
