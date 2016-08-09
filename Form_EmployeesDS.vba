Option Compare Database


'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()

    ' Set user role properties
    UserRoleProperties (EmployeeRole)

End Sub


'------------------------------------------------------------
' UserRoleProperties
'
'------------------------------------------------------------
Private Sub UserRoleProperties(EmpRole As String)
    On Error Resume Next

    If (EmpRole = AdminLevel) Then
        ' Allow admin to view and set roles choices except devel
        ' (Admin cannot make self Devel)
        Me.cboRole.RowSource = SalesLevel & ";" & ProdLevel & ";" & AdminLevel

        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
    ElseIf (EmpRole = DevelLevel) Then
        ' Hide Delete Commit button for production
        Me.cboRole.RowSource = SalesLevel & ";" & ProdLevel & ";" & AdminLevel & ";" & DevelLevel

        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
    Else
        Me.AllowAdditions = False
        Me.AllowDeletions = False
        Me.AllowEdits = False
    End If
End Sub
