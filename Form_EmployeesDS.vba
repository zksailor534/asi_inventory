Option Compare Database


'------------------------------------------------------------
' Form_Open
'
'------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

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
        Me.Role.RowSource = "'" & SalesLevel & "';'" & ProdLevel & "';'" & AdminLevel & "'"

        ' Hide version column from admin
        Me.ID.ColumnHidden = False
        Me.Login.ColumnHidden = False
        Me.Password.ColumnHidden = False
        Me.Role.ColumnHidden = False
        Me.Active.ColumnHidden = False
        Me.DefaultCategory.ColumnHidden = False
        Me.Version.ColumnHidden = True

        ' Allow to make any changes
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
    ElseIf (EmpRole = DevelLevel) Then
        ' Allow devel to set all roles
        Me.Role.RowSource = "'" & SalesLevel & "';'" & ProdLevel & "';'" & AdminLevel & "';'" & DevelLevel & "'"

        ' Do not hide any columns
        Me.ID.ColumnHidden = False
        Me.Login.ColumnHidden = False
        Me.Password.ColumnHidden = False
        Me.Role.ColumnHidden = False
        Me.Active.ColumnHidden = False
        Me.DefaultCategory.ColumnHidden = False
        Me.Version.ColumnHidden = False

        ' Allow to make any changes
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
    Else
        ' Hide all columns but names and emails
        Me.ID.ColumnHidden = True
        Me.Login.ColumnHidden = True
        Me.Password.ColumnHidden = True
        Me.Role.ColumnHidden = True
        Me.Active.ColumnHidden = True
        Me.DefaultCategory.ColumnHidden = True
        Me.Version.ColumnHidden = True

        ' Do not allow any changes
        Me.AllowAdditions = False
        Me.AllowDeletions = False
        Me.AllowEdits = False
    End If
End Sub
