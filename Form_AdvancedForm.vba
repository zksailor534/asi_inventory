Option Compare Database
Option Explicit


'------------------------------------------------------------
' Form_Open
'
'------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

    UserRoleSettings

End Sub


'------------------------------------------------------------
' UserRoleSettings
'
'------------------------------------------------------------
Private Sub UserRoleSettings()
    If (EmployeeRole = SalesLevel) Then
        nvbPrice.Enabled = False
        nvbCategories.Enabled = False
        nvbProducts.Enabled = False
        nvbEmployees.Enabled = False
        nvbUtilities.Enabled = False
    ElseIf (EmployeeRole = ProdLevel) Then
        nvbPrice.Enabled = False
        nvbCategories.Enabled = False
        nvbProducts.Enabled = False
        nvbEmployees.Enabled = False
        nvbUtilities.Enabled = False
    ElseIf (EmployeeRole = AdminLevel) Then
        nvbPrice.Enabled = True
        nvbCategories.Enabled = True
        nvbProducts.Enabled = True
        nvbEmployees.Enabled = True
        nvbUtilities.Enabled = False
    ElseIf (EmployeeRole = DevelLevel) Then
        nvbPrice.Enabled = True
        nvbCategories.Enabled = True
        nvbProducts.Enabled = True
        nvbEmployees.Enabled = True
        nvbUtilities.Enabled = True
    End If
End Sub
