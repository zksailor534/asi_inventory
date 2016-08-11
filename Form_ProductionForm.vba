Option Compare Database


'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
    setUserPermissions
End Sub


'------------------------------------------------------------
' setUserPermissions
'
'------------------------------------------------------------
Private Sub setUserPermissions()
    If (EmployeeRole = SalesLevel) Then
        nvbManageInventory.Enabled = False
        nvbManageInventory.Visible = True
        nvbAddItem.Enabled = False
        nvbAddItem.Visible = True
        nvbManageCommits.Enabled = True
        nvbManageCommits.Visible = True
        nvbCustomize.Enabled = True
        nvbCustomize.Visible = True
    ElseIf (EmployeeRole = ProdLevel) Then
        nvbManageInventory.Enabled = True
        nvbManageInventory.Visible = True
        nvbAddItem.Enabled = True
        nvbAddItem.Visible = True
        nvbManageCommits.Enabled = True
        nvbManageCommits.Visible = True
        nvbCustomize.Enabled = True
        nvbCustomize.Visible = True
    ElseIf (EmployeeRole = AdminLevel) Then
        nvbManageInventory.Enabled = True
        nvbManageInventory.Visible = True
        nvbAddItem.Enabled = True
        nvbAddItem.Visible = True
        nvbManageCommits.Enabled = True
        nvbManageCommits.Visible = True
        nvbCustomize.Enabled = True
        nvbCustomize.Visible = True
    ElseIf (EmployeeRole = DevelLevel) Then
        nvbManageInventory.Enabled = True
        nvbManageInventory.Visible = True
        nvbAddItem.Enabled = True
        nvbAddItem.Visible = True
        nvbManageCommits.Enabled = True
        nvbManageCommits.Visible = True
        nvbCustomize.Enabled = True
        nvbCustomize.Visible = True
    End If
End Sub
