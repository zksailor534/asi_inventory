Option Compare Database


'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
    screenSize
    setUserPermissions
End Sub


'------------------------------------------------------------
' Form_Resize
'
'------------------------------------------------------------
Private Sub Form_Resize()
    screenSize
End Sub


'------------------------------------------------------------
' screenSize
'
'------------------------------------------------------------
Private Sub screenSize()
    On Error Resume Next
    Me.Form.Width = Round(Me.InsideWidth * 0.9)
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
