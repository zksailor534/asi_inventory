Option Compare Database


Private Sub Form_Load()
    screenSize
    setUserPermissions
End Sub


Private Sub Form_Resize()
    screenSize
End Sub


Private Sub screenSize()
    On Error Resume Next
    Me.Form.Width = Round(Me.InsideWidth * 0.9)
End Sub


Private Sub setUserPermissions()
    If (EmployeeRole = SalesLevel) Then
        nvbSearch.Enabled = True
        nvbSearch.Visible = True
        nvbAddItem.Enabled = False
        nvbAddItem.Visible = True
        nvbManageInventory.Enabled = False
        nvbManageInventory.Visible = True
        nvbManageCommits.Enabled = True
        nvbManageCommits.Visible = True
        nvbCustomize.Enabled = True
        nvbCustomize.Visible = True
        nvbUtilities.Enabled = False
        nvbUtilities.Visible = False
    ElseIf (EmployeeRole = ProdLevel) Then
        nvbSearch.Enabled = True
        nvbSearch.Visible = True
        nvbAddItem.Enabled = True
        nvbAddItem.Visible = True
        nvbManageInventory.Enabled = True
        nvbManageInventory.Visible = True
        nvbManageCommits.Enabled = True
        nvbManageCommits.Visible = True
        nvbCustomize.Enabled = True
        nvbCustomize.Visible = True
        nvbUtilities.Enabled = False
        nvbUtilities.Visible = False
    ElseIf (EmployeeRole = AdminLevel) Then
        nvbSearch.Enabled = True
        nvbSearch.Visible = True
        nvbAddItem.Enabled = True
        nvbAddItem.Visible = True
        nvbManageInventory.Enabled = True
        nvbManageInventory.Visible = True
        nvbManageCommits.Enabled = True
        nvbManageCommits.Visible = True
        nvbCustomize.Enabled = True
        nvbCustomize.Visible = True
        nvbUtilities.Enabled = False
        nvbUtilities.Visible = False
    ElseIf (EmployeeRole = DevelLevel) Then
        nvbSearch.Enabled = True
        nvbSearch.Visible = True
        nvbAddItem.Enabled = True
        nvbAddItem.Visible = True
        nvbManageInventory.Enabled = True
        nvbManageInventory.Visible = True
        nvbManageCommits.Enabled = True
        nvbManageCommits.Visible = True
        nvbCustomize.Enabled = True
        nvbCustomize.Visible = True
        nvbUtilities.Enabled = True
        nvbUtilities.Visible = True
    End If
End Sub
