Option Compare Database
Option Explicit


'------------------------------------------------------------
' Form_Open
'
'------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

    open_db
    screenSize
    setUserPermissions

End Sub


'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    On Error GoTo 0
    Utilities.LoadSettings "American Surplus"
    Utilities.ConfirmLogin
    screenSize

Form_Load_Exit:
    Exit Sub

Form_Load_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume Form_Load_Exit

End Sub


'------------------------------------------------------------
' cmdNotyou_Click
'
'------------------------------------------------------------
Private Sub cmdNotyou_Click()
On Error GoTo cmdNotyou_Click_Err

    ValidLogin = False
    DoCmd.OpenForm LoginForm, acNormal, "", "", , acDialog
    Utilities.CompleteLogin

cmdNotyou_Click_Exit:
    Exit Sub

cmdNotyou_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdNotyou_Click_Exit

End Sub


'------------------------------------------------------------
' screenSize
'
'------------------------------------------------------------
Private Sub screenSize()
    On Error Resume Next
    Me.Form.Width = Round(Me.WindowWidth * 0.9)
End Sub


'------------------------------------------------------------
' Form_Resize
'
'------------------------------------------------------------
Private Sub Form_Resize()
    screenSize
End Sub


'------------------------------------------------------------
' nvcNavigationControl_Enter
'
'------------------------------------------------------------
Private Sub nvcNavigationControl_Enter()
    DoCmd.SetWarnings False
End Sub


'------------------------------------------------------------
' setUserPermissions
'
'------------------------------------------------------------
Private Sub setUserPermissions()
    If (EmployeeRole = SalesLevel) Then
        nvbInventory.Enabled = True
        nvbAdvanced.Visible = False
        nvbAdvanced.Enabled = False
    ElseIf (EmployeeRole = ProdLevel) Then
        nvbInventory.Enabled = True
        nvbAdvanced.Visible = False
        nvbAdvanced.Enabled = False
    ElseIf (EmployeeRole = AdminLevel) Then
        nvbInventory.Enabled = True
        nvbAdvanced.Visible = True
        nvbAdvanced.Enabled = True
    ElseIf (EmployeeRole = DevelLevel) Then
        nvbInventory.Enabled = True
        nvbAdvanced.Visible = True
        nvbAdvanced.Enabled = True
    End If
End Sub
