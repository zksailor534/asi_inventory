Option Compare Database
Option Explicit


'------------------------------------------------------------
' Form_Open
'
'------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

    open_db
    screenSize
    UserRoleSettings

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
' UserRoleSettings
'
'------------------------------------------------------------
Private Sub UserRoleSettings()
    If (EmployeeRole = SalesLevel) Then
        nvbSales.Enabled = True
        nvbProduction.Enabled = False
        nvbAdmin.Visible = False
        nvbAdmin.Enabled = False
    ElseIf (EmployeeRole = ProdLevel) Then
        Me.nvbProduction.SetFocus
        SendKeys "{ENTER}", 0
        nvbSales.Enabled = False
        nvbProduction.Enabled = True
        nvbAdmin.Visible = False
        nvbAdmin.Enabled = False
    ElseIf (EmployeeRole = AdminLevel) Then
        Me.nvbProduction.SetFocus
        SendKeys "{ENTER}", 0
        nvbSales.Enabled = True
        nvbProduction.Enabled = True
        nvbAdmin.Visible = True
        nvbAdmin.Enabled = True
    ElseIf (EmployeeRole = DevelLevel) Then
        nvbSales.Enabled = True
        nvbProduction.Enabled = True
        nvbAdmin.Visible = True
        nvbAdmin.Enabled = True
    End If
End Sub
