Option Compare Database
Option Explicit


'------------------------------------------------------------
' Form_Open
'
'------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

    open_db
    ScreenWidth = Round(Me.WindowWidth - 325)
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
' Form_Resize
'
'------------------------------------------------------------
Private Sub Form_Resize()
On Error Resume Next
    ScreenWidth = Round(Me.WindowWidth - 325)
    If Utilities.ProcedureExists(Me!NavigationSubform.Form!NavigationSubform.Form, "SetScreenSize") Then
         Me!NavigationSubform.Form!NavigationSubform.Form.SetScreenSize
    End If

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
    If (EmployeeRole = SalesLevel) Or (EmployeeRole = ProdLevel) Then
        nvbInventory.Enabled = True
        nvbAdmin.Visible = False
        nvbAdmin.Enabled = False
    ElseIf (EmployeeRole = AdminLevel) Or (EmployeeRole = DevelLevel) Then
        nvbInventory.Enabled = True
        nvbAdmin.Visible = True
        nvbAdmin.Enabled = True
    End If
End Sub
