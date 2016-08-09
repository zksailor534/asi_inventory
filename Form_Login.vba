Option Compare Database

'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    DoCmd.GoToControl "lstNames"

Form_Load_Exit:
    Exit Sub

Form_Load_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume Form_Load_Exit

End Sub


'------------------------------------------------------------
' lstNames_DblClick
'
'------------------------------------------------------------
Private Sub lstNames_DblClick(Cancel As Integer)
On Error GoTo lstNames_DblClick_Err

    txtPassword.SetFocus

lstNames_DblClick_Exit:
    Exit Sub

lstNames_DblClick_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume lstNames_DblClick_Exit

End Sub


'------------------------------------------------------------
' cmdLogin_Click
'
'------------------------------------------------------------
Private Sub cmdLogin_Click()
On Error GoTo cmdLogin_Click_Err

    On Error GoTo 0
    If (IsNull(lstNames)) Then
        Beep
        MsgBox "Select an employee.", vbOKOnly, ""
    ElseIf (IsNull(txtPassword)) Then
        Beep
        MsgBox "Enter a password.", vbOKOnly, ""
    Else
        EmployeeID = lstNames
    End If

    If (EmployeePassword = txtPassword) Then
        ValidLogin = True
        DoCmd.Close , ""
    Else
        Beep
        MsgBox "Incorrect password.", vbOKOnly, ""
    End If

cmdLogin_Click_Exit:
    Exit Sub

cmdLogin_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdLogin_Click_Exit

End Sub


'------------------------------------------------------------
' txtPassword_AfterUpdate
' Simulate click
'------------------------------------------------------------
Private Sub txtPassword_AfterUpdate()
    cmdLogin_Click
End Sub
