Option Compare Database
Option Explicit


'------------------------------------------------------------
' RecalcCommitButton_Click
'
'------------------------------------------------------------
Private Sub Form_Load()
    If EmployeeRole <> DevelLevel Then
        CloseFormsButton.Enabled = False
        CloseFormsButton.Visible = False
    End If
End Sub


'------------------------------------------------------------
' RecalcCommitButton_Click
'
'------------------------------------------------------------
Private Sub RecalcCommitButton_Click()

    Utilities.RecalculateCommit

End Sub


'------------------------------------------------------------
' CloseFormsButton_Click
'
'------------------------------------------------------------
Private Sub CloseFormsButton_Click()

    If EmployeeRole = DevelLevel Then
        DoCmd.Close acForm, MainForm, acSaveNo
        DoCmd.SelectObject acTable, , True
    Else
        CloseFormsButton.Enabled = False
    End If

End Sub
