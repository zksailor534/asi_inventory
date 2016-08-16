Option Compare Database
Option Explicit


'------------------------------------------------------------
' Form_Open
'
'------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

    ' Maximize the form
    DoCmd.Maximize

    ' Set screen view properties
    sbfrmOrderSearch.Form.DatasheetFontHeight = 10
    SetScreenSize

    ' By default display all commits and status
    If (Len(commitSelectStatus) > 0) Then
        StatusSelect = commitSelectStatus
    Else
        StatusSelect = "Active"
    End If
    Me.sbfrmOrderSearch.Form.Status.ColumnHidden = False

    ' Engage filter from Committed status
    SalesOrderFiltered = ""
    CurrentSalesOrder = ""
    Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered, StatusSelect)
    Me.sbfrmOrderSearch.Form.OrderBy = "DateActive DESC"
    Me.sbfrmOrderSearch.Form.FilterOn = True

    ' Set button status according to commit view
    ButtonStatus EmployeeRole, StatusSelect

outNow:
    SalesOrderFiltered.SetFocus

End Sub

'------------------------------------------------------------
' SalesOrderFiltered_AfterUpdate
'
'------------------------------------------------------------
Private Sub SalesOrderFiltered_AfterUpdate()
    If (SalesOrderFiltered <> "") Then
        OrderFilterButton_Click
    End If
End Sub


'------------------------------------------------------------
' OrderFilterButton_Click
'
'------------------------------------------------------------
Private Sub OrderFilterButton_Click()
    If IsNull(SalesOrderFiltered) Or (SalesOrderFiltered = "") Then
        SalesOrderFiltered.SetFocus
    Else
        If ValidSalesOrder(SalesOrderFiltered) Then
            ' Engage filter from sales order selection
            Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered, StatusSelect)
            Me.sbfrmOrderSearch.Form.OrderBy = "DateActive DESC"
            Me.sbfrmOrderSearch.Form.FilterOn = True
            CurrentSalesOrder = SalesOrderFiltered
        Else
            SalesOrderFiltered = ""
            SalesOrderFiltered.SetFocus
        End If
    End If
End Sub


'------------------------------------------------------------
' ClearFilterButton_Click
'
'------------------------------------------------------------
Private Sub ClearFilterButton_Click()
    SalesOrderFiltered = ""
    CurrentSalesOrder = ""
    Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered, StatusSelect)
    Me.sbfrmOrderSearch.Form.OrderBy = "DateActive DESC"
    Me.sbfrmOrderSearch.Form.FilterOn = True
    Me.sbfrmOrderSearch.Form.Requery
End Sub


'------------------------------------------------------------
' StatusSelect_AfterUpdate
'
'------------------------------------------------------------
Private Sub StatusSelect_AfterUpdate()
    commitSelectStatus = StatusSelect
    Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered, StatusSelect)
    Me.sbfrmOrderSearch.Form.FilterOn = True
    Me.sbfrmOrderSearch.Form.Requery
    ButtonStatus EmployeeRole, StatusSelect
End Sub


'------------------------------------------------------------
' ManageCommitButton_Click
'
'------------------------------------------------------------
Private Sub ManageCommitButton_Click()

    If IsNull(CurrentCommitID) Or (CurrentCommitID = 0) Then
        MsgBox "No commitment selected:" & vbCrLf & "Please select commitment to edit", , "Invalid Commit"
        Exit Sub
    Else
        If (EmployeeRole = SalesLevel) And EmployeeLogin <> SalesOrderUser(CurrentSalesOrder) Then
            MsgBox "Invalid User:" & vbCrLf & "Unable to edit Commit of other user", , "Invalid User"
            Exit Sub
        End If

        DoCmd.OpenForm OrderCommitManageForm, , , , , acDialog
        Me.sbfrmOrderSearch.Form.Requery
    End If

End Sub


'------------------------------------------------------------
' CancelCommitButton_Click
'
'------------------------------------------------------------
Private Sub CancelCommitButton_Click()
    Dim rstCommit As DAO.Recordset
    Dim decommitAll As Integer

    If (EmployeeRole = SalesLevel) Then
        MsgBox "Invalid User:" & vbCrLf & "Not authorized to Cancel Commitment", , "Invalid User"
        Exit Sub
    End If

    open_db
    Set rstCommit = db.OpenRecordset("SELECT * FROM " & CommitQuery & _
        " WHERE " & sbfrmOrderSearch.Form.Filter)

    If (Utilities.RecordCheck(rstCommit, "Reference", CurrentSalesOrder)) And _
        (Utilities.RecordCheck(rstCommit, "Status", "A")) Then

        decommitAll = MsgBox("Do you want to cancel ALL commitments in current view?", _
            vbYesNo, "Cancel Items")

        If decommitAll = vbNo Then
           GoTo CancelCommitButton_Click_Exit
        ElseIf decommitAll = vbYes Then
            Call Utilities.Commit_Cancel(rstCommit)

            Utilities.OperationEntry rstCommit!ID, "Commit", _
                "Cancelled Commitments from Sales Order " & CurrentSalesOrder

            Me.sbfrmOrderSearch.Form.Requery
        End If
    Else
        MsgBox "Invalid Selection:" & vbCrLf & _
            " - Select single sales order" & vbCrLf & _
            " - Status must be Active", , "Invalid Sales Order"
        GoTo CancelCommitButton_Click_Exit
    End If

CancelCommitButton_Click_Exit:
    rstCommit.Close
    Set rstCommit = Nothing

End Sub


'------------------------------------------------------------
' ReactivateCommitButton_Click
'
'------------------------------------------------------------
Private Sub ReactivateCommitButton_Click()
    Dim rstCommit As DAO.Recordset
    Dim reactivateAll As Integer

    If (EmployeeRole = SalesLevel) Then
        MsgBox "Invalid User:" & vbCrLf & "Not authorized to Reactivate Commitment", , "Invalid User"
        Exit Sub
    End If

    open_db
    Set rstCommit = db.OpenRecordset("SELECT * FROM " & CommitQuery & _
        " WHERE " & sbfrmOrderSearch.Form.Filter)

    If (Utilities.RecordCheck(rstCommit, "Reference", CurrentSalesOrder)) And _
        Not (Utilities.RecordCheck(rstCommit, "Status", "A")) Then

        reactivateAll = MsgBox("Do you want to reactivate ALL commitments in current view?", _
            vbYesNo, "Reactivate Items")

        If reactivateAll = vbNo Then
           GoTo ReactivateCommitButton_Click_Exit
        ElseIf reactivateAll = vbYes Then
            Call Utilities.Commit_Reactivate(rstCommit)

            Utilities.OperationEntry rstCommit!ID, "Commit", _
                "Reactivated Commitments from Sales Order " & CurrentSalesOrder

            Me.sbfrmOrderSearch.Form.Requery
        End If
    Else
        MsgBox "Invalid Selection:" & vbCrLf & _
            " - Select single sales order" & vbCrLf & _
            " - Status cannot be Active", , "Invalid Sales Order"
        GoTo ReactivateCommitButton_Click_Exit
    End If

ReactivateCommitButton_Click_Exit:
    rstCommit.Close
    Set rstCommit = Nothing

End Sub


'------------------------------------------------------------
' CompleteCommitButton_Click
'
'------------------------------------------------------------
Private Sub CompleteCommitButton_Click()
    Dim rstCommit As DAO.Recordset
    Dim completeAll As Integer

    If (EmployeeRole = SalesLevel) Then
        MsgBox "Invalid User:" & vbCrLf & "Not authorized to Complete Commitment", , "Invalid User"
        Exit Sub
    End If

    open_db
    Set rstCommit = db.OpenRecordset("SELECT * FROM " & CommitQuery & _
        " WHERE " & sbfrmOrderSearch.Form.Filter)

    If (Utilities.RecordCheck(rstCommit, "Reference", CurrentSalesOrder)) And _
        (Utilities.RecordCheck(rstCommit, "Status", "A")) Then

        completeAll = MsgBox("Do you want to complete ALL commitments in current view?", _
            vbYesNo, "Complete Items")

        If completeAll = vbNo Then
           GoTo CompleteCommitButton_Click_Exit
        ElseIf completeAll = vbYes Then
            Call Utilities.Commit_Complete(rstCommit)

            Utilities.OperationEntry rstCommit!ID, "Commit", _
                "Completed Commitments from Sales Order " & CurrentSalesOrder

            Me.sbfrmOrderSearch.Form.Requery
        End If
    Else
        MsgBox "Invalid Selection:" & vbCrLf & _
            " - Select single sales order" & vbCrLf & _
            " - Status must be Active", , "Invalid Sales Order"
        GoTo CompleteCommitButton_Click_Exit
    End If

CompleteCommitButton_Click_Exit:
    rstCommit.Close
    Set rstCommit = Nothing

End Sub


'------------------------------------------------------------
' SetScreenSize
'
'------------------------------------------------------------
Public Sub SetScreenSize()
    Me.sbfrmOrderSearch.Left = 0
    Me.sbfrmOrderSearch.Top = 0
    Me.sbfrmOrderSearch.Width = ScreenWidth
    Me.sbfrmOrderSearch.Height = Round(Me.WindowHeight * 0.95)
End Sub


'------------------------------------------------------------
' ValidSalesOrder
'
'------------------------------------------------------------
Private Function ValidSalesOrder(OrderNumber As String) As Boolean
    On Error Resume Next
    If Not (IsNull(DLookup("Reference", CommitDB, "[Reference]='" & OrderNumber & "'"))) Then
        ValidSalesOrder = True
    Else
        MsgBox "Invalid Sales Order: Not Found", vbOKOnly, "Invalid Sales Order"
        ValidSalesOrder = False
    End If
End Function


'------------------------------------------------------------
' SalesOrderUser
'
'------------------------------------------------------------
Public Function SalesOrderUser(OrderNumber As String) As String
    Dim User As String
    On Error Resume Next
    User = DLookup("OperatorActive", CommitDB, "[Reference]='" & OrderNumber & "'")
    If Not (IsNull(User)) Then
        SalesOrderUser = User
    Else
        SalesOrderUser = ""
    End If
End Function


'------------------------------------------------------------
' GetFilter
'
'------------------------------------------------------------
Private Function GetFilter(Role As String, Order As String, Status As String) As String
    On Error Resume Next
    ' Active
    If (Status = "Active") Then
        GetFilter = "[Status]='A'"
    ' Complete
    ElseIf (Status = "Complete") Then
        GetFilter = "[Status]='C'"
    ' Cancelled
    ElseIf (Status = "Cancelled") Then
        GetFilter = "[Status]='X'"
    ' All
    ElseIf (Status = "All") Then
        GetFilter = "[Status] LIKE '*'"
    End If

    ' User Filter
    If (Role = SalesLevel) Then
        GetFilter = GetFilter & " AND [OperatorActive]='" & EmployeeLogin & "'"
    End If

    ' Order Filter
    If (Order <> "") Then
        GetFilter = GetFilter & " AND [Reference]= '" & Order & "'"
    End If
End Function


'------------------------------------------------------------
' ButtonStatus
'
'------------------------------------------------------------
Private Sub ButtonStatus(Role As String, Status As String)
    On Error Resume Next

    ' Active
    If (Status = "Active") Then
        CancelCommitButton.Enabled = True
        ReactivateCommitButton.Enabled = False
        CompleteCommitButton.Enabled = True
    ' Complete
    ElseIf (Status = "Complete") Then
        CancelCommitButton.Enabled = False
        ReactivateCommitButton.Enabled = True
        CompleteCommitButton.Enabled = False
    ' Cancelled
    ElseIf (Status = "Cancelled") Then
        CancelCommitButton.Enabled = False
        ReactivateCommitButton.Enabled = True
        CompleteCommitButton.Enabled = False
    ' All
    ElseIf (Status = "All") Then
        CancelCommitButton.Enabled = True
        ReactivateCommitButton.Enabled = True
        CompleteCommitButton.Enabled = True
    End If

    ' Sales users cannot complete
    If (Role = SalesLevel) Then
        CompleteCommitButton.Enabled = False
    End If

End Sub
