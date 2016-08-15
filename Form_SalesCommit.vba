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

    ' Default is only active commits
    ActiveToggle.Value = True
    ButtonStatus
    Me.sbfrmOrderSearch.Form.Controls("Status").ColumnHidden = True

    ' Engage filter from Committed status
    SalesOrderFiltered = ""
    CurrentSalesOrder = ""
    Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered)
    Me.sbfrmOrderSearch.Form.OrderBy = "DateActive DESC"
    Me.sbfrmOrderSearch.Form.FilterOn = True

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
            Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered)
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
    Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered)
    Me.sbfrmOrderSearch.Form.OrderBy = "DateActive DESC"
    Me.sbfrmOrderSearch.Form.FilterOn = True
    Me.sbfrmOrderSearch.Form.Requery
End Sub


'------------------------------------------------------------
' ActiveToggle_Click
'
'------------------------------------------------------------
Private Sub ActiveToggle_Click()
    Me.sbfrmOrderSearch.Form.FilterOn = False

    If (ActiveToggle.Value = True) Then
        Me.sbfrmOrderSearch.Form.Controls("Status").ColumnHidden = True
        ButtonStatus
    Else
        Me.sbfrmOrderSearch.Form.Controls("Status").ColumnHidden = False
        ButtonStatus
    End If

    Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered)
    Me.sbfrmOrderSearch.Form.FilterOn = True
    Me.sbfrmOrderSearch.Form.Requery
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
        If (EmployeeRole = SalesLevel) And (EmployeeLogin <> SalesOrderUser(CurrentSalesOrder)) Then
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

    If (EmployeeRole = SalesLevel) And (EmployeeLogin <> SalesOrderUser(CurrentSalesOrder)) Then
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
' SetScreenSize
'
'------------------------------------------------------------
Private Sub SetScreenSize()
    On Error Resume Next
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
Private Function GetFilter(Role As String, Order As String) As String
    On Error Resume Next
    ' Active, No Order
    If (Role = SalesLevel) And (ActiveToggle.Value = True) And (Order = "") Then
        GetFilter = "[Status]='A' AND [OperatorActive]='" & EmployeeLogin & "'"
    ' Not Active, No Order
    ElseIf (Role = SalesLevel) And (ActiveToggle.Value = False) And (Order = "") Then
        GetFilter = "[OperatorActive]='" & EmployeeLogin & "'"
    ' Active, Order
    ElseIf (Role = SalesLevel) And (ActiveToggle.Value = True) And (Order <> "") Then
        GetFilter = "[Status]='A' AND [OperatorActive]='" & EmployeeLogin & "' AND [Reference]= '" & Order & "'"
    ' Not Active, Order
    ElseIf (Role = SalesLevel) And (ActiveToggle.Value = False) And (Order <> "") Then
        GetFilter = "[OperatorActive]='" & EmployeeLogin & "' AND [Reference]= '" & Order & "'"
    ' Manager Roles, Active, No Order
    ElseIf (Role <> SalesLevel) And (ActiveToggle.Value = True) And (Order = "") Then
        GetFilter = "[Status]='A'"
    ' Manager Roles, Not Active, No Order
    ElseIf (Role <> SalesLevel) And (ActiveToggle.Value = False) And (Order = "") Then
        GetFilter = ""
    ' Manager Roles, Active, Order
    ElseIf (Role <> SalesLevel) And (ActiveToggle.Value = True) And (Order <> "") Then
        GetFilter = "[Status]='A' AND [Reference]= '" & Order & "'"
    ' Manager Roles, Not Active, Order
    ElseIf (Role <> SalesLevel) And (ActiveToggle.Value = False) And (Order <> "") Then
        GetFilter = "[Reference]= '" & Order & "'"
    End If
End Function


'------------------------------------------------------------
' ButtonStatus
'
'------------------------------------------------------------
Private Sub ButtonStatus()
    On Error Resume Next

    ' Active
    If (ActiveToggle.Value = True) Then
        CancelCommitButton.Enabled = True
    ' Complete
    Else
        CancelCommitButton.Enabled = False
    End If

End Sub
