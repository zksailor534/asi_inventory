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
    StatusSelect = "All"
    Me.sbfrmOrderSearch.Form.Status.ColumnHidden = False

    ' Engage filter from Committed status
    SalesOrderFiltered = ""
    CurrentSalesOrder = ""
    Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered, StatusSelect)
    Me.sbfrmOrderSearch.Form.OrderBy = "DateActive DESC"
    Me.sbfrmOrderSearch.Form.FilterOn = True

    ' Set user role properties
    UserRoleProperties (EmployeeRole)

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
    Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered, StatusSelect)
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
        If EmployeeLogin <> SalesOrderUser(CurrentSalesOrder) Then
            MsgBox "Invalid User:" & vbCrLf & "Unable to edit Commit of other user", , "Invalid User"
            Exit Sub
        End If
    End If

    If CurrentSalesOrder <> "" Then
        decommitAll = MsgBox("Do you want to Cancel commitments for ALL items in Sales Order " & CurrentSalesOrder _
            & "?", vbYesNo, "Decommit All Items")

        If decommitAll = vbNo Then
           Exit Sub
        ElseIf decommitAll = vbYes Then
            open_db
            Set rstCommit = db.OpenRecordset("SELECT * FROM " & CommitQuery & _
                " WHERE [Reference] = '" & CurrentSalesOrder & "' AND [Status] IN ('A')")
            Call Utilities.Commit_Cancel(rstCommit)

            Utilities.OperationEntry rstCommit!ID, "Commit", _
                "Cancelled All Commitments from Sales Order " & CurrentSalesOrder

            rstCommit.Close
            Set rstCommit = Nothing
            Me.sbfrmOrderSearch.Form.Requery
        End If
    Else
        MsgBox "No Sales Order Selected", , "Invalid Sales Order"
        Exit Sub
    End If
    Me.sbfrmOrderSearch.Form.Requery
End Sub


'------------------------------------------------------------
' ReactivateCommitButton_Click
'
'------------------------------------------------------------
Private Sub ReactivateCommitButton_Click()
    Dim rstCommit As DAO.Recordset
    Dim reactivateAll As Integer

    If (EmployeeRole = SalesLevel) Then
        MsgBox "Invalid User:" & vbCrLf & "Not allowed to Reactivate Commitment", , "Invalid User"
        Exit Sub
    End If

    If CurrentSalesOrder <> "" Then
        reactivateAll = MsgBox("Do you want to reactivate ALL commitments in Sales Order " & CurrentSalesOrder _
            & "?", vbYesNo, "Reactivate All Items")

        If reactivateAll = vbNo Then
           Exit Sub
        ElseIf reactivateAll = vbYes Then
            open_db
            Set rstCommit = db.OpenRecordset("SELECT * FROM " & CommitQuery & _
                " WHERE [Reference] = '" & CurrentSalesOrder & "' AND [Status] IN ('C','X')")
            Call Utilities.Commit_Reactivate(rstCommit)

            Utilities.OperationEntry rstCommit!ID, "Commit", _
                "Reactivated All Commitments from Sales Order " & CurrentSalesOrder

            rstCommit.Close
            Set rstCommit = Nothing
            Me.sbfrmOrderSearch.Form.Requery
        End If
    Else
        MsgBox "No Sales Order Selected", , "Invalid Sales Order"
        Exit Sub
    End If
End Sub


'------------------------------------------------------------
' CompleteCommitButton_Click
'
'------------------------------------------------------------
Private Sub CompleteCommitButton_Click()
    Dim rstCommit As DAO.Recordset
    Dim completeAll As Integer

    If (EmployeeRole = SalesLevel) Then
        MsgBox "Invalid User:" & vbCrLf & "Unable to Complete Commit", , "Invalid User"
        Exit Sub
    End If

    If CurrentSalesOrder <> "" Then
        completeAll = MsgBox("Do you want to Complete commitments for ALL items in Sales Order " & CurrentSalesOrder _
            & "?", vbYesNo, "Complete All Items")

        If completeAll = vbNo Then
           Exit Sub
        ElseIf completeAll = vbYes Then
            open_db
            Set rstCommit = db.OpenRecordset("SELECT * FROM " & CommitQuery & _
                " WHERE [Reference] = '" & CurrentSalesOrder & "' AND [Status] = 'A'")
            Call Utilities.Commit_Complete(rstCommit)

            Utilities.OperationEntry rstCommit!ID, "Commit", _
                "Completed All Commitments from Sales Order " & CurrentSalesOrder

            rstCommit.Close
            Set rstCommit = Nothing
            Me.sbfrmOrderSearch.Form.Requery
        End If
    Else
        MsgBox "No Sales Order Selected", , "Invalid Sales Order"
        Exit Sub
    End If
End Sub


'------------------------------------------------------------
' Form_Resize
'
'------------------------------------------------------------
Private Sub Form_Resize()
    SetScreenSize
End Sub


'------------------------------------------------------------
' SetScreenSize
'
'------------------------------------------------------------
Private Sub SetScreenSize()
    On Error Resume Next
    Me.sbfrmOrderSearch.Left = 0
    Me.sbfrmOrderSearch.Top = 0
    Me.sbfrmOrderSearch.Width = Round(Me.WindowWidth)
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

    ' Order Filter
    If (Order <> "") Then
        GetFilter = GetFilter & " AND [Reference]= '" & Order & "'"
    End If
End Function


'------------------------------------------------------------
' UserRoleProperties
'
'------------------------------------------------------------
Private Sub UserRoleProperties(Role As String)
    On Error Resume Next

    If (Role = SalesLevel) Then
        ' Hide Complete Commit button for salespeople
        CancelCommitButton.Visible = True
        CancelCommitButton.Enabled = True
        ReactivateCommitButton.Visible = False
        ReactivateCommitButton.Enabled = False
        CompleteCommitButton.Visible = False
        CompleteCommitButton.Enabled = False
    ElseIf (Role = ProdLevel) Then
        CancelCommitButton.Visible = True
        CancelCommitButton.Enabled = True
        ReactivateCommitButton.Visible = True
        ReactivateCommitButton.Enabled = True
        CompleteCommitButton.Visible = True
        CompleteCommitButton.Enabled = True
    Else
        CancelCommitButton.Visible = True
        CancelCommitButton.Enabled = True
        ReactivateCommitButton.Visible = True
        ReactivateCommitButton.Enabled = True
        CompleteCommitButton.Visible = True
        CompleteCommitButton.Enabled = True
    End If
End Sub
