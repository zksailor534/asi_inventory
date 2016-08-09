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
    Me.sbfrmOrderSearch.Form.Controls("Status").ColumnHidden = True

    ' Engage filter from Committed status
    SalesOrderFiltered = ""
    CurrentSalesOrder = ""
    Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered)
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
    Me.sbfrmOrderSearch.Form.Filter = GetFilter(EmployeeRole, SalesOrderFiltered)
    Me.sbfrmOrderSearch.Form.FilterOn = True

    If (ActiveToggle.Value = True) Then
        Me.sbfrmOrderSearch.Form.Controls("Status").ColumnHidden = True
    Else
        Me.sbfrmOrderSearch.Form.Controls("Status").ColumnHidden = False
    End If

    Me.sbfrmOrderSearch.Form.Requery
End Sub


'------------------------------------------------------------
' ManageCommitButton_Click
'
'------------------------------------------------------------
Private Sub ManageCommitButton_Click()

    If IsNull(CurrentCommitID) Or (CurrentCommitID = 0) Then
        MsgBox "No commitment selected:" & vbCrLf & "Please select commitment to edit", , "Invalid Commit"
        Debug.Print "CurrentCommitID", CurrentCommitID
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
' DeleteCommitButton_Click
'
'------------------------------------------------------------
Private Sub DeleteCommitButton_Click()
    Dim rstCommit As DAO.Recordset
    Dim decommitAll As Integer

    If (EmployeeRole = SalesLevel) Then
        If EmployeeLogin <> SalesOrderUser(CurrentSalesOrder) Then
            MsgBox "Invalid User:" & vbCrLf & "Unable to edit Commit of other user", , "Invalid User"
            Exit Sub
        End If
    End If

    If CurrentSalesOrder <> "" Then
        decommitAll = MsgBox("Do you want to remove commitments for ALL items in Sales Order " & CurrentSalesOrder _
            & "?", vbYesNo, "Decommit All Items")

        If decommitAll = vbNo Then
           Exit Sub
        ElseIf decommitAll = vbYes Then
            open_db
            Set rstCommit = db.OpenRecordset("SELECT * FROM " & CommitQuery & _
                " WHERE [Reference] = '" & CurrentSalesOrder & "' AND [Status] = 'A'")
            Call Utilities.Commit_Cancel(rstCommit)
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
' CompleteCommitButton_Click
'
'------------------------------------------------------------
Private Sub CompleteCommitButton_Click()
    Dim rstCommit As DAO.Recordset
    Dim completeAll As Integer

    If (EmployeeRole = SalesLevel) Then
        MsgBox "Invalid User:" & vbCrLf & "Unable to complete Commit", , "Invalid User"
        Exit Sub
    End If

    If CurrentSalesOrder <> "" Then
        completeAll = MsgBox("Do you want to complete commitments for ALL items in Sales Order " & CurrentSalesOrder _
            & "?", vbYesNo, "Complete All Items")

        If completeAll = vbNo Then
           Exit Sub
        ElseIf completeAll = vbYes Then
            open_db
            Set rstCommit = db.OpenRecordset("SELECT * FROM " & CommitQuery & _
                " WHERE [Reference] = '" & CurrentSalesOrder & "' AND [Status] = 'A'")
            Call Utilities.Commit_Complete(rstCommit)
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
Private Function GetFilter(Role As String, Order As String) As String
    On Error Resume Next
    ' Sales Role, Active, No Order
    If (Role = SalesLevel) And (ActiveToggle.Value = True) And (Order = "") Then
        GetFilter = "[Status]='A' AND [OperatorActive]='" & EmployeeLogin & "'"
    ' Any Role, Active, No Order
    ElseIf (Role <> SalesLevel) And (ActiveToggle.Value = True) And (Order = "") Then
        GetFilter = "[Status]='A'"
    ' Sales Role, Not Active, No Order
    ElseIf (Role = SalesLevel) And (ActiveToggle.Value = False) And (Order = "") Then
        GetFilter = "[OperatorActive]='" & EmployeeLogin & "'"
    ' Any Role, Not Active, No Order
    ElseIf (Role <> SalesLevel) And (ActiveToggle.Value = False) And (Order = "") Then
        GetFilter = "[Status] IS NOT NULL"
    ' Sales Role, Active, Order
    ElseIf (Role = SalesLevel) And (ActiveToggle.Value = True) And (Order <> "") Then
        GetFilter = "[Status]='A' AND [OperatorActive]='" & EmployeeLogin & "' AND [Reference]= '" & Order & "'"
    ' Any Role, Active, Order
    ElseIf (Role <> SalesLevel) And (ActiveToggle.Value = True) And (Order <> "") Then
        GetFilter = "[Status]='A' AND [Reference]= '" & Order & "'"
    ' Sales Role, Not Active, Order
    ElseIf (Role = SalesLevel) And (ActiveToggle.Value = False) And (Order <> "") Then
        GetFilter = "[OperatorActive]='" & EmployeeLogin & "' AND [Reference]= '" & Order & "'"
    ' Any Role, Not Active, Order
    ElseIf (Role <> SalesLevel) And (ActiveToggle.Value = False) And (Order <> "") Then
        GetFilter = "[Reference]= '" & Order & "'"
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
        DeleteCommitButton.Visible = True
        DeleteCommitButton.Enabled = True
        CompleteCommitButton.Visible = False
        CompleteCommitButton.Enabled = False
    ElseIf (Role = ProdLevel) Then
        ' Hide Delete Commit button for production
        DeleteCommitButton.Visible = False
        DeleteCommitButton.Enabled = False
        CompleteCommitButton.Visible = True
        CompleteCommitButton.Enabled = True
    Else
        DeleteCommitButton.Visible = True
        DeleteCommitButton.Enabled = True
        CompleteCommitButton.Visible = True
        CompleteCommitButton.Enabled = True
    End If
End Sub
