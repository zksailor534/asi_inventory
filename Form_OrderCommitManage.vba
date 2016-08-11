Option Compare Database
Option Explicit

Private rstCommit As DAO.Recordset

'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    On Error GoTo 0

    ' Check for valid commit ID
    If CurrentCommitID <= 0 Then
        MsgBox "Invalid Commit ID", vbOKOnly, "Invalid Commit"
        DoCmd.Close
        GoTo Form_Load_Exit
    End If

    ' Check for valid user
    If (EmployeeLogin = "") Then
        MsgBox "User is invalid:" & vbCrLf & "Login as valid user", vbOKOnly, "Invalid User"
        DoCmd.Close
        GoTo Form_Load_Exit
    End If

    SetVisibility

    ' Open recordset
    open_db
    Set rstCommit = db.OpenRecordset("SELECT * FROM " & CommitQuery & " WHERE [CommitID] = " & CurrentCommitID)

    ' Check for valid commit status
    If (rstCommit!Status <> "A") And (rstCommit!Status <> "X") And (rstCommit!Status <> "C") Then
        MsgBox "Commit status is invalid:" & vbCrLf & "Commit Status must be 'A', 'X', or 'C'" _
            , vbOKOnly, "Invalid Commit Status"
        DoCmd.Close
        GoTo Form_Load_Exit
    End If

    ' By default, quantity adjust check is false (not checked)
    QtyAdjustCheck.Value = False
    NewQuantity.Enabled = False
    NewQuantity.Locked = True
    Utilities.FieldAvailableRemove Me.Controls("NewQuantity")

    FillFields

    SetTitleStatus

Form_Load_Exit:
    Exit Sub

Form_Load_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume Form_Load_Exit

End Sub


'------------------------------------------------------------
' Form_Close
'
'------------------------------------------------------------
Private Sub Form_Close()
    ' Clean Up
    rstCommit.Close
    Set rstCommit = Nothing
End Sub


'------------------------------------------------------------
' cmdSave_Click
'
'------------------------------------------------------------
Private Sub cmdSave_Click()
On Error GoTo cmdSave_Click_Err

    ' -------------------------------------------------------------------
    ' Error Handling
    ' -------------------------------------------------------------------
    On Error GoTo 0

    ' Check if anything has changed
    If (CLng(QtyCommitted) = rstCommit!QtyCommitted _
        And SalesOrder = rstCommit!Reference) Then
        GoTo cmdSave_NoChange_Err
    End If

    ' Check for valid user
    If (EmployeeLogin = "") Then
        GoTo cmdSave_User_Err
    ElseIf ((EmployeeRole = SalesLevel) And (EmployeeLogin <> CommitUser)) Then
        GoTo cmdSave_User_Err
    End If

    ' Check for valid Sales Order
    If (SalesOrder = "") Then
        GoTo cmdSave_SalesOrder_Err
    End If

    ' Check for valid Commit Quantity
    ' Must be numeric, positive, and less than or equal to available
    If Not (IsNumeric(QtyCommitted = "")) Then
        GoTo cmdSave_QtyCommitted_Err
    ElseIf CLng(QtyCommitted) <= 0 Then
        GoTo cmdSave_QtyCommitted_Err
    ElseIf CLng(QtyCommitted) > (rstCommit!OnHand + rstCommit!OnOrder - _
        rstCommit!Committed + rstCommit!QtyCommitted - CLng(QtyCommitted)) Then
        GoTo cmdSave_QtyCommitted_Err
    ElseIf (QtyAdjustCheck.Value = True) Then
        If CLng(QtyCommitted) > (NewQuantity + rstCommit!OnOrder - _
            rstCommit!Committed + rstCommit!QtyCommitted - CLng(QtyCommitted)) Then
            GoTo cmdSave_QtyCommitted_Err
        End If
    End If

    ' Adjust quantity if called for
    If (QtyAdjustCheck.Value = True) Then
        If (CLng(NewQuantity) >= (Committed - rstCommit!QtyCommitted + QtyCommitted)) Then
            Utilities.OperationEntry rstCommit!ID, "Inventory", _
                "Changed OnHand from " & rstCommit!OnHand & " to " & NewQuantity & _
                    " in Commit " & CurrentCommitID
        Else
            GoTo cmdSave_QtyAdjust_Err
        End If
    End If

    ' Update Sales order and Commit Quantity fields
    With rstCommit
        .Edit
        !DateActive = Now()
        !OperatorActive = EmployeeLogin
        !Reference = SalesOrder
        !Committed = Committed - !QtyCommitted + QtyCommitted
        !QtyCommitted = QtyCommitted
        If (NewQuantity <> OnHand) Then
            !OnHand = NewQuantity
            !LastOper = EmployeeLogin
            !LastDate = Now()
        End If
        .Update
    End With

    GoTo cmdSave_Success

cmdSave_Click_Exit:
    DoCmd.Close
    Exit Sub

cmdSave_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdSave_Click_Exit

cmdSave_Success:
    MsgBox "Successful Save:" & vbCrLf & _
        "Commit " & CurrentCommitID & " for Sales Order " & SalesOrder, vbOKOnly, "Success"
    GoTo cmdSave_Click_Exit

cmdSave_NoChange_Err:
    MsgBox "Nothing Has Changed", vbOKOnly, "Unable to Save"
    Exit Sub

cmdSave_User_Err:
    MsgBox "User is invalid:" & vbCrLf & "Login as valid user", vbOKOnly, "Unable to Save"
    Exit Sub

cmdSave_QtyAdjust_Err:
    MsgBox "Error: Unable is Adjust Quantity" & vbCrLf & "New quantity must be >= 0" _
        , vbOKOnly, "Unable to Adjust Quantity"
    Exit Sub

cmdSave_QtyCommitted_Err:
    MsgBox "Commit Quantity is invalid", vbOKOnly, "Unable to Save"
    Exit Sub

cmdSave_SalesOrder_Err:
    MsgBox "Sales Order is invalid", vbOKOnly, "Unable to Save"
    Exit Sub

End Sub


'------------------------------------------------------------
' cmdDelete_Click
'
'------------------------------------------------------------
Private Sub cmdDelete_Click()
On Error GoTo cmdDelete_Click_Err

    Dim success As Boolean

    ' Check for valid user
    If (EmployeeLogin = "") Then
        GoTo cmdDelete_User_Err
    ElseIf (EmployeeRole = SalesLevel) And (rstCommit!OperatorActive <> EmployeeLogin) Then
        GoTo cmdDelete_User_Err
    End If

    ' Check if on hand quantity adjustment is valid
    If (QtyAdjustCheck.Value = True) Then
        If (CLng(NewQuantity) >= (rstCommit!Committed - rstCommit!QtyCommitted)) Then
            Utilities.OperationEntry rstCommit!ID, "Inventory", _
                "Changed OnHand from " & rstCommit!OnHand & " to " & NewQuantity & _
                    " after Commit " & CurrentCommitID & " Deletion"
        Else
            GoTo cmdDelete_QtyAdjust_Err
        End If
    End If

    ' Delete commit
    success = Utilities.Commit_Cancel(rstCommit)
    If Not (success) Then
        GoTo cmdDelete_Delete_Err
    End If

    ' Adjust quantity if called for
    If (QtyAdjustCheck.Value = True) Then
        With rstCommit
            .Edit
            !OnHand = NewQuantity
            !LastOper = EmployeeLogin
            !LastDate = Now()
            .Update
        End With
    End If

    GoTo cmdDelete_Success

cmdDelete_Click_Exit:
    DoCmd.Close
    Exit Sub

cmdDelete_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdDelete_Click_Exit

cmdDelete_Success:
    MsgBox "Successful Removal:" & vbCrLf & _
        "Commit " & CurrentCommitID & " for Sales Order " & SalesOrder, vbOKOnly, "Success"
    GoTo cmdDelete_Click_Exit

cmdDelete_User_Err:
    MsgBox "Error: User is invalid" & vbCrLf & "Login as valid user", vbOKOnly, "Unable to Delete"
    Exit Sub

cmdDelete_QtyAdjust_Err:
    MsgBox "Error: Unable is Adjust Quantity" & vbCrLf & "New quantity must be greater than Committed Quantity" _
        , vbOKOnly, "Unable to Adjust Quantity"
    Exit Sub

cmdDelete_Delete_Err:
    MsgBox "Error: Unable to Delete", vbOKOnly, "Unable to Delete"
    Exit Sub

End Sub


'------------------------------------------------------------
' cmdComplete_Click
'
'------------------------------------------------------------
Private Sub cmdComplete_Click()
On Error GoTo cmdComplete_Click_Err

    Dim success As Boolean

    ' Check for valid user
    If (EmployeeLogin = "") Or (EmployeeRole = SalesLevel) Then
        GoTo cmdComplete_User_Err
    End If

    ' Check if on hand quantity adjustment is valid
    If (QtyAdjustCheck.Value = True) Then
        If (CLng(NewQuantity) >= (rstCommit!Committed - rstCommit!QtyCommitted)) Then
            Utilities.OperationEntry rstCommit!ID, "Inventory", _
                "Changed OnHand from " & (rstCommit!OnHand - rstCommit!QtyCommitted) & " to " & NewQuantity & _
                    " after Commit " & CurrentCommitID & " Completion"
        Else
            GoTo cmdComplete_QtyAdjust_Err
        End If
    End If

    ' Complete commit
    success = Utilities.Commit_Complete(rstCommit)
    If Not (success) Then
        GoTo cmdComplete_Complete_Err
    End If

    ' Adjust quantity if called for
    If (QtyAdjustCheck.Value = True) Then
        With rstCommit
            .Edit
            !OnHand = NewQuantity
            !LastOper = EmployeeLogin
            !LastDate = Now()
            .Update
        End With
    End If

    GoTo cmdComplete_Success

cmdComplete_Click_Exit:
    DoCmd.Close
    Exit Sub

cmdComplete_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdComplete_Click_Exit

cmdComplete_Success:
    MsgBox "Successful Completion:" & vbCrLf & _
        "Commit " & CurrentCommitID & " for Sales Order " & SalesOrder, vbOKOnly, "Success"
    GoTo cmdComplete_Click_Exit

cmdComplete_User_Err:
    MsgBox "Error: User is invalid" & vbCrLf & "Login as valid user", vbOKOnly, "Unable to Save"
    Exit Sub

cmdComplete_QtyAdjust_Err:
    MsgBox "Error: Unable is Adjust Quantity" & vbCrLf & "New quantity must be greater than Committed Quantity" _
        , vbOKOnly, "Unable to Adjust Quantity"
    Exit Sub

cmdComplete_Complete_Err:
    MsgBox "Error: Unable to Complete", vbOKOnly, "Unable to Complete"
    Exit Sub

End Sub


'------------------------------------------------------------
' cmdClose_Click
'
'------------------------------------------------------------
Private Sub cmdClose_Click()
    DoCmd.Close
End Sub


'------------------------------------------------------------
' QtyAdjustCheck_Click
'
'------------------------------------------------------------
Private Sub QtyAdjustCheck_Click()
    If (QtyAdjustCheck.Value = True) Then
        NewQuantity.Enabled = True
        NewQuantity.Locked = False
        Utilities.FieldAvailableSet Me.Controls("NewQuantity")
    Else
        NewQuantity.Enabled = False
        NewQuantity.Locked = True
        Utilities.FieldAvailableRemove Me.Controls("NewQuantity")
    End If
End Sub


'------------------------------------------------------------
' FillFields
'
'------------------------------------------------------------
Private Sub FillFields()
    Status = Nz(rstCommit!Status, "")
    txtSalesOrderTitle = Nz(rstCommit!Reference, "")
    SalesOrder = Nz(rstCommit!Reference, "")
    QtyCommitted = Nz(rstCommit!QtyCommitted, "")
    CommitUser = Nz(rstCommit!OperatorActive, "")
    CommitDate = Nz(rstCommit!DateActive, "")
    Product = Nz(rstCommit!Product, "")
    RecordID = Nz(rstCommit!RecordID, "")
    Category = Nz(rstCommit!Category, "")
    Manufacturer = Nz(rstCommit!Manufacturer, "")
    Style = Nz(rstCommit!Style, "")
    Color = Nz(rstCommit!Color, "")
    Condition = Nz(rstCommit!Condition, "")
    Vendor = Nz(rstCommit!Vendor, "")
    Description = Nz(rstCommit!Description, "")
    Location = Nz(rstCommit!Location, "")
    OnHand = Nz(rstCommit!OnHand, "")
    Committed = Nz(rstCommit!Committed, "")
    NewQuantity = Nz(rstCommit!OnHand, "")

    ' Fill in Last Change Date and user based on status
    If (Status = "A") Then
        LastUser = ""
        LastDate = ""
    ElseIf (Status = "X") Then
        LastUser = Nz(rstCommit!OperatorCancel, "")
        LastDate = Nz(rstCommit!DateCancel, "")
    ElseIf (Status = "P") Then
        LastUser = Nz(rstCommit!OperatorPicked, "")
        LastDate = Nz(rstCommit!DatePicked, "")
    ElseIf (Status = "C") Then
        LastUser = Nz(rstCommit!OperatorComplete, "")
        LastDate = Nz(rstCommit!DateComplete, "")
    End If

End Sub


'------------------------------------------------------------
' SetTitleStatus
'
'------------------------------------------------------------
Private Sub SetTitleStatus()
    If Me.Status = "A" Then
        Me.lblStatusTitle.Caption = "Active"
    ElseIf Me.Status = "X" Then
        Me.lblStatusTitle.Caption = "Cancelled"
    ElseIf Me.Status = "P" Then
        Me.lblStatusTitle.Caption = "Picked"
    ElseIf Me.Status = "C" Then
        Me.lblStatusTitle.Caption = "Completed"
    Else
        Me.lblStatusTitle.Caption = ""
    End If
End Sub


'------------------------------------------------------------
' SetVisibility
'
'------------------------------------------------------------
Private Sub SetVisibility()
    ' Employee Role settings
    If (EmployeeRole = SalesLevel) Then
        cmdComplete.Enabled = False
        cmdComplete.Visible = False
        cmdDelete.Enabled = True
        cmdDelete.Visible = True
        lblQtyAdjust.Visible = False
        QtyAdjustCheck.Visible = False
        QtyAdjustCheck.Enabled = False
        lblNewQuantity.Visible = False
        NewQuantity.Visible = False
        NewQuantity.Enabled = False
    ElseIf (EmployeeRole = ProdLevel) Then
        cmdComplete.Enabled = True
        cmdComplete.Visible = True
        cmdDelete.Enabled = True
        cmdDelete.Visible = True
        lblQtyAdjust.Visible = True
        QtyAdjustCheck.Visible = True
        QtyAdjustCheck.Enabled = True
        lblNewQuantity.Visible = True
        NewQuantity.Visible = True
        NewQuantity.Enabled = False
    Else
        cmdComplete.Enabled = True
        cmdComplete.Visible = True
        cmdDelete.Enabled = True
        cmdDelete.Visible = True
        lblQtyAdjust.Visible = True
        QtyAdjustCheck.Visible = True
        QtyAdjustCheck.Enabled = True
        lblNewQuantity.Visible = True
        NewQuantity.Visible = True
        NewQuantity.Enabled = False
    End If

    ' Order status settings
    If Me.Status <> "A" Then
        lblQtyAdjust.Visible = False
        QtyAdjustCheck.Visible = False
        QtyAdjustCheck.Enabled = False
        lblNewQuantity.Visible = False
        NewQuantity.Visible = False
        NewQuantity.Enabled = False
    End If
End Sub
