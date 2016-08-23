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

    ' Open recordsets
    open_db
    Set rstCommit = db.OpenRecordset("SELECT * FROM " & CommitQuery & " WHERE [CommitID] = " & CurrentCommitID)

    ' Check for valid commit status
    If (rstCommit!Status <> "A") And (rstCommit!Status <> "X") And (rstCommit!Status <> "P") And (rstCommit!Status <> "C") Then
        MsgBox "Commit status is invalid:" & vbCrLf & "Commit Status must be 'A', 'X', 'P', or 'C'" _
            , vbOKOnly, "Invalid Commit Status"
        DoCmd.Close
        GoTo Form_Load_Exit
    End If

    ' By default, location adjust check is false (not checked)
    LocAdjustCheck.Value = False
    Location.Enabled = False
    Location.Locked = True
    Utilities.FieldAvailableRemove Me.Controls("Location")

    ' By default, quantity adjust check is false (not checked)
    QtyAdjustCheck.Value = False
    OnHand.Enabled = False
    OnHand.Locked = True
    Utilities.FieldAvailableRemove Me.Controls("OnHand")

    ' Populate fields
    FillFields
    SetVisibility
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

    Dim saveCheck As Integer

    ' Check if anything has changed
    If (QtyCommitted = rstCommit!QtyCommitted _
        And SalesOrder = rstCommit!Reference) Then
        GoTo cmdSave_NoChange_Err
    End If

    ' Check for valid user
    If (EmployeeLogin = "") Then
        GoTo cmdSave_User_Err
    ElseIf ((EmployeeRole = SalesLevel) And (EmployeeLogin <> CommitUser)) Then
        GoTo cmdSave_User_Err
    End If

    ' Validate user command
    saveCheck = MsgBox("Do you want to Save changes to commitment " & CurrentCommitID & " in Sales Order " & CurrentSalesOrder _
        & "?", vbYesNo, "Save Commit")
    If saveCheck = vbNo Then
       Exit Sub
    End If

    ' Check for valid Sales Order
    If (SalesOrder = "") Then
        GoTo cmdSave_SalesOrder_Err
    End If

    ' Check for valid Commit Quantity
    ' Must be numeric, positive, and less than or equal to available
    If Not (IsNumeric(QtyCommitted)) Then
        GoTo cmdSave_QtyCommitted_Err
    ElseIf CLng(QtyCommitted) <= 0 Then
        GoTo cmdSave_QtyCommitted_Err
    ElseIf CLng(QtyCommitted) > (rstCommit!OnHand + rstCommit!OnOrder - _
        rstCommit!Committed + rstCommit!QtyCommitted) Then
        GoTo cmdSave_QtyCommitted_Err
    ElseIf (QtyAdjustCheck.Value = True) Then
        If CLng(QtyCommitted) > (CLng(OnHand) + rstCommit!OnOrder - _
            rstCommit!Committed + rstCommit!QtyCommitted - CLng(QtyCommitted)) Then
            GoTo cmdSave_QtyCommitted_Err
        End If
    End If

    ' Check for valid adjust quantity if called for
    If (QtyAdjustCheck.Value = True) Then
        If ((CLng(OnHand) + rstCommit!OnOrder) < (Committed - rstCommit!QtyCommitted + QtyCommitted)) Then
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

        ' Adjust Location if different
        If (Location <> !Location) Then
            Utilities.OperationEntry rstCommit!ID, "Inventory", _
                "Changed Location from " & rstCommit!Location & " to " & Location & _
                    " in Commit " & CurrentCommitID, "Move"
            !Location = Location
            !LastOper = EmployeeLogin
            !LastDate = Now()
        End If

        ' Adjust Quantity if different
        If (OnHand <> !OnHand) Then
            Utilities.OperationEntry rstCommit!ID, "Inventory", _
                "Changed OnHand from " & rstCommit!OnHand & " to " & OnHand & _
                    " in Commit " & CurrentCommitID, "Count"
            !OnHand = OnHand
            !LastOper = EmployeeLogin
            !LastDate = Now()
        End If

        .Update
    End With

    GoTo cmdSave_Success

cmdSave_Click_Exit:
    FillFields
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
' cmdCancel_Click
'
'------------------------------------------------------------
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_Click_Err

    Dim success As Boolean
    Dim cancelCheck As Integer

    ' Check for valid user
    If (EmployeeLogin = "") Then
        GoTo cmdCancel_User_Err
    ElseIf (EmployeeRole = SalesLevel) And (rstCommit!OperatorActive <> EmployeeLogin) Then
        GoTo cmdCancel_User_Err
    End If

    ' Validate user command
    cancelCheck = MsgBox("Do you want to Cancel commitment " & CurrentCommitID & " in Sales Order " & CurrentSalesOrder _
        & "?", vbYesNo, "Cancel Commit")
    If cancelCheck = vbNo Then
       Exit Sub
    End If

    ' Cancel commit
    success = Utilities.Commit_Cancel(rstCommit)
    If Not (success) Then
        GoTo cmdCancel_Err
    Else
        Utilities.OperationEntry rstCommit!ID, "Commit", _
            "Cancelled Commitment " & CurrentCommitID & " from Sales Order " & CurrentSalesOrder
    End If

    ' Adjust quantity if called for
    If (QtyAdjustCheck.Value = True) Or (LocAdjustCheck.Value = True) Then
        With rstCommit
            .Edit

            ' Adjust Location if different
            If (Location <> !Location) Then
                Utilities.OperationEntry rstCommit!ID, "Inventory", _
                    "Changed Location from " & rstCommit!Location & " to " & Location & _
                    " after Commit " & CurrentCommitID & " Cancellation", "Move"
                !Location = Location
                !LastOper = EmployeeLogin
                !LastDate = Now()
            End If

            ' Adjust Quantity if different
            If (OnHand <> !OnHand) Then
                Utilities.OperationEntry rstCommit!ID, "Inventory", _
                    "Changed OnHand from " & rstCommit!OnHand & " to " & OnHand & _
                    " after Commit " & CurrentCommitID & " Cancellation", "Count"
                !OnHand = OnHand
                !LastOper = EmployeeLogin
                !LastDate = Now()
            End If

            .Update
        End With
    End If

    GoTo cmdCancel_Success

cmdCancel_Click_Exit:
    DoCmd.Close
    Exit Sub

cmdCancel_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdCancel_Click_Exit

cmdCancel_Success:
    MsgBox "Successful Cancellation:" & vbCrLf & _
        "Commit " & CurrentCommitID & " for Sales Order " & SalesOrder, vbOKOnly, "Success"
    GoTo cmdCancel_Click_Exit

cmdCancel_User_Err:
    MsgBox "Error: User is invalid" & vbCrLf & "Login as valid user", vbOKOnly, "Unable to Cancel"
    Exit Sub

cmdCancel_Err:
    MsgBox "Error: Unable to Cancel Commitment", vbOKOnly, "Unable to Cancel"
    Exit Sub

End Sub


'------------------------------------------------------------
' cmdActive_Click
'
'------------------------------------------------------------
Private Sub cmdActive_Click()
On Error GoTo cmdActive_Click_Err

    Dim success As Boolean
    Dim activeCheck As Integer

    ' Check for valid user
    If (EmployeeLogin = "") Or (EmployeeRole = SalesLevel) Then
        GoTo cmdActive_User_Err
    End If

    ' Validate user command
    activeCheck = MsgBox("Do you want to Reactivate commitment " & CurrentCommitID & " in Sales Order " & CurrentSalesOrder _
        & "?", vbYesNo, "Reactivate Commit")
    If activeCheck = vbNo Then
       Exit Sub
    End If

    ' Reactivate committed item
    success = Utilities.Commit_Reactivate(rstCommit)
    If Not (success) Then
        GoTo cmdActive_Err
    Else
        Utilities.OperationEntry rstCommit!ID, "Commit", _
            "Reactivated Commitment " & CurrentCommitID & " from Sales Order " & CurrentSalesOrder
    End If

    GoTo cmdActive_Success

cmdActive_Click_Exit:
    DoCmd.Close
    Exit Sub

cmdActive_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdActive_Click_Exit

cmdActive_Success:
    MsgBox "Successfully Reactivated Item Commitment:" & vbCrLf & _
        "Commit " & CurrentCommitID & " for Sales Order " & SalesOrder, vbOKOnly, "Success"
    GoTo cmdActive_Click_Exit

cmdActive_User_Err:
    MsgBox "Error: User is invalid" & vbCrLf & "Login as valid user", vbOKOnly, "Unable to Save"
    Exit Sub

cmdActive_Err:
    MsgBox "Error: Unable to Reactivate Commitment", vbOKOnly, "Unable to Reactivate"
    Exit Sub

End Sub


'------------------------------------------------------------
' cmdComplete_Click
'
'------------------------------------------------------------
Private Sub cmdComplete_Click()
On Error GoTo cmdComplete_Click_Err

    Dim success As Boolean
    Dim completeCheck As Integer

    ' Check for valid user
    If (EmployeeLogin = "") Or (EmployeeRole = SalesLevel) Then
        GoTo cmdComplete_User_Err
    End If

    ' Validate user command
    completeCheck = MsgBox("Do you want to Complete commitment " & CurrentCommitID & " in Sales Order " & CurrentSalesOrder _
        & "?", vbYesNo, "Complete Commit")
    If completeCheck = vbNo Then
       Exit Sub
    End If

    ' Check if on hand quantity adjustment is valid
    If (QtyAdjustCheck.Value = True) Then
        If ((CLng(OnHand) + rstCommit!OnOrder) < (rstCommit!Committed - rstCommit!QtyCommitted)) Then
            GoTo cmdComplete_QtyAdjust_Err
        End If
    End If

    ' Complete commit
    success = Utilities.Commit_Complete(rstCommit)
    If Not (success) Then
        GoTo cmdComplete_Complete_Err
    Else
        Utilities.OperationEntry rstCommit!ID, "Commit", _
            "Completed Commitment " & CurrentCommitID & " from Sales Order " & CurrentSalesOrder
    End If

    ' Adjust quantity if called for
    If (QtyAdjustCheck.Value = True) Then
        With rstCommit
            .Edit

            ' Adjust Location if different
            If (Location <> !Location) Then
                If ((Trim(Location) = "") Or IsNull(Location)) Then
                    Utilities.OperationEntry rstCommit!ID, "Inventory", _
                        "Changed Location from " & rstCommit!Location & " to " & Location & _
                        " after Commit " & CurrentCommitID & " Completion", "Move"
                    !Location = Location
                    !LastOper = EmployeeLogin
                    !LastDate = Now()
                Else
                    MsgBox "Unable to change location", , "Invalid Location"
                End If
            End If

            ' Adjust Quantity if different
            If (OnHand <> !OnHand) Then
                If (Quantity = "") Or IsNull(Quantity) Then
                    Utilities.OperationEntry rstCommit!ID, "Inventory", _
                        "Changed OnHand from " & rstCommit!OnHand & " to " & OnHand & _
                        " after Commit " & CurrentCommitID & " Completion", "Count"
                    !OnHand = OnHand
                    !LastOper = EmployeeLogin
                    !LastDate = Now()
                Else
                    MsgBox "Unable to change OnHand quantity", , "Invalid Quantity"
                End If
            End If

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
' Image_DblClick
'
'------------------------------------------------------------
Private Sub Image_DblClick(Cancel As Integer)
    If Utilities.FileExists(ImagePath) Then
        Utilities.SendMessage True, , , ImagePath
    End If
End Sub

'------------------------------------------------------------
' LocAdjustCheck_Click
'
'------------------------------------------------------------
Private Sub LocAdjustCheck_Click()
    If (LocAdjustCheck.Value = True) Then
        Location.Enabled = True
        Location.Locked = False
        Utilities.FieldAvailableSet Me.Controls("Location")
    Else
        Location.Enabled = False
        Location.Locked = True
        Utilities.FieldAvailableRemove Me.Controls("Location")
    End If
End Sub

'------------------------------------------------------------
' QtyAdjustCheck_Click
'
'------------------------------------------------------------
Private Sub QtyAdjustCheck_Click()
    If (QtyAdjustCheck.Value = True) Then
        OnHand.Enabled = True
        OnHand.Locked = False
        Utilities.FieldAvailableSet Me.Controls("OnHand")
    Else
        OnHand.Enabled = False
        OnHand.Locked = True
        Utilities.FieldAvailableRemove Me.Controls("OnHand")
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
    ItemLength = Nz(rstCommit!ItemLength, "")
    ItemWidth = Nz(rstCommit!ItemWidth, "")
    ItemHeight = Nz(rstCommit!ItemHeight, "")
    ItemDepth = Nz(rstCommit!ItemDepth, "")

    DisplayImage Nz(rstCommit!ImagePath, "")

    Location = Nz(rstCommit!Location, "")
    OnHand = Nz(rstCommit!OnHand, "")
    Committed = Nz(rstCommit!Committed, "")

    ' Fill in Last Change Date and user based on status
    If (Status = "A") Then
        LastUser = ""
        LastDate = ""
    ElseIf (Status = "X") Then
        LastUser = Nz(rstCommit!OperatorCancel, "")
        LastDate = Nz(rstCommit!DateCancel, "")
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

    ' Can only reactivate Change Completed or Cancelled Orders
    If (Status = "C" Or Status = "X") Then
        lblLocAdjust.Visible = False
        LocAdjustCheck.Visible = False
        LocAdjustCheck.Enabled = False
        lblQtyAdjust.Visible = False
        QtyAdjustCheck.Visible = False
        QtyAdjustCheck.Enabled = False
        cmdSave.Enabled = False
        cmdSave.Visible = False
        cmdComplete.Enabled = False
        cmdComplete.Visible = False

        If (EmployeeRole = SalesLevel) Then
            cmdActive.Enabled = False
            cmdActive.Visible = False
        Else
            cmdActive.Enabled = True
            cmdActive.Visible = True
        End If

        cmdCancel.Enabled = False
        cmdCancel.Visible = False
    ' Active Orders
    ElseIf (Status = "A") Then
        If (EmployeeRole = SalesLevel) Then
        ' Employee Sales Role settings
            cmdComplete.Enabled = False
            cmdComplete.Visible = False
            cmdActive.Enabled = False
            cmdActive.Visible = False
            cmdCancel.Enabled = True
            cmdCancel.Visible = True
            lblLocAdjust.Visible = False
            LocAdjustCheck.Visible = False
            LocAdjustCheck.Enabled = False
            lblQtyAdjust.Visible = False
            QtyAdjustCheck.Visible = False
            QtyAdjustCheck.Enabled = False
        ElseIf (EmployeeRole = ProdLevel) Then
        ' Employee Production Role settings
            cmdComplete.Enabled = True
            cmdComplete.Visible = True
            cmdActive.Enabled = False
            cmdActive.Visible = False
            cmdCancel.Enabled = True
            cmdCancel.Visible = True
            lblLocAdjust.Visible = True
            LocAdjustCheck.Visible = True
            LocAdjustCheck.Enabled = True
            lblQtyAdjust.Visible = True
            QtyAdjustCheck.Visible = True
            QtyAdjustCheck.Enabled = True
        Else
        ' Employee Manager Role settings
            cmdComplete.Enabled = True
            cmdComplete.Visible = True
            cmdActive.Enabled = False
            cmdActive.Visible = False
            cmdCancel.Enabled = True
            cmdCancel.Visible = True
            lblLocAdjust.Visible = True
            LocAdjustCheck.Visible = True
            LocAdjustCheck.Enabled = True
            lblQtyAdjust.Visible = True
            QtyAdjustCheck.Visible = True
            QtyAdjustCheck.Enabled = True
        End If
    End If
End Sub


'------------------------------------------------------------
' DisplayImage
'
'------------------------------------------------------------
Private Sub DisplayImage(path As String)
    Dim fileExtension As String

    fileExtension = LCase(Right$(path, Len(path) - InStrRev(path, ".")))

    If Utilities.FileExists(path) And _
        ((fileExtension = "gif") Or (fileExtension = "png") Or _
        (fileExtension = "jpg")) Then
        Image.Picture = path
        ImagePath = path
    Else
        ImagePath = ""
        Image.Picture = ""
    End If
End Sub
