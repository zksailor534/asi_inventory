Option Compare Database

'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    On Error GoTo 0
    If CurrentItemID > 0 Then
        Me.Filter = "[ID] = " & CurrentItemID
        Me.FilterOn = True
    Else
        MsgBox "Invalid Item ID", vbOKOnly, "Invalid ID"
        DoCmd.Close
        Exit Sub
    End If

    ' Open database if not open
    open_db

Form_Load_Exit:
    Exit Sub

Form_Load_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume Form_Load_Exit

End Sub


'------------------------------------------------------------
' cmdCommit_Click
'
'------------------------------------------------------------
Private Sub cmdCommit_Click()

    Dim commitQuantity As Integer
    Dim commitPrompt, salesOrderPrompt As String
    Dim qtyAvailable As Integer
    Dim rstCommitted As DAO.Recordset

    qtyAvailable = Me.OnHand + Me.OnOrder - Me.Committed
    ' Check quantity available
    If (qtyAvailable <= 0) Then
        MsgBox "Available Quantity must be greater than Zero" & vbCrLf & _
            "Contact Administrator", , "Invalid Quantity Available"
        GoTo cmdCommit_Click_Exit
    End If

    ' Prompt for commit quantity
    commitPrompt = InputBox("Quantity to Commit?" & vbCrLf & vbCrLf & _
        "Record ID: " & RecordID & vbCrLf & _
        "Product: " & Product & vbCrLf & _
        "Quantity Available: " & qtyAvailable & vbCrLf, _
        "Quantity", "")

    If (commitPrompt = "") Then
        ' Deal with cancel selection
        GoTo cmdCommit_Click_Exit
    ElseIf Not (IsNumeric(commitPrompt)) Then
        ' Verify response is numeric
        MsgBox "Commit Quantity must be numeric", , "Invalid Commit Quantity"
        GoTo cmdCommit_Click_Exit
    End If

    commitQuantity = CInt(commitPrompt)

    ' Verify chosen quantity is available
    If (commitQuantity > qtyAvailable) Then
        MsgBox "Commit Quantity cannot exceed " & qtyAvailable, vbOKOnly, "Invalid Commit Quantity"
        GoTo cmdCommit_Click_Exit
    End If

    ' Verify chosen quantity is positive
    If (commitQuantity <= 0) Then
        MsgBox "Commit Quantity cannot be less than or equal to zero (0) ", vbOKOnly, "Invalid Commit Quantity"
        GoTo cmdCommit_Click_Exit
    End If

    ' Open the recordset
    Set rstCommitted = db.OpenRecordset(CommitDB)

cmdCommit_getSO:
    salesOrderPrompt = InputBox("Sales Order for Commit", "Sales Order", "")

    If (salesOrderPrompt = "") Then
        GoTo cmdCommit_Click_Exit
    ElseIf "A" & Trim(salesOrderPrompt) = "A" Then
        ' If nothing is entered or if all blanks are entered go back and get it again
        MsgBox "Retry: Enter Sales Order Number or Customer", vbOKOnly, "Invalid Data"
        GoTo cmdCommit_getSO
    End If

    ' Add the committed record
    With rstCommitted
        .AddNew
        !Location = Location
        !Reference = salesOrderPrompt
        !ItemId = ItemId
        !Status = "A"
        !QtyCommitted = commitQuantity
        !OperatorActive = EmployeeLogin
        !DateActive = Now()
        .Update
    End With

    ' Update the warehouse form
    Me.Committed = Me.Committed + commitQuantity
    Me.LastOper = EmployeeLogin
    Me.LastDate = Now()

cmdCommit_Click_Exit:
    Exit Sub

cmdCommit_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdCommit_Click_Exit

End Sub


'------------------------------------------------------------
' cmdClose_Click
'
'------------------------------------------------------------
Private Sub cmdClose_Click()
On Error GoTo cmdClose_Click_Err

    ' -------------------------------------------------------------------
    ' Error Handling
    ' -------------------------------------------------------------------
    On Error GoTo cmdClose_Click_Err
    DoCmd.Close , ""

cmdClose_Click_Exit:
    Exit Sub

cmdClose_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdClose_Click_Exit

End Sub
