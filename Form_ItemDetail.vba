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

    ' Prompt for commit quantity
    commitPrompt = InputBox("Quantity", "Quantity to Commit?", 1)

    ' Verify response is numeric
    If Not (IsNumeric(commitPrompt)) Then
        MsgBox "Commit Quantity must be numeric", , "Invalid Commit Quantity"
        GoTo cmdCommit_Click_Exit
    End If

    commitQuantity = CInt(commitPrompt)
    qtyAvailable = Me.OnHand + Me.OnOrder - Me.Committed

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

    ' Change the qty committed in the warehouse form
    Committed = Committed + commitQuantity



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
