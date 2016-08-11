Option Compare Database

Private rstInventory As DAO.Recordset

'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    On Error GoTo 0

    ' Check for valid item ID
    If Not (CurrentItemID > 0) Then
        MsgBox "Invalid Item ID", vbOKOnly, "Invalid ID"
        DoCmd.Close
        Exit Sub
    End If

    ' Filter form for info
    Me.Filter = "[ID] = " & CurrentItemID
    Me.FilterOn = True

    ' Open recordset
    open_db
    Set rstInventory = db.OpenRecordset("SELECT * FROM " & InventoryDB & " WHERE [ItemID] = " & CurrentItemID)

    FillFields

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
    rstInventory.Close
    Set rstInventory = Nothing
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

    ' Check for valid field entries
    If ValidateFields Then
        SaveInventory
    Else
        GoTo cmdSave_Click_Exit
    End If
    MsgBox "Item Successfully Saved!", , "Save Complete"

cmdSave_Click_Exit:
    Exit Sub

cmdSave_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdSave_Click_Exit

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


'------------------------------------------------------------
' ValidateFields
'
'------------------------------------------------------------
Private Function ValidateFields() As Boolean

    ValidateFields = True

    ' Check for valid user
    If (EmployeeRole = SalesLevel) Then
        MsgBox "User is invalid:" & vbCrLf & "Login as valid user", vbOKOnly, "Invalid User"
        ValidateFields = False
        GoTo ExitNow
    End If

    ' Check if anything has changed
    If (Location = rstInventory!Location) And CLng(OnHand) = rstInventory!OnHand And _
        CLng(OnOrder) = rstInventory!OnOrder Then
        MsgBox "Invalid Save:" & vbCrLf & "Nothing has changed", vbOKOnly, "Invalid Save"
        Utilities.FieldErrorClear Me.Controls("Location")
        Utilities.FieldErrorClear Me.Controls("OnHand")
        Utilities.FieldErrorClear Me.Controls("OnOrder")
        ValidateFields = False
        GoTo ExitNow
    End If

    ' Check for valid Location
    If (Location = "") Then
        Utilities.FieldErrorSet Me.Controls("Location")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("Location")
    End If

    ' Check for valid On Hand Quantity
    If (OnHand = "") Or Not (IsNumeric(OnHand)) Or _
        CLng(OnHand) < CLng(Committed) Then
        Utilities.FieldErrorSet Me.Controls("OnHand")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("OnHand")
    End If

    ' Check for valid On Order Quantity
    If (OnOrder = "") Or Not (IsNumeric(OnOrder)) Or CLng(OnOrder) < 0 Then
        Utilities.FieldErrorSet Me.Controls("OnOrder")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("OnOrder")
    End If

ExitNow:
    Exit Function

End Function


'------------------------------------------------------------
' FillFields
'
'------------------------------------------------------------
Private Sub FillFields()
    Dim empID As Long
    Location = Nz(rstInventory!Location)
    OnHand = Nz(rstInventory!OnHand)
    OnOrder = Nz(rstInventory!OnOrder)

    If Not IsNull(rstInventory!LastOper) Then
        empID = Utilities.GetEmployeeID(rstInventory!LastOper)
        If (empID <> 0) Then
            LastOper = Utilities.GetEmployeeName(empID)
        Else
            LastOper = Nz(rstInventory!LastOper, "")
        End If
    End If
End Sub


'------------------------------------------------------------
' SaveInventory
'
'------------------------------------------------------------
Private Sub SaveInventory()
    Dim change As String

    change = ""

    If (rstInventory!Location <> Location) Then
        change = change & "Changed Location from " & rstInventory!Location & " to " & Location & ";"
    End If

    If (rstInventory!OnHand <> OnHand) Then
        change = change & "Changed OnHand from " & rstInventory!OnHand & " to " & OnHand & ";"
    End If

    If (rstInventory!OnOrder <> OnOrder) Then
        change = change & "Changed OnOrder from " & rstInventory!OnOrder & " to " & OnOrder & ";"
    End If

    If (change <> "") Then
        ' Save change record
        Utilities.OperationEntry CurrentItemID, "Inventory", change

        ' Save Inventory Record
        With rstInventory
            .Edit
            !Location = Location
            !OnHand = OnHand
            !OnOrder = OnOrder
            !LastOper = EmployeeLogin
            !LastDate = Now()
            .Update
        End With
    End If
End Sub
