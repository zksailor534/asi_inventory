Option Compare Database
Option Explicit

Private rstItem As DAO.Recordset
Private rstInventory As DAO.Recordset


'------------------------------------------------------------
' Form_Open
'
'------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Form_Open_Err

    On Error GoTo 0

    ' Check for valid item ID
    If Not (CurrentItemID > 0) Then
        MsgBox "Invalid Item ID", vbOKOnly, "Invalid ID"
        DoCmd.Close
        Exit Sub
    End If

    ' Check for valid employee
    If Not (EmployeeID > 0) Then
        MsgBox "User is invalid:" & vbCrLf & "Login as valid user", vbOKOnly, "Invalid User"
        DoCmd.Close
        Exit Sub
    End If

    ' Open the recordsets
    open_db
    Set rstItem = db.OpenRecordset("SELECT TOP 1 * FROM " & ItemDB & " WHERE [ID] = " & CurrentItemID)
    Set rstInventory = db.OpenRecordset("SELECT TOP 1 * FROM " & InventoryDB & " WHERE [ItemID] = " & CurrentItemID)

    FillFields

Form_Open_Exit:
    Exit Sub

Form_Open_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume Form_Open_Exit

End Sub


'------------------------------------------------------------
' Form_Unload
'
'------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    ' Clean up
    'rstItem.Close
    'Set rstItem = Nothing
End Sub


'------------------------------------------------------------
' cmdCommit_Click
'
'------------------------------------------------------------
Private Sub cmdCommit_Click()
On Error GoTo cmdCommit_Click_Err

    Dim commitQuantity As Integer
    Dim commitPrompt, salesOrderPrompt As String
    Dim qtyAvailable As Integer
    Dim rstCommitted As DAO.Recordset

    qtyAvailable = OnHand + OnOrder - Committed

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
        !ItemID = CurrentItemID
        !Status = "A"
        !QtyCommitted = commitQuantity
        !OperatorActive = EmployeeLogin
        !DateActive = Now()
        .Update
    End With

    ' Update the warehouse record
    With rstInventory
        .MoveFirst
        .Edit
        !Committed = !Committed + commitQuantity
        .Update
    End With

    ' Update the form
    Committed = Committed + commitQuantity
    Available = Available - commitQuantity
    LastOper = EmployeeLogin
    LastDate = Now()

    MsgBox "Item Successfully Committed", vbOKOnly, "Commit Successful"

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


'------------------------------------------------------------
' Image_DblClick
'
'------------------------------------------------------------
Private Sub Image_DblClick(Cancel As Integer)
    Utilities.SendMessage True, , , rstItem!ImagePath
End Sub


'------------------------------------------------------------
' FillFields
'
'------------------------------------------------------------
Private Sub FillFields()
    Dim empID As Long

    ProductNameHeader = Nz(rstItem!Product, "")
    Product = Nz(rstItem!Product, "")

    Category = Nz(rstItem!Category, "")

    RecordID = Nz(rstItem!RecordID, "")
    Manufacturer = Nz(rstItem!Manufacturer, "")
    Style = Nz(rstItem!Style, "")
    SuggSellingPrice = Nz(rstItem!SuggSellingPrice, "")
    Focus = Nz(rstItem!Focus, "")
    Color = Nz(rstItem!Color, "")
    Condition = Nz(rstItem!Condition, "")
    Vendor = Nz(rstItem!Vendor, "")
    Description = Nz(rstItem!Description, "")
    ItemLength = Nz(rstItem!ItemLength, "")
    ItemWidth = Nz(rstItem!ItemWidth, "")
    ItemHeight = Nz(rstItem!ItemHeight, "")
    ItemDepth = Nz(rstItem!ItemDepth, "")
    Capacity = Nz(rstItem!Capacity, "")
    Column = Nz(rstItem!Column, "")
    BoltPattern = Nz(rstItem!BoltPattern, "")
    HolePattern = Nz(rstItem!HolePattern, "")
    RollerCenter = Nz(rstItem!RollerCenter, "")
    Diameter = Nz(rstItem!Diameter, "")
    DriveType = Nz(rstItem!DriveType, "")
    Degree = Nz(rstItem!Degree, "")
    Volts = Nz(rstItem!Volts, "")
    AmpHR = Nz(rstItem!AmpHR, "")
    Phase = Nz(rstItem!Phase, "")
    Serial = Nz(rstItem!Serial, "")
    NumSteps = Nz(rstItem!NumSteps, "")
    LowerLiftHeight = Nz(rstItem!LowerLiftHeight, "")
    NumStruts = Nz(rstItem!NumStruts, "")
    QtyDoors = Nz(rstItem!QtyDoors, "")
    TopStepHeight = Nz(rstItem!TopStepHeight, "")
    TopLiftHeight = Nz(rstItem!TopLiftHeight, "")

    ' Create Operator and Date
    If Not IsNull(rstItem!CreateOper) Then
        empID = Utilities.GetEmployeeID(rstItem!CreateOper)
        If (empID <> 0) Then
            CreateOper = Utilities.GetEmployeeName(empID)
        Else
            CreateOper = Nz(rstItem!CreateOper, "")
        End If
    End If
    CreateDate = Nz(rstItem!CreateDate, "")

    ' Load image
    Image.Picture = Nz(rstItem!ImagePath, "")

    ' Order History
    CommitHistory.Form.RecordSource = "SELECT * FROM " & CommitQuery & " WHERE ID=" & CurrentItemID & ";"
    CommitHistory.Form.Requery

    ' Inventory
    Location = Nz(rstInventory!Location, "")
    OrigQty = Nz(rstInventory!OrigQty, "")
    OnOrder = Nz(rstInventory!OnOrder, "")
    OnHand = Nz(rstInventory!OnHand, "")
    Committed = Nz(rstInventory!Committed, "")
    Available = Nz(rstInventory!OnHand + rstInventory!OnOrder - rstInventory!Committed, "")

    ' Last change user or create user
    If Not IsNull(rstInventory!LastOper) Then
        empID = Utilities.GetEmployeeID(rstInventory!LastOper)
        If (empID <> 0) Then
            LastOper = Utilities.GetEmployeeName(empID)
        Else
            LastOper = Nz(rstInventory!LastOper, "")
        End If
    End If

    ' Last change date
    If rstItem!LastChangeDate <> 0 Then
        LastDate = Nz(rstItem!LastChangeDate)
    ElseIf rstItem!CreateDate <> 0 Then
        LastDate = Nz(rstItem!CreateDate)
    End If

End Sub
