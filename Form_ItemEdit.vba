Option Compare Database

Private rstItem As DAO.Recordset

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

    ' Check for valid employee
    If Not (EmployeeID > 0) Then
        MsgBox "User is invalid:" & vbCrLf & "Login as valid user", vbOKOnly, "Invalid User"
        DoCmd.Close
        Exit Sub
    End If

    ' Open the recordset
    open_db
    Set rstItem = db.OpenRecordset("SELECT * FROM " & ItemDB & " WHERE [ID] = " & CurrentItemID)

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
    If Not (rstItem Is Nothing) Then
        ' Clean Up
        rstItem.Close
        Set rstItem = Nothing
    End If
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
        SaveItem
        MsgBox "Item Successfully Saved!", , "Save Complete"
        FillFields
    Else
        MsgBox "Unable to save", , "Save Failed"
    End If

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

    DoCmd.Close
    Exit Sub

cmdClose_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical

End Sub


'------------------------------------------------------------
' Category_AfterUpdate
'
'------------------------------------------------------------
Private Sub Category_AfterUpdate()
    If (Category <> "") Then
        updateProductList
        updateManufacturerList
    End If
End Sub


'------------------------------------------------------------
' Product_AfterUpdate
'
'------------------------------------------------------------
Private Sub Product_AfterUpdate()
    If (Product <> "") Then
        ProductNameHeader = Product
        updateStyleList
        updateColumnList
    End If
End Sub


'------------------------------------------------------------
' updateProductList
' Update Product Combo box with all products in category
'------------------------------------------------------------
Private Sub updateProductList()
    Dim categoryID As Long
    Dim sqlQuery As String

    categoryID = Utilities.GetCategoryID(Category)
    If (categoryID <> 0) Then
        sqlQuery = "SELECT ProductName FROM " & ProductDB & " WHERE Category.Value = " & categoryID
        Product.RowSource = sqlQuery
    End If
End Sub


'------------------------------------------------------------
' updateStyleList
' Update Style Combo box with all products in category
'------------------------------------------------------------
Private Sub updateStyleList()
    Dim sqlQuery As String

    If (Product <> "") Then
        sqlQuery = "SELECT DISTINCT Style FROM " & ItemDB & _
            " WHERE Product = '" & Product & "' AND Style <> ''" & _
            " ORDER BY Style;"
        Style.RowSource = sqlQuery
    Else
        Style.RowSource = ""
    End If
End Sub


'------------------------------------------------------------
' updateManufacturerList
' Update Manufacturer Combo box with all Manufacturers of category
'------------------------------------------------------------
Private Sub updateManufacturerList()
    Dim sqlQuery As String

    If (Category <> "") Then
        sqlQuery = "SELECT DISTINCT Manufacturer FROM " & ItemDB & _
            " WHERE Category = '" & Category & "' AND Manufacturer <> ''" & _
            " ORDER BY Manufacturer;"
        Manufacturer.RowSource = sqlQuery
    Else
        Manufacturer.RowSource = ""
    End If
End Sub


'------------------------------------------------------------
' updateColumnList
' Update Column Combo box with all column values
'------------------------------------------------------------
Private Sub updateColumnList()
    Dim sqlQuery As String

    If (Product <> "") Then
        sqlQuery = "SELECT DISTINCT Column FROM " & ItemDB & _
            " WHERE Product = '" & Product & "' AND Column <> ''" & _
            " ORDER BY Column;"
        Column.RowSource = sqlQuery
    Else
        Column.RowSource = ""
    End If
End Sub


'------------------------------------------------------------
' FillFields
'
'------------------------------------------------------------
Private Sub FillFields()
    Dim empID As Long

    Product = Nz(rstItem!Product, "")
    If (Product <> "") Then
        updateStyleList
        updateColumnList
    End If

    Category = Nz(rstItem!Category, "")
    If (Category <> "") Then
        updateProductList
        updateManufacturerList
    End If

    RecordID = Nz(rstItem!RecordID, "")
    Manufacturer = Nz(rstItem!Manufacturer, "")
    Style = Nz(rstItem!Style, "")
    SuggSellingPrice = Nz(rstItem!SuggSellingPrice, "")
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
    Gauge = Nz(rstItem!Gauge, "")
    NumSteps = Nz(rstItem!NumSteps, "")
    LowerLiftHeight = Nz(rstItem!LowerLiftHeight, "")
    NumStruts = Nz(rstItem!NumStruts, "")
    QtyDoors = Nz(rstItem!QtyDoors, "")
    TopStepHeight = Nz(rstItem!TopStepHeight, "")
    TopLiftHeight = Nz(rstItem!TopLiftHeight, "")
    CreateOper = Nz(rstItem!CreateOper, "")
    CreateDate = Nz(rstItem!CreateDate, "")

    ' Last change user or create user
    If Not IsNull(rstItem!LastChangeOper) Then
        empID = Utilities.GetEmployeeID(rstItem!LastChangeOper)
        If (empID <> 0) Then
            LastOper = Utilities.GetEmployeeName(empID)
        Else
            LastOper = Nz(rstItem!LastChangeOper, "")
        End If
    ElseIf Not IsNull(rstItem!CreateOper) Then
        empID = Utilities.GetEmployeeID(rstItem!CreateOper)
        If (empID <> 0) Then
            LastOper = Utilities.GetEmployeeName(empID)
        Else
            LastOper = Nz(rstItem!CreateOper, "")
        End If
    End If

    ' Last change date
    If rstItem!LastChangeDate <> 0 Then
        LastDate = Nz(rstItem!LastChangeDate)
    ElseIf rstItem!CreateDate <> 0 Then
        LastDate = Nz(rstItem!CreateDate)
    End If
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

    ' Check for valid Category
    If Not (Utilities.IsValidCategory(Category)) Then
        Utilities.FieldErrorSet Me.Controls("Category")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("Category")
    End If

    ' Check for valid Product (TBD: expand for Product List)
    If (Product = "") Then
        Utilities.FieldErrorSet Me.Controls("Product")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("Product")
    End If

    ' Check for valid Record ID
    If Not (Utilities.IsValidRecordID(RecordID)) Then
        Utilities.FieldErrorSet Me.Controls("RecordID")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("RecordID")
    End If

ExitNow:
    Exit Function

End Function


'------------------------------------------------------------
' SaveItem
'
'------------------------------------------------------------
Private Sub SaveItem()
    ' Save Item Record
    With rstItem
        .MoveFirst
        .Edit
        !Category = Category
        !RecordID = RecordID
        !Manufacturer = Manufacturer
        !Product = Product
        !Style = Style

        If (ItemLength = "") Then
            !ItemLength = Null
        Else
            !ItemLength = CLng(ItemLength)
        End If

        If (ItemWidth = "") Then
            !ItemWidth = Null
        Else
            !ItemWidth = CLng(ItemWidth)
        End If

        If (ItemHeight = "") Then
            !ItemHeight = Null
        Else
            !ItemHeight = CLng(ItemHeight)
        End If

        If (ItemDepth = "") Then
            !ItemDepth = Null
        Else
            !ItemDepth = CLng(ItemDepth)
        End If

        !RollerCenter = RollerCenter
        !Column = Column
        !Color = Color
        !Condition = Condition
        !Description = Description
        !Vendor = Vendor

        If (SuggSellingPrice = "") Then
            !SuggSellingPrice = Null
        Else
            !SuggSellingPrice = CCur(SuggSellingPrice)
        End If

        !Capacity = Capacity
        !BoltPattern = BoltPattern
        !HolePattern = HolePattern
        !Diameter = Diameter
        !Degree = Degree
        !DriveType = DriveType
        !Gauge = Gauge
        !NumStruts = NumStruts
        !NumSteps = NumSteps
        !Volts = Volts
        !Phase = Phase
        !AmpHR = AmpHR
        !Serial = Serial
        !QtyDoors = QtyDoors
        !TopLiftHeight = TopLiftHeight
        !LowerLiftHeight = LowerLiftHeight
        !TopStepHeight = TopStepHeight
        !LastChangeDate = Now()
        !LastChangeOper = EmployeeLogin
        .Update
    End With
End Sub
