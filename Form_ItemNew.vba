Option Compare Database

Private rstNewItem As DAO.Recordset
Private rstNewInventory As DAO.Recordset

'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    On Error GoTo 0

    ' Open recordsets
    open_db
    Set rstNewItem = db.OpenRecordset(ItemDB, Options:=dbSeeChanges)
    Set rstNewInventory = db.OpenRecordset(InventoryDB)

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
    rstNewItem.Close
    rstNewInventory.Close
    Set rstNewItem = Nothing
    Set rstNewInventory = Nothing
End Sub


'------------------------------------------------------------
' cmdGenerateRecordID_Click
'
'------------------------------------------------------------
Private Sub cmdGenerateRecordID_Click()

    If (Category <> "") Then
        DoCmd.OpenForm GenerateRecordIDForm, , , , , , Category
    Else
        DoCmd.OpenForm GenerateRecordIDForm
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
        SaveInventory
        MsgBox "Save Completed", , "Saved"
        ClearFields
    Else
        GoTo cmdSave_Click_Exit
    End If

cmdSave_Click_Exit:
    If Not (Utilities.HasParent(Me)) Then
        DoCmd.Close
    End If
    Exit Sub

cmdSave_Click_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume cmdSave_Click_Exit

End Sub


'------------------------------------------------------------
' Category_AfterUpdate
'
'------------------------------------------------------------
Private Sub Category_AfterUpdate()
    If (Category <> "") Then
        updatePrefix
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
' cmdNewProduct_Click
'
'------------------------------------------------------------
'Private Sub cmdNewProduct_Click()
'On Error GoTo cmdNewProduct_Click_Err
'
'    On Error GoTo 0
'    TempVars.Remove "Suppliers_ID"
'    DoCmd.OpenForm "SupplierDetails", acNormal, "", "", acAdd, acDialog
'    DoCmd.Requery "SupplierID"
'    If (Not (IsNull(TempVars!Suppliers_ID))) Then
'        DoCmd.SetProperty "SupplierID", , TempVars!Suppliers_ID
'    End If
'    TempVars.Remove "Suppliers_ID"
'
'
'cmdNewSupplier_Click_Exit:
'    Exit Sub
'
'cmdNewSupplier_Click_Err:
'    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
'    Resume cmdNewSupplier_Click_Exit
'
'End Sub


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

    If (Product <> "") Then
        sqlQuery = "SELECT DISTINCT Manufacturer FROM " & ItemDB & _
            " WHERE Category.Value = " & categoryID & _
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
            " WHERE Product = '" & Product & "'" & _
            " ORDER BY Column;"
        Column.RowSource = sqlQuery
    Else
        Column.RowSource = ""
    End If
End Sub


'------------------------------------------------------------
' updatePrefix
' Update Prefix text box with prefix associated with Category
'------------------------------------------------------------
Private Sub updatePrefix()
    Prefix = DLookup("Prefix", CategoryQuery, "CategoryName= '" & Category & "'")
End Sub


'------------------------------------------------------------
' cmdNewProduct_Click
'
'------------------------------------------------------------
Private Sub cmdNewRecordID_Click()
    If (Prefix <> "") Then
        RecordID = Utilities.NewRecordID(Prefix, 1)
    Else
        MsgBox "No Record ID Prefix Provided.", vbOKOnly
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

    ' Check for valid Location
    If (Location = "") Then
        Utilities.FieldErrorSet Me.Controls("Location")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("Location")
    End If

    ' Check for valid Quantity
    If (Quantity = "") Or Not (IsNumeric(Quantity)) Or CLng(Quantity) <= 0 Then
        Utilities.FieldErrorSet Me.Controls("Quantity")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("Quantity")
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
'Private Sub SaveItem()
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


'------------------------------------------------------------
' SaveItem
'
'------------------------------------------------------------
Private Sub SaveItem()
    ' Save Item Record
    With rstNewItem
        .AddNew
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
        !CreateDate = Now()
        !CreateOper = EmployeeLogin
        !LastChangeDate = Now()
        !LastChangeOper = EmployeeLogin
        .Update
    End With

    ' Get new record ID
    rstNewItem.Bookmark = rstNewItem.LastModified
    ItemId = rstNewItem("ID")
End Sub


'------------------------------------------------------------
' SaveInventory
'
'------------------------------------------------------------
Private Sub SaveInventory()
    ' Save Inventory Record
    With rstNewInventory
        .AddNew
        !ItemId = ItemId
        !Location = Location
        !OnHand = Quantity
        !OrigQty = Quantity
        !LastOper = EmployeeLogin
        !LastDate = Now()
        .Update
    End With
End Sub


'------------------------------------------------------------
' ClearFields
'
'------------------------------------------------------------
Private Sub ClearFields()

    Product = ""
    RecordID = ""
    Prefix = ""
    Category = ""
    Manufacturer = ""
    Style = ""
    SuggSellingPrice = ""
    Color = ""
    Condition = ""
    Vendor = ""
    Description = ""
    ItemLength = ""
    ItemWidth = ""
    ItemHeight = ""
    ItemDepth = ""
    Capacity = ""
    Column = ""
    BoltPattern = ""
    HolePattern = ""
    RollerCenter = ""
    Diameter = ""
    DriveType = ""
    Degree = ""
    Volts = ""
    AmpHR = ""
    Phase = ""
    Serial = ""
    Gauge = ""
    NumSteps = ""
    LowerLiftHeight = ""
    NumStruts = ""
    QtyDoors = ""
    TopStepHeight = ""
    TopLiftHeight = ""
    Location = ""
    Quantity = ""

End Sub
