Option Compare Database

Private rstItem As DAO.Recordset
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
    Set rstItem = db.OpenRecordset(ItemDB, Options:=dbSeeChanges)
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
    rstItem.Close
    rstNewInventory.Close
    Set rstItem = Nothing
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
        If (RecordID.ListIndex = -1) Then
            SaveNewItem
        Else
            Set rstItem = db.OpenRecordset("SELECT TOP 1 * FROM " & ItemDB & _
                " WHERE [RecordID] = '" & RecordID & "'")
            SaveReservedItem
        End If
        SaveInventory
        MsgBox "Item Successfully Saved!", , "Save Complete"
        ClearFields
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
' cmdBrowse_Click
'
'------------------------------------------------------------
Private Sub cmdBrowse_Click()
    Dim strChoice As String

    strChoice = FileSelection

    DisplayImage (strChoice)
End Sub


'------------------------------------------------------------
' Category_AfterUpdate
'
'------------------------------------------------------------
Private Sub Category_AfterUpdate()
    If (Category <> "") Then
        updatePrefix
        updateReservedRecordIDs
        updateProductList
        updateManufacturerList
        updateVendorList
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
    Dim CategoryID As Long
    Dim sqlQuery As String

    CategoryID = Utilities.GetCategoryID(Category)
    If (CategoryID <> 0) Then
        sqlQuery = "SELECT ProductName FROM " & ProductQuery & " WHERE Products.Category.Value = " & CategoryID & _
            " ORDER BY ProductName;"
        Product.RowSource = sqlQuery
    End If
End Sub


'------------------------------------------------------------
' updateReservedRecordIDs
' Update Record ID Combo box with any reserved Record IDs
'------------------------------------------------------------
Private Sub updateReservedRecordIDs()
    Dim CategoryID As Long
    Dim sqlQuery As String

    CategoryID = Utilities.GetCategoryID(Category)
    If (CategoryID <> 0) Then
        sqlQuery = "SELECT [RecordID],[VENDOR] FROM " & ItemDB & _
            " WHERE Category = '" & Category & "' AND [Vendor] = 'RESERVED'" & _
            " AND CreateOper = '" & EmployeeLogin & "'"
        RecordID.RowSource = sqlQuery
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
' updateVendorList
' Update Vendor Combo box with all Vendors of category
'------------------------------------------------------------
Private Sub updateVendorList()
    Dim sqlQuery As String

    If (Category <> "") Then
        sqlQuery = "SELECT DISTINCT Vendor FROM " & ItemDB & _
            " WHERE Category = '" & Category & "' AND Vendor <> ''" & _
            " ORDER BY Vendor;"
        Vendor.RowSource = sqlQuery
    Else
        Vendor.RowSource = ""
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
' updatePrefix
' Update Prefix text box with prefix associated with Category
'------------------------------------------------------------
Private Sub updatePrefix()
    Prefix = DLookup("Prefix", CategoryQuery, "CategoryName= '" & Category & "'")
End Sub


'------------------------------------------------------------
' cmdNewRecordID_Click
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
    Dim locationPrompt As Integer
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
        If (Location = "INBOUND") Then
            ' Check if item is on order
            locationPrompt = MsgBox("Is this Item On Order?", vbYesNo, "On Order")
            If (locationPrompt = vbNo) Then
                MsgBox "Choose a valid Location", , "Invalid Location"
                Utilities.FieldErrorSet Me.Controls("Location")
                ValidateFields = False
            Else
                Utilities.FieldErrorClear Me.Controls("Location")
            End If
        Else
            Utilities.FieldErrorClear Me.Controls("Location")
        End If
    End If

    ' Check for valid Quantity
    If (Quantity = "") Or IsNull(Quantity) Then
        Utilities.FieldErrorSet Me.Controls("Quantity")
        ValidateFields = False
    ElseIf Not (IsNumeric(Quantity)) Or CLng(Quantity) <= 0 Then
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

    ' Check for valid Product
    If Not (Utilities.IsValidProduct(Product)) Then
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

    ' Check for valid Manufacturer
    If (Len(Manufacturer) > 25) Then
        Utilities.FieldErrorSet Me.Controls("Manufacturer")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("Manufacturer")
    End If

    ' Check for valid Vendor
    If (Len(Vendor) > 25) Then
        Utilities.FieldErrorSet Me.Controls("Vendor")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("Vendor")
    End If

    ' Check for valid suggested price
    If (SuggSellingPrice <> "") Then
        If Not (IsNumeric(SuggSellingPrice)) Then
            Utilities.FieldErrorSet Me.Controls("SuggSellingPrice")
            ValidateFields = False
        Else
            Utilities.FieldErrorClear Me.Controls("SuggSellingPrice")
        End If
    End If

ExitNow:
    Exit Function

End Function


'------------------------------------------------------------
' SaveNewItem
'
'------------------------------------------------------------
Private Sub SaveNewItem()
On Error GoTo SaveNewItem_Err
    ' Save Item Record
    With rstItem
        .AddNew
        !Category = Category
        !RecordID = RecordID
        !Manufacturer = Manufacturer
        !Product = Product
        !Style = Style

        If IsNull(ItemLength) Or Not (IsNumeric(ItemLength)) Then
            !ItemLength = Null
        Else
            !ItemLength = CSng(ItemLength)
        End If

        If IsNull(ItemWidth) Or Not (IsNumeric(ItemWidth)) Then
            !ItemWidth = Null
        Else
            !ItemWidth = CSng(ItemWidth)
        End If

        If IsNull(ItemHeight) Or Not (IsNumeric(ItemHeight)) Then
            !ItemHeight = Null
        Else
            !ItemHeight = CSng(ItemHeight)
        End If

        If IsNull(ItemDepth) Or Not (IsNumeric(ItemDepth)) Then
            !ItemDepth = Null
        Else
            !ItemDepth = CSng(ItemDepth)
        End If

        !RollerCenter = RollerCenter
        !Column = Column
        !Color = Color
        !Condition = Condition
        !Description = Description
        !Vendor = Vendor

        If IsNull(SuggSellingPrice) Or Not (IsNumeric(SuggSellingPrice)) Then
            !SuggSellingPrice = Null
        Else
            !SuggSellingPrice = CCur(SuggSellingPrice)
        End If

        !Focus = Focus
        !ImagePath = ImagePath
        !Capacity = Capacity
        !BoltPattern = BoltPattern
        !HolePattern = HolePattern
        !Diameter = Diameter
        !Degree = Degree
        !DriveType = DriveType
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
    rstItem.Bookmark = rstItem.LastModified
    ItemID = rstItem("ID")

SaveNewItem_Exit:
    Exit Sub

SaveNewItem_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume SaveNewItem_Exit

End Sub


'------------------------------------------------------------
' SaveReservedItem
'
'------------------------------------------------------------
Private Sub SaveReservedItem()
On Error GoTo SaveReservedItem_Err
    ' Save Item Record
    With rstItem
        .MoveFirst
        .Edit
        !Category = Category
        !RecordID = RecordID
        !Manufacturer = Manufacturer
        !Product = Product
        !Style = Style

        If IsNull(ItemLength) Or Not (IsNumeric(ItemLength)) Then
            !ItemLength = Null
        Else
            !ItemLength = CSng(ItemLength)
        End If

        If IsNull(ItemWidth) Or Not (IsNumeric(ItemWidth)) Then
            !ItemWidth = Null
        Else
            !ItemWidth = CSng(ItemWidth)
        End If

        If IsNull(ItemHeight) Or Not (IsNumeric(ItemHeight)) Then
            !ItemHeight = Null
        Else
            !ItemHeight = CSng(ItemHeight)
        End If

        If IsNull(ItemDepth) Or Not (IsNumeric(ItemDepth)) Then
            !ItemDepth = Null
        Else
            !ItemDepth = CSng(ItemDepth)
        End If

        !RollerCenter = RollerCenter
        !Column = Column
        !Color = Color
        !Condition = Condition
        !Description = Description
        !Vendor = Vendor

        If IsNull(SuggSellingPrice) Or Not (IsNumeric(SuggSellingPrice)) Then
            !SuggSellingPrice = Null
        Else
            !SuggSellingPrice = CCur(SuggSellingPrice)
        End If

        !Focus = Focus
        !ImagePath = ImagePath
        !Capacity = Capacity
        !BoltPattern = BoltPattern
        !HolePattern = HolePattern
        !Diameter = Diameter
        !Degree = Degree
        !DriveType = DriveType
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

    ' Save new item ID for Inventory entry
    rstItem.Bookmark = rstItem.LastModified
    ItemID = rstItem("ID")

SaveReservedItem_Exit:
    Exit Sub

SaveReservedItem_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume SaveReservedItem_Exit

End Sub


'------------------------------------------------------------
' SaveInventory
'
'------------------------------------------------------------
Private Sub SaveInventory()
    ' Save Inventory Record
    With rstNewInventory
        .AddNew
        !ItemID = ItemID
        !Location = Location

        If (Location = "INBOUND") Then
            !OnOrder = Quantity
        Else
            !OnHand = Quantity
        End If

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
    ProductNameHeader = ""
    ItemID = ""
    RecordID = ""
    Prefix = ""
    Category = ""
    Manufacturer = ""
    Style = ""
    SuggSellingPrice = ""
    Color = ""
    Condition = ""
    Vendor = ""
    Focus = No
    ImagePath = ""
    Image.Picture = ""
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
    NumSteps = ""
    LowerLiftHeight = ""
    NumStruts = ""
    QtyDoors = ""
    TopStepHeight = ""
    TopLiftHeight = ""
    Location = ""
    Quantity = ""

End Sub


'------------------------------------------------------------
' FileSelection
'
'------------------------------------------------------------
Private Function FileSelection() As String
    Dim objFD As Object
    Dim strOut As String

    strOut = vbNullString
    Set objFD = Application.FileDialog(msoFileDialogFilePicker)
    If objFD.Show = -1 Then
        strOut = objFD.SelectedItems(1)
    End If
    Set objFD = Nothing
    FileSelection = strOut
End Function


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
