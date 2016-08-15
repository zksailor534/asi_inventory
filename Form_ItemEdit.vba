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
    Set rstItem = db.OpenRecordset("SELECT TOP 1 * FROM " & ItemDB & " WHERE [ID]=" & CurrentItemID)

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
' cmdNewRecordID_Click
'
'------------------------------------------------------------
Private Sub cmdNewRecordID_Click()
    If (Prefix <> "") Then
        Prefix = UCase(Prefix)
        RecordID = Utilities.NewRecordID(Prefix, 1)
    Else
        MsgBox "No Record ID Prefix Provided.", vbOKOnly
    End If
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
        UpdateFieldVisibility
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
        sqlQuery = "SELECT ProductName FROM " & ProductQuery & " WHERE Category.Value = " & CategoryID
        Product.RowSource = sqlQuery
    Else
        Product.RowSource = ""
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
    ProductNameHeader = Nz(rstItem!Product, "")
    If (Product <> "") Then
        updateStyleList
        updateColumnList
        UpdateFieldVisibility
    End If

    Category = Nz(rstItem!Category, "")
    If (Category <> "") Then
        updateProductList
        updateManufacturerList
    End If

    Prefix = Nz(Utilities.GetRecordPrefix(rstItem!RecordID), "")
    RecordID = Nz(rstItem!RecordID, "")
    Style = Nz(rstItem!Style, "")
    Manufacturer = Nz(rstItem!Manufacturer, "")
    SuggSellingPrice = Nz(rstItem!SuggSellingPrice, "")
    Focus = Nz(rstItem!Focus, "")
    Color = Nz(rstItem!Color, "")
    Condition = Nz(rstItem!Condition, "")
    Vendor = Nz(rstItem!Vendor, "")
    Description = Nz(rstItem!Description, "")

    DisplayImage Nz(rstItem!ImagePath, "")

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
    If Not IsNull(rstItem!CreateOper) Then
        empID = Utilities.GetEmployeeID(rstItem!CreateOper)
        If (empID <> 0) Then
            CreateOper = Utilities.GetEmployeeName(empID)
        Else
            CreateOper = Nz(rstItem!CreateOper, "")
        End If
    End If
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

    ' Check for valid Product
    If Not (Utilities.IsValidProduct(Product)) Then
        Utilities.FieldErrorSet Me.Controls("Product")
        ValidateFields = False
    Else
        Utilities.FieldErrorClear Me.Controls("Product")
    End If

    ' Check for valid Record ID
    If (RecordID <> rstItem!RecordID) Then
        If Not (Utilities.IsValidRecordID(RecordID)) Then
            Utilities.FieldErrorSet Me.Controls("RecordID")
            ValidateFields = False
        Else
            Utilities.FieldErrorClear Me.Controls("RecordID")
        End If
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


'------------------------------------------------------------
' UpdateFieldVisibility
'
'------------------------------------------------------------
Private Sub UpdateFieldVisibility()
    Dim sqlQuery As String
    Dim formCntrl As Control

    ' Set column visibility and order
    For Each formCntrl In Me.Controls
        If (formCntrl.ControlType = acComboBox) Or (formCntrl.ControlType = acTextBox) Or _
            (formCntrl.ControlType = acCheckBox) Then
            formCntrl.Enabled = Utilities.ProductFieldVisibility(Product, formCntrl.Name)
        End If
    Next
End Sub
