Option Compare Database

'------------------------------------------------------------
' American Surplus Inventory Database
' Author: Nathanael Greene
' Current Revision: 2.2.1
' Revision Date: 10/19/2015
'
' Revision History:
'   2.0.0:  Initial Release replaces legacy database
'           Complete GUI overhaul
'           Introduction of product-based structure
'           Add commit management for all users
'           Add Generate Record ID tools
'   2.0.1:  Bug fixes (ItemEdit, ItemNew, ItemInventoryManage,
'               Main, CategoriesEdit)
'   2.0.2:  Bug fixes (ItemEdit, ItemNew) - invalid null in
'               numeric inputs
'   2.0.3:  Bug fixes (Commit_Cancel, ItemNew) - cancel had
'               wrong sign (committing more)
'               Added Record ID search to Inventory Manage
'               Added scroll bar to ItemNew
'   2.0.4:  Bug fix (Utilities) - Commit_Cancel & Commit_Complete
'               missing recordset reference
'   2.1.0:  Bug fix (ItemInventoryManage) - Operations changes
'               overwrite; save L,W,H,D as Single
'           Bug fix (Utilities) - Commit_complete allowed <0
'               values after complete, added check
'           Bug fix (qryItemWarehouse) - Did not display OnOrder
'               items
'           Bug fix (ItemEdit) - Did not allow some current RecordIDs
'           Added RecalculateOriginalQuantities, ReclaimRecordIDs
'           Upgraded NewRecordID calculation
'           Added Record ID reservation system
'           Add Inbound item toggle to InventoryManage (replace Print)
'   2.1.1:  Bug fix (Commit_Complete)
'   2.1.2:  Bug fix (NewRecordID) - not finding next record
'               after full list
'           Bug fix (CategoriesDS) - Field12 not being updated in form
'           Bug fix (EmployeesDS) - Roles selector not working
'   2.1.3:  Bug fix (Commit_Complete) - not allowed to complete
'               when Committed > OnHand (+ Onorder added)
'   2.2.0:  New: Reorganized program into Sales, Production, Admin
'               Changed Print Range to print whole screen
'               Incorporated SW Version recording
'               Eliminated Subcategory field in Items
'   2.2.1:  Bug fix - re-link to database backend
'   2.2.2:  Bug fix (ProductionInventory) - Fix RecordID Filter
'               (ItemNew) - Vendor and Manufacturer field limits
'------------------------------------------------------------

'------------------------------------------------------------
' Global constants
'
'------------------------------------------------------------
Public Const ReleaseVersion As String = "2.2.2"
''' User Roles
Public Const DevelLevel As String = "Devel"
Public Const AdminLevel As String = "Admin"
Public Const ProdLevel As String = "Prod"
Public Const SalesLevel As String = "Sales"
''' Database Table Names
Public Const ItemDB As String = "Items"
Public Const InventoryDB As String = "Inventory"
Public Const ProductDB As String = "Products"
Public Const CategoryDB As String = "Categories"
Public Const CommitDB As String = "CommittedTable"
Public Const OperationDB As String = "Operations"
Public Const EmployeeDB As String = "Employees"
Public Const SettingsDB As String = "Settings"
''' Query Names
Public Const WarehouseQuery As String = "qryItemWarehouse"
Public Const CategoryQuery As String = "qryCategoryList"
Public Const CommitQuery As String = "qryItemCommit"
Public Const ProductQuery As String = "qrySubProducts"
''' Form Names
Public Const MainForm As String = "Main"
Public Const LoginForm As String = "Login"
Public Const SalesForm As String = "SalesForm"
Public Const SalesSearch As String = "SalesInventory"
Public Const SalesSearchSplit As String = "SalesInventorySplit"
Public Const InventoryManageForm As String = "InventoryManage"
Public Const InventoryOrderForm As String = "InventoryOrder"
Public Const ItemDetailForm As String = "ItemDetail"
Public Const ItemEditForm As String = "ItemEdit"
Public Const ItemInventoryManageForm As String = "ItemInventoryManage"
Public Const PrintRangeForm As String = "PrintRange"
Public Const CategoriesEditForm As String = "CategoriesEdit"
Public Const GenerateRecordIDForm As String = "GenerateRecordID"
Public Const OrderCommitManageForm As String = "OrderCommitManage"

'------------------------------------------------------------
' Application properties
'
'------------------------------------------------------------
Public db As Database
Public CurrentItemID As Long
Public CurrentCommitID As Long
Public CurrentSalesOrder As String
Public PrintCategorySelected As String
Public PrintFilter As String
Public CompanyName As String
Public CompanyAddress As String
Public CompanyCity As String
Public CompanyStateProvince As String
Public CompanyZipPostal As String
Public CompanyEmail As String
Public CompanyWebsite As String
Public CompanyPhone As String
Public CompanyFax As String
Private pvEmployeeID As Long
Public EmployeeName As String
Public EmployeeLogin As String
Public EmployeePassword As String
Public EmployeeRole As String
Public EmployeeCategory As String
Public ValidLogin As Boolean
Public ScreenWidth As Long

Public Property Get EmployeeID() As Long
    EmployeeID = pvEmployeeID
End Property

Public Property Let EmployeeID(Value As Long)
    pvEmployeeID = Value
    EmployeeName = GetEmployeeName(Value)
    EmployeeLogin = GetEmployeeLogin(Value)
    EmployeePassword = GetEmployeePassword(Value)
    EmployeeRole = GetEmployeeRole(Value)
    EmployeeCategory = GetEmployeeDefaultCategory(Value)
End Property


'------------------------------------------------------------
' open_db
'
'------------------------------------------------------------
Public Function open_db()
    '*** Open the database
    Set db = CurrentDb
End Function


'------------------------------------------------------------
' LoadSettings
'
'------------------------------------------------------------
Public Sub LoadSettings(Company As String)
On Error GoTo LoadSettings_Err

    CompanyName = Nz(DLookup("Company", SettingsDB, "[Company]='" & Company & "'"), "")
    CompanyAddress = Nz(DLookup("Address", SettingsDB, "[Company]='" & Company & "'"), "")
    CompanyCity = Nz(DLookup("City", SettingsDB, "[Company]='" & Company & "'"), "")
    CompanyStateProvince = Nz(DLookup("StateProvince", SettingsDB, "[Company]='" & Company & "'"), "")
    CompanyZipPostal = Nz(DLookup("ZipPostal", SettingsDB, "[Company]='" & Company & "'"), "")
    CompanyEmail = Nz(DLookup("Email", SettingsDB, "[Company]='" & Company & "'"), "")
    CompanyWebsite = Nz(DLookup("WebPage", SettingsDB, "[Company]='" & Company & "'"), "")
    CompanyPhone = Nz(DLookup("BusinessPhone", SettingsDB, "[Company]='" & Company & "'"), "")
    CompanyFax = Nz(DLookup("BusinessFax", SettingsDB, "[Company]='" & Company & "'"), "")

LoadSettings_Exit:
    Exit Sub

LoadSettings_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume LoadSettings_Exit
End Sub


'------------------------------------------------------------
' ConfirmLogin
'
'------------------------------------------------------------
Function ConfirmLogin()
On Error GoTo ConfirmLogin_Err

    ' ------------------------------------------------------------------
    ' Bail out of we already have a current user
    ' ------------------------------------------------------------------
    If (ValidLogin) Then
        Forms(MainForm)!lblCurrentEmployeeName.Caption = "Hello, " & EmployeeName
        Forms(MainForm)!lblVersion.Caption = "Version " & ReleaseVersion
        Exit Function
    Else
        DoCmd.OpenForm LoginForm, acNormal, "", "", , acDialog
        CompleteLogin
    End If

ConfirmLogin_Exit:
    Exit Function

ConfirmLogin_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume ConfirmLogin_Exit

End Function


'------------------------------------------------------------
' CompleteLogin
'
'------------------------------------------------------------
Function CompleteLogin()
On Error GoTo CompleteLogin_Err

    On Error GoTo 0
    DoCmd.Close acForm, MainForm, acSaveNo
    DoCmd.OpenForm MainForm
    Forms(MainForm)!lblCurrentEmployeeName.Caption = "Hello, " & EmployeeName
    Forms(MainForm)!lblVersion.Caption = "Version " & ReleaseVersion
    SetEmployeeVersion (EmployeeID)
    Err.Clear

CompleteLogin_Exit:
    Exit Function

CompleteLogin_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume CompleteLogin_Exit

End Function


'------------------------------------------------------------
' GetEmployeeID
'
'------------------------------------------------------------
Public Function GetEmployeeID(Login As String) As Long
    On Error GoTo ErrHandler
    Dim ID As Long
    ID = DLookup("ID", EmployeeDB, "[Login]='" & Login & "'")
    If IsNull(ID) Then
        GetEmployeeID = 0
    Else
        GetEmployeeID = ID
    End If
    Exit Function
ErrHandler:
    GetEmployeeID = 0
    Exit Function
End Function


'------------------------------------------------------------
' GetEmployeeName
'
'------------------------------------------------------------
Public Function GetEmployeeName(ID As Long)
    GetEmployeeName = DLookup("FullName", EmployeeDB, "[ID]=" & ID)
End Function


'------------------------------------------------------------
' GetEmployeeLogin
'
'------------------------------------------------------------
Private Function GetEmployeeLogin(ID As Long)
    GetEmployeeLogin = DLookup("Login", EmployeeDB, "[ID]=" & ID)
End Function


'------------------------------------------------------------
' GetEmployeePassword
'
'------------------------------------------------------------
Private Function GetEmployeePassword(ID As Long)
    GetEmployeePassword = DLookup("Password", EmployeeDB, "[ID]=" & ID)
End Function


'------------------------------------------------------------
' GetEmployeeRole
'
'------------------------------------------------------------
Private Function GetEmployeeRole(ID As Long)
    GetEmployeeRole = DLookup("Role", EmployeeDB, "[ID]=" & ID)
End Function


'------------------------------------------------------------
' GetEmployeeRole
'
'------------------------------------------------------------
Private Function GetEmployeeDefaultCategory(ID As Long)
    GetEmployeeDefaultCategory = DLookup("DefaultCategory", EmployeeDB, "[ID]=" & ID)
End Function


'------------------------------------------------------------
' SetEmployeeVersion
'
'------------------------------------------------------------
Private Function SetEmployeeVersion(ID As Long)
    Dim rst As DAO.Recordset

    open_db
    Set rst = db.OpenRecordset("SELECT * FROM " & EmployeeDB & " WHERE [ID]=" & ID)

    ' Set version in Employee DB
    rst.Edit
    rst!Version = ReleaseVersion
    rst.Update

    ' Clean up
    rst.Close
    Set rst = Nothing

End Function


'------------------------------------------------------------
' SelectFirstCategory
'
'------------------------------------------------------------
Public Function SelectFirstCategory() As String
    Dim rst As DAO.Recordset

    open_db
    Set rst = db.OpenRecordset("SELECT CategoryName FROM " & CategoryQuery)

    If rst.RecordCount > 0 Then
        ' Return first category
        rst.MoveFirst
        SelectFirstCategory = rst!CategoryName
    Else
        ' Return asterisk
        SelectFirstCategory = "*"
    End If

    rst.Close
    Set rst = Nothing
End Function


'------------------------------------------------------------
' CategoryFieldOrder
'
'------------------------------------------------------------
Public Function CategoryFieldOrder(CategoryName As String, fieldName As String) As Long
    Dim rst As DAO.Recordset

    CategoryFieldOrder = -1
    open_db
    Set rst = db.OpenRecordset("SELECT [Field1],[Field2],[Field3],[Field4],[Field5],[Field6]," & _
        "[Field7],[Field8],[Field9],[Field10],[Field11],[Field12] FROM " & CategoryDB & " WHERE [CategoryName]='" & _
        CategoryName & "' AND [User]=" & EmployeeID)

    If rst.RecordCount <> 1 Then
        ' Fall back to default
        Set rst = db.OpenRecordset("SELECT TOP 1 [Field1],[Field2],[Field3],[Field4],[Field5],[Field6],[Field7]," & _
            "[Field8],[Field9],[Field10],[Field11],[Field12] FROM " & CategoryDB & " WHERE [CategoryName]='" & _
            CategoryName & "' AND [User] IS NULL")
    End If

    If rst.RecordCount = 1 Then
        rst.MoveFirst
        For ii = 0 To rst.Fields.count - 1
            If (rst.Fields(ii) <> "") Then
                If (rst.Fields(ii) = fieldName) Then
                    CategoryFieldOrder = ii
                    GoTo ExitNow
                End If
            End If
        Next ii
    Else
        MsgBox "Unable to determine Category Field Order", , "Invalid Category Field Order"
    End If

ExitNow:
    rst.Close
    Set rst = Nothing
End Function


'------------------------------------------------------------
' NewRecordID
'
'------------------------------------------------------------
Public Function NewRecordID(strRecordPrefix As String, lowBound As Long) As String

    Dim rst As DAO.Recordset
    Dim strSql As String
    Dim arrayDim As Boolean
    Dim ind As Integer
    Dim recordNumber As Long
    Dim recordArray() As Long
    Dim sortedArray() As Long

    strSql = "SELECT DISTINCT [RecordID]" & " FROM " & ItemDB & _
        " WHERE [RecordID]" & " LIKE '" & strRecordPrefix & "-*'" & _
        " ORDER BY [RecordID];"

    open_db
    Set rst = db.OpenRecordset(strSql)

    If rst.RecordCount >= 1 Then
        With rst
            rst.MoveFirst
            Do While Not .EOF
                recordNumber = StripRecordID(strRecordPrefix, !RecordID)
                If (lowBound > recordNumber) Then
                    ' Skip numbers below low bound
                    rst.MoveNext
                ElseIf (lowBound = recordNumber) Then
                    ' When low bound record is found, move to next
                    lowBound = lowBound + 1
                    rst.MoveNext
                    If .EOF Then
                        ' Capture next record
                        NewRecordID = strRecordPrefix & "-" & Format(lowBound, "0000")
                        GoTo CleanUp
                    End If
                Else
                    ' No record matches low bound, use it
                    NewRecordID = strRecordPrefix & "-" & Format(lowBound, "0000")
                    GoTo CleanUp
                End If
            Loop
        End With
    Else
        ' No records found, use low bound
        NewRecordID = strRecordPrefix & "-" & Format(lowBound, "0000")
    End If

CleanUp:
    rst.Close
    Set rst = Nothing
    Set db = Nothing

End Function


'------------------------------------------------------------
' FieldExists
'
'------------------------------------------------------------
Public Function FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Dim nLen As Long

    On Error GoTo Failed
    With DBEngine(0)(0).TableDefs(tableName)
        .Fields.Refresh
        nLen = Len(.Fields(fieldName).Name)

        If nLen > 0 Then FieldExists = True

    End With
    Exit Function
Failed:
    If Err.Number = 3265 Then Err.Clear ' Error 3265 : Item not found in this collection.
    FieldExists = False
End Function


'------------------------------------------------------------
' RecalculateCommit
'
'------------------------------------------------------------
Public Sub RecalculateCommit()

    Dim rstCommit, rstInventory As DAO.Recordset
    Dim sqlQuery As String
    Dim qty, count As Long

    ' Open database
    open_db

    ' Open recordsets
    Set rstInventory = db.OpenRecordset(InventoryDB)

    ' Initialize progress meter
    rstInventory.MoveLast
    count = rstInventory.RecordCount
    SysCmd acSysCmdInitMeter, "Recalculating commits...", count

    ' Loop over all items in inventory
    rstInventory.MoveFirst
    For progress_amount = 1 To count
        ' Update the progress meter
         SysCmd acSysCmdUpdateMeter, progress_amount

        ' Open active commit table entries for given item
        sqlQuery = "SELECT QtyCommitted FROM " & CommitDB & _
            " WHERE [Status] = 'A' AND [ItemID] = " & rstInventory!ItemID
        Set rstCommit = db.OpenRecordset(sqlQuery)

        If Not (rstCommit.EOF) Then
            qty = 0

            ' Get sum of active commit entries
            rstCommit.MoveFirst
            Do While Not rstCommit.EOF
                qty = qty + CInt(rstCommit!QtyCommitted)
                rstCommit.MoveNext
            Loop

            ' Compare calculated commits with Inventory table
            If (CInt(rstInventory!Committed) <> qty) Then
                rstInventory.Edit
                rstInventory!Committed = qty
                rstInventory.Update
            End If
        End If
        rstInventory.MoveNext
   Next progress_amount

   ' Remove the progress meter.
   SysCmd acSysCmdRemoveMeter

    ' Clean up
    rstCommit.Close
    rstInventory.Close
    Set rstCommit = Nothing
    Set rstInventory = Nothing
End Sub


'------------------------------------------------------------
' RecalculateOriginalQuantity
'
'------------------------------------------------------------
Public Sub RecalculateOriginalQuantity()

    Dim rstCommit, rstInventory As DAO.Recordset
    Dim sqlQuery As String
    Dim qty, count As Long

    ' Open database
    open_db

    ' Open recordsets
    Set rstInventory = db.OpenRecordset(InventoryDB)

    ' Initialize progress meter
    rstInventory.MoveLast
    count = rstInventory.RecordCount
    SysCmd acSysCmdInitMeter, "Recalculating original quantities...", count

    ' Loop over all items in inventory
    rstInventory.MoveFirst
    For progress_amount = 1 To count
        ' Update the progress meter
         SysCmd acSysCmdUpdateMeter, progress_amount

        ' Open active commit table entries for given item
        sqlQuery = "SELECT QtyCommitted FROM " & CommitDB & _
            " WHERE [Status] = 'C' AND [ItemID] = " & rstInventory!ItemID
        Set rstCommit = db.OpenRecordset(sqlQuery)

        If Not (rstCommit.EOF) Then
            ' Get current quantity
            qty = rstInventory!OnHand

            ' Get sum of active commit entries
            rstCommit.MoveFirst
            Do While Not rstCommit.EOF
                qty = qty + CInt(rstCommit!QtyCommitted)
                rstCommit.MoveNext
            Loop

            ' Compare calculated quantity with Inventory table
            If IsNull(rstInventory!OrigQty) Then
                rstInventory.Edit
                rstInventory!OrigQty = qty
                rstInventory.Update
            ElseIf (CInt(rstInventory!OrigQty) <> qty) Then
                rstInventory.Edit
                rstInventory!OrigQty = qty
                rstInventory.Update
            End If
        ElseIf IsNull(rstInventory!OrigQty) Then
            rstInventory.Edit
            rstInventory!OrigQty = rstInventory!OnHand
            rstInventory.Update
        ElseIf (CInt(rstInventory!OnHand) <> CInt(rstInventory!OrigQty)) Then
            rstInventory.Edit
            rstInventory!OrigQty = rstInventory!OnHand
            rstInventory.Update
        End If
        rstInventory.MoveNext
   Next progress_amount

   ' Remove the progress meter.
   SysCmd acSysCmdRemoveMeter

    ' Clean up
    rstCommit.Close
    rstInventory.Close
    Set rstCommit = Nothing
    Set rstInventory = Nothing
End Sub


'------------------------------------------------------------
' ReclaimRecordIDs
'
'------------------------------------------------------------
Public Sub ReclaimRecordIDs()

    Dim rstCommit, rstInventory, rstItem As DAO.Recordset
    Dim invQuery, comQuery As String
    Dim count As Long
    Dim pregress_amount As Integer

    ' Open database
    open_db

    ' Open item recordset
    Set rstItem = db.OpenRecordset("SELECT * FROM " & ItemDB & " ORDER BY ID;")

    ' Initialize progress meter
    rstItem.MoveLast
    count = rstItem.RecordCount
    SysCmd acSysCmdInitMeter, "Reclaiming unused Record IDs...", count

    ' Loop over all items
    rstItem.MoveFirst
    For progress_amount = 1 To count
        ' Update the progress meter
         SysCmd acSysCmdUpdateMeter, progress_amount

        ' Check if Record ID has already been reclaimed
        If Not (rstItem!RecordID = "---") Then

            ' Get commit record
            comQuery = "SELECT TOP 1 * FROM " & CommitDB & " WHERE [ItemID] = " & rstItem!ID
            Set rstCommit = db.OpenRecordset(comQuery)

            If (rstCommit.RecordCount = 0) Then

                ' Get inventory record
                invQuery = "SELECT TOP 1 * FROM " & InventoryDB & " WHERE [ItemID] = " & rstItem!ID
                Set rstInventory = db.OpenRecordset(invQuery)

                If (rstInventory.RecordCount = 0) Then
                    ' Remove record if no inventory or commit entry exists
                    rstItem.Delete
                ElseIf (rstInventory!OnHand = 0) Then
                    ' Set RecordID to "---" if no quantity remains
                    rstItem.Edit
                    rstItem!RecordID = "---"
                    rstItem.Update
                End If
            End If
        End If

        ' Go to the next record
         rstItem.MoveNext
   Next progress_amount

   ' Remove the progress meter
   SysCmd acSysCmdRemoveMeter

    ' Clean up
    rstItem.Close
    rstCommit.Close
    rstInventory.Close
    Set rstCommit = Nothing
    Set rstInventory = Nothing
    Set rstItem = Nothing
End Sub


'------------------------------------------------------------
' StripRecordID
' Removes prefix and (suffix) from Record ID, returns Long
'------------------------------------------------------------
Public Function StripRecordID(strRecordPrefix As String, strRecordID As String) As Long

    Dim regEx1 As New RegExp
    Dim regEx2 As New RegExp
    Dim regexReplace As String

    regEx1.Pattern = "^" & strRecordPrefix & "-"
    regEx2.Pattern = "[^0-9]$"
    regEx1.IgnoreCase = True
    regEx2.IgnoreCase = True
    regexReplace = ""


    StripRecordID = CLng(regEx2.Replace(regEx1.Replace(strRecordID, regexReplace), regexReplace))

End Function


'------------------------------------------------------------
' IsValidCategory
' Check if given string is valid category
'------------------------------------------------------------
Public Function IsValidCategory(Category As String) As Boolean
    On Error GoTo ErrHandler
    If (Category = DLookup("CategoryName", CategoryQuery, "[CategoryName]='" & Category & "'")) Then
        IsValidCategory = True
    Else
        IsValidCategory = False
    End If
    Exit Function
ErrHandler:
    IsValidCategory = False
    Exit Function
End Function


'------------------------------------------------------------
' IsValidProduct
' Check if given string is valid product
'------------------------------------------------------------
Public Function IsValidProduct(Product As String) As Boolean
    On Error GoTo ErrHandler
    If (Product = DLookup("ProductName", ProductDB, "[ProductName]='" & Product & "'")) Then
        IsValidProduct = True
    Else
        IsValidProduct = False
    End If
    Exit Function
ErrHandler:
    IsValidProduct = False
    Exit Function
End Function


'------------------------------------------------------------
' IsValidRecordID
' Check if given string is valid Record ID
'------------------------------------------------------------
Public Function IsValidRecordID(RecordID As String) As Boolean
    On Error GoTo ErrHandler
    Dim regEx As New RegExp

    regEx.IgnoreCase = False
    regEx.Multiline = False
    regEx.Pattern = "^[A-Z]{1,4}-\d{4}$"
    If (regEx.Test(RecordID)) Then
        IsValidRecordID = True
    Else
        IsValidRecordID = False
    End If
    Exit Function
ErrHandler:
    IsValidRecordID = False
    Exit Function
End Function


'------------------------------------------------------------
' RecordExists
' Check if given record exists in field in given table
'------------------------------------------------------------
Public Function RecordExists(Table As String, Field As String, Record As String) As Boolean
    On Error GoTo ErrHandler
    If (Record = DLookup(Field, Table, "[" & Field & "]='" & Record & "'")) Then
        RecordExists = True
    Else
        RecordExists = False
    End If
    Exit Function
ErrHandler:
    RecordExists = False
    Exit Function
End Function


'------------------------------------------------------------
' HasParent
' Check if given form has a parent form
'------------------------------------------------------------
Public Function HasParent(FormRef As Form) As Boolean
On Error GoTo ErrHandler
  HasParent = TypeName(FormRef.Parent.Name) = "String"
  Exit Function
ErrHandler:
End Function


'------------------------------------------------------------
' GetCategoryID
'
'------------------------------------------------------------
Public Function GetCategoryID(Name As String) As Long
    On Error GoTo ErrHandler
    Dim ID As Long
    ID = DLookup("CategoryID", CategoryQuery, "[CategoryName]='" & Name & "'")
    If IsNull(ID) Then
        GetCategoryID = 0
    Else
        GetCategoryID = ID
    End If
    Exit Function
ErrHandler:
    GetCategoryID = 0
    Exit Function
End Function


'------------------------------------------------------------
' FieldAvailableSet
'
'------------------------------------------------------------
Public Sub FieldAvailableSet(ByRef formCntrl As Control)
    formCntrl.BackShade = 100 ' Full brightness
    formCntrl.BackThemeColorIndex = 1 ' Background 1
End Sub


'------------------------------------------------------------
' FieldAvailableRemove
'
'------------------------------------------------------------
Public Sub FieldAvailableRemove(ByRef formCntrl As Control)
    formCntrl.BackShade = 95 ' Darker 5%
    formCntrl.BackThemeColorIndex = 1 ' Background 1
End Sub


'------------------------------------------------------------
' FieldErrorSet
'
'------------------------------------------------------------
Public Sub FieldErrorSet(ByRef formCntrl As Control)
    formCntrl.BorderColor = 2366701 ' Red
    formCntrl.BackColor = 13421823 ' Light Red
End Sub


'------------------------------------------------------------
' FieldErrorClear
'
'------------------------------------------------------------
Public Sub FieldErrorClear(ByRef formCntrl As Control)
    formCntrl.BorderThemeColorIndex = 1 ' Background 1
    formCntrl.BorderShade = 65 ' Darker 35%
    formCntrl.BackThemeColorIndex = 1 ' Background 1
End Sub


'------------------------------------------------------------
' Commit_Cancel
'
'------------------------------------------------------------
Public Function Commit_Cancel(rst As DAO.Recordset) As Boolean
    On Error GoTo Commit_Cancel_Err

    rst.MoveFirst
    Do While Not rst.EOF
        If (rst!Status = "A") Then
            ' Check existing available quantity
            If (rst!OnHand < 0) Or (rst!OnHand < rst!Committed) Then
                GoTo Quantity_Err
            ElseIf (rst!Committed <= 0) Or (rst!QtyCommitted <= 0) Then
                GoTo Quantity_Err
            End If

            With rst
                .Edit
                !DateCancel = Now()
                !Status = "X"
                !OperatorCancel = EmployeeLogin
                !Committed = !Committed - !QtyCommitted
                !LastOper = EmployeeLogin
                !LastDate = Now()
                .Update
            End With
        Else
            GoTo Commit_Cancel_Err
        End If
        rst.MoveNext
    Loop
    Commit_Cancel = True
    rst.MoveFirst

Commit_Cancel_Exit:
    Exit Function

Commit_Cancel_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Commit_Cancel = False
    GoTo Commit_Cancel_Exit

Status_Err:
    MsgBox "Error: " & vbCrLf & "Cannot Delete Commit " & rst!ID & vbCrLf & _
        "Commit must be active", , "Status Error"
    Commit_Cancel = False
    GoTo Commit_Cancel_Exit

Quantity_Err:
    MsgBox "Error: " & vbCrLf & "Cannot Delete Commit " & rst!ID & vbCrLf & _
        "Commit Quantity or Item Available Quantity are invalid", , "Quantity Error"
    Commit_Cancel = False
    GoTo Commit_Cancel_Exit

End Function


'------------------------------------------------------------
' Commit_Complete
'
'------------------------------------------------------------
Public Function Commit_Complete(ByRef rst As DAO.Recordset) As Boolean
    On Error GoTo Commit_Complete_Err

    rst.MoveFirst
    Do While Not rst.EOF
        If (rst!Status = "A") Then
            ' Check existing available quantity
            If (rst!OnHand < 0) Or ((rst!OnHand + rst!OnOrder) < rst!Committed) Then
                GoTo Quantity_Err
            ' Check commit quantities are valid
            ElseIf (rst!Committed <= 0) Or (rst!QtyCommitted <= 0) Then
                GoTo Quantity_Err
            ' Check if current commit quantity is valid
            ElseIf ((rst!Committed - rst!QtyCommitted) < 0) Or ((rst!OnHand - rst!QtyCommitted) < 0) Then
                GoTo Quantity_Err
            End If

            With rst
                .Edit
                !DateComplete = Now()
                !Status = "C"
                !OperatorComplete = EmployeeLogin
                !Committed = !Committed - !QtyCommitted
                !OnHand = !OnHand - !QtyCommitted
                !LastOper = EmployeeLogin
                !LastDate = Now()
                .Update
            End With
        Else
            GoTo Status_Err
        End If
        rst.MoveNext
    Loop
    Commit_Complete = True
    rst.MoveFirst

Commit_Complete_Exit:
    Exit Function

Commit_Complete_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Commit_Complete = False
    GoTo Commit_Complete_Exit

Status_Err:
    MsgBox "Error: " & vbCrLf & "Cannot Complete Commit " & rst!ID & vbCrLf & _
        "Commit must be active", , "Status Error"
    Commit_Complete = False
    GoTo Commit_Complete_Exit

Quantity_Err:
    MsgBox "Error: " & vbCrLf & "Cannot Complete Commit " & rst!ID & vbCrLf & _
        "Commit Quantity or Item Available Quantity are invalid", , "Quantity Error"
    Commit_Complete = False
    GoTo Commit_Complete_Exit

End Function


'------------------------------------------------------------
' OperationEntry
'
'------------------------------------------------------------
Public Sub OperationEntry(ItemID As Long, Operation As String, Description As String)
    On Error GoTo OperationEntry_Err

    Dim rst As Recordset

    Set rst = db.OpenRecordset("SELECT * FROM " & OperationDB)

    With rst
        .AddNew
        !ItemID = ItemID
        !Operator = EmployeeLogin
        !Date = Now()
        !Operation = Operation
        !Description = Description
        .Update
    End With

OperationEntry_Exit:
    rst.Close
    Set rst = Nothing
    Exit Sub

OperationEntry_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume OperationEntry_Exit
End Sub


'------------------------------------------------------------
' RecordIDReserve
'
'------------------------------------------------------------
Public Function RecordIDReserve(ID As String, Category As String) As Boolean
    On Error GoTo RecordIDReserve_Err

    Dim rst As Recordset

    open_db
    Set rst = db.OpenRecordset(ItemDB)

    ' Check to make sure that Record ID doesn't exist
    If IsNull(DLookup("RecordID", ItemDB, "[RecordID]='" & ID & "'")) Then
        ' Proceed to reserve Record ID
        With rst
            .AddNew
            !RecordID = ID
            !Category = Category
            !Vendor = "RESERVED"
            !CreateDate = Now()
            !CreateOper = EmployeeLogin
            .Update
        End With
        RecordIDReserve = True
    Else
        MsgBox "Error: " & vbCrLf & "Cannot Reserve Record ID " & vbCrLf & _
            "Record ID already exists", , "Cannot Reserve"
        RecordIDReserve = False
        GoTo RecordIDReserve_Exit
    End If

RecordIDReserve_Exit:
    rst.Close
    Set rst = Nothing
    Exit Function

RecordIDReserve_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume RecordIDReserve_Exit
End Function
