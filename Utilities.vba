Option Compare Database

'------------------------------------------------------------
' American Surplus Inventory Database
' Author: Nathanael Greene
' Current Revision: 2.6.0
' Revision Date: 07/20/2016
'
' Revision History:
'   2.6.0:  New: Reorganized program, removed separation between Sales / Production
'           New (ItemNew, ItemEdit): Added validation for Description length
'   2.5.3:  New (ProductionInventory) Enabled Record ID search across categories
'   2.5.2:  Bug fix: (Utilities) IsFileName wrong variable 'path' -> 'strFile'
'   2.5.1:  Bug fix: (*) Set SetScreenSize subroutine to Public
'   2.5.0:  New (ItemNew, Utilities) Disable ItemNew fields based on product
'           New (ProductionCommit, SalesCommit) '* All' based on filtered view
'           New (ProductionCommit, SalesCommit) Disable '* All' buttons based on commit view
'           Bug fix (***Inventory***) Fix SetScreensize
'           Bug fix (Utilities) Fix Reclaim Record IDs
'   2.4.2:  New (Utilities, ***Inventory***) Remember last category
'           New (Utilities, ProductionCommit) Remember commit state
'           Bug fix (Utilities) Recalculate Commits
'           Bug fix (***Inventory***) Remove SetScreensize from Form_Open
'   2.4.1:  New (OrderCommitManage) Add confirmations for all buttons
'           New (Update) Update form details new features
'           Bug fix (OrderCommitManage, ProductionCommit) Complete not visible
'           Bug fix (SalesInventory, ProductionInventory) Force sort by Focus
'   2.4.0:  New (OrderCommitManage): Add reactivate command
'           New (ProductionCommit): Add reactivate command
'           New (InventorySearchSubForm): Add Focus field and set order by
'           New (ProductionCommit): Add reactivate command
'           New: Added AdminInventoryPrice form to set Prices in bulk
'           New (ItemDetail): cmdCommit to prompt if OnOrder
'           New (ItemNew): Removed restriction of reserved Record IDs to user
'           New (OrderCommitManage): Add adjust location command
'           New (Login): Added Exit button
'           Bug fix (OrderCommitManage) cmdSave QtyCommitted
'           Bug fix (ItemInventoryManage) OnOrder not considered with commits
'           Bug fix (Utilities, ItemNew, ItemEdit) capitalize RecordID prefix
'           Bug fix (Utilities, ItemNew, ItemEdit) handle missing images
'   2.3.3:  Bug fix: (ItemDetail) Missing photo causes load fail
'           Bug fix: Changed Description fields to Plain Text
'   2.3.2:  Bug fix: Missing references from ASIdev
'   2.3.1:  Bug fix (OrderCommitManage): need to use Query instead
'               of individual tables (Item, Inv, Commit)
'   2.3.0:  New: Updated ItemDetail with new design and Image field
'           New: Updated ItemNew with new design and Image field
'           New: Updated ItemEdit with new design and Image field
'           New: Added GetRecordPrefix, SendMessage to Utilities
'   2.2.3:  Bug fix (SalesInventorySplit) - Show Available only
'               (SalesInventory) - Switch OnHand to Available
'           New: Add Last User and Date fields to OrderCommitManage
'   2.2.2:  Bug fix (ProductionInventory) - Fix RecordID Filter
'               (ItemNew) - Vendor and Manufacturer field limits
'   2.2.1:  Bug fix - re-link to database backend
'   2.2.0:  New: Reorganized program into Sales, Production, Admin
'               Changed Print Range to print whole screen
'               Incorporated SW Version recording
'               Eliminated Subcategory field in Items
'   2.1.3:  Bug fix (Commit_Complete) - not allowed to complete
'               when Committed > OnHand (+ Onorder added)
'   2.1.2:  Bug fix (NewRecordID) - not finding next record
'               after full list
'           Bug fix (CategoriesDS) - Field12 not being updated in form
'           Bug fix (EmployeesDS) - Roles selector not working
'   2.1.1:  Bug fix (Commit_Complete)
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
'   2.0.4:  Bug fix (Utilities) - Commit_Cancel & Commit_Complete
'               missing recordset reference
'   2.0.3:  Bug fixes (Commit_Cancel, ItemNew) - cancel had
'               wrong sign (committing more)
'               Added Record ID search to Inventory Manage
'               Added scroll bar to ItemNew
'   2.0.2:  Bug fixes (ItemEdit, ItemNew) - invalid null in
'               numeric inputs
'   2.0.1:  Bug fixes (ItemEdit, ItemNew, ItemInventoryManage,
'               Main, CategoriesEdit)
'   2.0.0:  Initial Release replaces legacy database
'           Complete GUI overhaul
'           Introduction of product-based structure
'           Add commit management for all users
'           Add Generate Record ID tools
'------------------------------------------------------------

'------------------------------------------------------------
' Global constants
'
'------------------------------------------------------------
Public Const ReleaseVersion As String = "2.6.0"
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
Public Const InventoryForm As String = "InventoryForm"
Public Const InventoryManageForm As String = "InventoryManage"
Public Const InventoryManageSplitForm As String = "InventoryManageSplit"
Public Const InventoryOrderForm As String = "InventoryOrder"
Public Const ItemDetailForm As String = "ItemDetail"
Public Const ItemEditForm As String = "ItemEdit"
Public Const ItemInventoryManageForm As String = "ItemInventoryManage"
Public Const PrintRangeForm As String = "PrintRange"
Public Const CategoriesEditForm As String = "CategoriesEdit"
Public Const GenerateRecordIDForm As String = "GenerateRecordID"
Public Const OrderCommitManageForm As String = "OrderCommitManage"
Public Const UpdateForm As String = "Update"

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
Public CompanyVersion As String
Private pvEmployeeID As Long
Public EmployeeName As String
Public EmployeeLogin As String
Public EmployeePassword As String
Public EmployeeRole As String
Public EmployeeCategory As String
Public EmployeeVersion As String
Public ValidLogin As Boolean
Public ScreenWidth As Long
Public searchCategory As String
Public searchCategoryBottom As String
Public commitSelectStatus As String
Public ProductFields() As String

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
    EmployeeVersion = GetEmployeeVersion(Value)
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
    CompanyVersion = Nz(DLookup("Build", SettingsDB, "[Company]='" & Company & "'"), "")

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

    ' Check for update
    VersionCheck

    ' Reset employee use parameters
    searchCategory = ""
    searchCategoryBottom = ""
    commitSelectStatus = ""

    DoCmd.OpenForm MainForm
    Forms(MainForm)!lblCurrentEmployeeName.Caption = "Hello, " & EmployeeName
    Forms(MainForm)!lblVersion.Caption = "Version " & ReleaseVersion
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
' GetEmployeeDefaultCategory
'
'------------------------------------------------------------
Private Function GetEmployeeDefaultCategory(ID As Long)
    GetEmployeeDefaultCategory = DLookup("DefaultCategory", EmployeeDB, "[ID]=" & ID)
End Function


'------------------------------------------------------------
' GetEmployeeVersion
'
'------------------------------------------------------------
Private Function GetEmployeeVersion(ID As Long)
    GetEmployeeVersion = DLookup("Version", EmployeeDB, "[ID]=" & ID)
End Function


'------------------------------------------------------------
' VersionCheck
'
'------------------------------------------------------------
Public Sub VersionCheck()
On Error GoTo VersionCheck_Err

    Dim updateNow As Integer
    Dim upgradeLoc As String

    upgradeLoc = "F:\data\ACCESS"

    ' Running updated version
    If (ReleaseVersion <> CompanyVersion) Then
        updateNow = MsgBox("Running outdated version" & vbCrLf & _
            "Current Version is: " & CompanyVersion & vbCrLf & vbCrLf & _
            "Would you like to upgrade now?", vbYesNo, "Update Version")

        If updateNow = vbYes Then
            DoCmd.Close acForm, "Main"
            Shell "C:\WINDOWS\explorer.exe """ & upgradeLoc & "", vbNormalFocus
            DoCmd.Quit
        End If

    ElseIf (EmployeeVersion <> ReleaseVersion) Then
        DoCmd.OpenForm UpdateForm, acNormal, "", "", , acDialog
    End If
    SetEmployeeVersion (EmployeeID)

VersionCheck_Exit:
    Exit Sub

VersionCheck_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume VersionCheck_Exit

End Sub


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
        "[Field7],[Field8],[Field9],[Field10],[Field11],[Field12],[Field13],[Field14],[Field15]" & _
        " FROM " & CategoryDB & " WHERE [CategoryName]='" & CategoryName & "' AND [User]=" & EmployeeID)

    If rst.RecordCount <> 1 Then
        ' Fall back to default
        Set rst = db.OpenRecordset("SELECT TOP 1 [Field1],[Field2],[Field3],[Field4],[Field5],[Field6],[Field7]," & _
            "[Field8],[Field9],[Field10],[Field11],[Field12],[Field13],[Field14],[Field15] FROM " & _
            CategoryDB & " WHERE [CategoryName]='" & CategoryName & "' AND [User] IS NULL")
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
' ProductFieldVisibility
'
'------------------------------------------------------------
Public Function ProductFieldVisibility(ProductName As String, fieldName As String) As Boolean
    Dim rst As DAO.Recordset
    Dim ii, fCount As Integer

    ProductFieldVisibility = True
    open_db

    If IsVarArrayEmpty(ProductFields) Then
        Set rst = db.OpenRecordset("SELECT TOP 1 * FROM " & ProductDB)

        ' Define number of Product Fields
        fCount = 0
        For ii = 0 To rst.Fields.count - 1
            If (rst.Fields(ii).Type = 1) Then
                fCount = fCount + 1
            End If
        Next ii
        ReDim ProductFields(fCount - 1)

        ' Fill array of Product Fields
        fCount = 0
        For ii = 0 To rst.Fields.count - 1
            If (rst.Fields(ii).Type = 1) Then
                ProductFields(fCount) = rst.Fields(ii).Name
                fCount = fCount + 1
            End If
        Next ii

        rst.Close
        Set rst = Nothing
    End If

    Set rst = db.OpenRecordset("SELECT TOP 1 * FROM " & ProductDB & " WHERE [ProductName]='" & ProductName & "'")
    For ii = 0 To UBound(ProductFields)
        If (ProductFields(ii) = fieldName) Then
            If (rst(fieldName) = True) Then
                ProductFieldVisibility = True
                Exit For
            ElseIf (rst(fieldName) = False) Then
                ProductFieldVisibility = False
                Exit For
            End If
        End If
    Next ii
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
                        NewRecordID = UCase(strRecordPrefix) & "-" & Format(lowBound, "0000")
                        GoTo CleanUp
                    End If
                Else
                    ' No record matches low bound, use it
                    NewRecordID = UCase(strRecordPrefix) & "-" & Format(lowBound, "0000")
                    GoTo CleanUp
                End If
            Loop
        End With
    Else
        ' No records found, use low bound
        NewRecordID = UCase(strRecordPrefix) & "-" & Format(lowBound, "0000")
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
        Else
            ' No active commits exist
            rstInventory.Edit
            rstInventory!Committed = 0
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
    Dim count, reclaim As Long
    Dim pregress_amount As Integer

    ' Open database
    open_db
    reclaim = 0

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

        ' Check if Record ID has already been reclaimed or is reserved
        If Not (rstItem!RecordID = "---") And Not (rstItem!Vendor = "RESERVED") Then

            ' Get commit record
            comQuery = "SELECT TOP 1 * FROM " & CommitDB & " WHERE [ItemID] = " & rstItem!ID & ";"
            Set rstCommit = db.OpenRecordset(comQuery)

            ' Get inventory record
            invQuery = "SELECT TOP 1 * FROM " & InventoryDB & " WHERE [ItemId] = " & rstItem!ID & ";"
            Set rstInventory = db.OpenRecordset(invQuery)

            If (rstInventory.RecordCount = 0) And (rstCommit.RecordCount = 0) Then
                ' Remove record if no inventory or commit entry exists
                rstItem.Delete
                reclaim = reclaim + 1
            ElseIf (rstInventory.RecordCount = 0) Then
                ' Set RecordID to "---" if no inventory record remains
                rstItem.Edit
                rstItem!RecordID = "---"
                rstItem.Update
                reclaim = reclaim + 1
            ElseIf (rstInventory!OnHand = 0) Then
                ' Set RecordID to "---" if no quantity remains
                rstItem.Edit
                rstItem!RecordID = "---"
                rstItem.Update
                reclaim = reclaim + 1
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

    MsgBox reclaim & " Record IDs Reclaimed"
End Sub


'------------------------------------------------------------
' StripRecordID
' Removes prefix and (suffix) from Record ID, returns Long
'------------------------------------------------------------
Public Function StripRecordID(strRecordPrefix As String, strRecordID As String) As Long

    Dim regEx1 As New RegExp
    Dim regEx2 As New RegExp
    Dim regexReplace As String

    regEx1.Pattern = "^" & UCase(strRecordPrefix) & "-"
    regEx2.Pattern = "[^0-9]$"
    regEx1.IgnoreCase = True
    regEx2.IgnoreCase = True
    regexReplace = ""


    StripRecordID = CLng(regEx2.Replace(regEx1.Replace(strRecordID, regexReplace), regexReplace))

End Function


'------------------------------------------------------------
' GetRecordPrefix
' Retrieves prefix from Record ID
'------------------------------------------------------------
Public Function GetRecordPrefix(strRecordID As String) As String

    Dim regEx As New RegExp
    Dim regEx2 As New RegExp
    Dim regexReplace As String

    regEx.Pattern = "-.*$"
    regEx.IgnoreCase = True
    regexReplace = ""

    GetRecordPrefix = UCase(regEx.Replace(strRecordID, regexReplace))

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
            ' Flag invalid inventory quantities
            If (rst!OnHand < 0) Or (rst!OnOrder < 0) Then
                GoTo Quantity_Err
            ' Flag invalid commit quantities
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
    MsgBox "Error: " & vbCrLf & "Cannot Cancel Commit " & rst!ID & vbCrLf & _
        "Commit must be Active", , "Status Error"
    Commit_Cancel = False
    GoTo Commit_Cancel_Exit

Quantity_Err:
    MsgBox "Error: " & vbCrLf & "Cannot Cancel Commit " & rst!ID & vbCrLf & _
        "Commit Quantity or Item Available Quantity are invalid", , "Quantity Error"
    Commit_Cancel = False
    GoTo Commit_Cancel_Exit

End Function


'------------------------------------------------------------
' Commit_Reactivate
'
'------------------------------------------------------------
Public Function Commit_Reactivate(ByRef rst As DAO.Recordset) As Boolean
    On Error GoTo Commit_Reactivate_Err

    rst.MoveFirst
    Do While Not rst.EOF
        If (rst!Status = "A") Then
            GoTo Status_Err
        ElseIf (rst!Status = "C") Then
            With rst
                .Edit
                !DateComplete = Null
                !Status = "A"
                !OperatorComplete = ""
                !Committed = !Committed + !QtyCommitted
                !OnHand = !OnHand + !QtyCommitted
                !LastOper = EmployeeLogin
                !LastDate = Now()
                .Update
            End With
        ElseIf (rst!Status = "X") Then
            With rst
                .Edit
                !DateCancel = Null
                !Status = "A"
                !OperatorCancel = ""
                !Committed = !Committed + !QtyCommitted
                !LastOper = EmployeeLogin
                !LastDate = Now()
                .Update
            End With
        End If
        rst.MoveNext
    Loop
    Commit_Reactivate = True
    rst.MoveFirst

Commit_Reactivate_Exit:
    Exit Function

Commit_Reactivate_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Commit_Reactivate = False
    GoTo Commit_Reactivate_Exit

Status_Err:
    MsgBox "Error: " & vbCrLf & "Cannot Reactivate Commitment " & rst!ID & vbCrLf & _
        "Commit must be not be active", , "Status Error"
    Commit_Reactivate = False
    GoTo Commit_Reactivate_Exit

End Function


'------------------------------------------------------------
' Commit_Complete
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
        "Commit must be Active", , "Status Error"
    Commit_Complete = False
    GoTo Commit_Complete_Exit

Quantity_Err:
    MsgBox "Error: " & vbCrLf & "Cannot Complete Commit " & rst!ID & vbCrLf & _
        "Commit Quantity or Item Available Quantity are invalid", , "Quantity Error"
    Commit_Complete = False
    GoTo Commit_Complete_Exit

End Function


'------------------------------------------------------------
' RecordCheck
' Returns true if given field condition matches all records
' in field in Recordset
'------------------------------------------------------------
Public Function RecordCheck(ByRef rst As DAO.Recordset, Field As String, Condition As String) As Boolean
    On Error Resume Next

    rst.MoveFirst
    Do While Not rst.EOF
        If (rst.Fields(Field) = Condition) Then
            rst.MoveNext
        Else
            RecordCheck = False
            GoTo RecordCheck_Exit
        End If
    Loop
    RecordCheck = True

RecordCheck_Exit:
    Exit Function

End Function


'------------------------------------------------------------
' OperationEntry
'
'------------------------------------------------------------
Public Sub OperationEntry(ItemID As Long, Operation As String, Description As String, _
    Optional Reason As String)
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
        If (Reason <> "") Then
            !Reason = Reason
        End If
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


'------------------------------------------------------------
' SendMessage
'
'------------------------------------------------------------
Sub SendMessage(DisplayMsg As Boolean, _
    Optional strRecipient As String, _
    Optional strSubject As String, _
    Optional AttachmentPath As String)

    Dim objOutlook As Outlook.Application
    Dim objOutlookMsg As Outlook.MailItem
    Dim objOutlookRecip As Outlook.Recipient
    Dim objOutlookAttach As Object

    ' Create the Outlook session
    Set objOutlook = CreateObject("Outlook.Application")

    ' Create the message
    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)

    With objOutlookMsg
        ' Add the To recipient(s) to the message
        If (strRecipient <> "") Then
            Set objOutlookRecip = .Recipients.Add(strRecipient)
            objOutlookRecip.Type = olTo
        End If

        ' Set the Subject, Body, and Importance of the message
        .Subject = strSubject

        ' Add attachments to the message
        If Not IsMissing(AttachmentPath) Then
            Set objOutlookAttach = .Attachments.Add(AttachmentPath)
        End If

        ' Resolve each Recipient's name
        For Each objOutlookRecip In .Recipients
            objOutlookRecip.Resolve
        Next

        ' Should we display the message before sending?
        If DisplayMsg Then
            .Display
        Else
            .Save
        End If
    End With

    Set objOutlookMsg = Nothing
    Set objOutlook = Nothing
    Set objOutlookRecip = Nothing
    Set objOutlookAttach = Nothing

End Sub


'------------------------------------------------------------
' FileExists
'
'------------------------------------------------------------
Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    If IsFileName(strFile) Then
        FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
    Else
        FileExists = False
    End If
End Function


'------------------------------------------------------------
' IsFileName
'
'------------------------------------------------------------
Function IsFileName(ByVal strFile As String) As Boolean
    'Purpose:   Return True if the input string has a file extension
    'Arguments: strFile: File name to look at

    IsFileName = (Len(Right$(strFile, Len(strFile) - InStrRev(strFile, "."))) > 0)
End Function


'------------------------------------------------------------
' IsVarArrayEmpty
'
'------------------------------------------------------------
Function IsVarArrayEmpty(anArray As Variant)
    Dim i As Integer

    On Error Resume Next
        i = UBound(anArray, 1)
    If Err.Number = 0 Then
        IsVarArrayEmpty = False
    Else
        IsVarArrayEmpty = True
    End If

End Function


'------------------------------------------------------------
' ProcedureExists
' Check if procedure exists in form module
'------------------------------------------------------------
Function ProcedureExists(ProcedureForm As Access.Form, _
    ProcedureName As String) As Boolean

    Dim m As Module, p As Integer
    ProcedureExists = True
    On Error Resume Next

    Set m = ProcedureForm.Module
    p = m.ProcBodyLine(ProcedureName, vbext_pk_Proc)
    If Err.Number <> 35 Then
        Exit Function
    End If
    ProcedureExists = False
End Function
