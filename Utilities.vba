Option Compare Database

'------------------------------------------------------------
' American Surplus Inventory Database
' Author: Nathanael Greene
' Current Revision: 2.0
' Revision Date: 09/12/2015
'
' Revision History:
'   2.0:    Initial Release replaces legacy database
'           Complete GUI overhaul
'           Introduction of product-based structure
'           Add Generate Record ID tools
'------------------------------------------------------------

'------------------------------------------------------------
' Global constants
'
'------------------------------------------------------------
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
Public Const CategoryQuery As String = "qryCategoryList"
Public Const CommitQuery As String = "qryItemCommit"
''' Form Names
Public Const MainForm As String = "Main"
Public Const LoginForm As String = "Login"
Public Const InventoryForm As String = "InventoryForm"
Public Const InventorySearchForm As String = "InventorySearch"
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
    Forms(MainForm)!nvbInventory.SetFocus
    Err.Clear

CompleteLogin_Exit:
    Exit Function

CompleteLogin_Err:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume CompleteLogin_Exit

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

    ' Get total number of records
    rstInventory.MoveLast
    count = rstInventory.RecordCount
    rstInventory.MoveFirst

    ' Initialize progress meter
    SysCmd acSysCmdInitMeter, "Recalculating commits...", count

    ' Loop over all items in inventory
    Do While Not rstInventory.EOF
        ' Update progress meter
        SysCmd acSysCmdUpdateMeter, rstInventory.RecordCount

        ' Open active commit table entries for given item
        sqlQuery = "SELECT QtyCommitted FROM " & CommitDB & _
            " WHERE [Status] = 'A' AND [ItemID] = " & rstInventory!ItemId
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
    Loop

   ' Remove the progress meter.
   SysCmd acSysCmdRemoveMeter

    ' Clean up
    rstCommit.Close
    rstInventory.Close
    Set rstCommit = Nothing
    Set rstInventory = Nothing
End Sub


'------------------------------------------------------------
' NewRecordID
'
'------------------------------------------------------------
Public Function NewRecordID(strRecordPrefix As String, lowBound As Long) As String

    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim arrayDim As Boolean
    Dim ind As Integer
    Dim recordArray() As Long
    Dim sortedArray() As Long

    strSQL = "SELECT [RecordID]" & _
        " FROM " & ItemDB & _
        " WHERE [RecordID]" & _
        " LIKE '" & strRecordPrefix & "-*'" & _
        " ORDER BY [RecordID]"
    open_db
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rst.RecordCount > 1 Then
        With rst
            rst.MoveFirst
            Do While Not .EOF
                If arrayDim = True Then
                    ReDim Preserve recordArray(1 To UBound(recordArray) + 1) As Long
                Else
                    ReDim recordArray(1 To 1) As Long
                    arrayDim = True
                End If
                recordArray(UBound(recordArray)) = StripRecordID(strRecordPrefix, !RecordID)
                rst.MoveNext
            Loop
        End With
    ElseIf rst.RecordCount = 1 Then
        If StripRecordID(strRecordPrefix, rst!RecordID) = lowBound Then
            lowBound = lowBound + 1
        End If
        NewRecordID = strRecordPrefix & "-" & Format(lowBound, "0000")
    ElseIf rst.RecordCount = 0 Then
        NewRecordID = strRecordPrefix & "-" & Format(lowBound, "0000")
    End If

    If (IsInitialized(recordArray)) Then
        sortedArray = RemoveDups(recordArray)
        For ind = lowBound To UBound(sortedArray)
            If (sortedArray(ind) >= lowBound) Then
                If (lowBound <> sortedArray(ind)) Then
                    NewRecordID = strRecordPrefix & "-" & Format(lowBound, "0000")
                    Exit For
                Else
                    lowBound = lowBound + 1
                End If
            Else
                Exit For
            End If
        Next ind
    Else
        NewRecordID = strRecordPrefix & "-" & Format(lowBound, "0000")
    End If

CleanUp:
    rst.Close
    Set rst = Nothing
    Set db = Nothing

End Function


'------------------------------------------------------------
' BubbleSrt
' Sort array
'------------------------------------------------------------
Private Function BubbleSrt(ArrayIn() As Long, Ascending As Boolean) As Long()

    Dim SrtTemp As Long
    Dim i As Long
    Dim j As Long


    If Ascending = True Then
        For i = LBound(ArrayIn) To UBound(ArrayIn)
             For j = i + 1 To UBound(ArrayIn)
                 If ArrayIn(i) > ArrayIn(j) Then
                     SrtTemp = ArrayIn(j)
                     ArrayIn(j) = ArrayIn(i)
                     ArrayIn(i) = SrtTemp
                 End If
             Next j
         Next i
    Else
        For i = LBound(ArrayIn) To UBound(ArrayIn)
             For j = i + 1 To UBound(ArrayIn)
                 If ArrayIn(i) < ArrayIn(j) Then
                     SrtTemp = ArrayIn(j)
                     ArrayIn(j) = ArrayIn(i)
                     ArrayIn(i) = SrtTemp
                 End If
             Next j
         Next i
    End If

    BubbleSrt = RemoveDups(ArrayIn)

End Function


'------------------------------------------------------------
' RemoveDups
' Remove duplicates from array
'------------------------------------------------------------
Private Function RemoveDups(intArray() As Long) As Long()

    Dim old_i As Integer
    Dim last_i As Integer
    Dim low_i, high_i As Integer
    Dim result() As Long

    ' Get the lower and upper bounds
    low_i = LBound(intArray)
    high_i = UBound(intArray)

    ' Make the result array.
    ReDim result(low_i To high_i)

    ' Copy the first item into the result array.
    result(low_i) = intArray(low_i)

    ' Copy the other items
    last_i = low_i
    For old_i = (low_i + 1) To high_i
        If result(last_i) <> intArray(old_i) Then
            ' No duplicate found
            last_i = last_i + 1
            result(last_i) = intArray(old_i)
        End If
    Next old_i

    ' Remove unused entries from the result array.
    ReDim Preserve result(low_i To last_i)

    ' Return the result array.
    RemoveDups = result
End Function


'------------------------------------------------------------
' IsInitialized
' Check if array of Long integers is initialized
'------------------------------------------------------------
Private Function IsInitialized(arr() As Long) As Boolean
    On Error GoTo ErrHandler
    Dim nUbound As Long
    nUbound = UBound(arr)
    IsInitialized = True
    Exit Function
ErrHandler:
    IsInitialized = False
    Exit Function
End Function


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
            With rst
                .Edit
                !DateCancel = Now()
                !Status = "X"
                !OperatorCancel = EmployeeLogin
                !Committed = !Committed + !QtyCommitted
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
    Resume Commit_Cancel_Exit
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
            GoTo Commit_Complete_Err
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
    Resume Commit_Complete_Exit
End Function


'------------------------------------------------------------
' OperationEntry
'
'------------------------------------------------------------
Public Sub OperationEntry(ItemId As Long, Operation As String, Description As String)
    On Error GoTo OperationEntry_Err

    Dim rst As Recordset

    Set rst = db.OpenRecordset("SELECT * FROM " & OperationDB)

    With rst
        .AddNew
        !ItemId = ItemId
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
