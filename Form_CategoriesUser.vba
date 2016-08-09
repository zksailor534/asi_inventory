Option Compare Database
Option Explicit

Private rstCategory As DAO.Recordset

'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    Dim DefaultCategory As String

    ' Set selected category
    DefaultCategory = "Acc / Rack"

    SetScreenSize

    ' Open the recordset
    open_db
    Set rstCategory = db.OpenRecordset("SELECT TOP 1 * FROM " & CategoryDB & " WHERE [User] = " & EmployeeID)

    ' Hide User and ID Columns
    Me.sbfrmDS!ID.ColumnHidden = True
    Me.sbfrmDS!User.ColumnHidden = True

    ' Filter only on user custom field orders
    Me.sbfrmDS.Form.Filter = "[User] = " & EmployeeID
    Me.sbfrmDS.Form.FilterOn = True

    If rstCategory.RecordCount > 0 Then
        FillFields
    End If

    ' Fill Field source combo boxes
    FieldSources

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
    Set rstCategory = Nothing
End Sub


'------------------------------------------------------------
' Category_AfterUpdate
'
'------------------------------------------------------------
Private Sub Category_AfterUpdate()
    ' Select user customization if available
    Set rstCategory = db.OpenRecordset("SELECT * FROM " & CategoryDB & _
        " WHERE [CategoryName] = '" & Category & "' AND [User] = " & EmployeeID)

    ' Fall back to default category
    If rstCategory.RecordCount = 0 Then
        Set rstCategory = db.OpenRecordset("SELECT * FROM " & CategoryDB & _
            " WHERE [CategoryName] = '" & Category & "' AND [User] IS NULL")
        FillFields
        AddNewButton.Enabled = True
        SaveButton.Enabled = False
    Else
        FillFields
        AddNewButton.Enabled = False
        SaveButton.Enabled = True
    End If
End Sub


'------------------------------------------------------------
' SaveButton_Click
'
'------------------------------------------------------------
Private Sub AddNewButton_Click()
    Dim saveCategory As Integer
    saveCategory = MsgBox("Do you want to save this Category field order?", vbYesNo, "Save Customization")
    If saveCategory = vbYes Then
        NewItem
    End If

    ' Update subform
    Me.sbfrmDS.Requery
End Sub


'------------------------------------------------------------
' SaveButton_Click
'
'------------------------------------------------------------
Private Sub SaveButton_Click()
    Dim saveCategory As Integer
    saveCategory = MsgBox("Do you want to save changes to this Category field order?", vbYesNo, "Save Customization")
    If saveCategory = vbYes Then
        SaveItem
    End If

    ' Update subform
    Me.sbfrmDS.Requery
End Sub


'------------------------------------------------------------
' DeleteButton_Click
'
'------------------------------------------------------------
Private Sub DeleteButton_Click()
    Dim delCategory As Integer
    delCategory = MsgBox("Do you want to delete this Category field order?", vbYesNo, "Delete Customization")
    If delCategory = vbYes Then
        DeleteItem
    End If
End Sub


'------------------------------------------------------------
' FillFields
'
'------------------------------------------------------------
Private Sub FillFields()

    If rstCategory.RecordCount <> 0 Then
        Category = Nz(rstCategory!CategoryName, "")
        Prefix = Nz(rstCategory!Prefix, "")
        Field1 = Nz(rstCategory!Field1, "")
        Field2 = Nz(rstCategory!Field2, "")
        Field3 = Nz(rstCategory!Field3, "")
        Field4 = Nz(rstCategory!Field4, "")
        Field5 = Nz(rstCategory!Field5, "")
        Field6 = Nz(rstCategory!Field6, "")
        Field7 = Nz(rstCategory!Field7, "")
        Field8 = Nz(rstCategory!Field8, "")
        Field9 = Nz(rstCategory!Field9, "")
        Field10 = Nz(rstCategory!Field10, "")
        Field11 = Nz(rstCategory!Field11, "")
        Field12 = Nz(rstCategory!Field12, "")
    End If

End Sub


'------------------------------------------------------------
' FieldSources
'
'------------------------------------------------------------
Private Sub FieldSources()

    Dim rst As DAO.Recordset
    Set rst = db.OpenRecordset("SELECT * FROM " & ItemDB)

    ' Determine field lists
    Dim ii As Integer
    Dim ss As String
    For ii = 0 To rst.Fields.count - 1
        If Len(ss) > 0 Then
            ss = ss & ";'" & rst.Fields(ii).Name & "'"
        Else
            ss = "'" & rst.Fields(ii).Name & "'"
        End If
    Next ii

    ' Set field lists
    Field1.RowSource = ss
    Field2.RowSource = ss
    Field3.RowSource = ss
    Field4.RowSource = ss
    Field5.RowSource = ss
    Field6.RowSource = ss
    Field7.RowSource = ss
    Field8.RowSource = ss
    Field9.RowSource = ss
    Field10.RowSource = ss
    Field11.RowSource = ss
    Field12.RowSource = ss
End Sub


'------------------------------------------------------------
' NewItem
'
'------------------------------------------------------------
Private Sub NewItem()
    Dim rst As DAO.Recordset
    If (EmployeeID > 0) Then
        Set rst = db.OpenRecordset("SELECT * FROM " & CategoryDB & _
            " WHERE [CategoryName] = '" & Category & "' AND [User] = " & EmployeeID)
        If rst.RecordCount = 0 Then
            ' Save New Category Record
            With rstCategory
                .AddNew
                !CategoryName = Category
                !Prefix = Prefix
                !User = EmployeeID
                !Field1 = Field1
                !Field2 = Field2
                !Field3 = Field3
                !Field4 = Field4
                !Field5 = Field5
                !Field6 = Field6
                !Field7 = Field7
                !Field8 = Field8
                !Field9 = Field9
                !Field10 = Field10
                !Field11 = Field11
                !Field12 = Field12
                .Update
            End With
        Else
            MsgBox "Record exists:" & vbCrLf & "Use SAVE button instead", , "Record Exists"
            GoTo CleanUp
        End If
    Else
        MsgBox "Invalid user:" & vbCrLf & "Sign in again and retry", , "Invalid User"
        Exit Sub
    End If
CleanUp:
    rst.Close
    Set rst = Nothing
    Exit Sub
End Sub


'------------------------------------------------------------
' SaveItem
'
'------------------------------------------------------------
Private Sub SaveItem()
    If (EmployeeID > 0) Then
        If rstCategory!User.Value = EmployeeID Then
            ' Save Category Record
            With rstCategory
                .Edit
                !CategoryName = Category
                !Prefix = Prefix
                !User = EmployeeID
                !Field1 = Field1
                !Field2 = Field2
                !Field3 = Field3
                !Field4 = Field4
                !Field5 = Field5
                !Field6 = Field6
                !Field7 = Field7
                !Field8 = Field8
                !Field9 = Field9
                !Field10 = Field10
                !Field11 = Field11
                !Field12 = Field12
                .Update
            End With
        Else
            MsgBox "Invalid user:" & vbCrLf & "Not allowed to edit customizations for other users", , "Invalid User"
            Exit Sub
        End If
    Else
        MsgBox "Invalid user:" & vbCrLf & "Sign in again and retry", , "Invalid User"
        Exit Sub
    End If
End Sub


'------------------------------------------------------------
' DeleteItem
'
'------------------------------------------------------------
Private Sub DeleteItem()
    If Not (EmployeeID > 0) Then
        If rstCategory!User.Value = EmployeeID Then
            ' Delete Category Record
            rstCategory.Delete
        Else
            MsgBox "Invalid user:" & vbCrLf & "Not allowed to delete customizations for other users", , "Invalid User"
            Exit Sub
        End If
    Else
        MsgBox "Invalid user:" & vbCrLf & "Sign in again and retry", , "Invalid User"
        Exit Sub
    End If
End Sub


'------------------------------------------------------------
' SetScreenSize
'
'------------------------------------------------------------
Private Sub SetScreenSize()
    On Error Resume Next
    Me.sbfrmDS.Left = 0
    Me.sbfrmDS.Top = 0
    Me.sbfrmDS.Width = Round(Me.WindowWidth)
    Me.sbfrmDS.Height = Round(Me.WindowHeight) - 3400
End Sub
