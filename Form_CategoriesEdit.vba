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
    DefaultCategory = Utilities.SelectFirstCategory

    SetScreenSize

    ' Open the recordset
    open_db
    Set rstCategory = db.OpenRecordset("SELECT * FROM " & CategoryDB & _
        " WHERE [CategoryName] LIKE '" & DefaultCategory & "' AND [User] IS NULL")

    ' Hide User Column
    Me.sbfrmDS!User.ColumnHidden = True

    ' Filter only on categories, not user custom field order
    Me.sbfrmDS.Form.Filter = "[User] IS NULL"
    Me.sbfrmDS.Form.FilterOn = True

    FillFields

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
' SaveButton_Click
'
'------------------------------------------------------------
Private Sub AddNewButton_Click()
    Dim saveCategory As Integer
    saveCategory = MsgBox("Do you want to save this New Category?", vbYesNo, "Save Category")
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
    saveCategory = MsgBox("Do you want to save changes to this Category?", vbYesNo, "Save Category")
    If saveCategory = vbYes Then
        Set rstCategory = db.OpenRecordset("SELECT * FROM " & CategoryDB & _
            " WHERE [CategoryName] = '" & Category & "' AND [User] IS NULL")
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
    delCategory = MsgBox("Do you want to delete this Category?", vbYesNo, "Delete Category")
    If delCategory = vbYes Then
        Set rstCategory = db.OpenRecordset("SELECT * FROM " & CategoryDB & _
            " WHERE [CategoryName] = '" & Category & "' AND [User] IS NULL")
        DeleteItem
    End If
End Sub


'------------------------------------------------------------
' FillFields
'
'------------------------------------------------------------
Private Sub FillFields()

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
    If Not (Utilities.RecordExists(CategoryDB, "CategoryName", Category)) Then
        ' Save New Category Record
        With rstCategory
            .AddNew
            !CategoryName = Category
            !Prefix = Prefix
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
            !User = Null
            .Update
        End With
    Else
        MsgBox "Record exists:" & vbCrLf & "Use SAVE button instead", , "Record Exists"
        Exit Sub
    End If
End Sub


'------------------------------------------------------------
' SaveItem
'
'------------------------------------------------------------
Private Sub SaveItem()
    If Utilities.RecordExists(CategoryDB, "CategoryName", Category) Then
        ' Save Category Record
        With rstCategory
            .Edit
            !CategoryName = Category
            !Prefix = Prefix
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
        MsgBox "Record does not exist:" & vbCrLf & "Use ADD NEW button instead", , "Record Does Not Exist"
        Exit Sub
    End If
End Sub


'------------------------------------------------------------
' DeleteItem
'
'------------------------------------------------------------
Private Sub DeleteItem()
    ' Delete Category Record
    rstCategory.Delete
End Sub


'------------------------------------------------------------
' SetScreenSize
'
'------------------------------------------------------------
Public Sub SetScreenSize()
    On Error Resume Next
    Me.sbfrmDS.Left = 0
    Me.sbfrmDS.Top = 0
    Me.sbfrmDS.Width = ScreenWidth
    Me.sbfrmDS.Height = Round(Me.WindowHeight) - 3400
End Sub
