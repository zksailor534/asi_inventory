Option Compare Database

Private Sub Form_Load()
    ' Input the pre-selected Category
    cboCategory = Me.OpenArgs

    ' Get the record prefix for selected category
    txtPrefix = DLookup("Prefix", CategoryQuery, "CategoryName= '" & cboCategory & "'")
End Sub

Private Sub cboCategory_AfterUpdate()
    ' Get the record prefix for selected category
    txtPrefix = DLookup("Prefix", CategoryQuery, "CategoryName= '" & cboCategory & "'")
End Sub

Private Sub cmdGetID_Click()
    If (txtPrefix <> "") Then
        ' Get the IDs
        lblID1.Caption = Utilities.NewRecordID(txtPrefix, 1)
        lblID2.Caption = Utilities.NewRecordID(txtPrefix, Utilities.StripRecordID(txtPrefix, lblID1.Caption) + 1)
        lblID3.Caption = Utilities.NewRecordID(txtPrefix, Utilities.StripRecordID(txtPrefix, lblID2.Caption) + 1)
    Else
        MsgBox "Invalid Record Prefix", vbOKOnly, "Invalid Prefix"
    End If
End Sub
