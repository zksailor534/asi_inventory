Option Compare Database


'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
    ' Input the pre-selected Category
    cboCategory = Me.OpenArgs

    ' Get the record prefix for selected category
    txtPrefix = DLookup("Prefix", CategoryQuery, "CategoryName= '" & cboCategory & "'")

    ' Clear Record ID list
    lstRecordIDs.RowSource = vbNullString
End Sub


'------------------------------------------------------------
' cboCategory_AfterUpdate
'
'------------------------------------------------------------
Private Sub cboCategory_AfterUpdate()
    ' Get the record prefix for selected category
    txtPrefix = DLookup("Prefix", CategoryQuery, "CategoryName= '" & cboCategory & "'")
End Sub


'------------------------------------------------------------
' cmdGetID_Click
'
'------------------------------------------------------------
Private Sub cmdGetID_Click()
    Dim progressAmount As Integer
    Dim startNumber As Long
    Dim newID As String

    ' Clear old list and re-initialize
    lstRecordIDs.RowSource = vbNullString
    startNumber = 1

    If (IsNull(txtPrefix) Or (txtPrefix = "")) Then
        MsgBox "Invalid Record Prefix", vbOKOnly, "Invalid Prefix"
    ElseIf IsNull(cboNumRecordIDs.Value) Then
        MsgBox "Invalid Number of Records", vbOKOnly, "Invalid Number of Records"
    Else
        ' Get the IDs
        SysCmd acSysCmdInitMeter, "Generating Record IDs...", cboNumRecordIDs.Value

        ' Loop over number of items
        For progressAmount = 1 To cboNumRecordIDs.Value
            ' Update the progress meter
             SysCmd acSysCmdUpdateMeter, progressAmount

             ' Retrieve Record ID
             newID = Utilities.NewRecordID(txtPrefix, startNumber)

             ' Add new ID to list
             lstRecordIDs.AddItem newID

             ' Move to next start number
             startNumber = Utilities.StripRecordID(txtPrefix, newID) + 1

       Next progressAmount

       ' Remove the progress meter
       SysCmd acSysCmdRemoveMeter

    End If
End Sub


'------------------------------------------------------------
' cmdReserveSel_Click
'
'------------------------------------------------------------
Private Sub cmdReserveSel_Click()
    Dim lngRow As Integer
    Dim success As Boolean
    If (lstRecordIDs.ItemsSelected.count = 0) Then
        MsgBox "No Record IDs Selected... Try Again", , "Nothing Selected"
    Else
        With Me.lstRecordIDs
            For lngRow = 0 To .ListCount - 1
                If .Selected(lngRow) Then
                    success = Utilities.RecordIDReserve(.Column(0, lngRow), cboCategory)
                    If Not (success) Then
                        MsgBox "Error: " & vbCrLf & "Unable to reserve Record ID " & _
                            .Column(0, lngRow) & vbCrLf & "Terminating Save"
                        Exit Sub
                    End If
                End If
            Next lngRow
        End With
        If success Then
            MsgBox "Successfully saved all Record IDs"
        End If
    End If
End Sub


'------------------------------------------------------------
' cmdReserveAll_Click
'
'------------------------------------------------------------
Private Sub cmdReserveAll_Click()
    Dim lngRow As Integer
    Dim success As Boolean
    If (lstRecordIDs.ListCount = 0) Then
        MsgBox "No Record IDs Generated... Try Again", , "Nothing Generated"
    Else
        With Me.lstRecordIDs
            For lngRow = 0 To .ListCount - 1
                success = Utilities.RecordIDReserve(.Column(0, lngRow), cboCategory)
                If Not (success) Then
                    MsgBox "Error: " & vbCrLf & "Unable to reserve Record ID " & _
                        .Column(0, lngRow) & vbCrLf & "Terminating Save"
                    Exit Sub
                End If
            Next lngRow
        End With
        If success Then
            MsgBox "Successfully saved all Record IDs"
        End If
    End If
End Sub
