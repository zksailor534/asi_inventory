Option Compare Database
Option Explicit


'------------------------------------------------------------
' Form_Open
'
'------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

    Dim subForm As Form

    ' Maximize the form
    DoCmd.Maximize

    ' Set selected category
    If (EmployeeCategory = "") Then
        CategorySelected = Utilities.SelectFirstCategory
    Else
        CategorySelected = EmployeeCategory
    End If

    ' Set the subform properties
    Set subForm = sbfrmInvSearch.Form

    ' Set screen view properties
    subForm.DatasheetFontHeight = 10
    SetScreenSize

    ' Set visibility for the extra fields in warehouse portion of query
    subForm.Controls("CreateDate").ColumnHidden = True
    subForm.Controls("CreateOper").ColumnHidden = True
    subForm.Controls("LastOper").ColumnHidden = True
    subForm.Controls("LastDate").ColumnHidden = True

    ' Default Stock Selected combobox to On Hand
    Me.StockSelected.Value = "On Hand"
    subForm.Controls("OnOrder").ColumnHidden = True

    ' Engage filter from category selection
    subForm.Filter = "[Category]= '" & CategorySelected & "' AND [OnHand] > 0"
    subForm.FilterOn = True

    ' Set column visibility
    SetColumnVisibility

    CategorySelected.SetFocus

End Sub


'------------------------------------------------------------
' CategorySelected_AfterUpdate
'
'------------------------------------------------------------
Private Sub CategorySelected_AfterUpdate()

    Dim subForm As Access.Form
    Set subForm = sbfrmInvSearch.Form

    ' Engage filter from category selection
    subForm.Filter = "[Category]= '" & CategorySelected & "' AND [OnHand] > 0"
    subForm.FilterOn = True

    SetColumnVisibility

    CategorySelected.SetFocus

End Sub


'------------------------------------------------------------
' RecordIDFilter_AfterUpdate
'
'------------------------------------------------------------
Private Sub RecordIDFilter_AfterUpdate()
    If Not (IsNull(RecordIDFilter)) Or (RecordIDFilter <> "") Then
        IDFilterButton_Click
    End If
End Sub


'------------------------------------------------------------
' IDFilterButton_Click
'
'------------------------------------------------------------
Private Sub IDFilterButton_Click()
    If IsNull(RecordIDFilter) Or (RecordIDFilter = "") Then
        RecordIDFilter.SetFocus
    Else
        If ValidRecordID(RecordIDFilter) Then
            ' Engage filter from RecordID selection
            Me.sbfrmInvSearch.Form.Filter = "[Category]= '" & CategorySelected & "' AND [OnHand] > 0" & _
                "' AND [RecordID] = '" & RecordIDFilter & "'"
            Me.sbfrmInvSearch.Form.FilterOn = True
            CurrentItemID = GetRecordItemID(RecordIDFilter)
        Else
            RecordIDFilter = ""
            RecordIDFilter.SetFocus
        End If
    End If
End Sub


'------------------------------------------------------------
' ClearFilterButton_Click
'
'------------------------------------------------------------
Private Sub ClearFilterButton_Click()
    RecordIDFilter = ""
    CurrentSalesOrder = ""
    Me.sbfrmInvSearch.Form.Filter = "[Category]= '" & CategorySelected & "' AND [OnHand] > 0"
    Me.sbfrmInvSearch.Form.FilterOn = True
    Me.sbfrmInvSearch.Form.Requery
End Sub


'------------------------------------------------------------
' StockSelected_AfterUpdate
'
'------------------------------------------------------------
Private Sub StockSelected_AfterUpdate()
    If (StockSelected.Value = "On Hand") Then
        Me.sbfrmInvSearch.Form.Filter = "[Category]= '" & CategorySelected & "' AND [OnHand] > 0"
        Me.sbfrmInvSearch.Form.FilterOn = True
        Me.sbfrmInvSearch.Form.Controls("OnOrder").ColumnHidden = True
        Me.sbfrmInvSearch.Form.Requery
    ElseIf (StockSelected.Value = "Inbound") Then
        Me.sbfrmInvSearch.Form.Filter = "[OnOrder] > 0"
        Me.sbfrmInvSearch.Form.FilterOn = True
        Me.sbfrmInvSearch.Form.Controls("OnOrder").ColumnHidden = False
        Me.sbfrmInvSearch.Form.Controls("OnOrder").ColumnOrder = 5
        Me.sbfrmInvSearch.Form.Requery
    ElseIf (StockSelected.Value = "Out of Stock") Then
        Me.sbfrmInvSearch.Form.Filter = "[Category]= '" & CategorySelected & "' AND [OnHand] <= 0"
        Me.sbfrmInvSearch.Form.FilterOn = True
        Me.sbfrmInvSearch.Form.Controls("OnOrder").ColumnHidden = True
        Me.sbfrmInvSearch.Form.Requery
    End If
End Sub


'------------------------------------------------------------
' ViewItemButton_Click
'
'------------------------------------------------------------
Private Sub ViewItemButton_Click()

    If (CurrentItemID <> 0) Then
        DoCmd.OpenForm ItemDetailForm
    Else
        MsgBox "Please select record to view"
        Exit Sub
    End If

End Sub


'------------------------------------------------------------
' EditItemButton_Click
'
'------------------------------------------------------------
Private Sub EditItemButton_Click()

    If (CurrentItemID > 0) Then
        DoCmd.OpenForm ItemEditForm, acNormal, , , , acDialog
        Me.sbfrmInvSearch.Form.Requery
    Else
        Debug.Print "CurrentItemID", CurrentItemID
        MsgBox "Please select record to edit"
        Exit Sub
    End If

End Sub


'------------------------------------------------------------
' ManageInvButton_Click
'
'------------------------------------------------------------
Private Sub ManageInvButton_Click()

    If (CurrentItemID > 0) Then
        DoCmd.OpenForm ItemInventoryManageForm, acNormal, , , , acDialog
        Me.sbfrmInvSearch.Form.Requery
    Else
        MsgBox "Please select record to edit"
        Exit Sub
    End If

End Sub


'------------------------------------------------------------
' Form_Resize
'
'------------------------------------------------------------
Private Sub Form_Resize()
    SetScreenSize
End Sub


'------------------------------------------------------------
' SetScreenSize
'
'------------------------------------------------------------
Private Sub SetScreenSize()
    On Error Resume Next
    Me.sbfrmInvSearch.Left = 0
    Me.sbfrmInvSearch.Top = 0
    Me.sbfrmInvSearch.Width = Round(Me.WindowWidth)
    Me.sbfrmInvSearch.Height = Round(Me.WindowHeight * 0.95)
End Sub


'------------------------------------------------------------
' SetColumnVisibility
'
'------------------------------------------------------------
Private Sub SetColumnVisibility()

    Dim subForm As Form
    Dim formCntrl As Control
    Dim LeftColumns, ColumnValue, ii As Integer

    Set subForm = sbfrmInvSearch.Form
    LeftColumns = 6

    ' Set column visibility and order
    For Each formCntrl In subForm.Controls
        If formCntrl.ControlType <> acLabel Then
            ColumnValue = Utilities.CategoryFieldOrder(CategorySelected, formCntrl.Name)
            If ColumnValue <> -1 Then
                ' Column is visible
                formCntrl.ColumnHidden = False

                ' Set column order
                formCntrl.ColumnOrder = LeftColumns + (ColumnValue * 2)

                ' Set column width to fit, except Description and Condition
                If (formCntrl.Name = "Description") Then
                    formCntrl.ColumnWidth = 2500
                ElseIf (formCntrl.Name = "Condition") Then
                    formCntrl.ColumnWidth = 250
                Else
                    formCntrl.ColumnWidth = -2
                End If
            Else
                ' Hide column
                formCntrl.ColumnHidden = True
            End If
        End If
    Next

    ' Set Warehouse column locations
    subForm.Location.ColumnOrder = 1
    subForm.RecordID.ColumnOrder = 2
    subForm.OnHand.ColumnOrder = 3
    subForm.Available.ColumnOrder = 4

    ' Set Warehouse column sizes
    subForm.Location.ColumnWidth = -2
    subForm.RecordID.ColumnWidth = 900
    subForm.OnHand.ColumnWidth = 550
    subForm.Available.ColumnWidth = 550

    ' Set Warehouse column visibility
    subForm.Location.ColumnHidden = False
    subForm.RecordID.ColumnHidden = False
    subForm.OnHand.ColumnHidden = False
    subForm.Available.ColumnHidden = False

End Sub


'------------------------------------------------------------
' ValidRecordID
'
'------------------------------------------------------------
Private Function ValidRecordID(RecordID As String) As Boolean
    On Error Resume Next
    If Not (IsNull(DLookup("RecordID", WarehouseQuery, "[Category]= '" & CategorySelected & _
        "' AND [RecordID]='" & RecordID & "'"))) Then
        ValidRecordID = True
    Else
        MsgBox "Invalid Record ID: Not Found", vbOKOnly, "Invalid Record ID"
        ValidRecordID = False
    End If
End Function


'------------------------------------------------------------
' GetRecordItemID
'
'------------------------------------------------------------
Private Function GetRecordItemID(RecordID As String) As Long
    On Error Resume Next

    Dim ItemId As Long

    ItemId = DLookup("ID", WarehouseQuery, "[Category]= '" & CategorySelected & _
        "' AND [RecordID]='" & RecordID & "'")
    If Not (IsNull(ItemId)) Then
        GetRecordItemID = ItemId
    Else
        GetRecordItemID = 0
    End If
End Function
