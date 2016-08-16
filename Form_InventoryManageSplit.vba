Option Compare Database
Option Explicit


'------------------------------------------------------------
' Form_Open
'
'------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

    Dim subForm1, subForm2 As Form

    ' Maximize the form
    DoCmd.Maximize

    ' Set selected top category
    If (Len(searchCategory) > 0) Then
        CategorySelectedTop = searchCategory
    ElseIf (Len(EmployeeCategory) > 0) Then
        CategorySelectedTop = EmployeeCategory
    Else
        CategorySelectedTop = Utilities.SelectFirstCategory
    End If

    ' Set selected bottom category
    If (Len(searchCategoryBottom) > 0) Then
        CategorySelectedBottom = searchCategoryBottom
    ElseIf (Len(EmployeeCategory) > 0) Then
        CategorySelectedBottom = EmployeeCategory
    Else
        CategorySelectedBottom = Utilities.SelectFirstCategory
    End If

    ' Set the subform properties
    Set subForm1 = sbfrmInvSearch.Form
    Set subForm2 = sbfrmInvSearchBottom.Form

    ' Set screen view properties
    SetScreenSize
    subForm1.DatasheetFontHeight = 10
    subForm2.DatasheetFontHeight = 10

    ' Set visibility for the extra fields in warehouse portion of query
    subForm1.Controls("CreateDate").ColumnHidden = True
    subForm1.Controls("CreateOper").ColumnHidden = True
    subForm1.Controls("LastOper").ColumnHidden = True
    subForm1.Controls("LastDate").ColumnHidden = True
    subForm2.Controls("CreateDate").ColumnHidden = True
    subForm2.Controls("CreateOper").ColumnHidden = True
    subForm2.Controls("LastOper").ColumnHidden = True
    subForm2.Controls("LastDate").ColumnHidden = True

    ' Engage filter from category selection
    subForm1.Filter = "[Category]= '" & CategorySelectedTop & "' AND [OnHand] > 0"
    subForm1.FilterOn = True
    subForm2.Filter = "[Category]= '" & CategorySelectedBottom & "' AND [OnHand] > 0"
    subForm2.FilterOn = True

    ' Engage default sorting
    subForm1.OrderBy = "[Focus]"
    subForm1.OrderByOn = True
    subForm2.OrderBy = "[Focus]"
    subForm2.OrderByOn = True

    ' Set column visibility
    SetColumnVisibility

outNow:
    CategorySelectedTop.SetFocus

End Sub


'------------------------------------------------------------
' CategorySelectedTop_AfterUpdate
'
'------------------------------------------------------------
Private Sub CategorySelectedTop_AfterUpdate()

    Dim subForm As Access.Form
    Set subForm = sbfrmInvSearch.Form

    ' Engage filter from category selection
    searchCategory = CategorySelectedTop
    subForm.Filter = "[Category]= '" & CategorySelectedTop & "' AND [OnHand] > 0"
    subForm.FilterOn = True

    ' Engage default sorting
    subForm.OrderBy = "[Focus]"
    subForm.OrderByOn = True

    SetColumnVisibility

outNow:
    CategorySelectedTop.SetFocus

End Sub


'------------------------------------------------------------
' CategorySelectedBottom_AfterUpdate
'
'------------------------------------------------------------
Private Sub CategorySelectedBottom_AfterUpdate()

    Dim subForm As Access.Form
    Set subForm = sbfrmInvSearchBottom.Form

    ' Engage filter from category selection
    searchCategoryBottom = CategorySelectedBottom
    subForm.Filter = "[Category]= '" & CategorySelectedBottom & "' AND [OnHand] > 0"
    subForm.FilterOn = True

    ' Engage default sorting
    subForm.OrderBy = "[Focus]"
    subForm.OrderByOn = True

    SetColumnVisibility

outNow:
    CategorySelectedBottom.SetFocus

End Sub


'------------------------------------------------------------
' ViewItemButton_Click
'
'------------------------------------------------------------
Private Sub ViewItemButton_Click()

    If Not (IsNull(CurrentItemID)) Then
        DoCmd.OpenForm ItemDetailForm
    Else
        MsgBox "Please select record to view"
        Exit Sub
    End If

End Sub


'------------------------------------------------------------
' cmdCloseSplit_Click
'
'------------------------------------------------------------
Private Sub cmdCloseSplit_Click()
    If Utilities.HasParent(Me) Then
        If Me.Parent.Name = InventoryForm Then
            Me.Parent!nvbManageInventory.NavigationTargetName = InventoryManageForm
            Me.Parent!nvbManageInventory.SetFocus
            SendKeys "{ENTER}", 0
        End If
    Else
        DoCmd.Close
    End If
End Sub


'------------------------------------------------------------
' SetScreenSize
'
'------------------------------------------------------------
Public Sub SetScreenSize()
    ' Top datasheet
    Me.sbfrmInvSearch.Left = 0
    Me.sbfrmInvSearch.Top = 0
    Me.sbfrmInvSearch.Width = ScreenWidth
    Me.sbfrmInvSearch.Height = Round(Me.WindowHeight * 0.95) / 2
    ' Bottom datasheet
    Me.sbfrmInvSearchBottom.Left = 0
    Me.sbfrmInvSearchBottom.Top = Round(Me.WindowHeight * 0.95) / 2 + 100
    Me.sbfrmInvSearchBottom.Width = ScreenWidth
    Me.sbfrmInvSearchBottom.Height = Round(Me.WindowHeight * 0.95) / 2
    ' Divider line
    Me.lineDivider.Top = Round(Me.WindowHeight * 0.95) / 2 + 55
    Me.lineDivider.Width = ScreenWidth
End Sub


'------------------------------------------------------------
' SetColumnVisibility
'
'------------------------------------------------------------
Private Sub SetColumnVisibility()

    Dim subForm As Form
    Dim formCntrl As Control
    Dim LeftColumns, ColumnValue As Integer
    LeftColumns = 6

    '*** Top Subform
    Set subForm = sbfrmInvSearch.Form

    ' Set column visibility and order
    For Each formCntrl In subForm.Controls
        If formCntrl.ControlType <> acLabel Then
            ColumnValue = Utilities.CategoryFieldOrder(CategorySelectedTop, formCntrl.Name)
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
    subForm.OnHand.ColumnWidth = 700
    subForm.Available.ColumnWidth = 700

    ' Set Warehouse column visibility
    subForm.Location.ColumnHidden = False
    subForm.RecordID.ColumnHidden = False
    subForm.OnHand.ColumnHidden = False
    subForm.Available.ColumnHidden = False

    '*** Bottom Subform
    Set subForm = sbfrmInvSearchBottom.Form

    ' Set column visibility and order
    For Each formCntrl In subForm.Controls
        If formCntrl.ControlType <> acLabel Then
            ColumnValue = Utilities.CategoryFieldOrder(CategorySelectedBottom, formCntrl.Name)
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
    subForm.OnHand.ColumnWidth = 700
    subForm.Available.ColumnWidth = 700

    ' Set Warehouse column visibility
    subForm.Location.ColumnHidden = False
    subForm.RecordID.ColumnHidden = False
    subForm.OnHand.ColumnHidden = False
    subForm.Available.ColumnHidden = False

End Sub
