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

    ' Set selected category
    If (EmployeeCategory = "") Then
        CategorySelectedTop = Utilities.SelectFirstCategory
        CategorySelectedBottom = Utilities.SelectFirstCategory
    Else
        CategorySelectedTop = EmployeeCategory
        CategorySelectedBottom = EmployeeCategory
    End If

    ' Set the subform properties
    Set subForm1 = sbfrmInvSearch.Form
    Set subForm2 = sbfrmInvSearchBottom.Form

    ' Set screen view properties
    subForm1.DatasheetFontHeight = 10
    subForm2.DatasheetFontHeight = 10
    SetScreenSize

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
    subForm.Filter = "[Category]= '" & CategorySelectedTop & "' AND [OnHand] > 0"
    subForm.FilterOn = True

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
    subForm.Filter = "[Category]= '" & CategorySelectedBottom & "' AND [OnHand] > 0"
    subForm.FilterOn = True

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
        If Me.Parent.Name = SalesForm Then
            Me.Parent!nvbSearch.NavigationTargetName = SalesSearch
            Me.Parent!nvbSearch.SetFocus
            SendKeys "{ENTER}", 0
        End If
    Else
        DoCmd.Close
    End If
End Sub


'------------------------------------------------------------
' SortOrderEntered_AfterUpdate
'
'------------------------------------------------------------
Private Sub SortOrderEntered_AfterUpdate()

    Dim subFormCntrl As Control, sbForm As Access.Form

    Set subFormCntrl = sbfrmInvSearch

    Set sbForm = sbfrmInvSearch.Form
    sbForm.OrderBy = SortOrderEntered
    sbForm.OrderByOn = True

    SortOrderEntered.SetFocus

    Set subFormCntrl = Nothing
    Set sbForm = Nothing

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
    ' Top datasheet
    Me.sbfrmInvSearch.Left = 0
    Me.sbfrmInvSearch.Top = 0
    Me.sbfrmInvSearch.Width = Round(Me.WindowWidth)
    Me.sbfrmInvSearch.Height = Round(Me.WindowHeight * 0.95) / 2
    ' Bottom datasheet
    Me.sbfrmInvSearchBottom.Left = 0
    Me.sbfrmInvSearchBottom.Top = Round(Me.WindowHeight * 0.95) / 2 + 100
    Me.sbfrmInvSearchBottom.Width = Round(Me.WindowWidth)
    Me.sbfrmInvSearchBottom.Height = Round(Me.WindowHeight * 0.95) / 2
    ' Divider line
    Me.lineDivider.Top = Round(Me.WindowHeight * 0.95) / 2 + 55
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
