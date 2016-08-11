Option Compare Database
Option Explicit

'Private searchCategory As String

'------------------------------------------------------------
' Form_Open
'
'------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

    Dim subForm As Form

    ' Maximize the form
    DoCmd.Maximize

    ' Set selected category
    If (Len(searchCategory) > 0) Then
        CategorySelected = searchCategory
    ElseIf (Len(EmployeeCategory) > 0) Then
        CategorySelected = EmployeeCategory
    Else
        CategorySelected = Utilities.SelectFirstCategory
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
    searchCategory = CategorySelected
    subForm.Filter = "[Category]= '" & searchCategory & "' AND [OnHand] > 0"
    subForm.FilterOn = True

    ' Engage default sorting
    subForm.OrderBy = "[Focus]"
    subForm.OrderByOn = True

    ' Set column visibility
    SetColumnVisibility

    setUserPermissions

outNow:
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
    searchCategory = CategorySelected
    subForm.Filter = "[Category]= '" & searchCategory & "' AND [OnHand] > 0"
    Me.StockSelected.Value = "On Hand"
    subForm.FilterOn = True

    ' Engage default sorting
    subForm.OrderBy = "[Focus]"
    subForm.OrderByOn = True

    SetColumnVisibility

outNow:
    CategorySelected.SetFocus

End Sub


'------------------------------------------------------------
' StockSelected_AfterUpdate
'
'------------------------------------------------------------
Private Sub StockSelected_AfterUpdate()
    If (StockSelected.Value = "On Hand") Then
        Me.CategorySelected = searchCategory
        Me.sbfrmInvSearch.Form.Filter = "[Category]= '" & searchCategory & "' AND [OnHand] > 0"
        Me.sbfrmInvSearch.Form.FilterOn = True
        Me.sbfrmInvSearch.Form.Controls("OnOrder").ColumnHidden = True
        Me.sbfrmInvSearch.Form.OrderBy = "[Focus]"
        Me.sbfrmInvSearch.Form.OrderByOn = True
        Me.sbfrmInvSearch.Form.Requery
    ElseIf (StockSelected.Value = "Inbound Only") Then
        Me.CategorySelected = ""
        Me.sbfrmInvSearch.Form.Filter = "[OnOrder] > 0"
        Me.sbfrmInvSearch.Form.FilterOn = True
        Me.sbfrmInvSearch.Form.Controls("OnOrder").ColumnHidden = False
        Me.sbfrmInvSearch.Form.Controls("OnOrder").ColumnOrder = 5
        Me.sbfrmInvSearch.Form.OrderBy = "[Focus]"
        Me.sbfrmInvSearch.Form.OrderByOn = True
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
' PrintButton_Click
'
'------------------------------------------------------------
Private Sub PrintButton_Click()
    On Error GoTo PrintButton_ErrHandler

    PrintCategorySelected = searchCategory
    PrintFilter = sbfrmInvSearch.Form.Filter

    ' Open PrintRange form
    DoCmd.OpenForm PrintRangeForm, acFormDS, , , , , acWindowNormal

PrintButton_Exit:
    PrintFilter = ""
    PrintCategorySelected = ""
    Exit Sub

PrintButton_ErrHandler:
    MsgBox "Error in SelRecsBtn_Click( ) in" & vbCrLf & Me.Name & " form." & vbCrLf & vbCrLf & "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description
    Err.Clear
    GoTo PrintButton_Exit

End Sub


'------------------------------------------------------------
' SplitScreenButton_Click
'
'------------------------------------------------------------
Private Sub SplitScreenButton_Click()
    If (Utilities.HasParent(Me)) Then
        If (Me.Parent.Name = SalesForm) Then
            Me.Parent!nvbSearch.NavigationTargetName = SalesSearchSplit
            Me.Parent!nvbSearch.SetFocus
            SendKeys "{ENTER}", 0
        Else
            SplitScreenButton.Enabled = False
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
    Me.sbfrmInvSearch.Left = 0
    Me.sbfrmInvSearch.Top = 0
    Me.sbfrmInvSearch.Width = ScreenWidth
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
            ColumnValue = Utilities.CategoryFieldOrder(searchCategory, formCntrl.Name)
            If ColumnValue <> -1 Then
                ' Column is visible
                formCntrl.ColumnHidden = False

                ' Set column order (doubled columnvalue to ensure that order is correct)
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


'------------------------------------------------------------
' setUserPermissions
'
'------------------------------------------------------------
Private Sub setUserPermissions()
    If (EmployeeRole = SalesLevel) Then
        ViewItemButton.Enabled = True
        ViewItemButton.Visible = True
        PrintButton.Enabled = True
        PrintButton.Visible = True
        SplitScreenButton.Enabled = True
        SplitScreenButton.Visible = True
    ElseIf (EmployeeRole = ProdLevel) Then
        ViewItemButton.Enabled = True
        ViewItemButton.Visible = True
        PrintButton.Enabled = False
        PrintButton.Visible = False
        SplitScreenButton.Enabled = False
        SplitScreenButton.Visible = False
    ElseIf (EmployeeRole = AdminLevel) Then
        ViewItemButton.Enabled = True
        ViewItemButton.Visible = True
        PrintButton.Enabled = True
        PrintButton.Visible = True
        SplitScreenButton.Enabled = True
        SplitScreenButton.Visible = True
    ElseIf (EmployeeRole = DevelLevel) Then
        ViewItemButton.Enabled = True
        ViewItemButton.Visible = True
        PrintButton.Enabled = True
        PrintButton.Visible = True
        SplitScreenButton.Enabled = True
        SplitScreenButton.Visible = True
    End If
End Sub
