Option Compare Database
Option Explicit

Private searchCategory As String

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
    Set subForm = sbfrmInvPrice.Form

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

    ' Set column visibility
    SetColumnVisibility

outNow:
    CategorySelected.SetFocus

End Sub


'------------------------------------------------------------
' CategorySelected_AfterUpdate
'
'------------------------------------------------------------
Private Sub CategorySelected_AfterUpdate()

    Dim subForm As Access.Form
    Set subForm = sbfrmInvPrice.Form

    ' Engage filter from category selection
    searchCategory = CategorySelected
    subForm.Filter = "[Category]= '" & searchCategory & "' AND [OnHand] > 0"
    Me.StockSelected.Value = "On Hand"
    subForm.FilterOn = True

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
        Me.sbfrmInvSearch.Form.Requery
    ElseIf (StockSelected.Value = "Inbound Only") Then
        Me.CategorySelected = ""
        Me.sbfrmInvSearch.Form.Filter = "[OnOrder] > 0"
        Me.sbfrmInvSearch.Form.FilterOn = True
        Me.sbfrmInvSearch.Form.Controls("OnOrder").ColumnHidden = False
        Me.sbfrmInvSearch.Form.Controls("OnOrder").ColumnOrder = 5
        Me.sbfrmInvSearch.Form.Requery
    End If
End Sub


'------------------------------------------------------------
' PrintButton_Click
'
'------------------------------------------------------------
Private Sub PrintButton_Click()
    On Error GoTo PrintButton_ErrHandler

    PrintCategorySelected = searchCategory
    PrintFilter = sbfrmInvPrice.Form.Filter

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
' SetScreenSize
'
'------------------------------------------------------------
Private Sub SetScreenSize()
    On Error Resume Next
    Me.sbfrmInvPrice.Left = 0
    Me.sbfrmInvPrice.Top = 0
    Me.sbfrmInvPrice.Width = ScreenWidth
    Me.sbfrmInvPrice.Height = Round(Me.WindowHeight * 0.95)
End Sub


'------------------------------------------------------------
' SetColumnVisibility
'
'------------------------------------------------------------
Private Sub SetColumnVisibility()

    Dim subForm As Form
    Dim formCntrl As Control
    Dim LeftColumns, ColumnValue, ii As Integer

    Set subForm = sbfrmInvPrice.Form
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

    ' Set Price column location
    subForm.SuggSellingPrice.ColumnOrder = 5

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
