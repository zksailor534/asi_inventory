Option Compare Database
Option Explicit


Private Sub Form_Open(Cancel As Integer)

    Dim formCntrl As Control
    Dim LeftColumns, ColumnValue, ii As Integer

    LeftColumns = 6

    ' Set column visibility and order
    For Each formCntrl In Me.Form.Controls
        If formCntrl.ControlType <> acLabel Then
            ColumnValue = Utilities.CategoryFieldOrder(PrintCategorySelected, formCntrl.Name)
            If ColumnValue <> -1 Then
                ' Column is visible
                formCntrl.ColumnHidden = False

                ' Set column width to fit
                formCntrl.ColumnOrder = ColumnValue + LeftColumns

                ' Set column width to fit
                formCntrl.ColumnWidth = -2
            Else
                ' Hide column
                formCntrl.ColumnHidden = True
            End If
        End If
    Next

    ' Set Warehouse column locations
    Me.Location.ColumnOrder = 1
    Me.RecordID.ColumnOrder = 2
    Me.OnHand.ColumnOrder = 3
    Me.Available.ColumnOrder = 4

    ' Set Warehouse column sizes
    Me.Location.ColumnWidth = -2
    Me.RecordID.ColumnWidth = 900
    Me.OnHand.ColumnWidth = 550
    Me.Available.ColumnWidth = 550

    ' Set Warehouse column visibility
    Me.Location.ColumnHidden = False
    Me.RecordID.ColumnHidden = False
    Me.OnHand.ColumnHidden = False
    Me.Available.ColumnHidden = False

End Sub
