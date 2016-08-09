Option Compare Database
Option Explicit

Dim CatEdit As Boolean


'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
    If (Utilities.HasParent(Me)) Then
        If (Me.Parent.Name = CategoriesEditForm) Then
            CatEdit = True
        Else
            CatEdit = False
        End If
    End If

    ' Set column widths
    SetColumnWidth
End Sub

Private Sub ID_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub CategoryName_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Prefix_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field1_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field2_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field3_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field4_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field5_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field6_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field7_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field8_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field9_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field10_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub

Private Sub Field11_Click()
    If (CatEdit) Then
        Me.Parent!Category = Me.CategoryName
        Me.Parent!Prefix = Me.Prefix
        Me.Parent!Field1 = Me.Field1
        Me.Parent!Field2 = Me.Field2
        Me.Parent!Field3 = Me.Field3
        Me.Parent!Field4 = Me.Field4
        Me.Parent!Field5 = Me.Field5
        Me.Parent!Field6 = Me.Field6
        Me.Parent!Field7 = Me.Field7
        Me.Parent!Field8 = Me.Field8
        Me.Parent!Field9 = Me.Field9
        Me.Parent!Field10 = Me.Field10
        Me.Parent!Field11 = Me.Field11
    End If
End Sub


'------------------------------------------------------------
' SetColumnWidth
'
'------------------------------------------------------------
Private Sub SetColumnWidth()

    Dim formCntrl As Control

    ' Set column widths to fit
    For Each formCntrl In Me.Controls
        If formCntrl.ControlType <> acLabel Then
            If (formCntrl.Name = "ID") Then
                formCntrl.ColumnWidth = 500
            ElseIf (formCntrl.Name = "CategoryName") Then
                formCntrl.ColumnWidth = -2
            ElseIf (formCntrl.Name = "Prefix") Then
                formCntrl.ColumnWidth = -2
            Else
                formCntrl.ColumnWidth = -2
            End If
        End If
    Next

End Sub
