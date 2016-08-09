Option Compare Database
Option Explicit

Private m_nHt As Long

Public Property Let DSSelHeight(nHt As Long)
    m_nHt = nHt
End Property

Public Property Get DSSelHeight() As Long
    DSSelHeight = m_nHt
End Property

Private Sub Form_Click()
    On Error Resume Next
    m_nHt = Me.SelHeight
    Exit Sub

ErrHandler:
    MsgBox "Error in Form_Click( ) in" & vbCrLf & Me.Name & " form." & vbCrLf & vbCrLf & "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description
    Err.Clear
End Sub

Private Sub ID_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Category_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Location_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Order_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Product_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Style_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub QtyCommitted_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub RecordID_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Manufacturer_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Description_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub OperatorActive_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Color_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Condition_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Vendor_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub OnHand_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub Committed_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub

Private Sub CommitID_Click()
    CurrentCommitID = Me.CommitID
    CurrentSalesOrder = Me.Reference
End Sub
