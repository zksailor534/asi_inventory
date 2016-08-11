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
    On Error GoTo ErrHandler
    m_nHt = Me.SelHeight
    CurrentItemID = Me.ID
    Exit Sub

ErrHandler:
    MsgBox "Error in Form_Click( ) in" & vbCrLf & Me.Name & " form." & vbCrLf & vbCrLf & "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description
    Err.Clear
End Sub

Private Sub Location_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub RecordID_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub OnHand_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Available_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub ItemLength_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Manufacturer_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub NumSteps_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub RollerCenter_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub ItemDepth_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub ItemWidth_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub ItemHeight_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub SubCategory_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Model_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Style_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Color_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Column_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Condition_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub NumStruts_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Description_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Vendor_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub CreateDate_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub SuggSellingPrice_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Capacity_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub BoltPattern_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub HoleCenter_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Diameter_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Degree_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub DriveType_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Gauge_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Volts_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Phase_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub AmpHR_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub Serial_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub QtyDoors_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub TopLiftHeight_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub LowerLiftHeight_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub TopStepHeight_Click()
    CurrentItemID = Me.ID
End Sub

Private Sub ID_Click()
    CurrentItemID = Me.ID
End Sub
