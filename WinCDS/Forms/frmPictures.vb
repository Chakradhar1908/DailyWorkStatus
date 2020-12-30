Public Class frmPictures
    Public Enum dbPicType
        dbpty_Sales = 0
        dbpty_PO = 1
        dbpty_ServiceParts = 2
        dbpty_ServiceCalls = 1
    End Enum

    Private mType As dbPicType, mRef As String, mLoc As Long
    Private NewIndex As Long

    Public Sub LoadPicturesByRef(ByVal pType As dbPicType, ByVal pRef As String, Optional ByVal pLoc As String = 0)
        mType = pType
        mRef = pRef
        If pLoc = 0 Then pLoc = StoresSld
        mLoc = pLoc

        UpdateCaption
        UpdateData
        Show vbModal
End Sub

End Class