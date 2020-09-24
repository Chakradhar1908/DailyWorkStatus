Public Class URLDL
    Dim P As frmProgress, P2 As Object, Start As Integer
    'Implements olelib.IBindStatusCallback
    Public Function DownloadFileProgress(ByVal URL As String, ByVal LocalFile As String, Optional ByRef ErrCode As Integer = 0, Optional ByRef AltPrg As Object = Nothing, Optional ByRef AltPrg2 As Object = Nothing) As Boolean
        On Error Resume Next
        P = New frmProgress
        'P.AltPrg = AltPrg
        P2 = AltPrg2
        Start = P2.Value

        'ErrCode = olelib.URLDownloadToFile(Nothing, URL, LocalFile, 0, Me)
        'DownloadFileProgress = (ErrCode = olelib.S_OK)
    End Function
End Class
