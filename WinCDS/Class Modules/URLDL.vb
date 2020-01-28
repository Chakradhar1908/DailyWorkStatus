Public Class URLDL
    Public Function DownloadFileProgress(ByVal URL As String, ByVal LocalFile As String, Optional ByRef ErrCode As Long, Optional ByRef AltPrg As Object, Optional ByRef AltPrg2 As Object) As Boolean
        On Error Resume Next
  Set P = New frmProgress
  Set P.AltPrg = AltPrg
  Set P2 = AltPrg2
  Start = P2.Value

        ErrCode = olelib.URLDownloadToFile(Nothing, URL, LocalFile, 0, Me)
        DownloadFileProgress = (ErrCode = olelib.S_OK)
    End Function

End Class
