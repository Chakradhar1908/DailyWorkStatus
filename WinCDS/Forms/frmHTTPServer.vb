Public Class frmHTTPServer
    Public Property HTTPPort() As Long
        Get
            HTTPPort = mHTTPPort
            If HTTPPort = 0 Then HTTPPort = 8080
        End Get
        Set(value As Long)
            StopHTTP
            mHTTPPort = value
            StartHTTP
        End Set
    End Property

    Public Sub StartHTTP()
        On Error Resume Next
        StopHTTP

        sck(0).LocalPort = HTTPPort ' set this to the port you want the server to listen on...
        sck(0).Listen

        DoEvents

        If sck(0).State = sckListening Then lblFileProgress(0) = "00 Listening"
    End Sub

End Class