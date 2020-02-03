Public Class frmHTTPServer
    Private mHTTPPort As Long
    Public Property HTTPPort() As Long
        Get
            HTTPPort = mHTTPPort
            If HTTPPort = 0 Then HTTPPort = 8080
        End Get
        Set(value As Long)
            StopHTTP
            mHTTPPort = value
            StartHTTP()
        End Set
    End Property

    Public Sub StartHTTP()
        On Error Resume Next
        StopHTTP()

        'sck(0).LocalPort = HTTPPort ' set this to the port you want the server to listen on...
        'sck(0).Listen
        sck0.LocalPort = HTTPPort ' set this to the port you want the server to listen on...
        sck0.Listen()
        Application.DoEvents()

        Dim state As AxMSWinsockLib.AxWinsock.State

        'If sck0.State = sckListening Then lblFileProgress(0) = "00 Listening"
        If sck0.CtlState = MSWinsockLib.StateConstants.sckListening Then lblFileProgress0.Text = "00 Listening"
    End Sub

    Public Sub StopHTTP()
        On Error Resume Next
        'sck(0).Close
        sck0.Close()
    End Sub

End Class