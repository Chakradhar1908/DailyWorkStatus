Public Class frmWinsock
    Private mBufferData As Boolean, mBufferText As String
    Public UseSSL As Boolean
    Private mDefaultTimeout As Integer, mTimeOutExpiry As Integer, mUnLoadOnClose As Boolean
    Public Function UploadString(ByVal Str As String, ByVal URL As String, Optional ByVal MimeType As String = "text/plain", Optional ByVal FileName As String = "file.txt") As String
        Dim strHttp As String
        Dim U As modURLHelper.URLExtract

        U = ExtractUrl(URL)
        ' build the HTTP request
        strHttp = BuildFileUploadRequest(Str, U, FileName, FileName, MimeType)

        ' assign the protocol host and port
        Protocol = MSWinsockLib.ProtocolConstants.sckTCPProtocol
        BufferData = True
        Connect(U.Host, 80)

        StartTimeout(15)
        Do While State <> MSWinsockLib.StateConstants.sckConnected
            Application.DoEvents()
            If Timeout Then Exit Do
        Loop

        If Not State = MSWinsockLib.StateConstants.sckConnected Then Exit Function
        SendData(strHttp)

        StartTimeout(10)
        Do While State = MSWinsockLib.StateConstants.sckConnected
            Application.DoEvents()
            If Timeout Then Exit Do
        Loop

        UploadString = BufferText
    End Function

    Public ReadOnly Property BufferText() As String
        Get
            BufferText = mBufferText
        End Get
    End Property

    Public Sub SendData(ByVal vData As String)
        If Not UseSSL Then
            Sock.SendData(vData)
        Else
            SSLSend(Sock, vData)
        End If
    End Sub

    Public ReadOnly Property Timeout() As Boolean
        Get
            If mTimeOutExpiry = 0 Then Exit Property                      ' Never timeout when timer not set up
            If GetTickCount() > mTimeOutExpiry Then Timeout = True        ' Querying checks the stop time
        End Get
    End Property

    Public ReadOnly Property State() As MSWinsockLib.StateConstants
        Get
            'State = Sock.State
            State = Sock.CtlState
        End Get
    End Property

    Private Sub StartTimeout(Optional ByVal Delay As Integer = -1, Optional ByVal Milliseconds As Integer = 0)
        If Delay = -1 Then Delay = DefaultTimeout
        mTimeOutExpiry = 0                                            ' clear it!
        Delay = Delay * 1000 + Milliseconds                           ' Usually, we want seconds... occasionally, milli.  Do the math first
        If Delay < 1 Then Exit Sub                                    ' Send 0 (zero) to clear it (stop the timer)
        mTimeOutExpiry = GetTickCount() + Delay                       ' Set when the timeout will expire
    End Sub

    Public Property DefaultTimeout() As Integer
        Get
            DefaultTimeout = mDefaultTimeout
        End Get
        Set(value As Integer)
            mDefaultTimeout = FitRange(1, value, 60)
        End Set
    End Property

    Public Sub Connect(Optional ByVal RemoteHost As String = "", Optional ByVal RemotePort As Integer = 0)
        Sock.Connect(RemoteHost, RemotePort)
    End Sub

    Public Property BufferData() As Boolean
        Get
            BufferData = mBufferData
        End Get
        Set(value As Boolean)
            ClearBuffer()
            mBufferData = value
        End Set
    End Property

    Public Sub ClearBuffer()
        mBufferText = ""
    End Sub

    Public Property Protocol() As MSWinsockLib.ProtocolConstants
        Get
            Protocol = Sock.Protocol
        End Get
        Set(value As MSWinsockLib.ProtocolConstants)
            Sock.Protocol = value
        End Set
    End Property

    Private Function BuildFileUploadRequest(ByRef strData As String, ByRef URL As modURLHelper.URLExtract, ByVal UploadName As String, ByVal FileName As String, ByVal MimeType As String) As String
        Dim strHttp As String ' holds the entire HTTP request
        Dim strBoundary As String 'the boundary between each entity
        Dim strBody As String ' holds the body of the HTTP request
        Dim lngLength As Integer ' the length of the HTTP request

        ' create a boundary consisting of a random string
        strBoundary = RandomAlphaNumString(10) '(32)

        ' create the body of the http request in the form
        '
        ' --boundary
        ' Content-Disposition: form-data; name="UploadName"; filename="FileName"
        ' Content-Type: MimeType
        '
        ' file data here
        '--boundary--
        strBody = "--" & strBoundary & vbCrLf
        strBody = strBody & "Content-Disposition: form-data; name=""txtFile""; filename=""" & FileName & """" & vbCrLf
        strBody = strBody & "Content-Type: " & MimeType & vbCrLf
        strBody = strBody & vbCrLf & strData
        strBody = strBody & vbCrLf & "--" & strBoundary & "--"

        ' find the length of the request body - this is required for the
        ' Content-Length header
        lngLength = Len(strBody)

        ' construct the HTTP request in the form:
        '
        ' POST /path/to/reosurce HTTP/1.0
        ' Host: host
        ' Content-Type: multipart-form-data, boundary=boundary
        ' Content-Length: len(strbody)
        '
        ' HTTP request body

        strHttp = "POST " & URL.URI & IIf(URL.Query <> "", "?" & URL.Query, "") & " HTTP/1.0" & vbCrLf
        strHttp = strHttp & "Host: " & URL.Host & vbCrLf
        strHttp = strHttp & "Content-Type: multipart/form-data; boundary=" & strBoundary & vbCrLf
        strHttp = strHttp & "Content-Length: " & lngLength & vbCrLf2
        strHttp = strHttp & strBody

        BuildFileUploadRequest = strHttp
    End Function
End Class