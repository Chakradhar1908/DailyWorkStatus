Module modExternalIP
    Private mIPStore As String
    Private Const IPServer As String = "www.wincdspro.com"
    Private Const IPServerURL As String = "http://" & IPServer & "/ip.php"
    Private LastInternetCheck As Date
    Public Const INTERNET_URL_MONITOR As String = "http://www.google.com"

    Public ReadOnly Property IP_CONTROL() As Boolean
        Get
            IP_CONTROL = True
            'BFH20170227 - External IP Tracking temporarily disabled because the IP getter is lagging badly
            '  IP_CONTROL = False
        End Get
    End Property

    Public Function ExternalIPAddress() As String
        '::::ExternalIPAddress
        ':::SUMMARY
        ': External IP Address
        ':::DESCRIPTION
        ': This function is used to get External IP Address.
        ':::RETURN
        ': String
        On Error Resume Next
        If Not IP_CONTROL Then ExternalIPAddress = "[UNKNOWN]" : Exit Function

        If mIPStore <> "" Then
            ExternalIPAddress = mIPStore
            Exit Function
        End If

        'ExternalIPAddress = INETGET(IPServerURL)
        mIPStore = ExternalIPAddress
        If mIPStore = "" Then mIPStore = "[UNKNOWN]"

        ' Spoofing
        '  ExternalIPAddress = CDS_IP_ADDRESS
        '  ExternalIPAddress = BFH_IP_ADDRESS
        '  ExternalIPAddress = BFH_IP_ADDRESS2
    End Function

    Public Function MonitorInternet() As Boolean
        '::::MonitorInternet
        ':::SUMMARY
        ': Monitor Internet Up-ness
        ':::DESCRIPTION
        ': Check internet availability with self-limiting so it can be called in a non-limited loop.
        ':::RETURN
        ': Boolean - Returns True.

        Dim S As String
        If Not DateAfter(Now, DateAdd("m", 30, LastInternetCheck)) Then Exit Function

        S = DateTimeStamp() & ": " & ArrangeString(GetLocalComputerName, 16) & " - " & ArrangeString(ExternalIPAddress, 15) & " - " & ArrangeString(IIf(InternetIsConnected, "Online", "** OFFLINE"), 10)
        LogFile("Persistence", S)
        LastInternetCheck = Now
        MonitorInternet = True
    End Function

    Private Function InternetIsConnected() As Boolean
        InternetIsConnected = (DownloadURLToString(INTERNET_URL_MONITOR) <> "")
    End Function
End Module
