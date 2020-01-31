Module modExternalIP
    Private mIPStore As String
    Private Const IPServer As String = "www.wincdspro.com"
    Private Const IPServerURL As String = "http://" & IPServer & "/ip.php"
    Private LastInternetCheck As Date
    Public Const INTERNET_URL_MONITOR As String = "http://www.google.com"
    Private IPActions() As ExternalIPAddresss, IPCount As Integer

    Private Enum IPAction
        IPAct_NoAction
        IPAct_DisableLicense
        IPAct_SetStoreName
        IPAct_SetLicense
    End Enum
    Private Structure ExternalIPAddresss
        Dim IP As String
        Dim Computer As String
        Dim Identifier As String
        Dim StoreCount As Integer
        Dim Action As IPAction
    End Structure

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

    Public Function IPAddressIsBanned(Optional ByVal IP As String = "") As Boolean
        ':::SUMMARY
        ': Whether an IP address is banned
        ':::DESCRIPTION
        ': Returns true if the IP address is banned.
        ':::PARAMETERS
        ': - IP - Indicates the IP Address.
        ':::RETURN
        ': Boolean
        Dim I As Integer
        If IP = "" Then IP = ExternalIPAddress()

        InitAddressList()

        For I = 1 To IPCount
            If IPActions(I).Action = IPAction.IPAct_DisableLicense Then
                If IPActions(I).IP = IP Then IPAddressIsBanned = True : Exit Function
            End If
        Next
    End Function

    Private Function BanIPAddress(ByVal IP As String, Optional ByVal Reason As String = "DEMO") As Boolean
        BanIPAddress = TrackIPAddress(IP, Reason, IPAction.IPAct_DisableLicense)
    End Function

    Private Function ResetStoreName(ByVal IP As String, ByVal Computer As String, ByVal Name As String) As Boolean
        ResetStoreName = TrackIPAddress(IP, Name, IPAction.IPAct_SetStoreName, Computer)
    End Function

    Private Function TrackIPAddress(ByVal IP As String, ByVal ID As String, ByVal Action As IPAction, Optional ByVal Computer As String = "", Optional ByVal StoreCount As Integer = 0) As Boolean
        IPCount = IPCount + 1
        ReDim Preserve IPActions(0 To IPCount - 1)
        IPActions(IPCount - 1).IP = IP
        IPActions(IPCount - 1).Computer = Computer
        IPActions(IPCount - 1).Identifier = ID
        IPActions(IPCount - 1).Action = Action
        IPActions(IPCount - 1).StoreCount = StoreCount
        TrackIPAddress = True
    End Function

    Private Function InitAddressList() As Boolean
        '  BanIPAddress CDS_IP_ADDRESS
        '  BanIPAddress BFH_IP_ADDRESS
        '  BanIPAddress BFH_IP_ADDRESS2

        BanIPAddress("50.160.27.136")      ' RL-LAPTOP
        BanIPAddress("59.178.154.136")     ' RAVINDERCHOUHAN
        BanIPAddress("67.140.83.132") ' LYNDA-PC
        BanIPAddress("72.175.36.131")  ' OFFICE-PC
        BanIPAddress("73.41.136.24")   ' ASUS-PC
        BanIPAddress("75.44.90.105")    ' HOME
        '  BanIPAddress "76.23.26.79"        ' OWNER-15704AB8D, TIFFANY-PC
        BanIPAddress("107.10.52.6")     ' SHORVATH
        BanIPAddress("108.188.84.96")    ' LUXFURNITURE-PC
        BanIPAddress("184.166.101.1")      ' WIN10-THINK14
        '  BanIPAddress "204.228.144.205"    ' SOUTHPAW

        ResetStoreName("50.249.151.66", "OFFICE-3", "BF MYERS FURNITURE")
        ResetStoreName("67.42.171.194", "LAPTOP-3BQO50HQ", "ROCKY MOUNTAIN DESIGN WAREHOUSE")
        ResetStoreName("68.41.198.49", "WRKDT001B24Z8P1", "HOUSE OF BEDROOMS")
        ResetStoreName("68.56.21.168", "SHOWROOM", "HOUSE OF BEDROOMS KIDS")
        ResetStoreName("68.199.179.101", "GITI-HP", "DECORATIVE TOUCH PATIO")
        ResetStoreName("71.84.222.11", "ORDER", "BARRS FURNITURE")
        ResetStoreName("71.143.252.75", "MICKY-PC", "THORNTONS HOME FURNISHINGS")
        ResetStoreName("72.2.244.58", "ANNIE-PC", "CORVINS FURNITURE OF ETOWN")
        ResetStoreName("72.2.244.58", "NANCY-PC", "CORVINS FURNITURE OF ETOWN")
        ResetStoreName("72.177.112.34", "DOUBLER-PC", "DOUBLE R DRY GOODS")
        ResetStoreName("96.56.36.82", "IVORI-PC", "ROGERS FURNITURE")
        ResetStoreName("104.232.176.46", "LOTT-DEPOT-1-PC", "LOTT FURNITURE")
        ResetStoreName("173.13.88.54", "STATION4", "ADAMS FURNITURE")

        '  ResetStoreLicense "50.249.151.66", 3, "BF MYERS FURNITURE"

        ' ...
        InitAddressList = True
    End Function

End Module
