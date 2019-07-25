Module modExternalIP
    Private mIPStore As String
    Private Const IPServer As String = "www.wincdspro.com"
    Private Const IPServerURL As String = "http://" & IPServer & "/ip.php"

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

End Module
