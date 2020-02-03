Imports VBA
Module modDevelopment
    Private MANUAL_DEV_MODE As Boolean
    Public Const EXE_CHANNEL_PRODUCTION As String = "production"
    Public Const EXE_CHANNEL_BETA As String = "beta"
    Public Const EXE_CHANNEL_ID_BETA As String = "0"
    Public Const EXE_CHANNEL_ALPHA As String = "alpha"
    Public Const EXE_CHANNEL_ID_ALPHA As String = "1"
    Public Const EXE_CHANNEL_DEV As String = "dev"
    Public Const EXE_CHANNEL_ID_DEV As String = "2"

    Private Const UnitCount as integer = 11

    Public Function IsCDSComputer(Optional ByRef Unit As String = "", Optional ByVal SetIt As Boolean = False) As Boolean
        ' it may be a good idea to change these values every now and again to prevent their abuse.
        ' resetting all of them will result in no computers registering as CDS Computers.
        ' You would probably want to re-initiate the values onto the relevant computers after doing so.
        ' used function :?createuniqueid
        '
        ' Network Scan:
        '* 10.16.0.15  2C:D0:5A:17:B4:A2 47     JERRY-LAPTOP
        '* 10.16.0.26  B8:E8:56:A9:33:F2 11     (??)
        '* 10.16.0.11  00:15:5D:05:17:01 2      WEBSERVER1
        '* 10.16.0.23  00:0D:C5:11:4C:8C 2      (Echostar Global)
        '* 10.16.0.20  10:78:D2:DB:24:F7 3      INVENTORY2
        '* 10.16.0.2   00:15:5D:05:17:00 0      DC1
        '* 10.16.0.18  68:A8:6D:3F:BE:9C 3      MACBOOKPRO-8354
        '* 10.16.0.10  E0:69:95:EB:DF:E4 0      (virtual server 1)
        Dim GID As String = "", X As String, Y As String, Z As String
        Dim F As String, G As String, I as integer, Rec As String = ""
        Dim Mac As String, Ma2 As String

        Const UnitCount as integer = 12

        Select Case UCase(Unit)
            Case "1", "LAPTOP", "JERRY-LAPTOP"
                Rec = "LAPTOP"
                GID = "{00596825-6A46-AE48-8984-1E21EA03C430}"
                Mac = "2C:D0:5A:17:B4:A2"
            Case "2", "INVENTORY", "INVENTORY2"
                Rec = "INVENTORY2"
                GID = "{A1B8761D-C87B-E442-8F19-06F376D408AE}"
                Mac = "10:78:D2:DB:24:F7"
            Case "3", "DC1"
                Rec = "DC1"
                GID = "{640F8EE6-DB41-804A-8173-BFB239DBFB11}"
                Mac = "00:15:5D:05:17:00"
            Case "4", "WEBSERVER", "WEBSERVER1"
                Rec = "WEBSERVER1"
                GID = "{0B940077-202F-834B-AE14-C77702F8A97A}"
                Mac = "00:15:5D:05:17:01"
            Case "5", "MACBOOKPRO"
                Rec = "MACBOOKPRO"
                GID = "{2A134559-B6E9-FB41-9AEB-15000B6D4B26}"
                Mac = "68:A8:6D:3F:BE:9C"
            Case "6", "HOOGTERP-HP"
                Rec = "HOOGTERP-HP"
                GID = "{746E4FDD-19D6-CA4F-9F9B-FC342ACA15BE}"
                Mac = "68:A3:C4:09:33:00"
            Case "7", "AZARIAH"
                Rec = "AZARIAH"
                GID = "{48A354C7-83E7-6E44-9D3E-9ECF20F8B9AD}"
                Mac = "00:1C:DF:34:42:04"
                Ma2 = "6C:62:6D:83:83:31"
            Case "8", "GLORY"
                Rec = "GLORY"
                GID = "{DE8183D6-0F07-2D41-A6C6-2DBFE4E50064}"
                Mac = "00:0F:66:F2:B6:28"
                Ma2 = "00:0D:56:5F:E3:E2"
            Case "9", "BEN-LAPTOP"
                Rec = "BEN-LAPTOP"
                GID = "{D447E13C-B537-CA47-AC73-FA24DD42F2C6}"
                Mac = "00:04:23:6B:0B:6F"
            Case "10", "KRISHNA"
                Rec = "KRISHNA"
                GID = "{A504A7FE-0521-3748-9C08-6CA4FA9159F4}"
                Mac = "C8:D3:FF:34:F1:72"
            Case "11", "WINCDSDEV1"
                Rec = "WINCDSDEV1"
                GID = "{E29CEB4E-A0C1-524D-B7B7-E7F33C22ED18}"
                Mac = "18:60:24:71:1E:09"
            Case "12", "CHANDARKAR"
                Rec = "CHAKRADHAR"
                GID = "{77EC3AEE-3417-2D4D-A8EF-D1092EC938C3}"
                Mac = ""
            Case "PROTOTYPE" : IsCDSComputer = IsCDSComputer(1) Or IsCDSComputer(6) Or IsCDSComputer(7) Or IsCDSComputer(8) Or IsCDSComputer(9) : Exit Function
            Case "0", "", "NONE"
                If Not SetIt Then
                    For I = 1 To UnitCount
                        Rec = I
                        If IsCDSComputer(Rec) Then Unit = Rec : IsCDSComputer = True : Exit Function
                    Next
                    Exit Function
                End If
                Rec = ""
                GID = ""
                Mac = BlankMacAddress()
                Ma2 = BlankMacAddress()

            Case Else : Err.Raise("Unknown CDS Station: " & Unit)
        End Select

        F = LocalCDSDataFolder()
        G = WinCDSFolder() & "Access.dat"
        If SetIt Then
            If GID <> "" Then
                '      WriteStoreSetting 1, iniSection_StoreSettings, "CompID", GID
                WriteFile(F, GID, True, True)
                WriteFile(G, GID, True, True)
                SaveCDSSetting("CompID", GID, , True)
            Else
                On Error Resume Next
                '      WriteStoreSetting 1, iniSection_StoreSettings, "CompID", GID
                Kill(F)
                Kill(G)
                SaveCDSSetting("CompID", "", , True)
            End If
        End If
        '  X = Trim(ReadFile(F))

        X = GID
        Y = GetCDSSetting("CompID", "", "", True)
        Z = Trim(ReadFile(G))

        IsCDSComputer = (GID <> "") And (X = GID) And (Y = GID) And (Z = GID)

        '  If IsCDSComputer Then
        'BFH20141110 Mac Address Check caused immediate crash on several computers
        '    If Mac <> BlankMacAddress And Mac <> GetMacAddress And Ma2 <> GetMacAddress Then
        ''      If MsgBox("Physical Machine is recorded as CDS Computer [" & rec & "] but does not match physical address." & vbCrLf & "Allow Authentication?", vbOKCancel, "CDS Computer Authentication", , , , , BACKDOOR_PASSWORD, False) = vbCancel Then
        '        IsCDSComputer "", True
        '        IsCDSComputer = False
        '        rec = ""
        '        Exit Function
        ''      End If
        '    End If
        Unit = Rec
        '  End If
    End Function
    Public Function IsDevelopment() As Boolean
        On Error Resume Next
        IsDevelopment = False
        If IsDevelopmentMANUAL() Then IsDevelopment = True : Exit Function
        If IsDevelopmentSTANDARD() Then IsDevelopment = True : Exit Function
        '  If IsDevChannel Then IsDevelopment = True: Exit Function
        Err.Clear()
    End Function
    Public Function IsDevelopmentMANUAL(Optional ByVal doSet As VbTriState = vbUseDefault) As Boolean
        If doSet = vbTrue Then MANUAL_DEV_MODE = True
        If doSet = vbFalse Then MANUAL_DEV_MODE = False

        IsDevelopmentMANUAL = MANUAL_DEV_MODE
    End Function
    Public Function IsDevelopmentSTANDARD() As Boolean
        '  IsDevelopmentSTANDARD = ReadFile(UpdateFolder & "DEV.TXT") = "DEVELOPMENT"
        IsDevelopmentSTANDARD = FileExists(UpdateFolder & "DEV.TXT")
    End Function
    Public Function ExeChannelDescriptor(Optional ByVal DevChannel As String = "production") As String
        Select Case LCase(DevChannel)
            Case EXE_CHANNEL_BETA : ExeChannelDescriptor = EXE_CHANNEL_ID_BETA
            Case EXE_CHANNEL_ALPHA : ExeChannelDescriptor = EXE_CHANNEL_ID_ALPHA
            Case EXE_CHANNEL_DEV : ExeChannelDescriptor = EXE_CHANNEL_ID_DEV
            Case Else : ExeChannelDescriptor = ""
        End Select
    End Function

    Public Function IsBetaChannel() As Boolean
        IsBetaChannel = WinCDSRevisionNumber() = EXE_CHANNEL_ID_BETA
    End Function

    '###EXECHANNEL
    Public Function ExeChannelName(Optional ByVal RevisionNumber As String = "#") As String
        If RevisionNumber = "#" Then RevisionNumber = WinCDSRevisionNumber()

        Select Case RevisionNumber
            Case EXE_CHANNEL_ID_BETA : ExeChannelName = EXE_CHANNEL_BETA
            Case EXE_CHANNEL_ID_ALPHA : ExeChannelName = EXE_CHANNEL_ALPHA
            Case EXE_CHANNEL_ID_DEV : ExeChannelName = EXE_CHANNEL_DEV
            Case Else : ExeChannelName = EXE_CHANNEL_PRODUCTION
        End Select
    End Function

    '###EXECHANNEL
    Public Function ExeChannelNameColor(Optional ByVal RevisionNumber As String = "#") As Color
        If RevisionNumber = "#" Then RevisionNumber = WinCDSRevisionNumber()

        Select Case RevisionNumber
            Case EXE_CHANNEL_ID_BETA : ExeChannelNameColor = Color.FromArgb(128, 128, 255)
            Case EXE_CHANNEL_ID_ALPHA : ExeChannelNameColor = Color.Red
            Case EXE_CHANNEL_ID_DEV : ExeChannelNameColor = Color.Yellow
            Case Else : ExeChannelNameColor = Color.FromArgb(128, 128, 255)
        End Select
    End Function

    Public Function SetDevMode(Optional ByVal State As Integer = 0) As Boolean
        Select Case State
            Case 0 : MANUAL_DEV_MODE = False : KillDevModeFile()
            Case 1 : MANUAL_DEV_MODE = False : CreateDevModeFile()
            Case 2 : MANUAL_DEV_MODE = True : KillDevModeFile()
        End Select
    End Function

    Public Function KillDevModeFile() As Boolean
        On Error Resume Next
        KillDevModeFile = IsDevelopmentSTANDARD()
        Kill(UpdateFolder() & "DEV.TXT")
        If KillDevModeFile Then KillDevModeFile = Not IsDevelopmentSTANDARD()
    End Function

    Public Function CreateDevModeFile() As Boolean
        WriteFile(UpdateFolder() & "DEV.TXT", "DEVELOPMENT", True, True)
    End Function

End Module
