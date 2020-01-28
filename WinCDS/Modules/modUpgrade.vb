Module modUpgrade
    Public UpdateMsg As String, UpdateFileList As String
    Public UpgradeNoMessages As Boolean
    Public Function GetInstallDir(ByVal vKEY As String) As String
        Dim R As String
        If Left(vKEY, 1) = "$" Then
            Select Case LCase(vKEY)
                Case "$apppath", "$appfolder" : GetInstallDir = AppFolder()
                Case "$winsysdir", "$winsyspath"
                    R = GetWindowsDir() & "\SysWOW64"
                    If DirExists(R) Then
                        GetInstallDir = R
                    Else
                        GetInstallDir = GetWindowsSystemDir()
                    End If
                Case "$windir", "$winpath" : GetInstallDir = GetWindowsDir()
                Case "$fontdir", "$winfontdir" : GetInstallDir = SpecialFolder(FolderEnum.feWindowsFonts)
                Case "$px" : GetInstallDir = PXFolder()
                Case "$fx" : GetInstallDir = FXFolder()
                Case "$fxcontrol" : GetInstallDir = FXControlFolder()
                Case "$fxwallpaper" : GetInstallDir = FXWallpaperFolder()
                Case "$fxtaglayout" : GetInstallDir = TagLayoutFolder()
                Case "$invent" : GetInstallDir = InventFolder()
                Case "$store1" : GetInstallDir = StoreFolder(1)
                Case "$store2" : GetInstallDir = StoreFolder(2)
                Case "$store3" : GetInstallDir = StoreFolder(3)
                Case "$store4" : GetInstallDir = StoreFolder(4)
                Case "$store5" : GetInstallDir = StoreFolder(5)
                Case "$store6" : GetInstallDir = StoreFolder(6)
                Case "$store7" : GetInstallDir = StoreFolder(7)
                Case "$store8" : GetInstallDir = StoreFolder(8)
                Case "$store9" : GetInstallDir = StoreFolder(9)
                Case "$store10" : GetInstallDir = StoreFolder(10)
                Case "$store11" : GetInstallDir = StoreFolder(11)
                Case "$store12" : GetInstallDir = StoreFolder(12)
                Case "$store13" : GetInstallDir = StoreFolder(13)
                Case "$store14" : GetInstallDir = StoreFolder(14)
                Case "$store15" : GetInstallDir = StoreFolder(15)
                Case "$store16" : GetInstallDir = StoreFolder(16)
                Case "$store17" : GetInstallDir = StoreFolder(17)
                Case "$store18" : GetInstallDir = StoreFolder(18)
                Case "$store19" : GetInstallDir = StoreFolder(19)
                Case "$store20" : GetInstallDir = StoreFolder(20)
                Case "$store21" : GetInstallDir = StoreFolder(21)
                Case "$store22" : GetInstallDir = StoreFolder(22)
                Case "$store23" : GetInstallDir = StoreFolder(23)
                Case "$store24" : GetInstallDir = StoreFolder(24)
                Case "$store25" : GetInstallDir = StoreFolder(25)
                Case "$store26" : GetInstallDir = StoreFolder(26)
                Case "$store27" : GetInstallDir = StoreFolder(27)
                Case "$store28" : GetInstallDir = StoreFolder(28)
                Case "$store29" : GetInstallDir = StoreFolder(29)
                Case "$store30" : GetInstallDir = StoreFolder(30)
                Case "$store31" : GetInstallDir = StoreFolder(31)
                Case "$store32" : GetInstallDir = StoreFolder(32)
                Case "$cdsdata" : GetInstallDir = CDSDataFolder()
                Case "$qbcommon" : GetInstallDir = LocalProgramFilesFolder() & "Common Files\Intuit\QuickBooks\"
                Case "$pr" : GetInstallDir = PRFolder(, True)
                Case "$gl" : GetInstallDir = GLFolder(, True)
                Case "$ap" : GetInstallDir = APFolder(, True)
                Case "$bk" : GetInstallDir = BKFolder(, True)
                Case Else : GetInstallDir = ""
            End Select
        Else
            If vKEY = "" Then vKEY = AppFolder()   ' default to home dir
            GetInstallDir = vKEY
        End If
        GetInstallDir = CleanDir(GetInstallDir)
    End Function

    Public Function UpdateLog(ByVal M As String) As Boolean
        LogFile tLOG, M, False
End Function

    Public Function CurrentVersionURL() As String
        Dim A As String

        A = ""
        A = A & WebUpdateURL
        A = A & "currentversion.php"
        A = A & "?a=" & CDbl(Now)
        A = A & "&c=" & ProtectValueForURL(GetLocalComputerName)
        A = A & "&s=" & ProtectValueForURL(Trim(LCase(StoreSettings(1).Name)))
        A = A & "&k=" & LCase(License)
        A = A & "&p=" & "" '### Password -- Passed but not checked into StoreAllowed()..
        A = A & "&d=" & DateFormat(GetCurrentEXEDate(True), "-")
        A = A & ExtraURLParams

        CurrentVersionURL = URLEncode(A)
    End Function

    Public Function CurrentVersionLCL() As String
        CurrentVersionLCL = UpdateFolder() & "CurrentVersion.xml"
    End Function

    Public Function UMsgBox(ByRef Msg As String, Optional ByVal Style As VbMsgBoxStyle = vbInformation, Optional ByVal Caption As String = UMSG_TITLE_STAT) As VbMsgBoxResult
        ' here for automation..
        ' change this to an appropriate value to supress all msg boxes
        Dim MaxDur As Long
        Select Case Left(Msg, 8)
            Case "There ar" : MaxDur = 5
            Case "Note:  T" : MaxDur = 10
            Case "Please c" : MaxDur = 15
            Case Else : MaxDur = 30
        End Select

        UpdateLog "UMsgBox - " & Msg
  If Not UpgradeNoMessages Then UMsgBox = MsgBox(Msg, Style, Caption, , , MaxDur)
    End Function

    Public Function ScheduledUpdateToday(Optional ByVal D As Date = NullDate, Optional ByRef StoreName As String = "#") As Boolean
        If D = NullDate Then D = Date
        ScheduledUpdateToday = (Weekday(D, cdsFirstDayOfWeek) = UpdateDay(StoreName))
    End Function

    Public Function InstallUpgrade(ByVal fName As String, ByVal Path As String, ByVal DestPath As String, ByVal Install As String) As Boolean

        If UCase(fName) = WinCDSEXEName(True, True, True) Then ' WINCDS.EXE
            ' This will likely end the program, and will not return...
            If Not UpgradeReplaceWinCDS(fName, Path, DestPath) Then Exit Function
        Else
            If Not UpgradeReplaceFile(fName, Path, DestPath, True) Then Exit Function
        End If

        UpgradePerformInstall fName, DestPath, Install

  Select Case UCase(Right(fName, 3))
            Case "ZIP" : UpdateMsg = UpdateMsg & IIf(Len(UpdateMsg) > 0, ",", "") & "support modules"
            Case "CHM" : UpdateMsg = UpdateMsg & IIf(Len(UpdateMsg) > 0, ",", "") & "help file"
            Case "EXE" : UpdateMsg = UpdateMsg & IIf(Len(UpdateMsg) > 0, ",", "") & "accounting modules"
            Case "DLL" : UpdateMsg = UpdateMsg & IIf(Len(UpdateMsg) > 0, ",", "") & "program features"
            Case "INI" : UpdateMsg = UpdateMsg & IIf(Len(UpdateMsg) > 0, ",", "") & "configuration values"
            Case "TTF" : UpdateMsg = UpdateMsg & IIf(Len(UpdateMsg) > 0, ",", "") & "fonts"
        End Select
        UpdateFileList = UpdateFileList & IIf(Len(UpdateFileList) > 0, vbCrLf, "") & fName
        InstallUpgrade = True
    End Function

End Module
