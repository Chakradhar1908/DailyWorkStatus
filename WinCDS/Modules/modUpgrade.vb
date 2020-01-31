Module modUpgrade
    Public UpdateMsg As String, UpdateFileList As String
    Public UpgradeNoMessages As Boolean
    Private Const tLOG As String = "upgrade"
    Private Const UMSG_TITLE_STAT As String = "Upgrade Status"
    Private Const UMSG_TITLE_FAIL As String = "Upgrade Failure"
    Public InstallFontName As String
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
        LogFile(tLOG, M, False)
    End Function

    Public Function CurrentVersionURL() As String
        Dim A As String

        A = ""
        A = A & WebUpdateURL
        A = A & "currentversion.php"
        'A = A & "?a=" & CDbl(Now)
        A = A & "?a=" & Now
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

    Public Function UMsgBox(ByRef Msg As String, Optional ByVal Style As VBA.VbMsgBoxStyle = vbInformation, Optional ByVal Caption As String = UMSG_TITLE_STAT) As VBA.VbMsgBoxResult
        ' here for automation..
        ' change this to an appropriate value to supress all msg boxes
        Dim MaxDur As Integer
        Select Case Left(Msg, 8)
            Case "There ar" : MaxDur = 5
            Case "Note:  T" : MaxDur = 10
            Case "Please c" : MaxDur = 15
            Case Else : MaxDur = 30
        End Select

        UpdateLog("UMsgBox - " & Msg)
        If Not UpgradeNoMessages Then UMsgBox = MsgBox(Msg, Style, Caption, , , MaxDur)
    End Function

    Public Function ScheduledUpdateToday(Optional ByVal D As Date = NullDate, Optional ByRef StoreName As String = "#") As Boolean
        If D = NullDate Then D = Today
        ScheduledUpdateToday = (Weekday(D, cdsFirstDayOfWeek) = UpdateDay(StoreName))
    End Function

    Public Function InstallUpgrade(ByVal fName As String, ByVal Path As String, ByVal DestPath As String, ByVal Install As String) As Boolean
        If UCase(fName) = WinCDSEXEName(True, True, True) Then ' WINCDS.EXE
            ' This will likely end the program, and will not return...
            If Not UpgradeReplaceWinCDS(fName, Path, DestPath) Then Exit Function
        Else
            If Not UpgradeReplaceFile(fName, Path, DestPath, True) Then Exit Function
        End If

        UpgradePerformInstall(fName, DestPath, Install)

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

    Public Function NotifyUpgradeURL() As String
        NotifyUpgradeURL = CurrentVersionURL() & "&notify=1"
        '  Dim A As String
        '
        '  A = ""
        '  A = A & WebUpdateURL
        '  A = A & "CurrentVersion."
        '  A = A & "?a=" & CDbl(Now)
        '  A = A & "&c=" & ProtectValueForURL(GetLocalComputerName)
        '  A = A & "&s=" & ProtectValueForURL(LCase(Trim(StoreSettings(1).Name)))
        '  A = A & "&k=" & LCase(modStores.License)
        '  A = A & "&p=" & "" '### Password -- Passed but not checked into StoreAllowed()..
        '  A = A & "&d=" & DateFormat(GetCurrentEXEDate(True), "-")
        '  A = A & "&notify=1"
        '  A = A & ExtraURLParams
        '
        '  NotifyUpgradeURL = URLEncode(A)
    End Function

    Public Function ExtraURLParams() As String
        ExtraURLParams = ""
        'Exit Property  ' These are mostly for debugging..  They may be able to be disabled.
        ExtraURLParams = ExtraURLParams & "&version=" & SoftwareVersion(False)
        '  ExtraURLParams = ExtraURLParams & "&hash=" & SoftwareVersionHash()
        ExtraURLParams = ExtraURLParams & "&localtime=" & DateTimeStamp()
        '  ExtraURLParams = ExtraURLParams & "&dbg=" & IIf(UpgradeNoMessages, "S", "M") & IIf(ScheduledUpdateToday, "S", "M") & IIf(Not IsFormLoaded("frmUpgrade"), "S", "M")
        ExtraURLParams = ExtraURLParams & "&computer=" & GetLocalComputerName()
        ExtraURLParams = ExtraURLParams & "&osver=" & GetWinVerNumber()
    End Function

    Public Function UpdateDay(Optional ByVal SN As String = "#") As Integer
        On Error Resume Next
        If SN = "#" Then SN = StoreSettings(1).Name
        SN = UCase(SN)

        ' Updates are always run in the morning between 3-5a
        ' We don't want Saturday or Sunday morning, so we don't use those.
        Select Case Asc(Left(SN, 1))                        ' Monday - Friday  (1=sunday, 7 = saturday)
            Case "65" To "69" : UpdateDay = vbMonday          ' Monday Morning @ 3a
            Case "70" To "74" : UpdateDay = vbTuesday
            Case "75" To "79" : UpdateDay = vbWednesday
            Case "80" To "84" : UpdateDay = vbThursday
            Case "85" To "90" : UpdateDay = vbFriday          ' Friday Morning @ 3a
            Case Else : UpdateDay = vbFriday          ' If not a letter, use Friday
        End Select
    End Function

    Private Function UpgradeReplaceWinCDS(ByVal fName As String, ByVal Path As String, ByVal DestPath As String, Optional ByVal ReplaceEvenIfNewer As Boolean = True) As Boolean
        Dim toCopy As String, toReplace As String

        On Error Resume Next
        toCopy = Path & fName
        toReplace = DestPath & fName

        ' This sets up the WinCDS.Replace.exe file.
        ' It will handle the semafore, kill processes, and anything else needed.
        '  SetReplaceWinCDS toCopy
        '  ShellOut.RunFile toCopy
        MainModule.RestartProgram
        '  MainMenu.ShutDown True
        End  ' just in case (and visibility here)...  not needed b/c it's in MM.ShutDown



        '  If IsWin5 Then      ' This handles Windows XP...
        '    ShellOut.ShellOut AppFolder & WaitEXE & " -x"
        '    MainMenu.ShutDown True
        '    End
        '  End If
        '
        '' if WinCDS is less than 1MB, assume the download failed.
        '  If FileLen(toCopy) < FileSize_1MB Then Exit Function
        '
        '  UMsgBox "Please click OK and wait." & vbCrLf & "The program will restart automatically in approximately 60 seconds." & vbCrLf2 & "See yellow button below."
        '
        '' the program checks every 20 seconds for the shutdown semafore file so that should give
        '' all copies ample time to shut themselvs down.  then we try the KillProcess method.
        '' then we shutdown ourselves if we are still running..
        '' by the time wait.exe completes, all copies of WinCDS should be exited so that
        '' it can do the copy
        ''    ShellOut.ShellOut AppFolder & "wait.exe -s 60 cmd /c copy /Y """ & W & """ """ & X & """;" & X
        '
        '  If VerifyWaitScheduledTask Then
        '
        '    UpdateLog "Using Scheduled Task Wait Method (" & ReadStoreSetting(1, iniSection_Program, "SchTsk_Wait", "0") & ")..."
        '    If CommandLineUpdate Then
        '      ShellOut.ShellOut WaitEXE & " -x -y -q -mService"
        ''        RunWinCDSTask tskWait
        '    Else
        '      ' We must get around UAC.  We use a scheduled task in the current user run at highest privs.
        '      ' The requirement is that we copy WinCDS to program files folder.  But, we must also re-launch
        '      ' WinCDS in the current state, usually without ELEVATION, but we don't really know.
        '      ' The way to launch an ELEVATED task without UAC prompts is a scheduled task.
        '      ' But this can't re-launch the program without elevation (necessary for AutoVNC among other things)
        '      '    1.  The Scheduled task launches a hidden wait procedure ELEVATED to copy to programs folder
        '      '    2.  The Current user launches a second visible wait at USER PRIVS, and will relaunch WinCDS at same privs
        '      RunWinCDSTask tskWait
        '      ShellOut.ShellOut WaitEXE & " -c -s 40 " & GetShortName(WinCDSEXEFile(True))
        '    End If
        '  Else
        '    UpdateLog "Using Command Line Wait Method (" & ReadStoreSetting(1, iniSection_Program, "SchTsk_Wait", "0") & ")..."
        '    If CommandLineUpdate Then
        '      ShellOut.ShellOut WaitEXE & " -x -y -q -mService"
        '    Else
        '      ShellOut.ShellOut WaitEXE & " -x"
        '    End If
        '  End If
        '  ShutdownSemaforeFile CreateIt:=True
        '
        '  Dim R As Date
        '  R = DateAdd("s", 30, Now) ' this should cause it to wait 30 seconds before killing
        '  Do While DateAfter(Now, R, , "s"): DoEvents: Loop
        '  KillProcess WinCDSEXEName(True, True, True)
        '  MainMenu.ShutDown True
        '  End  ' just in case (and visibility here)...  not needed b/c it's in MM.ShutDown

    End Function

    Private Function UpgradeReplaceFile(ByVal fName As String, ByVal Path As String, ByVal DestPath As String, Optional ByVal ReplaceEvenIfNewer As Boolean = True) As Boolean
        Dim toCopy As String, toReplace As String

        On Error Resume Next
        toCopy = Path & fName
        toReplace = DestPath & fName

        If FileExists(toReplace) Then
            ' if they're the same size and version (non-version'd files are ""), do not copy..
            If FileLen(toCopy) = FileLen(toReplace) And FileVersion(toCopy) = FileVersion(toReplace) Then
                Debug.Print("UpgradeReplaceFile - " & GetFileBase(fName) & " - New and existing files are identical, skipped.")
                Exit Function
            End If

            If ReplaceEvenIfNewer Then
                If DateAfter(FileDateTime(toReplace), FileDateTime(toCopy)) Then
                    Debug.Print("UpgradeReplaceFile - " & GetFileBase(fName) & " - Existing file is newer, skipped.")
                    Exit Function
                End If
            End If

            DeleteFileIfExists(toReplace)          ' remove any previous file
        End If

        If FileExists(toReplace) Then
            UMsgBox("Could not remove previous version." & vbCrLf & toReplace & vbCrLf & "This file cannot be upgraded.", vbCritical, "Upgrade Failure")
            Exit Function
        End If

        FileCopy(toCopy, toReplace)

        ' Verify copy
        If Not FileExists(toReplace) Then
            UMsgBox("Could not copy over new version." & vbCrLf & toReplace & vbCrLf & "This file cannot be upgraded.", vbCritical, UMSG_TITLE_FAIL)
            Exit Function
        End If

        ' Verify file size
        If FileLen(toCopy) <> FileLen(toReplace) Then
            UMsgBox("New file size does not match after upgrade." & vbCrLf & toCopy & " -> " & toReplace & vbCrLf & "The file update liklely failed.", vbCritical, UMSG_TITLE_FAIL)
            Exit Function
        End If
        '
        UpgradeReplaceFile = True
    End Function

    Public Function UpgradePerformInstall(ByVal fName As String, ByVal DestPath As String, ByVal Install As String) As Boolean
        Dim toReplace As String, toReplaceBase As String
        Dim fBase As String, fExt As String
        Dim SysDir As String
        Dim T As String, R As String, Tmp As String

        DestPath = CleanDir(DestPath)

        fBase = GetFileBase(fName)
        fExt = GetFileExt(fName)
        toReplace = GetShortName(DestPath & fName)
        toReplaceBase = GetFileBase(toReplace, True, True) ' Keeps short name, but without extension

        Select Case LCase(Install)
            Case "$dllselfregister"
                PushDir(DestPath)
                RunFileAsAdmin(RegSvr32EXE & " /s " & fBase)
                PopDir
            Case "$tlbregister"
                PushDir(DestPath)
                ShellAndWait(RegAsmEXE & " /tlb:" & fBase & ".tlb /nologo")
                PopDir
            Case "$tlb"
                PushDir(GetWindowsSystemDir)
                ShellAndWait(RegTLIBEXE & " -q " & fBase)
                PopDir
            Case "$2com"
                ' regasm /tlb:optimroute.tlb optimroute.dll /nologo
                ' gacutil /i optimroute.dll /nologo
                'C:\WINDOWS\MICROS~1.NET\FRAMEW~1\V40~1.303\RegAsm.exe /tlb:optimroute.tlb optimroute.dll /nologo
                'C:\WINDOWS\MICROS~1.NET\FRAMEW~1\V11~1.432\GACUtil.exe /i optimroute.dll /nologo

                PushDir(DestPath)
                ShellAndWait(RegAsmEXE & " /tlb:" & fBase & ".tlb " & fBase & ".dll /nologo")
                'BFH20170119
                ' Apparently, the gacutil is only part of the SDK, and is not very useful...
                ' The above should be sufficient..
                '      ShellAndWait GACUtilEXE & " /i " & fBase & ".dll /nologo"
                PopDir
            Case "$unzip"
                frmBackUpGeneric.UnzipFiles(toReplace, DestPath, False)
                'Unload frmBackUpGeneric
                frmBackUpGeneric.Close()
                If UCase(T) = "AUTOVNC" Then CreateShortcutforAutoVNC
            Case "$exe"
                ShellOut.ShellOut(toReplace)
            Case "$installini"
                InstallINIToStoreSettings(toReplace, FitRange(1, Val(Right(GetFileBase(T), 2)), ActiveNoOfLocations))
            Case "$fontregister"
                InstallFontTTFToWindows(toReplace, InstallFontName)
            Case Else
                ' no special install required...
        End Select

        UpgradePerformInstall = True
    End Function

    Public Sub CreateShortcutforAutoVNC()
        Dim F As Object, MyShortcut As Object
        Dim IconFileName As String, TargetDir As String, EXEFileName As String

        Const OldLinkFileName As String = "Jerry's Connect.lnk"

        On Error GoTo NoDesktopIcon

        DeleteFileIfExists(AllUsersDesktopFolder & OldLinkFileName)

        F = CreateObject("WScript.Shell")
        IconFileName = AllUsersDesktopFolder & ProgramName & " Connect.lnk"
        TargetDir = WinCDSAutoVNCFolder
        EXEFileName = TargetDir & "Connect.cmd"

        If Dir(IconFileName) = "" Then
            MyShortcut = F.CreateShortcut(IconFileName) ' Create a shortcut object on the shared desktop
            MyShortcut.TargetPath = F.ExpandEnvironmentStrings(EXEFileName)
            MyShortcut.WorkingDirectory = TargetDir
            MyShortcut.WindowStyle = 4
            'Let MyShortcut.Description = "Description"
            MyShortcut.IconLocation = EXEFileName & ", 0" 'Put icon in here (optional)
            MyShortcut.Save
        End If
        F = Nothing
        Exit Sub

NoDesktopIcon:
        'Special Folders http://msdn.microsoft.com/library/default.asp?url=/library/en-us/script56/html/wsprospecialfolders.asp
    End Sub

    Public Function RegSvr32EXE() As String
        Dim R As String

        R = GetWindowsDir() & "\SysWOW64"
        If DirExists(R) Then
            RegSvr32EXE = R
        Else
            RegSvr32EXE = GetWindowsSystemDir()
        End If

        RegSvr32EXE = RegSvr32EXE & "\RegSvr32.exe"

        If FileExists(RegSvr32EXE) Then RegSvr32EXE = GetShortName(RegSvr32EXE)

        If Not FileExists(RegSvr32EXE) Then RegSvr32EXE = ""
    End Function

    Public Function RegAsmEXE() As String
        ' What we have to account for:
        '   Windows Dir
        '   .NET Version
        '   x86 vs x64
        Const RegASM As String = "RegAsm.exe"

        Const MSNETFW As String = "Microsoft.NET\Framework\"
        Const v1_0 As String = "v1.0.3705"
        Const v1_1 As String = "v1.1.4322"
        Const v2_0 As String = "v2.0.50727"
        Const v3_0 As String = "v3.0"
        Const v3_5 As String = "v3.5"
        Const v4_0 As String = "v4.0.30319"


        Dim P As String, T As String
        Dim Versions() As Object, L As Object

        P = CleanDir(GetWindowsDir) & MSNETFW
        'Versions = Array(v4_0, v3_5, v3_0, v2_0, v1_1, v1_0)
        Versions = {v4_0, v3_5, v3_0, v2_0, v1_1, v1_0}

        For Each L In Versions
            T = P & L & DIRSEP & RegASM
            If FileExists(T) Then
                RegAsmEXE = GetShortName(T)
                Exit Function
            End If
        Next
    End Function

    Public Function RegTLIBEXE() As String
        RegTLIBEXE = GetShortName(GetWindowsDir(True) & "RegTLIB.exe")
    End Function
End Module
