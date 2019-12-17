Imports System.IO
Imports System.Reflection
Imports stdole
Module MainModule
    Private mIsServer As TriState           ' Cache this high-use, non-trivial value
    Public Allow_ADODB_Errors As Boolean      ' Allow the database to continue after errors - for debugging and special cases.
    Public Const WinCDS_ProjectFilename As String = "WinCDS.vbp"
    Public Const WinCDSEXE_Base As String = "WinCDS"
    Public Const WinCDSEXE As String = WinCDSEXE_Base & ".exe"
    Public gblLastDeliveryDate As Date        ' This will make the last delivery date persist without keeping whole forms loaded.
    Public gblLastDeliveryDateEpoch As Date   ' This will make the last delivery date reset daily
    Public PrvKill As Boolean

    Public Sub HideSplash()
        If frmSplashIsLoaded Then frmSplash.Hide()
    End Sub

    Public Function IsFormLoaded(ByVal FrmName As String) As Boolean
        IsFormLoaded = False
        If FrmName = "frmSplash" Then
            If IsFormLoaded("frmSplash1") Then IsFormLoaded = True
            If IsFormLoaded("frmSplash2") Then IsFormLoaded = True
            If IsFormLoaded Then Exit Function
        End If
        If FrmName = "MainMenu" Then
            If IsFormLoaded("MainMenu1") Then IsFormLoaded = True
            If IsFormLoaded("MainMenu4") Then IsFormLoaded = True
            If IsFormLoaded Then Exit Function
        End If
        IsFormLoaded = Not (FormIsLoaded(FrmName) Is Nothing)
    End Function

    Public Function FormIsLoaded(ByVal FrmName As String) As Form
        'On Error Resume Next
        FormIsLoaded = Nothing

        For Each f As Form In My.Application.OpenForms
            If UCase(f.Name) = UCase(FrmName) Then
                FormIsLoaded = f
                Exit Function
            End If
        Next
    End Function
    Public Function GetDatabaseAtLocation(Optional ByVal Store As Integer = 0) As String
        If Store <= 0 Then Store = StoresSld
        GetDatabaseAtLocation = NewOrderFolder(Store) & "CDSDATA.MDB"
    End Function
    Public Function NewOrderFolder(Optional ByVal StoreNum As Integer = 0) As String
        NewOrderFolder = StoreFolder(StoreNum) & "NewOrder\"
    End Function
    Public Function StoreFolder(Optional ByVal StoreNum As Integer = 0, Optional ByVal doLocal As Boolean = False) As String
        If StoreNum = 0 Then StoreNum = StoresSld
        If StoreNum > Setup_MaxStores Or StoreNum < 1 Then StoreNum = 1
        StoreFolder = GetStation(doLocal:=doLocal) & "Store" & StoreNum & DIRSEP
    End Function
    Private Function GetStation(Optional ByVal doLocal As Boolean = False, Optional ByVal doStub As Boolean = True) As String
        'MsgBox "GetStation"
        'LogStartup "GetStation(" & doLocal & ", " & doStub & ") a"
        If doLocal Then         ' VB doesn't short-circuit, and we need it to here...
            GetStation = LocalRoot
        Else
            GetStation = IIf(IsServer(), LocalRoot, RemoteRoot)
        End If
        'LogStartup "GetStation(" & doLocal & ", " & doStub & "): b"
        'MsgBox "GetStation=" & GetStation
        If doStub Then
            ' this fails for remote stores...
            ' if we put it in all users\documents, we would have to share the documents folder
            ' if we put it in all users\shared documents, win7 uses "public", not "shared documents"
            '   in addition, the whole folder is completely shared, and cannot be protected at all (or it defeats the purpose of this folder)
            ' it would make sense to use the AppData folder for local, but not for remote
            ' so this is probably still back-burnered.
            ' CDSData should work for now, but it does still leave it so we have to run in full admin
            ' on Vista+
            '    If GetStation = "Cx:\" Then
            '      If mDataPath = "" Then mDataPath = WinCDSDataPath     ' cache this
            '      If DirExists(mDataPath & "Store1\") Then GetStation = WinCDSDataPath
            '    End If

            If DirExists(GetStation & WinCDSStubDir) Then GetStation = GetStation & WinCDSStubDir
        End If
        'LogStartup "GetStation(" & doLocal & ", " & doStub & "): c"
    End Function

    Public ReadOnly Property WinCDSStubDir() As String
        Get
            WinCDSStubDir = "CDSData\"
        End Get
    End Property

    ' Check the registry for a server/station indicator. If it is not present, prompt for one.
    Public Function IsServer(Optional ByVal Optimize As Boolean = True) As Boolean
        'MsgBox "IsServer"

        ' vbFalse=0, so this is NOT SET value.  vbTrue is server, vbUseDefault is workstation
        ' If they want optimization, make sure it's set.  Otherwise, calculate the value.
        If Optimize And mIsServer <> vbFalse Then
            IsServer = IIf(mIsServer = vbTrue, True, False)
            Exit Function
        End If

        IsServer = GetIsServer()
        mIsServer = IIf(IsServer, vbTrue, vbUseDefault)
    End Function

    Private Function GetIsServer() As Boolean
        Dim strServer As String, SL As Boolean
        On Error Resume Next
        strServer = LCase(GetCDSSetting("IsServer", ""))

        Select Case LCase(strServer)
            Case "station" : GetIsServer = False
            Case "server" : GetIsServer = True
            Case Else
                'MsgBox "strServer=" & strServer
                'MsgBox "I:? " & IIf(FileExists("I:\Invent\CDSInvent.mdb"), "YES", "NO")
                'MsgBox "I:STUB? " & IIf(FileExists("I:\" & WinCDSStubDir & "Invent\CDSInvent.mdb"), "YES", "NO")
                'MsgBox "--"
                'MsgBox "Lock=" & ServerLock

                If Not ServerLock() Then
                    If FileExists(RemoteRoot & "Invent\CDSInvent.mdb") Or FileExists(RemoteRoot & WinCDSStubDir & "Invent\CDSInvent.mdb") Then  ' Assume it's a station, it's got I: mapped to a valid server.
                        SetServer(False)
                        GetIsServer = False
                    ElseIf FileExists(LocalRoot & "Invent\CDSInvent.mdb") Or FileExists(LocalRoot & WinCDSStubDir & "Invent\CDSInvent.mdb") Then ' Assume it's a server, since it's set up as such..
                        SetServer(True)
                        GetIsServer = True
                    ElseIf MsgBox("Is this computer the server?", vbQuestion + vbYesNo) = vbYes Then
                        SetServer(True)
                        GetIsServer = True
                    Else
                        SetServer(False)
                        GetIsServer = False
                    End If
                End If
                'MsgBox "GIS=" & GetIsServer
        End Select

        ' Make sure the server is mapped if it's a workstation
        If Not GetIsServer And Not DriveMapped(RemoteDriveLetter) Then
            If Not VerifyIsServerOnWorkstation() Then
                SetServer(True)
                GetIsServer = True
            End If
        End If
    End Function
    Public Function SetServer(ByVal Server As Boolean)
        'MsgBox "SetServer: " & Server
        SetServer = Nothing
        mIsServer = vbFalse
        SaveCDSSetting("IsServer", IIf(Server, "server", "station"))
    End Function

    Public Function ServerLock(Optional ByVal doSet As TriState = vbUseDefault, Optional ByRef bIsServer As TriState = vbUseDefault) As Boolean
        Const X = "ServerLock.txt"
        Const Y = "LOCK-SERVER"
        Const Z = "LOCK-WORKSTATION"
        Dim FN As String, T As String

        On Error GoTo ServerLockedFailed
        bIsServer = vbUseDefault

        ' This routine is called from IsServer..
        ' That means, we cannot use a folder derived from IsServer (aka, CDSData).
        ' But, we don't really want to bury this in app data, do we?  It would make it near impossible to clear..
        ' But, do we have a choice?
        FN = CDSAppDataFolder() & X

        If doSet = vbTrue Then
            WriteFile(FN, IIf(IsServer(), Y, Z), True)
            WriteStoreSetting(1, IniSections_StoreSettings.iniSection_StoreSettings, "ServerLock", "True")
        ElseIf doSet = vbFalse Then
            On Error Resume Next
            Kill(FN)
        End If

        ServerLock = FileExists(FN)
        T = ReadFile(FN)

        If T = Y Then bIsServer = vbTrue
        If T = Z Then bIsServer = vbFalse
ServerLockedFailed:
    End Function
    Public Function WriteStoreSetting(ByVal nStoreNo As Integer, ByVal nSection As IniSections_StoreSettings, ByVal nKey As String, ByVal nValue As String) As String
        If nStoreNo = -1 Then  ' Allow "broadcast" setting, to save to all stores at once.
            Dim I As Integer
            For I = 1 To ActiveNoOfLocations
                WriteStoreSetting(I, nSection, nKey, nValue)
            Next
        End If
        WriteIniValue(StoreINIFile(nStoreNo), StoreSettingSectionKey(nSection), nKey, nValue, True)
        WriteStoreSetting = nValue
        ResetStoreSettings()
    End Function

    Public Function GetStoreNumber(ByVal DBName As String) As Integer
        Dim X As Integer
        ' Inverse of GetDatabaseAtLocation, particularly useful for patches.
        If UCase(DBName) Like "?:*\INVENT\CDSINVENT.MDB" Then
            GetStoreNumber = -1   ' Return -1 for inventory DB
        ElseIf Not UCase(DBName) Like "?:*\STORE*\NEWORDER\CDSDATA.MDB" Then
            GetStoreNumber = 1    ' Invalid database name?
        Else
            On Error Resume Next
            X = InStr(LCase(DBName), "\store") + 6
            GetStoreNumber = CLng(Mid(DBName, X, InStr(X, DBName, DIRSEP) - X))
        End If
    End Function
    Public Function BackupSemaforeFile(Optional ByVal CreateIt As Boolean = False, Optional ByVal ItExists As Boolean = False, Optional ByVal DeleteIt As Boolean = False) As Boolean
        Dim FN As String
        Dim X As String
        FN = StoreFolder(1) & "backup.txt"

        On Error Resume Next
        BackupSemaforeFile = True
        If CreateIt Then WriteFile(FN, "" & Now, True)
        If DeleteIt Then Kill(FN)
        If ItExists Then


            X = Trim(ReadFile(FN))
            If Not FileExists(FN) Then      ' if file doesn't exist
                BackupSemaforeFile = False
            ElseIf Not IsDate(X) Then                        ' or if it isn't a date
                BackupSemaforeFile = False
            ElseIf Math.Abs(DateDiff("n", Now, DateAdd("n", 2, X))) > 2 Then ' or if it's more than a couple minutes old
                BackupSemaforeFile = False
            Else
                BackupSemaforeFile = True
            End If
        End If
    End Function
    Public Function LocalCDSDataFolder() As String
        ' We sometimes want to bypass the IsServer check...  This allows direct access to the C-drive without the overhead
        LocalCDSDataFolder = LocalRootFolder & WinCDSStubDir
        If DirExists(LocalCDSDataFolder) Then Exit Function
        LocalCDSDataFolder = LocalRootFolder
    End Function
    Private Function VerifyIsServerOnWorkstation() As Boolean
        VerifyIsServerOnWorkstation = False ' Start out as workstation

        ' We have to give the user a choice here to make this computer the server,
        ' or they'll be trapped forever.

        ' not sure what the consequences are, but jk wanted this removed because many customers kept
        ' having this fail at startup and people being careless and designating server..
        ' bfh20050715
        Dim DoAsk As Boolean, ConfirmM As String
        DoAsk = True

        If DoAsk Then
            If MsgBox(ProgramName & " can't connect to the server." & vbCrLf & "Is this computer the server?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
                ConfirmM = ""
                ConfirmM = ConfirmM & "--- WARNING --- WARNING --- WARNING --- WARNING --- WARNING --- WARNING ---" & vbCrLf
                ConfirmM = ConfirmM & "" & vbCrLf
                ConfirmM = ConfirmM & "You are about to change this computer to be the server!!!" & vbCrLf
                ConfirmM = ConfirmM & "" & vbCrLf
                ConfirmM = ConfirmM & "This computer was perviously setup as a workstation and." & vbCrLf
                ConfirmM = ConfirmM & "you have selected to make this computer the server." & vbCrLf
                ConfirmM = ConfirmM & "If this is not actually the server, this program will not function and you will" & vbCrLf
                ConfirmM = ConfirmM & "need to enter Store Setup to switch it back to workstation mode." & vbCrLf
                ConfirmM = ConfirmM & "" & vbCrLf
                ConfirmM = ConfirmM & "Chances are, you do not want this option." & vbCrLf
                ConfirmM = ConfirmM & "" & vbCrLf
                ConfirmM = ConfirmM & "Before pressing OK, please confirm this settings change." & vbCrLf
                ConfirmM = ConfirmM & "If this is a workstation, please press Cancel and confirm that your is up and that" & vbCrLf
                ConfirmM = ConfirmM & "drive I is mapped to the server before restarting the software." & vbCrLf
                ConfirmM = ConfirmM & "" & vbCrLf
                ConfirmM = ConfirmM & "--- WARNING --- WARNING --- WARNING --- WARNING --- WARNING --- WARNING ---" & vbCrLf
                ConfirmM = ConfirmM & vbCrLf
                ConfirmM = ConfirmM & "If this is the server, enter the word SERVER in the box below." & vbCrLf

                If MsgBox(ConfirmM, vbOKCancel + vbDefaultButton2, "CONFIRM SERVER") = vbOK Then
                    ' This is the only clause that doesn't immediately end the program
                    VerifyIsServerOnWorkstation = True
                    Exit Function
                Else
                    End
                End If
            Else
                ConfirmM = ""
                ConfirmM = ConfirmM & "This program is configured to run in workstation mode and cannot" & vbCrLf
                ConfirmM = ConfirmM & "be run without the network running and Drive I mapped to the server." & vbCrLf
                ConfirmM = ConfirmM & vbCrLf
                ConfirmM = ConfirmM & "Please either configure your network or wait until it connects before restarting the software."
                MsgBox(ConfirmM, vbExclamation, "Exiting WinCDS")
                End
            End If
        Else
            MsgBox("Please make sure Drive I is mapped to the correct computer, then try again.", vbCritical)
            End
        End If
    End Function
    Public Function CDSAppDataFolder() As String
        Dim D As String
        D = AppDataFolder()
        D = D & CompanyName & DIRSEP
        If Not DirExists(D) Then MkDir(D)
        D = D & ProgramName & DIRSEP
        If Not DirExists(D) Then MkDir(D)

        CDSAppDataFolder = D
    End Function
    Public Function AppDataFolder() As String
        AppDataFolder = Environ("AppData") & DIRSEP
        If Not DirExists(AppDataFolder) Then AppDataFolder = CDSDataFolder()
    End Function
    Public Function CDSDataFolder(Optional ByVal doLocal As Boolean = False) As String
        If doLocal Then
            CDSDataFolder = LocalCDSDataFolder()
        Else
            CDSDataFolder = GetStation()
        End If
    End Function
    Public Function WinCDSDevFolder() As String
        WinCDSDevFolder = LocalRoot & "WinCDS\"
    End Function
    Public Function WinCDSFolder() As String
        WinCDSFolder = LocalProgramFilesFolder() & "WinCDS\"
    End Function
    Public Function LocalProgramFilesFolder(Optional ByVal Nox86 As Boolean = False) As String
        LocalProgramFilesFolder = ProgramFilesFolder(doLocal:=True, Nox86:=Nox86)
    End Function
    Public Function ProgramFilesFolder(Optional ByVal doLocal As Boolean = False, Optional ByVal Nox86 As Boolean = False) As String '###x86
        ProgramFilesFolder = GetStation(doLocal:=doLocal, doStub:=False) & "Program Files"
        If Not Nox86 Then
            If DirExists(ProgramFilesFolder & " (x86)") Then ProgramFilesFolder = ProgramFilesFolder & " (x86)"
        End If
        ProgramFilesFolder = ProgramFilesFolder & DIRSEP
        '  ProgramFilesFolder = GetStation(doLocal:=doLocal, doStub:=False) & "Program Files" & IIf(x86Exists(doLocal:=doLocal), " (x86)", "") & dirsep
    End Function
    Public Function GetDatabaseInventory() As String
        GetDatabaseInventory = InventFolder() & "CDSInvent.mdb"
    End Function
    Public Function InventFolder(Optional ByVal doLocal As Boolean = False) As String
        InventFolder = GetStation(doLocal) & "Invent\"
    End Function

    Public Function CustomerTermsMessageFile(Optional ByVal StoreNum As Integer = 0) As String
        If StoreNum = 0 Then StoreNum = StoresSld
        CustomerTermsMessageFile = FXFile("CustomerTerms" & IIf(StoreNum = 1, "", StoreNum) & ".rtf", , False)

        ' patch for old way of doing it.
        On Error Resume Next
        If StoreNum <> 1 Then
            If Dir(CustomerTermsMessageFile) = "" Then
#If False Then
      If Dir(CustomerTermsMessageFile(1)) <> "" Then
        FileCopy CustomerTermsMessageFile(1), CustomerTermsMessageFile
      End If
#Else
                CustomerTermsMessageFile = CustomerTermsMessageFile(1)
#End If
            End If
        End If

    End Function
    Public Function StorePolicyMessageFile(Optional ByVal StoreNum As Integer = 0) As String
        If StoreNum = 0 Then StoreNum = StoresSld
        StorePolicyMessageFile = FXFile("StorePolicy.rtf", , False)
    End Function
    Public Function FXFile(ByVal S As String, Optional ByVal SubF As String = "", Optional ByVal RequireExists As Boolean = True) As String
        Dim R As String, T As String
        If InStr(S, ":") = 0 Then
            If FileExists(FXFolder() & S) Then FXFile = FXFolder() & S : Exit Function
        Else
            If FileExists(S) Then FXFile = S : Exit Function
        End If

        If SubF = "" Then
            R = FXFile(S, FXControlFolder)
            If R <> "" Then FXFile = R : Exit Function

            R = FXFile(S, PXFolder)                     ' Why not check here too...  Old location..
            If R <> "" Then FXFile = R : Exit Function

            '    R = FXFile(S, FXControlFolder)
            '    If R <> "" Then FXFile = R: Exit Function



            SubF = FXFolder()
        End If

        FXFile = SubF & S
        If FileExists(FXFile) Then Exit Function


        FXFile = PXfile(S, SubF)
        If FXFile <> "" Then Exit Function

        If Not RequireExists Then
            FXFile = FXFolder() & S
        Else
            FXFile = ""
        End If
    End Function
    Public Function FXFolder(Optional ByVal doLocal As Boolean = False) As String
        FXFolder = CDSDataFolder(doLocal) & "InventFX\"
        If Not DirExists(FXFolder) Then FXFolder = PXFolder() : Exit Function
    End Function
    Public Function FXControlFolder(Optional ByVal doLocal As Boolean = False) As String
        FXControlFolder = FXFolder(doLocal) & "Control\"
        If Not DirExists(FXControlFolder) Then FXControlFolder = PXFolder() : Exit Function
    End Function
    Public Function PXFolder() As String
        PXFolder = GetStation() & "InventPX\"
    End Function
    Public Function PXfile(ByVal NFile As String, Optional ByVal SrcDir As String = "") As String
        PXfile = ""
        If SrcDir = "" Then SrcDir = PXFolder()
        If FileExists(NFile) Then PXfile = NFile : Exit Function
        If FileExists(SrcDir & NFile) Then PXfile = SrcDir & NFile : Exit Function
        If FileExists(SrcDir & NFile & ".jpg") Then PXfile = SrcDir & NFile & ".jpg" : Exit Function
        If FileExists(SrcDir & NFile & ".jpeg") Then PXfile = SrcDir & NFile & ".jpeg" : Exit Function
        If FileExists(SrcDir & NFile & ".bmp") Then PXfile = SrcDir & NFile & ".bmp" : Exit Function
        If FileExists(SrcDir & NFile & ".gif") Then PXfile = SrcDir & NFile & ".gif" : Exit Function
        If FileExists(SrcDir & NFile & ".png") Then PXfile = SrcDir & NFile & ".png" : Exit Function
    End Function
    Public Function ItemPXByRN(ByVal RN As Integer, Optional ByVal WithPath As Boolean = True, Optional ByVal ForceExt As String = "") As String
        Dim TF As String, SF As String
        SF = RN
        TF = PXFolder() & RN
        ItemPXByRN = IIf(WithPath, PXFolder, "") & RN
        If ForceExt <> "" Then ItemPXByRN = IIf(WithPath, TF, SF) & ForceExt : Exit Function
        If FileExists(TF & ".gif") Then ItemPXByRN = IIf(WithPath, TF, SF) & ".gif" : Exit Function
        If FileExists(TF & ".bmp") Then ItemPXByRN = IIf(WithPath, TF, SF) & ".bmp" : Exit Function
        If FileExists(TF & ".png") Then ItemPXByRN = IIf(WithPath, TF, SF) & ".png" : Exit Function
        ItemPXByRN = IIf(WithPath, TF, SF) & ".jpg"
    End Function

    Public Function LoadPictureStd(ByVal FileName As String) As StdPicture
        If FileExists(FileName) Then
            If FreeImage_IsAvailable() Then
                LoadPictureStd = LoadPictureEx(FileName)
            Else
                'LoadPictureStd = LoadPicture(FileName)
                LoadPictureStd = Image.FromFile(FileName)
            End If
        End If
    End Function
    Public Function UpdateFolder(Optional ByVal SubFolder As String = "") As String
        UpdateFolder = InventFolder(True) & "update\" & SubFolder
        If Right(UpdateFolder, 1) <> DIRSEP Then UpdateFolder = UpdateFolder & DIRSEP
        EnsureFolderExists(UpdateFolder, True)
    End Function
    Public Function TempFile(Optional ByVal UseFolder As String = "", Optional ByVal UsePrefix As String = "wincds_tmp_", Optional ByVal Extension As String = ".tmp", Optional ByVal TestWrite As Boolean = True) As String
        Dim FN As String, Res As String
        If UseFolder <> "" And Not DirExists(UseFolder) Then UseFolder = ""
        If UseFolder = "" Then UseFolder = UpdateFolder()
        If Right(UseFolder, 1) <> DIRSEP Then UseFolder = UseFolder & DIRSEP
        'FN = Replace(UsePrefix & CDbl(Now.ToString) & "_" & App.ThreadID & "_" & Random(999999), ".", "_")
        FN = Replace(UsePrefix & CDbl(Now.ToString) & "_" & AppDomain.GetCurrentThreadId & "_" & Random(999999), ".", "_")


        Do While FileExists(UseFolder & FN & ".tmp")
            FN = FN & Chr(Random(25) + Asc("a"))
        Loop
        TempFile = UseFolder & FN & Extension

        If TestWrite Then
            On Error GoTo TestWriteFailed
            WriteFile(TempFile, "TEST", True, True)
            On Error GoTo TestReadFailed
            Res = ReadFile(TempFile)
            If Res <> "TEST" Then MsgBox("Test write to temp file " & TempFile & " failed." & vbCrLf & "Result (Len=" & Len(Res) & "):" & vbCrLf & Res, vbCritical)
            On Error GoTo TestClearFailed
            Kill(TempFile)
        End If
        Exit Function

TestWriteFailed:
        MsgBox("Failed to write temp file " & TempFile & "." & vbCrLf & Err.Description, vbCritical)
        Exit Function
TestReadFailed:
        MsgBox("Failed to read temp file " & TempFile & "." & vbCrLf & Err.Description, vbCritical)
        Exit Function
TestClearFailed:
        If Err.Number = 53 Then
            Err.Clear()
            Resume Next
        End If

        'BFH20160627
        ' Jerry wanted this commented out.  Absolutely horrible idea.
        '  If IsDevelopment Then
        MsgBox("Failed to clear temp file " & TempFile & "." & vbCrLf & Err.Description, vbCritical)
        '  End If
        Exit Function
    End Function
    Public Function DevOutputFolder() As String
        DevOutputFolder = LocalDesktopFolder()
    End Function
    Public Function TempFolder(Optional ByVal UseFolder As String = "", Optional ByVal UsePrefix As String = "tmp_", Optional ByVal Extension As String = ".tmp", Optional ByVal TestWrite As Boolean = True) As String
        Dim FN As String, Res As String
        If UseFolder <> "" And Not DirExists(UseFolder) Then UseFolder = ""
        If UseFolder = "" Then UseFolder = UpdateFolder()
        If Right(UseFolder, 1) <> DIRSEP Then UseFolder = UseFolder & DIRSEP
        UseFolder = UseFolder & UsePrefix & DateTimeStamp()

        Do While FolderExists(UseFolder)
            UseFolder = AugmentByRightLetter(UseFolder, False)
        Loop

        MkDir(UseFolder)

        UseFolder = UseFolder & DIRSEP

        If TestWrite Then
            If Not CanWriteToFolder(UseFolder) Then MsgBox("Test write to temp folder " & UseFolder & "testwrite.txt" & " failed.", vbCritical)
        End If

        TempFolder = UseFolder
    End Function

    Public Function LocalDesktopFolder() As String
        On Error Resume Next
        Dim W
        W = CreateObject("WScript.Shell")
        LocalDesktopFolder = W.SpecialFolders("Desktop") & DIRSEP
        W = Nothing
    End Function
    Public Function WinCDSEXEFile(Optional ByVal Ext As Boolean = False, Optional ByVal Standard As Boolean = False, Optional ByVal DoShortIfExists As Boolean = False) As String
        Dim T As String
        If Standard Then T = WinCDSFolder() Else T = AppFolder()
        WinCDSEXEFile = T & WinCDSEXEName(Ext, Standard)
        If DoShortIfExists And FileExists(WinCDSEXEFile) Then WinCDSEXEFile = GetShortName(WinCDSEXEFile)
    End Function
    Public Function WinCDSEXEName(Optional ByVal Ext As Boolean = False, Optional ByVal Standard As Boolean = False, Optional ByVal DoUCase As Boolean = False) As String
        Dim Location As String
        Dim Appname As String

        If Standard Then
            WinCDSEXEName = IIf(Ext, WinCDSEXE, WinCDSEXE_Base)
        Else
            'WinCDSEXEName = IIf(Ext, App.EXEName & ".EXE", App.EXEName)
            Location = Assembly.GetExecutingAssembly().Location
            Appname = Path.GetFileName(Location)
            WinCDSEXEName = IIf(Ext, Appname & ".EXE", Appname)
        End If
        If DoUCase Then WinCDSEXEName = UCase(WinCDSEXEName)
    End Function

    Public Function XChargeFolder(Optional ByVal doLocal As Boolean = False) As String
        XChargeFolder = GetStation(doLocal) & "X-Charge\"
        If DirExists(XChargeFolder) Then Exit Function

        XChargeFolder = ProgramFilesFolder(doLocal) & "X-Charge\"
        If DirExists(XChargeFolder) Then Exit Function

        '  XChargeFolder = ""     ' just keep last attempt
    End Function
    Public Function ItemPictureByRN(ByVal RN As Integer) As StdPicture
        Dim S As String, DIB As Integer
        On Error Resume Next

        'BFH20170515 - Modified in the following ways:
        '  1.  Ported single instance to this global handler.
        '        - The global handler already existed.  Other portions of the program undoubtably use it.
        '        - Individual instances are always a pain.  If it can functionized, always do so.
        '        - Global handlers fix every instance in the software at once.  Set it and forget it.  Instance at a time is the bane of all good programming standards.
        '  2.  Added appropriate existence checks using FreeImage_IsAvailable
        '        - CDS Deployment is never guaranteed to have the instance.  Always check for the instance.The instance added the FreeImage library (good), but failed to check for it.  Added FreeImage_IsAvailable.
        '        - If the FreeImage DLL is available, we can use it.  Otherwise, fail in the predictable fashion (that is, use the standard VB function, and only THEN fail completely).
        '  3.  Used Generic Loader Wrapper instead of JPEG specific.
        '        - As per the FreeImage documentation, use of _LoadEx is recommended to provide a wide variety of formats (autodetect).
        '        - The proposed solution only allowed for JPEG format, and failed for BMP (as well as any other formats).  WinCDS fully supports BMP, and it must be handled.
        '  4.  Encapsulated and fixed the FreeImage_Unload call.
        '        - Do not use parentheses on a call that does not return a function.
        '        - The unload function must be included in this call, otherwise a memory leak will result.
        '  5.  Further encapsulated the new "replacement" for the VB6 LoadPicture call into LoadPictureStd
        '        - While FreeImage provides the LoadPictureEx, this is insufficient because it does not handle native VB loading under FreeImage failure.
        '        - It stands to reason that LoadPicture is used elsewhere in the software.  If so, they also suffer this drawback.
        '        - "Loading pictures" is the issue, not "by item".  Since this is the issue, we should abstract this portion of the logic to be useful in any situation.
        '        - Never code something twice.
        '
        ' Summary Points:
        '   - Always, always, always use global handler if possible.  Never revert to instance
        '     coding.  This will ALWAYS create headaches down the road, as if anything ever changes
        '     in the future, EVERY instance will have to be updated.  Simply update the handler.
        '   - For already deployed software, NEVER NEVER NEVER assume the library will simply "be there".  For
        '     WinCDS code, we push EXEs regularly, and DLLs should be pushed 1-2 weeks in advance of all EXE updates.
        '     Failure to deploy appropriate countermeasures in the case of DLL failure, especially in key systems,
        '     is always a recipie for disaster.  "Code Defensively".  Always.
        '   - Always execute full support, when reasonable.  It should be obvious by looking at the code that
        '     The proposed solution only supports JPEG format.  Yet, ItemPxByRN function, it is clear that the software
        '     is designed to support multiple image formats.  No one format is ever guaranteed, even by Ashley,
        '     and so the foresight to find the LoadEX function is essential.  It makes no sense to fix something once,
        '     only to have to fix it again three months down the road, because it wasn't done properly the first time.
        '   - Always be on the lookout for further abstraction.  By removing the "load picture" part, which now
        '     must also consider the FreeImage library... by removing that FROM the "by item" portion, you set yourself
        '     up for success on the next project, because your work is already done.  You no longer need to create
        '     Library functions, but can quickly progress into higher level development, because image handling
        '     is already accomplished.
        '
        '   The basic rule is this:  Never code something twice.
        '     If you can abstract it, do it.  If you can improve a handler, do it.  If you can prevent having to fix
        '     100 instances in two years, even if they were all coded perfectly at the time they were made, you will
        '     always win in the long run.  Think ahead.

        ItemPictureByRN = LoadPictureStd(ItemPXByRN(RN))
    End Function

    Public Sub Domain_exit()
        dbClose
    End Sub

    Public Function GetLastDeliveryDate() As Date
        On Error Resume Next
        'If CDbl(gblLastDeliveryDateEpoch) = 0 Or DateAfter(Now, DateAdd("d", 1, gblLastDeliveryDateEpoch)) Then
        If IsNothing(gblLastDeliveryDateEpoch) Or DateAfter(Now, DateAdd("d", 1, gblLastDeliveryDateEpoch)) Then
            SetLastDeliveryDate()
            'ElseIf CDbl(gblLastDeliveryDate) = 0 Then
        ElseIf IsNothing(gblLastDeliveryDate) Then
            SetLastDeliveryDate()
        End If
        GetLastDeliveryDate = gblLastDeliveryDate
    End Function

    Public Sub SetLastDeliveryDate(Optional ByVal Whenn As Date = NullDate)
        On Error Resume Next
        If Whenn = NullDate Then Whenn = Now
        gblLastDeliveryDateEpoch = Today
        gblLastDeliveryDate = Whenn
    End Sub

    Public Function UIOutputFolder() As String
        UIOutputFolder = IIf(IsDevelopment, LocalDesktopFolder, DevOutputFolder)
    End Function

    Public Function FXWallpaperFolder() As String
        FXWallpaperFolder = FXFolder() & "Wallpapers\"
        If Not DirExists(FXWallpaperFolder) Then FXWallpaperFolder = PXFolder() :
        Exit Function
    End Function

    Public Function TagLayoutFolder() As String
        TagLayoutFolder = FXFolder() & "TagLayouts\"
        If Not DirExists(TagLayoutFolder) Then TagLayoutFolder = PXFolder() :
        Exit Function
    End Function

    Public Function PRFolder(Optional ByVal Data As Boolean = False, Optional ByVal LocalOnly As Boolean = False) As String
        PRFolder = AccountingFolder("Payroll", Data, LocalOnly)
    End Function

    Public Function GLFolder(Optional ByVal Data As Boolean = False, Optional ByVal LocalOnly As Boolean = False) As String
        GLFolder = AccountingFolder("GenLedger", Data, LocalOnly)
    End Function

    Public Function APFolder(Optional ByVal Data As Boolean = False, Optional ByVal LocalOnly As Boolean = False) As String
        APFolder = AccountingFolder("Payable", Data, LocalOnly)
    End Function

    Public Function BKFolder(Optional ByVal Data As Boolean = False, Optional ByVal LocalOnly As Boolean = False) As String
        BKFolder = AccountingFolder("Banking", Data, LocalOnly)
    End Function

    Private Function AccountingFolder(ByVal Modulee As String, Optional ByVal Data As Boolean = False, Optional ByVal LocalOnly As Boolean = False) As String
        Dim S As String
        If Not IsServer() And Not LocalOnly Then
            S = GetStation() & Modulee & DIRSEP
            If DirExists(S) Then AccountingFolder = S : GoTo Finish
            S = GetStation(, False) & Modulee & DIRSEP
            If DirExists(S) Then AccountingFolder = S : GoTo Finish
            S = ProgramFilesFolder() & Modulee & DIRSEP
            If DirExists(S) Then AccountingFolder = S : GoTo Finish
        Else
            S = GetStation(LocalOnly) & Modulee & DIRSEP
            If DirExists(S) Then AccountingFolder = S : GoTo Finish
            AccountingFolder = IIf(LocalOnly, LocalProgramFilesFolder, ProgramFilesFolder) & Modulee & DIRSEP
        End If
Finish:
        If Data Then AccountingFolder = AccountingFolder & "Data\"
    End Function

    Public Function GetDatabaseAP(Optional ByVal Location as integer = 1) As String
        GetDatabaseAP = APFolder(True) & "L" & Location & "-AP.MDB"
    End Function
End Module
