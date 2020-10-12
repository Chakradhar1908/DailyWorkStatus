Imports System.IO
Imports System.Reflection
Imports stdole
Imports System.Threading
Imports Microsoft.Build.Framework.XamlTypes
Module MainModule
    Private mIsServer As TriState           ' Cache this high-use, non-trivial value
    Public Allow_ADODB_Errors As Boolean      ' Allow the database to continue after errors - for debugging and special cases.
    Public Const WinCDS_ProjectFilename As String = "WinCDS.vbp"
    Public Const WinCDSEXE_Base As String = "WinCDS"
    Public Const WinCDSEXE As String = WinCDSEXE_Base & ".exe"
    Public gblLastDeliveryDate As Date        ' This will make the last delivery date persist without keeping whole forms loaded.
    Public gblLastDeliveryDateEpoch As Date   ' This will make the last delivery date reset daily
    Public PrvKill As Boolean
    Public ProgramStart As Date               ' When the program started
    Private mQuickQuit As Boolean             ' used to bypass the quit confirmation msgbox
    Public ServiceMode As Boolean
    Public ProgramStarted As Boolean          ' Record that the program has fully started
    Private Const ELCkey1 As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    Private Const ELCkey2 As String = "EnableLinkedConnections"
    Public NoReplace As Boolean               ' Prevent Update re-run

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
            If MessageBox.Show(ProgramName & " can't connect to the server." & vbCrLf & "Is this computer the server?", "WinCDS", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
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

                If MessageBox.Show(ConfirmM, "CONFIRM SERVER", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2) = DialogResult.OK Then
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
                MessageBox.Show(ConfirmM, "Exiting WinCDS")
                End
            End If
        Else
            MessageBox.Show("Please make sure Drive I is mapped to the correct computer, then try again.", "WinCDS")
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
        'FN = Replace(UsePrefix & CDbl(Now.ToString) & "_" & AppDomain.GetCurrentThreadId & "_" & Random(999999), ".", "_")
        'FN = Replace(UsePrefix & CDbl(Now.ToOADate) & "_" & AppDomain.GetCurrentThreadId & "_" & Random(999999), ".", "_")
        FN = Replace(UsePrefix & CDbl(Now.ToOADate) & "_" & Thread.CurrentThread.ManagedThreadId & "_" & Random(999999), ".", "_")



        Do While FileExists(UseFolder & FN & ".tmp")
            FN = FN & Chr(Random(25) + Asc("a"))
        Loop
        TempFile = UseFolder & FN & Extension

        If TestWrite Then
            On Error GoTo TestWriteFailed
            WriteFile(TempFile, "TEST", True, True)
            On Error GoTo TestReadFailed
            Res = ReadFile(TempFile)
            If Res <> "TEST" Then MessageBox.Show("Test write to temp file " & TempFile & " failed." & vbCrLf & "Result (Len=" & Len(Res) & "):" & vbCrLf & Res, "WinCDS")
            On Error GoTo TestClearFailed
            Kill(TempFile)
        End If
        Exit Function

TestWriteFailed:
        MessageBox.Show("Failed to write temp file " & TempFile & "." & vbCrLf & Err.Description, "WinCDS")
        Exit Function
TestReadFailed:
        MessageBox.Show("Failed to read temp file " & TempFile & "." & vbCrLf & Err.Description, "WinCDS")
        Exit Function
TestClearFailed:
        If Err.Number = 53 Then
            Err.Clear()
            Resume Next
        End If

        'BFH20160627
        ' Jerry wanted this commented out.  Absolutely horrible idea.
        '  If IsDevelopment Then
        MessageBox.Show("Failed to clear temp file " & TempFile & "." & vbCrLf & Err.Description, "WinCDS")
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
            If Not CanWriteToFolder(UseFolder) Then MessageBox.Show("Test write to temp folder " & UseFolder & "testwrite.txt" & " failed.", "WinCDS")
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
        dbClose()
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

    Public Function GetDatabaseAP(Optional ByVal Location As Integer = 1) As String
        GetDatabaseAP = APFolder(True) & "L" & Location & "-AP.MDB"
    End Function

    Public Property QuickQuit() As Boolean
        Get
            QuickQuit = mQuickQuit
            If ReadStoreSetting(0, IniSections_StoreSettings.iniSection_StoreSettings, "QuickQuit") <> "" Then QuickQuit = True
        End Get
        Set(value As Boolean)
            mQuickQuit = value
        End Set
    End Property

    Public Function ShowHelp() As Boolean
        ShellOut.RunFile(WinCDSHelpFile)
        ShowHelp = True
    End Function

    Public Function ShutdownSemaforeFile(Optional ByVal CreateIt As Boolean = False, Optional ByVal ItExists As Boolean = False, Optional ByVal DeleteIt As Boolean = False)
        ShutdownSemaforeFile = StoreFolder(1) & "shutdown.txt"
        On Error Resume Next
        If CreateIt Then WriteFile(ShutdownSemaforeFile, "" & Today, True)
        If DeleteIt Then Kill(ShutdownSemaforeFile)
        If ItExists Then ShutdownSemaforeFile = FileExists(ShutdownSemaforeFile)
    End Function

    Public Function NightlyCleanup() As Boolean
        '  VerifyMailRecUnique 0, , True ' this cleans up any dangling mail record reservations..
    End Function

    Public Function WinCDSHelpFile(Optional ByVal Ext As Boolean = True, Optional ByVal wPath As Boolean = True, Optional ByVal Standardized As Boolean = False) As String
        Dim Std As String

        WinCDSHelpFile = ProgramName
        If Ext Then WinCDSHelpFile = WinCDSHelpFile & ".CHM"
        If wPath Then
            Std = FXFolder(True) & ProgramName & ".CHM"
            If Standardized Or FileExists(Std) Then
                WinCDSHelpFile = Std
            Else
                WinCDSHelpFile = AppFolder() & WinCDSHelpFile
            End If
        End If
    End Function

    Public Function UserFolder() As String
        UserFolder = ParentDirectory(LocalDesktopFolder)
    End Function

    Public Sub RestartProgram() ' 3/3/6
        On Error Resume Next
        Shell(WinCDSEXEFile, vbNormalFocus)
        MainMenu.ShutDown(True)
    End Sub

    Public Function AllUsersDesktopFolder() As String
        On Error Resume Next
        Dim W
        W = CreateObject("WScript.Shell")
        AllUsersDesktopFolder = W.SpecialFolders("AllUsersDesktop") & DIRSEP
        W = Nothing
    End Function

    Public Function WinCDSAutoVNCFolder() As String
        WinCDSAutoVNCFolder = CDSDataFolder(True) & "AutoVNC\"
        If DirExists(WinCDSAutoVNCFolder) Then Exit Function
        WinCDSAutoVNCFolder = AppFolder() & "AutoVNC\"
    End Function

    Public Function PhysicalInvOldFolder() As String
        PhysicalInvOldFolder = PhysicalInvFolder() & "Old Data\"
    End Function

    Public Function PhysicalInvFolder() As String
        PhysicalInvFolder = InventFolder() & "Physical\"
    End Function

    Public Function System32Folder(Optional ByVal IncludeTrailingBackslash As Boolean = False) As String
        System32Folder = GetWindowsSystemDir()

        If IncludeTrailingBackslash Then
            If Right(System32Folder, 1) <> DIRSEP Then System32Folder = System32Folder & DIRSEP
        End If
    End Function

    Public Function ConnectCMDFile() As String
        ConnectCMDFile = WinCDSAutoVNCFolder() & "Connect.cmd"
    End Function

    ' ======>>  This is the first procedure run in the program
    Public Sub Main()
        Dim StartupProcess As String
        Dim DoEnd As Boolean
        On Error GoTo StartupErrorHandler
        ProgramStart = Now
        LogStartup("INIT - ******************************************************")
        LogStartup("INIT - " & ProgramStart)
        LogStartup("INIT - " & SoftwareVersionForLog())
        EnableLinkedConnections()

        ActiveLog("Init::====== Startup: " & ProgramStart)
        If Command() = "" Then
            If Not IsIDE() Then
                If CheckReplaceWinCDS() Then Exit Sub
            End If
            If NotifyDemoExpired(Command) Then Exit Sub
        End If

        'PatchFxFolder
        'PatchFxFolder2


        If CheckStartupParameters(Command, DoEnd) Then
            LogStartup("INIT - Processed Command Line: " & Command())
            LogStartup("=================================================")
            If DoEnd Then End
            Exit Sub
        End If
        ActiveLog("Checked Startup Parameters -- normal run")

        LogStartup("Initializing Controls...")
        EnableLinkedConnections()
        DoInitCommonControls()

        LogStartup("IDE Check")
        If IsIDE() Then IDEStartup()

        LogStartup("Secure Startup")
        '  UpgradeStoreInformationFile
        SecureStartup()
        If Not IsDemo() Then frmUpgrade.NotifyUpgrade(False)

        '  LogStartup "Vista Check"
        '  CheckVistaAdminRights

        LogStartup("------------ Log Folders ------------")
        LogFolders()

        StartupProcess = "Initializing..." : LogStartup(StartupProcess)
        'frmSplash.DoStatus(StartupProcess)
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        SplashProgress(0, 100)

        LogStartup("OnDemand Startup")
        If Not IsDemo() Then OnDemandStartup()
        SplashProgress(1)

        ' use this to make the permission monitor come on automatically for development mode
        ' Usually, to keep it off, the first clause should be False.
        If IsDevelopment() Then
            If GetCDSSetting("Permission Monitor") <> "" Then
                'Load frmPermissionMonitor
                frmPermissionMonitor.LoadSettings()
                frmPermissionMonitor.Show()
            End If
        End If

        If Not IsWinCDSInCorrectFolder() Then
            LogStartup("WinCDS is in the incorrect folder: " & vbCrLf & ">>" & CurrentEXEDirectory() & vbCrLf & ">>" & WinCDSFolder())
            MessageBox.Show("WinCDS is not running in the correct folder:" & CurrentEXEDirectory() & vbCrLf2 & "Please be sure WinCDS is running in the following folder:" & vbCrLf & WinCDSFolder(), "WinCDS Path Warning")
        End If


        StartupProcess = "Attempting to restore I: Drive"
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        RestoreIDrive()                                     ' force the I: drive to connect if possible
        SplashProgress(3)

        ' Initialize store variables.
        StartupProcess = "Loading Settings..." : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        modPatches.MoveUserRegistryToSystem()                          ' Registry patch..
        SplashProgress(5)

        StartupProcess = "Verifying Settings..." : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        If License <> StoreSettings(1).loadedLicense Then
            If WinCDSLicenseValid(StoreSettings(1).loadedLicense) Then License = StoreSettings(1).loadedLicense
        End If
        If InstallmentLicense <> StoreSettings(1).loadedInstallmentLicense Then
            If InstallmentLicenseValid(StoreSettings(1).loadedInstallmentLicense) Then InstallmentLicense = StoreSettings(1).loadedInstallmentLicense
        End If
        SplashProgress(10)

        StartupProcess = "Preparing Main Store..." : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        LogStartup("Main -> LoadPermOptions")
        LoadPermOptions()                 ' Initialize the permission possibility areas.
        SplashProgress(15)

        StartupProcess = "Loading Program.." : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        SplashProgress(35)    'skinning takes longer
        'Load MainMenu
        If Not IsFormLoaded("MainMenu") Then
            'Load Practice
            Practice.StartupFailure()
            Exit Sub
        End If

        StartupProcess = "Detecting initial window state..." : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)

        If StoreSettings.bStartMaximized Then MainMenu.SetWindowState(VBRUN.FormWindowStateConstants.vbMaximized)
        SplashProgress(50)

        StartupProcess = "Loading Help File..." : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        'App.HelpFile = WinCDSHelpFile()
        'Dim b As New BaseProperty
        'b.HelpFile = WinCDSHelpFile()

        SplashProgress(60)


        StartupProcess = "Verifying Scheduled Task" : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        modScheduledTasks.CheckScheduledTasks()

        '  modScheduledTasks.CheckWinCDSServiceScheduledTask REMOVE:=True
        SplashProgress(61)



        StartupProcess = "Detecting Printers..." : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        OutputObject = Printer
        SplashProgress(65)

        StartupProcess = "Loading Program...." : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        MainMenu.Show()
        '  If MainMenu.WindowState = vbNormal Then MainMenu.Move Screen.Width / 2 - MainMenu.picBackground.Width / 2, Screen.Height / 2 - MainMenu.picBackground.Height / 2
        SplashProgress(70)

        StartupProcess = "Applying most recent patches..." : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        'MainMenu.MousePointer = vbHourglass
        MainMenu.Cursor = Cursors.WaitCursor
        AutoPatch() ' Check for required patches - This takes a few seconds every time the program loads... We should make it quicker if we can.
        'MainMenu.MousePointer = vbNormal
        MainMenu.Cursor = Cursors.Default
        SplashProgress(95)

        StartupProcess = "Verifying Program Features..." : LogStartup(StartupProcess)
        'frmSplash.DoStatus StartupProcess
        modMainMenu.frmSplash.DoStatus(StartupProcess)
        VerifyProgramFeatures()
        InitializeForCDSCustomers()

        SplashProgress(100)

        SplashProgress()
        StartupProcess = "" : LogStartup(StartupProcess)
        'frmSplash.DoClose
        modMainMenu.frmSplash.DoClose()

        ShowLicenseAgreement(False)
        ProgramStarted = True

        Exit Sub

        '  #If TESTING Then
        '    frmtblAPVendors.Show 1
        '    Exit Sub
        '  #End If
StartupErrorHandler:
        Dim ReShow As Boolean, DC As String, En As Integer
        DC = Err.Description
        En = Err.Number
        ReShow = False
        HideSplash()
        MessageBox.Show("Error in startup procedure (" & En & "): " & DC & vbCrLf & "Current Operation: " & StartupProcess, ProgramName & " Startup Error")
        If ReShow Then frmSplash.Show()
        Err.Clear()
        Resume Next
    End Sub

    Public Sub LogStartup(ByVal Msg As String)
        Dim FN As String

        If ProgramStarted Then Exit Sub

        '  MsgBox Msg & vbCrLf & "3" & vbCrLf & "FN: " & FN
        '  MsgBox "" & GetStation
        On Error Resume Next
        FN = "StartUp.txt"
        If True Then
            LogFile(FN, Msg, False)
        Else
            KillLog(FN) ' while disabled, delete the log
        End If
    End Sub

    Public Function EnableLinkedConnections() As Boolean
        ' Indeed, with UAC enabled you cannot access a network drive mapped in the normal mode from an app run elevated.
        '
        ' From the articles:
        '   http://woshub.com/how-to-access-mapped-network-drives-from-the-elevated-apps/
        '
        ' https://technet.microsoft.com/en-us/library/ee844140(v=ws.10).aspx

        ' Require admin..
        If IsWinXP() Then Exit Function
        If Not IsElevated() Then Exit Function

        EnableLinkedConnections = True
        If GetRegistrySetting(HKEYS.regHKLM, ELCkey1, ELCkey2) = "1" Then Exit Function

        'If GetRegistrySetting, ELCkey1, ELCkey2) = "1" Then Exit Function

        SaveRegistrySetting(HKEYS.regHKLM, ELCkey1, ELCkey2, 1, REG_TYPE.vtDWord)
        EnableLinkedConnections = GetRegistrySetting(HKEYS.regHKLM, ELCkey1, ELCkey2) = "1"

        If Not EnableLinkedConnections Then
            ErrMsg("Could not set EnableLinkedConnections--this could result in an endless computer reboot loop.")
            Exit Function
        End If

        If MessageBox.Show("This update requires a one-time computer restart." & vbCrLf & "Press OK to restart your computer.", "Restart Needed", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) = DialogResult.Cancel Then
            MessageBox.Show("If you are prompted whether this is the server or not, rebooting the computer should fix this.")
            Exit Function
        End If

        RestartComputer()
        End

        ' To test...
        '  DeleteRegistrySetting regHKLM, ELCkey1, ELCkey2
        ' An alternate method...
        '  ShellOut.ShellOut "reg add ""HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System"" /v ""EnableLinkedConnections"" /t REG_DWORD /d 0x00000001 /f"
    End Function

    Public Function CheckStartupParameters(ByVal Args As String, Optional ByRef DoEnd As Boolean = False) As Boolean
        ' returning 'True' from this function will cause the main program not to load.
        DoEnd = True ' Default is to end program on a flag.  This will only matter if this function returns a True result.

        If IsInStr(LCase(Args), "/?") Then
            Dim S As String
            S = ""
            S = S & "Options:" & vbCrLf
            S = S & "  /INIT          - Reserved for Installer"
            S = S & "  /COMMAND       - Open Command Prompt" & vbCrLf
            S = S & "  /NOREPLACE     - Prevent Software Install" & vbCrLf
            S = S & "  /STAYOPEN      - Leave Program Running Without Windows" & vbCrLf
            S = S & "  /NOKILL        - Disable Kill ug" & vbCrLf
            S = S & "  /PRACTICE      - Open Developer Panel" & vbCrLf
            S = S & "  /STAYOPEN      - Leave Program Running Without Windows" & vbCrLf
            S = S & "  /OFFLINEUPDATE - Upgrade WinCDS from standalone CD" & vbCrLf
            S = S & "  /TASKS         - Reset Scheduled Tasks" & vbCrLf
            S = S & "  /PROCESS       - Nightlky Command Line" & vbCrLf
            S = S & "  /SERVICE       - Nightly Service Request (Update & Backup)" & vbCrLf
            S = S & "  /UPDATE        - Check for nightly updates" & vbCrLf
            S = S & "  /BACKUP        - Do Cloud Backup (if necessary)" & vbCrLf
            S = S & "  /REPAIR        - Stand-alone compact/repair"
            S = S & "  /SERVER        - Start as server"
            S = S & "  /WORKSTATION   - Start as workstation"
            MessageBox.Show(S)
            End
        End If

        If IsInStr(LCase(Args), "/init") Then
            ' Be careful here... Don't use anything...  This is post-installer setup
            frmMapDrive.InitialSetup()
            End
        End If

        If IsInStr(LCase(Args), "/server") Then
            SetServer(True)
            CheckStartupParameters = False
            DoEnd = False
            Exit Function
        End If

        If IsInStr(LCase(Args), "/workstation") Then
            SetServer(False)
            CheckStartupParameters = False
            DoEnd = False
            Exit Function
        End If

        If IsInStr(LCase(Args), "/command") Then
            PracticeCommandPrompt.Show()
            CheckStartupParameters = True
            DoEnd = False
            Exit Function
        End If

        If IsInStr(LCase(Args), "/repair") Then
            DoCompactAndRepairDAO(True)
            CompactRepairJETAllDDBs(True)
            CheckStartupParameters = True
            DoEnd = False
            Exit Function
        End If

        If IsInStr(LCase(Args), "/noreplace") Then
            NoReplace = True
        End If

        If IsInStr(LCase(Args), "/stayopen") Then
            CheckStartupParameters = True
            'Load MainMenu4_Images                     ' Load something to keep it running..
            DoEnd = False
            Exit Function
        End If

        If IsInStr(LCase(Args), "/practice") Then
            LogStartup("Performing Practice Task")
            Practice.StartupFailure(True)
            CheckStartupParameters = True
            DoEnd = False
            LogStartup("Completed Practice Command-Line Task")
        End If
        If IsInStr(LCase(Args), "/nokill") Then PrvKill = True

        If IsInStr(LCase(Args), "/offlineupdate") Then
            LogStartup("Performing Offline Update Task")
            modOnDemand.DoOfflineUpdate()
            CheckStartupParameters = True
            DoEnd = True
            LogStartup("Completed Offline Update Command-Line Task")
        End If

        If IsDemo() Then Exit Function ' No notification, of course, as these are command line flags.

        If IsInStr(LCase(Args), "/tasks") Then
            LogStartup("Performing Scheduled Tasks Update:  Elevated? " & YesNo(IsElevated))
            modScheduledTasks.CheckScheduledTasks(True)
            CheckStartupParameters = True
            DoEnd = True
            LogStartup("Completed Scheduled Tasks Update")
        End If


        '  If IsInStr(LCase(Args), "/email") Then
        '    LogStartup "Performing Email Task"
        '    modAWS.DoCommandLineBackup
        '    CheckStartupParameters = True
        '    LogStartup "Completed Email Command-Line Task"
        '  End If
        '


        ' BFH20160720
        ' We actually don't care if they software is killed...
        ' Just let it update if it can.  Updates are controlled server-side, and if they're
        ' able to get an update from it, they should get one even if they're expired...
        '  If Not KillBug(True) Then
        If IsInStr(LCase(Args), "/service") Then
            LogStartup("/Service param passed")
            If IsExpired() Then
                LogStartup("/Service mode disabled for KILLED software")
                End
            End If
            If IsServer() Then Args = Args & "/process"
            Args = Args & " /update"
            If IsServer() Then Args = Args & " /backup"
            ServiceMode = True
        End If

        If IsInStr(LCase(Args), "/process") Then
            LogStartup("Performing Nightly Service Command-Line Task")
            modNightlyProcesses.NightlyProcesses()
            CheckStartupParameters = True
            LogStartup("Completed Nightly Service Command-Line Task")
        End If

        If IsInStr(LCase(Args), "/backup") Then
            LogStartup("Performing Backup Command-Line Task")
            modAWS.DoCommandLineBackup()
            CheckStartupParameters = True
            LogStartup("Completed Backup Command-Line Task")
        End If

        If IsInStr(LCase(Args), "/update") Then
            LogStartup("Performing Update Command-Line Task")
            '      modScheduledTasks.CheckScheduledTasks DateBefore(Date, #5/10/2016#)
            If Not IsDemo() Then
                frmUpgrade.DoCommandLineUpdate()
            End If
            CheckStartupParameters = True
            LogStartup("Copmleted Update Command-Line Task")
        End If
        '  End If

        If CheckStartupParameters Then
            LogStartup("Processed Startup Parameters - Returning Exit Code")
        Else
            LogStartup("Processed Startup Parameters - No Codes")
            DoEnd = False ' clear this, just in case.. doesn't really matter though.
        End If
    End Function

    Private Sub IDEStartup()
        RecordBuildDate()           ' Re-Write modBuildDate.bas every program start inside IDE - keeps it fresh
        CheckCompanyInformation()
        CheckCopyrightDate()
        CheckCertificateExpiration()

        IDEKillBugNotify()          ' notifies developers (ONLY) when kill bug will be expiring shortly
    End Sub

    Private Sub SecureStartup()
        If Not IsCDSComputer() Then SetDevMode(0)  ' not necessary...  customers seemed to be in dev mode.  this prevents it.

        If Not IsServer() And IsServerLocked() Then
            ServerLock(doSet:=vbTrue)
        End If

        ClearAllTempFiles()
        CheckSecureIPAddress()
        EnableLinkedConnections()
        EnableRichCHMContent()
        '  CreateCDSDataShare
        ForcePrinterSelection()

        'BFH20150514
        ' Causes error on Demo startup, which we would like to hide.
        ' This pushes the error further down the track if the update
        ' folder does not have full permissions, of course.
        '  TempFile

    End Sub

    Public Sub SplashProgress(Optional ByVal Value As Integer = -1, Optional ByVal Max As Integer = -1)
        If Not IsFormLoaded("frmSplash") Then Exit Sub
        'frmSplash.DoProgress Value, Max
        modMainMenu.frmSplash.DoProgress(Value, Max)
        Application.DoEvents()
    End Sub

    Public Function IsWinCDSInCorrectFolder() As Boolean
        IsWinCDSInCorrectFolder = DirEqual(CurrentEXEDirectory, WinCDSFolder) Or IsIDE()
    End Function

    Public Function RecordBuildDate() As Boolean
        Dim F As String
        Dim S As String, T As String
        Dim N As String, M As String
        Dim X As String, L As Object
        ' Only in IDE, of course
        If Not IsIDE() Then Exit Function

        ' Do not post if nothing changed.
        If DateEqual(BuildDate, Today) Then Exit Function

        N = vbCrLf : M = ""

        ' this is the entire contents of this file.. it will get compiled in
        S = ""
        S = S & M & ""
        'S = S & M & "Attribute VB_Name = ""modBuildDate"""
        'S = S & N & "Option Explicit"
        S = S & "Module modBuildDate"
        S = S & N & "' ***** WARNING: FILE IS GENERATED ON EACH COMPILE"
        S = S & N & "' DO NOT MODIFY THIS FILE.  FIND ANOTHER FILE TO MODIFY."
        S = S & N & "' YOUR CHANGES WILL BE DELETED AUTOMATICALLY"
        S = S & N
        S = S & N & "Public Function BuildDate() As String"
        S = S & N & "  BuildDate = """ & Today & """"
        S = S & N & "End Function"
        S = S & N
        S = S & N & "Public Function BuildTime() As String"
        S = S & N & "  BuildTime = """ & TimeValue(Now) & """"
        S = S & N & "End Function"
        S = S & N
        S = S & N & "Public Function BuildComputer() As String"
        S = S & N & "  BuildComputer = """ & GetLocalComputerName() & "\" & GetSystemUserName() & """"
        S = S & N & "End Function"
        S = S & N
        S = S & N & "Public Function BuildHistory() As String"
        S = S & N & "  Dim S As String"
        S = S & N
        X = ReadEntireFile(DistributionCSV)

        If X <> "" Then
            For Each L In Split(X, vbCrLf)
                S = S & N & "  S = S & vbCrLf & """ & Replace(EncodeBase64String(L), vbCrLf, "") & """"
            Next
        End If
        S = S & N
        S = S & N & "  BuildHistory = S"
        S = S & N & "End Function"
        S = S & N & "End Module"
        S = S & N


        'F = AppFolder() & "modBuildDate.bas"
        'F = Appfolder() & "modBuildDate.vb"
        Dim Applicationfolder As String
        Applicationfolder = AppFolder()
        F = Left(Applicationfolder, InStr(11, Applicationfolder, "\")) & "Modules\modBuildDate.vb"
        T = ReadFile(F)
        If NLTrim(T) <> NLTrim(S) Then
            WriteFile(F, S, True)
        End If
        RecordBuildDate = True
    End Function

    Public Function IsServerLocked() As Boolean
        IsServerLocked = UCase(Left(ReadStoreSetting(1, IniSections_StoreSettings.iniSection_StoreSettings, "ServerLock"), 1)) = "T"
    End Function

    Public Function ClearAllTempFiles() As Boolean
        Const TEMPFILE_PREFIX = "wincds_tmp_"
        Const TEMPFILE_EXTENSION = ".tmp"
        Const TEMPFILE_EXTENSION2 = ".tmp2"

        Const TEMPFOLDER_PREFIX = "tmp_"
        Dim F As String, T As String
        F = UpdateFolder()
        On Error Resume Next
        Kill(F & TEMPFILE_PREFIX & "*.*")
        Kill(F & "*" & TEMPFILE_EXTENSION)
        Kill(F & "*" & TEMPFILE_EXTENSION2)

        '  Do While True
        '    T = Dir(F & TEMPFOLDER_PREFIX & "*", vbDirectory)
        '    If T = "" Then Exit Function
        '    CleanPath F & T
        '  Loop
    End Function

    Public Function ForcePrinterSelection()
        Dim S As String
        S = ReadStoreSetting(StoresSld, IniSections_StoreSettings.iniSection_StoreSettings, "StartupPrinter")
        If S <> "" Then SetPrinter(S)
        'MsgBox "s=" & S & vbCrLf & Printer.DeviceName
    End Function

    Public Function ThisEXEFile() As String
        ThisEXEFile = WinCDSEXEFile(True, False, True)
    End Function

    Public Function ThisEXEName() As String
        ThisEXEName = WinCDSEXEName(True, False, False)
    End Function

    Public Function TestWriteFolder(ByVal UseFolder As String, Optional ByRef FailMsg As String = "") As Boolean
        Dim T As String
        Dim Res As String
        FailMsg = ""
        T = TempFile(UseFolder, , , False)
        On Error GoTo TestWriteFailed
        WriteFile(T, "TEST", True, True)
        On Error GoTo TestReadFailed
        Res = ReadFile(T)
        If Res <> "TEST" Then MessageBox.Show("Test write to temp file " & TempFile() & " failed." & vbCrLf & "Result (Len=" & Len(Res) & "):" & vbCrLf & Res)
        On Error GoTo TestClearFailed
        Kill(T)

        TestWriteFolder = True

TestWriteFailed:
        FailMsg = "Failed to write temp file " & TempFile() & "." & vbCrLf & Err.Description
        Exit Function
TestReadFailed:
        FailMsg = "Failed to read temp file " & TempFile() & "." & vbCrLf & Err.Description
        Exit Function
TestClearFailed:
        FailMsg = "Failed to clear temp file " & TempFile() & "." & vbCrLf & Err.Description
        Exit Function
    End Function

    Public ReadOnly Property UseScheduledTask() As Boolean
        Get
            'BFH20160423 - Default value set to 1
            UseScheduledTask = ReadStoreSetting(0, IniSections_StoreSettings.iniSection_StoreSettings, "UseScheduledTask", "1") <> ""
        End Get
    End Property

    Public Function LocalStoreFolder(Optional ByVal StoreNum As Integer = 0) As String
        LocalStoreFolder = StoreFolder(StoreNum:=StoreNum, doLocal:=True)
    End Function

    Public Function ReportsFolder(Optional ByVal SubFolder As String = "") As String
        ReportsFolder = InventFolder() & "Reports\"
        EnsureFolderExists(ReportsFolder, True)
        If SubFolder <> "" Then
            ReportsFolder = ReportsFolder & SubFolder
            If Right(ReportsFolder, 1) <> DIRSEP Then ReportsFolder = ReportsFolder & DIRSEP
            EnsureFolderExists(ReportsFolder, True)
        End If
    End Function

    Public Function DevelopmentFolder() As String
        DevelopmentFolder = WinCDSDevFolder() & "WinCDS\"
    End Function

    Public Function WaitEXEFile(Optional ByVal Ext As Boolean = True, Optional ByVal wPath As Boolean = True) As String
        WaitEXEFile = WinCDSDevFolder() & "Xtras\Wait\Wait.exe"
        If FileExists(WaitEXEFile) Then Exit Function
        WaitEXEFile = IIf(wPath, AppFolder, "") & "Wait" & IIf(Ext, ".exe", "")
    End Function

    Public Function DeliveryTicketMessageFile(Optional ByVal StoreNum As Long = 0) As String
        If StoreNum = 0 Then StoreNum = StoresSld
        DeliveryTicketMessageFile = FXFile("DeliveryTicket.rtf", , False)
        If Not FileExists(DeliveryTicketMessageFile) Then
            WriteFile DeliveryTicketMessageFile, "{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\froman\fprq2\fcharset0 Times New Roman;}{\f1\fnil\fcharset0 Courier New;}}" & vbCrLf
    WriteFile DeliveryTicketMessageFile, "{\*\generator Riched20 10.0.10240}\viewkind4\uc1 " & vbCrLf
    WriteFile DeliveryTicketMessageFile, "\pard\b\f0\fs16 All Items Received in good condition!\b0\f1\fs20\par" & vbCrLf
    WriteFile DeliveryTicketMessageFile, "}" & vbCrLf
    WriteFile DeliveryTicketMessageFile, Chr(0)


'    WriteFile DeliveryTicketMessageFile, "{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fnil\fcharset0 Calibri;}}" & vbCrLf
            '    WriteFile DeliveryTicketMessageFile, "{\*\generator Riched20 10.0.10240}\viewkind4\uc1" & vbCrLf
            '    WriteFile DeliveryTicketMessageFile, "\pard\sa200\sl276\slmult1\b\f0\fs22\lang9 Received in good condition!\b0\par" & vbCrLf
            '    WriteFile DeliveryTicketMessageFile, "}" & vbCrLf
        End If
    End Function

End Module
