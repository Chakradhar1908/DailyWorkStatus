Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Module modDiagnostics
    Public Enum eComponentIDs
        cidDAOEngine36 = 0

        cidMSExcel
        cidMSWord
        cidMSOutlook
        '  cidMSMapPoint    ' raises msgboxes on load
        cidMSComm

        cidScriptingDictionary
        cidScriptingFSO

        cidWScriptShell
        cidWScriptNetwork

        cidCDOMessage
        cidCDOConfiguration
        cidMAPISession
        cidMSMAPISession

        cidvbSendMail
        cidXChargeTransaction
        cidSCPinPadDevice

        cidQBFC5SM
        '  cidQBFC6SM
        '  cidQBFC7SM
        cidQBFC8SM
        '  cidQBFC9SM
        cidQBFC10SM
        '  cidQBFC11SM
        cidQBFC12SM
        cidQBFC13SM        'BFH20150205 this is the highest so far
        '  cidQBFC14SM
        '  cidQBFC15SM
        '  cidQBFC16SM
        '  cidQBFC17SM
        '  cidQBFC18SM
        '  cidQBFC19SM
        '  cidQBFC20SM


        cid_MAXPLUSONE

    End Enum

    Public Sub LogFolders()
        ActiveLog("Init::InventFolder: " & InventFolder(), 1)
        ActiveLog("Init::PXFolder: " & PXFolder(), 1)
        ActiveLog("Init::FXFolder: " & FXFolder(), 1)
        ActiveLog("Init::AppFolder: " & AppFolder(), 1)
        ActiveLog("Init::UpdateFolder: " & UpdateFolder(), 1)
        If IsDevelopment() Then ActiveLog("Init::DEV MODE", 1)

        LogInformationFiles()
    End Sub

    Public Function LogInformationFiles(Optional ByVal Which As String = "X") As String
        Const Y As Boolean = False
        Dim X As Boolean, S As String, F As String
        X = InStr(Which, "X")
        If X Or InStr(Which, "C") Then F = "Computer.txt" : S = ComputerInformationString()
        If X Or InStr(Which, "D") Then F = "Directories.txt" : S = DirectoryInformationString()
        If X Or InStr(Which, "M") Then F = "Modules.txt" : S = ModuleInformationString()
        If X Or InStr(Which, "R") Then F = "Registry.txt" : S = RegistrySettingsString()
        If Y Or InStr(Which, "O") Then F = "Components.txt" : S = ComponentSummary()
        If X Or InStr(Which, "G") Then F = "ConfigTable.txt" : S = ConfigTableString()
        If X Or InStr(Which, "P") Then F = "PrinterInfo.txt" : S = PrinterInformationString()
        WriteFile(LogFolder() & F, S, True)
        LogInformationFiles = S
    End Function

    Private Function A(ByVal S As String) As Boolean
        '  MsgBox "a:dbg:" & S
    End Function

    Public Function PrinterInformationString() As String
        Dim L, P As Printer
        Dim I As Integer, S As String, X As Integer

        Const N As String = vbCrLf
        Const M As String = ""
        On Error Resume Next

        ' This function is hanging for some of the local computers...
        Exit Function

        S = ""
        For Each L In Printers
            Application.DoEvents()
            I = I + 1
            P = L
            With P
                S = S & M & "[Printer #" & I & "]" & vbCrLf
                '      S = S & N & "hDC=" & .hDC
                S = S & N & "DeviceName=" & .DeviceName
                S = S & N & "Size=" & .Width & "x" & .Height
                S = S & N & "Scaled=" & .ScaleWidth & "x" & .ScaleHeight & " [" & DescribeScaleMode(.ScaleMode) & "]"
                '      S = S & N & "Inches=" & .ScaleX(.Width, vbTwips, vbInches) & "x" & .ScaleY(.Height, vbwips, vbInches)
                'S = S & N & "DriverName=" & .DriverName
                'S = S & N & "Port=" & .Port
                S = S & N & "PrintQuality=" & .PrintQuality
                S = S & N & "FontCount=" & .FontCount
                S = S & N & "Font=" & .Font.Size & " " & .Font.Name & IIf(.FontBold, " Bold", "") & IIf(.FontItalic, " Italic", "") & IIf(.FontUnderline, " Underline", "")
                S = S & N & "Orientation=" & IIf(.Orientation = vbPRORLandscape, "Landscape", "Portrait")
                S = S & N & "PaperBin=" & .PaperBin
                S = S & N & "PaperSize=" & .PaperSize
                'S = S & N & "Zoom=" & .Zoom
                S = S & N & "TPP=" & .TwipsPerPixelX & "x" & .TwipsPerPixelY
                S = S & N & ""
                S = S & N & ""
            End With
        Next

        PrinterInformationString = S
    End Function

    Public Function ComponentSummary() As String
        Dim I As Integer, S As String, E As Boolean
        Dim D As String, N As Integer
        Dim Tx As String, M As String, L As String
        I = 0

        L = vbCrLf
        Tx = ""

        Tx = Tx & M & SoftwareVersionForLog()
        Tx = Tx & L & "OLE Component Information:"
        Tx = Tx & L & ""
        Tx = Tx & L & "Report Date: " & Now
        Tx = Tx & L & "Computer Name: " & GetLocalComputerName()
        Tx = Tx & L & ""

        Do While True
            S = ComponentName(I)
            If S = "" Then Exit Do
            E = ComponentExists(I, D, N)
            Tx = Tx & L & Format(I, "000") & ": " & IIf(E, "+", "-") & S & IIf(E Or N = 429, "", " *** [" & Hex(N) & "] " & D)
            I = I + 1
        Loop
        ComponentSummary = Tx
    End Function

    Public Function ComponentExists(ByVal ComponentID As eComponentIDs, Optional ByRef vEMsg As String = "", Optional ByRef vENum As Integer = 0) As Boolean
        ComponentExists = ObjectExists(ComponentName(ComponentID), vEMsg, vENum)
    End Function

    Public Function ObjectExists(ByVal S As String, Optional ByRef vEMsg As String = "", Optional ByRef vENum As Integer = 0) As Boolean
        Dim R As Object
        If S = "" Then Exit Function
        On Error Resume Next
        R = CreateObject(S)
        ObjectExists = Not (R Is Nothing)
        If Not ObjectExists Then
            vEMsg = Err.Description
            vENum = Err.Number
        End If
        R = Nothing
    End Function

    Public Function ComponentName(ByVal ComponentID As eComponentIDs) As String
        Select Case ComponentID
            Case eComponentIDs.cidDAOEngine36 : ComponentName = "DAO.DBEngine.36"

            Case eComponentIDs.cidMSExcel : ComponentName = "Excel.Application"
            Case eComponentIDs.cidMSWord : ComponentName = "Word.Application"
            Case eComponentIDs.cidMSOutlook : ComponentName = "Outlook.Application"
'    Case cidMSMapPoint:             ComponentName = "Mappoint.Application"
            Case eComponentIDs.cidMSComm : ComponentName = "MSCommlib.MSComm"

            Case eComponentIDs.cidScriptingDictionary : ComponentName = "Scripting.Dictionary"
            Case eComponentIDs.cidScriptingFSO : ComponentName = "Scripting.FileSystemObject"

            Case eComponentIDs.cidWScriptShell : ComponentName = "WScript.Shell"
            Case eComponentIDs.cidWScriptNetwork : ComponentName = "WScript.Network"

            Case eComponentIDs.cidCDOMessage : ComponentName = "CDO.Message"
            Case eComponentIDs.cidCDOConfiguration : ComponentName = "CDO.Configuration"
            Case eComponentIDs.cidMAPISession : ComponentName = "MAPI.Session"
            Case eComponentIDs.cidMSMAPISession : ComponentName = "MSMAPI.MAPISession"

            Case eComponentIDs.cidvbSendMail : ComponentName = "vbSendMail.clsSendMail"
            Case eComponentIDs.cidXChargeTransaction : ComponentName = "XCTransaction2.XChargeTransaction"
            Case eComponentIDs.cidSCPinPadDevice : ComponentName = "PINPadDevice.PINPad"

            Case eComponentIDs.cidQBFC5SM : ComponentName = "QBFC5.QBSessionManager"
'    Case cidQBFC6SM:                ComponentName = "QBFC6.QBSessionManager"
'    Case cidQBFC7SM:                ComponentName = "QBFC7.QBSessionManager"
            Case eComponentIDs.cidQBFC8SM : ComponentName = "QBFC8.QBSessionManager"
'    Case cidQBFC9SM:                ComponentName = "QBFC9.QBSessionManager"
            Case eComponentIDs.cidQBFC10SM : ComponentName = "QBFC10.QBSessionManager"
'    Case cidQBFC11SM:               ComponentName = "QBFC11.QBSessionManager"
            Case eComponentIDs.cidQBFC12SM : ComponentName = "QBFC12.QBSessionManager"
            Case eComponentIDs.cidQBFC13SM : ComponentName = "QBFC13.QBSessionManager"
                '    Case cidQBFC14SM:               ComponentName = "QBFC14.QBSessionManager"
                '    Case cidQBFC15SM:               ComponentName = "QBFC15.QBSessionManager"
                '    Case cidQBFC16SM:               ComponentName = "QBFC16.QBSessionManager"
                '    Case cidQBFC17SM:               ComponentName = "QBFC17.QBSessionManager"
                '    Case cidQBFC18SM:               ComponentName = "QBFC18.QBSessionManager"
                '    Case cidQBFC19SM:               ComponentName = "QBFC19.QBSessionManager"
                '    Case cidQBFC20SM:               ComponentName = "QBFC20.QBSessionManager"

            Case Else : ComponentName = ""

        End Select
    End Function

    Public Function ModuleInformationString() As String
        Dim Tx As String, L As String
        On Error Resume Next

        A("Maa") : Tx = ""
        A("Mab") : Tx = Tx & SoftwareVersionForLog() & vbCrLf
        A("Mac") : Tx = Tx & "System Update and Module Information:" & vbCrLf
        A("Mad") : Tx = Tx & vbCrLf
        A("Mae") : Tx = Tx & "Report Date: " & Now & vbCrLf
        A("Maf") : Tx = Tx & "Computer Name: " & GetLocalComputerName() & vbCrLf
        If IsCDSComputer(L) Then Tx = Tx & "CDS Computer:  " & L & vbCrLf
        A("Mag") : Tx = Tx & "WinDir:        " & GetWindowsDir() & vbCrLf
        A("Mah") : Tx = Tx & "WinSysDir:     " & GetWindowsSystemDir() & vbCrLf
        A("Mai") : Tx = Tx & vbCrLf
        A("Maj") : Tx = Tx & "IRLib.dll:     " & YesNo(FileExists(AppFolder() & "IRLib.dll")) & vbCrLf
        A("Mak") : Tx = Tx & "vbZip10.dll:   " & YesNo(FileExists(AppFolder() & "vbzip10.dll")) & vbCrLf
        A("Mal") : Tx = Tx & "vbZip11.dll:   " & YesNo(FileExists(AppFolder() & "vbzip11.dll")) & vbCrLf
        A("Mam") : Tx = Tx & "SetACL.exe:    " & YesNo(ACLExists) & vbCrLf
        A("Man") : Tx = Tx & "FreeImage.dll: " & YesNo(FreeImage_IsAvailable) & vbCrLf


        ModuleInformationString = Tx
    End Function

    Public Function RegistrySettingsString() As String
        Dim Tx As String
        On Error Resume Next

        RegistrySettingsString = Tx
    End Function

    Public Function ConfigTableString() As String
        Dim Tx As String, R As ADODB.Recordset
        Dim N As Integer
        Const Max As Integer = 1000
        On Error Resume Next

        R = GetRecordsetBySQL("SELECT * FROM [Config]", , GetDatabaseInventory)

        Tx = ""
        Tx = Tx & "#" & SoftwareVersionForLog() & vbCrLf
        Tx = "[Config]" & vbCrLf
        Do While Not R.EOF
            N = N + 1
            If N > Max Then Exit Do
            Tx = Tx & R("FieldName").Value & "=" & R("Value").Value & vbCrLf
            R.MoveNext()
        Loop

        ConfigTableString = Tx
    End Function

    Public Function DirectoryInformationString() As String
        Dim Tx As String, L As String
        Dim N As String, M As String
        M = ""
        N = vbCrLf
        On Error Resume Next

        A("Daa") : Tx = ""
        A("Dab") : Tx = Tx & M & SoftwareVersionForLog()
        A("Dac") : Tx = Tx & N & "Directory Information:"
        A("Dad") : Tx = Tx & N & ""
        A("Dae") : Tx = Tx & N & "Report Date: " & Now
        A("Daf") : Tx = Tx & N & "Computer Name: " & GetLocalComputerName()
        A("Dag") : Tx = Tx & N & "EXE: " & WinCDSEXEFile(True)
        A("Dah") : Tx = Tx & N & ""
        A("Dai") : Tx = Tx & N & "WinCDS Folders:"
        A("Daj") : Tx = Tx & N & "  AppFolder:           " & AppFolder() '& ACL_FA(AppFolder)
        A("Dak") : Tx = Tx & N & "  DataFolder:          " & CDSDataFolder() '& ACL_FA(InventFolder)
        A("Dal") : Tx = Tx & N & "  InventFolder:        " & InventFolder() '& ACL_FA(InventFolder)
        A("Dam") : Tx = Tx & N & "  Store1Folder:        " & StoreFolder(1) '& ACL_FA(StoreFolder(1))
        A("Dan") : Tx = Tx & N & "  LStore1Folder:       " & LocalStoreFolder(1) '& ACL_FA(LocalStoreFolder(1))
        A("Dao") : Tx = Tx & N & "  New Order Folder:    " & NewOrderFolder(1) '& ACL_FA(NewOrderFolder(1))
        A("Dap") : Tx = Tx & N & "  PX Folder:           " & PXFolder() '& ACL_FA(PXFolder)
        A("DaP") : Tx = Tx & N & "  FX Folder:           " & FXFolder() '& ACL_FA(FXFolder)
        A("Daq") : Tx = Tx & N & "  Update Folder:       " & UpdateFolder() '& ACL_FA(UpdateFolder)
        A("Dar") : Tx = Tx & N & "  Reports Folder:      " & ReportsFolder '& ACL_FA(UpdateFolder)
        A("Das") : Tx = Tx & N & "  Program Files:       " & ProgramFilesFolder() '& ACL_FA(ProgramFilesFolder)
        A("Dat") : Tx = Tx & N & "  LPF:                 " & LocalProgramFilesFolder() '& ACL_FA(LocalProgramFilesFolder)
        A("Dau") : Tx = Tx & N & "  Desktop:             " & LocalDesktopFolder() '& ACL_FA(LocalDesktopFolder)
        A("Dav") : Tx = Tx & N & "  All Users Desktop:   " & AllUsersDesktopFolder() '& ACL_FA(AllUsersDesktopFolder)
        A("Daw") : Tx = Tx & N & "  WinCDS Folder:       " & WinCDSFolder() '& ACL_FA(WinCDSFolder)
        A("Dax") : Tx = Tx & N & "  Development Folder:  " & DevelopmentFolder '& ACL_FA(DevelopmentFolder)
        A("Day") : Tx = Tx & N & ""
        A("Daz") : Tx = Tx & N & "Windows Folders:"
        A("Dba") : Tx = Tx & N & "  WinDir:              " & SpecialFolder(FolderEnum.feWindows)
        A("Dbb") : Tx = Tx & N & "  Fonts:               " & SpecialFolder(FolderEnum.feWindowsFonts)
        A("Dbc") : Tx = Tx & N & "  Resources:           " & SpecialFolder(FolderEnum.feWindowsResources)
        A("Dbd") : Tx = Tx & N & "  System:              " & SpecialFolder(FolderEnum.feWindowsSystem)
        A("Dbe") : Tx = Tx & N & "  CD Burning:          " & SpecialFolder(FolderEnum.feCDBurnArea)
        A("Dbf") : Tx = Tx & N & ""
        A("Dbg") : Tx = Tx & N & "Program Files:"
        A("Dbh") : Tx = Tx & N & "  PF Dir:              " & SpecialFolder(FolderEnum.feProgramFiles)
        A("Dbi") : Tx = Tx & N & "  PF Common Dir:       " & SpecialFolder(FolderEnum.feProgramFilesCommon)
        A("Dbj") : Tx = Tx & N & ""
        A("Dbk") : Tx = Tx & N & "Local Folders:"
        A("Dbl") : Tx = Tx & N & "  AppData:             " & SpecialFolder(FolderEnum.feLocalAppData)
        A("Dbm") : Tx = Tx & N & "  CD Burning:          " & SpecialFolder(FolderEnum.feLocalCDBurning)
        A("Dbn") : Tx = Tx & N & "  History:             " & SpecialFolder(FolderEnum.feLocalHistory)
        A("Dbo") : Tx = Tx & N & "  Temp Internet Files: " & SpecialFolder(FolderEnum.feLocalTempInternetFiles)
        A("Dbp") : Tx = Tx & N & ""
        A("Dbq") : Tx = Tx & N & "Common Folders:"
        A("Dbr") : Tx = Tx & N & "  AdminTools:          " & SpecialFolder(FolderEnum.feCommonAdminTools)
        A("Dbs") : Tx = Tx & N & "  AppData:             " & SpecialFolder(FolderEnum.feCommonAppData)
        A("Dbt") : Tx = Tx & N & "  Desktop:             " & SpecialFolder(FolderEnum.feCommonDesktop)
        A("Dbu") : Tx = Tx & N & "  Docs:                " & SpecialFolder(FolderEnum.feCommonDocs)
        A("Dbv") : Tx = Tx & N & "  Music:               " & SpecialFolder(FolderEnum.feCommonMusic)
        A("Dbw") : Tx = Tx & N & "  Pics:                " & SpecialFolder(FolderEnum.feCommonPics)
        A("Dbx") : Tx = Tx & N & "  Start:               " & SpecialFolder(FolderEnum.feCommonStartMenu)
        A("Dby") : Tx = Tx & N & "  StartMenu:           " & SpecialFolder(FolderEnum.feCommonStartMenuPrograms)
        A("Dbz") : Tx = Tx & N & "  Templates:           " & SpecialFolder(FolderEnum.feCommonTemplates)
        A("Dca") : Tx = Tx & N & "  Videos:              " & SpecialFolder(FolderEnum.feCommonVideos)
        A("Dcb") : Tx = Tx & N & ""
        A("Dcc") : Tx = Tx & N & "User Folders:"
        A("Dcd") : Tx = Tx & N & "  User:              " & SpecialFolder(FolderEnum.feUser)
        A("Dce") : Tx = Tx & N & "  AdminTools:        " & SpecialFolder(FolderEnum.feUserAdminTools)
        A("Dcf") : Tx = Tx & N & "  AppData:           " & SpecialFolder(FolderEnum.feUserAppData)
        A("Dcg") : Tx = Tx & N & "  Cache:             " & SpecialFolder(FolderEnum.feUserCache)
        A("Dch") : Tx = Tx & N & "  Cookies:           " & SpecialFolder(FolderEnum.feUserCookies)
        A("Dci") : Tx = Tx & N & "  Desktop:           " & SpecialFolder(FolderEnum.feUserDesktop)
        A("Dcj") : Tx = Tx & N & "  Docs:              " & SpecialFolder(FolderEnum.feUserDocs)
        A("Dck") : Tx = Tx & N & "  Favorites:         " & SpecialFolder(FolderEnum.feUserFavorites)
        A("Dcl") : Tx = Tx & N & "  Music:             " & SpecialFolder(FolderEnum.feUserMusic)
        A("Dcm") : Tx = Tx & N & "  NetHood:           " & SpecialFolder(FolderEnum.feUserNetHood)
        A("Dcn") : Tx = Tx & N & "  Pics:              " & SpecialFolder(FolderEnum.feUserPics)
        A("Dco") : Tx = Tx & N & "  PrintHood:         " & SpecialFolder(FolderEnum.feUserPrintHood)
        A("Dcp") : Tx = Tx & N & "  Recent:            " & SpecialFolder(FolderEnum.feUserRecent)
        A("Dcq") : Tx = Tx & N & "  SendTo:            " & SpecialFolder(FolderEnum.feUserSendTo)
        A("Dcr") : Tx = Tx & N & "  StartMenu:         " & SpecialFolder(FolderEnum.feUserStartMenu)
        A("Dcs") : Tx = Tx & N & "  Start:             " & SpecialFolder(FolderEnum.feUserStartMenuPrograms)
        A("Dct") : Tx = Tx & N & "  Startup:           " & SpecialFolder(FolderEnum.feUserStartup)
        A("Dcu") : Tx = Tx & N & "  Templates:         " & SpecialFolder(FolderEnum.feUserTemplates)
        A("Dcv") : Tx = Tx & N & "  Videos:            " & SpecialFolder(FolderEnum.feUserVideos)

        DirectoryInformationString = Tx
    End Function
End Module
