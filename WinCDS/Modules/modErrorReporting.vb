Module modErrorReporting
    Private ErrorContext As String
    Private ErrorNumber as integer, ErrorDesc As String
    Private Const ErrPrefix As String = "err_"
    Private ProgramState As String
    Private VarDump As String
    Private CompInfo As String
    Public Function ReportError(Optional ByVal vErrorContext As String = "", Optional ByVal vErrorNumber as integer = 0, Optional ByVal vErrorDesc As String = "") As Boolean
        '::::ReportError
        ':::SUMMARY
        ': Generates a new WinCDS `crash dump`
        ':::DESCRIPTION
        ': Builds a zip file containing a representation of current program state.
        ':::RETURN
        ': Boolean - Returns True on Success
        ':::SEE ALSO
        ': CleanOldErrorReports
        Dim T As String, F As String
        Dim S as integer, X as integer

        Const DO_ZIP As Boolean = True

        On Error Resume Next
        ErrorContext = vErrorContext
        If vErrorNumber = 0 Then ErrorNumber = Err.Number Else ErrorNumber = vErrorNumber
        If vErrorDesc = "" Then ErrorDesc = Err.Description Else ErrorDesc = vErrorDesc

        S = GetTickCount
        ReportError_PrepareReport

        T = TempFolder(ErrorReportingFolder, ErrPrefix)
        F = GetFileName(TempFile(ErrorReportingFolder, ErrPrefix, ".zip"))
        ReportError_WriteToFolder(T)

        If DO_ZIP Then
            ZipFiles(T, ErrorReportingFolder, F)
            'RemoveFolder(T)
        End If

        X = GetTickCount
        Debug.Print("Error Reporting Cost: " & DescribeTimeDurationMS(X - S))

        ReportError = DiagnosticErrorUpload(ErrorReportingFolder & F)
        Debug.Print("Upload Cost: " & DescribeTimeDurationMS(GetTickCount - X))
    End Function
    Private Function ReportError_PrepareReport() As Boolean
        On Error Resume Next
        '  ProgressForm 0, 1, "Preparing Error Context...", vbAbortRetryIgnore, , , prgIndefinite
        ReportError_BuildScreenshots()
        'ProgramState = ReportError_BuildProgramState
        'VarDump = ReportError_BuildVarDump
        CompInfo = ComputerInformationString
        '  ProgressForm

        ReportError_PrepareReport = True
    End Function
    Private ReadOnly Property ErrorReportingFolder() As String
        Get
            ErrorReportingFolder = InventFolder()
        End Get
    End Property
    Private Function ReportError_WriteToFolder(ByVal Folder As String) As Boolean
        Dim L, I as integer
        'SavePicture(Screenshot, Folder & "Screenshot.bmp")
        'For I = 0 To Forms.Count - 1
        On Error Resume Next
        'SavePicture(pForms(I), Folder & "Form_" & Format(I + 1, "00") & ".bmp")
        On Error GoTo 0
        'WriteFile(Folder & "FormVars_" & Format(I + 1, "00") & ".txt", tForms(I))
        'Next

        WriteFile(Folder & "ProgramState.txt", ProgramState)
        WriteFile(Folder & "VarDump.txt", VarDump)
        WriteFile(Folder & "CompInfo.txt", CompInfo)

    End Function
    Private Function ReportError_BuildScreenshots() as integer
        Dim I as integer, L, Val As String
        Dim T As String
        Dim S() As String
        'Screenshot = CaptureScreen
        ReportError_BuildScreenshots = 1
        'ReDim pForms(0 To Forms.Count - 1)
        'ReDim tForms(0 To Forms.Count - 1)
        'For I = 0 To Forms.Count - 1
        '    pForms(I) = CaptureForm(Forms(I))
        '    ReportError_BuildScreenshots = ReportError_BuildScreenshots + 1
        '    tForms(I) = "FORM " & Format(I, "00") & ": " & Forms(I).Name
        '    S = TLIObjectMembers(Forms(I))
        For Each L In S
                If IsIn(L, "_Default", "Name", "MouseX", "MouseY", "CurrentMenuIndex") Then GoTo DoContinue
                On Error Resume Next
                Val = ""
            'Val = CallByName(Forms(I), L, vbGet)
            On Error Resume Next
            'If Val <> "" Then tForms(I) = tForms(I) & N & "     " & AlignString(Forms(I).Name & "." & L, 28, vbAlignRight) & " = " & Val
DoContinue:
            Next
        'Next
    End Function
    '    Private Function ReportError_BuildProgramState() As String
    '        S = ""
    '        S = S & M & ""
    '        S = S & M & SoftwareVersionForLog

    '        If ErrorContext <> "" Then S = S & N & "Context: " & ErrorContext
    '        If ErrorNumber <> 0 Or ErrorDesc <> "" Then S = S & N & "Error: [" & ErrorNumber & "] " & ErrorDesc
    '        S = S & N

    '        S = S & N & "Loaded Forms (" & Forms.Count & "): " & FormList
    '        If IsNotNothing(Screen) Then
    '            If IsNotNothing(Screen.ActiveForm) Then
    '                On Error Resume Next
    '                S = S & N & "Active Form: " & Screen.ActiveForm.Name
    '            End If
    '        End If

    '        S = S & N

    '        S = S & N & "Order=" & Order
    '        S = S & N & "Inven=" & Inven
    '        S = S & N & "ArSelect=" & ArSelect
    '        S = S & N & "Reports=" & Reports
    '        S = S & N & "Mail=" & Mail
    '        S = S & N & "PurchaseOrder=" & PurchaseOrder
    '        S = S & N & ""

    '        ReportError_BuildProgramState = S
    '    End Function
    '    Private Function ReportError_BuildVarDump(Optional ByVal fWrite As Boolean = False) As String

    '        Dim C as integer, I as integer, R As CDSFunc, Val As String, T As String
    '        Dim FL As String, LL, L, Lines as integer

    '        Const nV = "[NO VALUE]"
    '        S = ""
    '        S = S & M & ""
    '        S = S & M & "=================  Variable Dump  ================="
    '        InitCDSFunctions

    '#If True Then

    '        For I = 1 To CDSFunctionListCount
    '            With CDSFunctionList(I)
    '                If (.bFunction = cdsVariable Or .bFunction = cdsPropertyGet) And Not .bPrivate Then
    '                    On Error Resume Next
    '                    Val = nV
    '                    Val = CDSFunctionDispatch(.Name)
    '                    On Error GoTo 0
    '                    If Val <> nV And Trim(Val) <> "" Then
    '                        S = S & N & "     " & AlignString(Replace(.Module, ".bas", "") & "." & .Name, 28, vbAlignRight) & " = " & Val
    '                        Lines = Lines + 1
    '                    End If
    '                End If
    '            End With
    '        Next

    '#Else

    '  FL = CDSFunctionModules
    '  LL = Split(FL, vbCrLf)

    '  For Each L In LL
    '    Lines = 0
    '    T = "*** MODULE: " & L

    '    C = CDSFunctionListModuleCount(L)
    '    For I = 1 To C
    '      R = CDSFunctionListModule(L, I)
    '      With R
    '        If (.bFunction = cdsVariable Or .bFunction = cdsPropertyGet) And Not .bPrivate Then
    'On Error Resume Next
    '          Val = nV
    '          Val = CDSFunctionDispatch(R.Name)
    'On Error GoTo 0
    '          If Val <> nV And Trim(Val) <> "" Then
    '            T = T & N & "     " & AlignString(.Name, 28, vbAlignRight) & " = " & Val
    '            Lines = Lines + 1
    '          End If
    '        End If
    '      End With
    '    Next
    '    If Lines > 0 Then S = S & N & N & T
    '  Next

    '#End If

    '        If fWrite Then WriteFile(DevOutputFolder() & "var_dump.txt", S, True)
    '        ReportError_BuildVarDump = S
    '    End Function
    Public Function ComputerInformationString() As String
        On Error Resume Next
        Dim Tx As String, L As String, C As String
        Dim M As String, N As String
        Dim I as integer
        On Error Resume Next
        'Dim K As New frmWinsock

        N = vbCrLf
        M = ""

        Tx = ""
        Tx = Tx & M & ""
        'A "aaa": Tx = Tx & M & SoftwareVersionForLog
        'A "aab": Tx = Tx & N & "Application And Computer Information:"
        'A "aac": Tx = Tx & N & ""
        'A "aad": Tx = Tx & N & "Report Date: " & Now
        'Tx = Tx & N & "UAC Is Admin:  " & UACIsAdmin
        'A "aae": Tx = Tx & N & "Computer Name: " & GetLocalComputerName()
        'A "aa.": Tx = Tx & N & "External IP:   " & ExternalIPAddress()
        'A "aa,": Tx = Tx & N & "Apparent Store:" & KnownStoreNameByIP(C)
        'A "aa+": Tx = Tx & N & "Capabilities:  " & C
        'If IsCDSComputer(L) Then Tx = Tx & N & "CDS Computer:  " & L
        'A "aaf": Tx = Tx & N & "WinDir:        " & GetWindowsDir
        'A "aag": Tx = Tx & N & "WinSysDir:     " & GetWindowsSystemDir
        'A "aah": Tx = Tx & N & "User Dir:      " & UserFolder
        'A "aai": Tx = Tx & N & "All Users Dir: " & AllUsersFolder
        'A "aaj": Tx = Tx & N & "AppData Dir:   " & AppDataFolder()
        'A "aak": Tx = Tx & N & "Temp Dir:      " & GetTempDir
        'A "aal": Tx = Tx & N & "CDSAppData Dr: " & WinCDSDataPath
        'A "aam": Tx = Tx & N & "Sys User Name: " & GetSystemUserName
        'A "aan": Tx = Tx & N & "Local IP:      " & K.LocalIP '& " (MAC: " & GetMacAddress & ")"  ' GetMacAddress causes Full Crash on some XP machines.
        'A "aao": Tx = Tx & N & "Lcl HostName:  " & K.LocalHostName
        'If MainMenuIsLoaded Then
        '    A "aap": Tx = Tx & N & "Max Colors:    " & GetMaxColors(MainMenu.hDC)
        'End If
        'A "aaq": Tx = Tx & N & "Remote?        " & YesNo(SessionIsRemote)
        'A ".aq": Tx = Tx & N & "WinVer:        " & GetWinVer
        'A "aar": Tx = Tx & N & ""
        'A "aas": Tx = Tx & N & "Application:   " & App.Title & " (" & App.ProductName & ")"
        'A "aat": Tx = Tx & N & "Company Nme:   " & App.CompanyName
        'A "aau": Tx = Tx & N & "Version:       " & SoftwareVersion(False, False) & " (Build: " & WinCDSRevisionNumber & ") [*" & VersionRepresents & "*]"
        'A "aav": Tx = Tx & N & "EXE Date:      " & FileDateTime(WinCDSEXEFile(True))
        'A "aaw": Tx = Tx & N & "Build Date:    " & BuildDate()
        'A "aax": Tx = Tx & N & "License:       " & License
        'A "aay": Tx = Tx & N & "Installment:   " & InstallmentLicense
        'A "aaz": Tx = Tx & N & "Dev Mode:      " & DevelopmentModeDescriptor
        'A "aba": Tx = Tx & N & "Kill Date:     " & KillDate
        'If CrippleDate <> YearAdd(Of Date, "yyyy") Then Tx = Tx & N & "Cripple Date:  " & CrippleDate
        'A "abb": Tx = Tx & N & ""
        'A "abc": Tx = Tx & N & "Server:        " & YesNo(IsServer)
        'A "abd": Tx = Tx & N & "AppFolder:     " & AppFolder()
        'A "abe": Tx = Tx & N & "Invent Fld:    " & InventFolder()
        'A "abf": Tx = Tx & N & "Store1 Fld:    " & StoreFolder(1)
        'A "abg": Tx = Tx & N & "PXFolder:      " & PXFolder()
        'A "abG": Tx = Tx & N & "FXFolder:      " & FXFolder()
        'A "abh": Tx = Tx & N & "Update Fld:    " & UpdateFolder()
        'A "abi": Tx = Tx & N & "Reports Fld:   " & ReportsFolder
        'A "abj": Tx = Tx & N & "AP Fld:        " & APFolder
        'A "abk": Tx = Tx & N & "payable.exe:   " & FileAccountPayable
        'A "abl": Tx = Tx & N & ""
        'A "abm": Tx = Tx & N & "SetACL.exe:    " & YesNo(ACLExists)
        'A "abn": Tx = Tx & N & ""
        'A "abo": Tx = Tx & N & "Access Version "
        'A "abp": Tx = Tx & N & "Invent AV:     " & AccessVersion(GetDatabaseInventory)

        For I = 1 To Setup_MaxStores
            If LicensedNoOfStores >= I Then
                'Tx = Tx & N & "Store" & Format(I, "00") & " AV:    " & AccessVersion(GetDatabaseAtLocation(2))
            End If
        Next

        'A "abq": Tx = Tx & N & ""
        'A "abr": Tx = Tx & N & SoftwareCopyright
        'A "abs": Tx = Tx & N & ""

        'Unload K

        ComputerInformationString = Tx
    End Function

End Module
