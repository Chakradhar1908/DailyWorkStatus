Module ShellOut
    Public Const CREATE_NO_WINDOW = &H8000000
    Public Const NORMAL_PRIORITY_CLASS = &H20&
    Public Const INFINITE = -1&
    Private Const ASW As String = "AppShell.Form1.ShellAndWait: "
    Public LastProcessID As Integer
    Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Integer, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Integer, ByVal lpThreadAttributes As Integer, ByVal bInheritHandles As Integer, ByVal dwCreationFlags As Integer, ByVal lpEnvironment As Integer, ByVal lpCurrentDirectory As Integer, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Integer
    Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
    Declare Function CloseHandle Lib "kernel32" (hObject As Integer) As Boolean
    Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer 'vb6.0
    'Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess as integer, ByVal bInheritHandle as integer, ByVal dwProcessId as integer) As IntPtr  '-->vb.net
    Enum EnSW
        enSW_HIDE = 0
        enSW_NORMAL = 1
        enSW_MAXIMIZE = 3
        enSW_MINIMIZE = 6
    End Enum
    Public Structure PROCESS_INFORMATION
        Dim hProcess As Integer
        Dim hThread As Integer
        Dim dwProcessId As Integer
        Dim dwThreadId As Integer
    End Structure
    Public Structure STARTUPINFO
        Dim Cb As Integer
        Dim lpReserved As Integer ' !!! must be Long for Unicode string
        Dim lpDesktop As Integer  ' !!! must be Long for Unicode string
        Dim lpTitle As Integer    ' !!! must be Long for Unicode string
        Dim dwX As Integer
        Dim dwY As Integer
        Dim dwXSize As Integer
        Dim dwYSize As Integer
        Dim dwXCountChars As Integer
        Dim dwYCountChars As Integer
        Dim dwFillAttribute As Integer
        Dim dwFlags As Integer
        Dim wShowWindow As Integer
        Dim cbReserved2 As Integer
        Dim lpReserved2 As Integer
        Dim hStdInput As Integer
        Dim hStdOutput As Integer
        Dim hStdError As Integer
    End Structure
    Private Structure OSVERSIONINFO
        Dim dwOSVersionInfoSize As Integer
        Dim dwMajorVersion As Integer
        Dim dwMinorVersion As Integer
        Dim dwBuildNumber As Integer
        Dim dwPlatformId As Integer
        <VBFixedString(128)> Dim szCSDVersion As String
    End Structure
    Private Structure PROCESSENTRY32
        Dim dwSize As Integer
        Dim cntUsage As Integer
        Dim th32ProcessID As Integer
        Dim th32DefaultHeapID As Integer
        Dim th32ModuleID As Integer
        Dim cntThreads As Integer
        Dim th32ParentProcessID As Integer
        Dim pcPriClassBase As Integer
        Dim dwFlags As Integer
        <VBFixedString(260)> Dim szExeFile As String
    End Structure

    Public Sub RunWordpad(ByVal tFile As String, Optional ByVal WindowStyle As AppWinStyle = vbNormalFocus, Optional ByVal AndWait As Boolean = False)
        If AndWait Then
            ShellAndWait("write " & QuoteString(tFile), WindowStyle)
        Else
            DoShell("write " & QuoteString(tFile), WindowStyle)
        End If
    End Sub

    ' to allow for Shell.
    ' This routine shells out to another application and waits for it to exit.
    Public Sub ShellAndWait(AppToRun As String, Optional ByVal SW As EnSW = EnSW.enSW_NORMAL)
        Dim NameOfProc As PROCESS_INFORMATION
        Dim NameStart As STARTUPINFO
        Dim rc As Integer
        Dim e As New Exception

        LogFile("ShellAndWait", AppToRun, False)

        On Error GoTo ErrorRoutineErr
        NameStart.Cb = Len(NameStart)
        If SW = EnSW.enSW_HIDE Then
            rc = CreateProcessA(0&, AppToRun, 0&, 0&, CLng(SW), CREATE_NO_WINDOW, 0&, 0&, NameStart, NameOfProc)
        Else
            rc = CreateProcessA(0&, AppToRun, 0&, 0&, CLng(SW), NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc)
        End If
        LastProcessID = NameOfProc.dwProcessId
        rc = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
        rc = CloseHandle(NameOfProc.hProcess)

        'ErrorRoutineResume:
        '        Exit Sub
ErrorRoutineErr:
        MsgBox(ASW & e.Message)
        Resume Next
    End Sub

    Public Function DoShell(ByVal App As String, Optional ByVal WindowStyle As AppWinStyle = vbMinimizedFocus) As Integer
        LastProcessID = Shell(App, WindowStyle)
        DoShell = LastProcessID
    End Function

    Public Function RunCmdToOutput(ByVal cmd As String, Optional ByRef ErrStr As String = "", Optional ByVal AsAdmin As Boolean = False) As String
        On Error GoTo RunError
        Dim A As String, B As String, C As String
        Dim tLen As Integer, Iter As Integer
        A = TempFile()
        B = TempFile()

        If Not AsAdmin Then
            ShellAndWait("cmd /c " & cmd & " 1> " & A & " 2> " & B, EnSW.enSW_HIDE)
        Else
            C = TempFile(, , ".bat")
            WriteFile(C, cmd & " 1> " & A & " 2> " & B, True)
            RunFileAsAdmin(C, , EnSW.enSW_HIDE)
        End If

        Iter = 0
        Const MaxIter As Integer = 10
        Do While True
            tLen = FileLen(A)
            Sleep(800)
            If Iter > MaxIter Or FileLen(A) = tLen Then Exit Do
            Iter = Iter + 1
        Loop
        RunCmdToOutput = ReadEntireFileAndDelete(A)
        If Iter > MaxIter Then RunCmdToOutput = RunCmdToOutput & vbCrLf2 & "<<< OUTPUT TRUNCATED >>>"
        ErrStr = ReadEntireFileAndDelete(B)
        DeleteFileIfExists(C)
        Exit Function

RunError:
        RunCmdToOutput = ""
        ErrStr = "ShellOut.RunCmdToOutput: Command Execution Error - [" & Err.Number & "] " & Err.Description
    End Function

    Public Function RunFileAsAdmin(ByVal App As String, Optional ByVal nHwnd As Integer = 0, Optional ByVal WindowState As Integer = modAPI.SW_SHOWNORMAL) As Boolean
        If Not IsWinXP() Then
            RunShellExecuteAdmin(App, nHwnd, WindowState)
        Else
            ShellOut(App)
        End If
        RunFileAsAdmin = True
    End Function

    Public Sub ShellOut(ByVal App As String)
        Dim NameOfProc As PROCESS_INFORMATION
        Dim NameStart As STARTUPINFO
        Dim rc As Integer

        On Error GoTo ErrorRoutineErr
        NameStart.Cb = Len(NameStart)
        rc = CreateProcessA(0&, App, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc)
        LastProcessID = NameOfProc.dwProcessId
        '    rc = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
        '    rc = CloseHandle(NameOfProc.hProcess)

ErrorRoutineResume:
        Exit Sub
ErrorRoutineErr:
        MsgBox("AppShell.Form1.ShellOut: ", Err.Description)
        Resume Next
    End Sub

    Public Function ShellOut_URL(ByVal URL As String, Optional ByVal WaitForExit As Boolean = False) As Boolean
        Dim Res As Integer

        ' it is mandatory that the URL is prefixed with http:// or https://
        If Left(URL, 2) = "C:" Then
            URL = "file:///" & URL
        ElseIf InStr(1, URL, "http", vbTextCompare) <> 1 Then
            URL = "http://" & URL
        End If

        Res = ShellExecute(0&, "open", URL, vbNullString, vbNullString, vbNormalFocus)
        ShellOut_URL = (Res > 32)


        '  Dim IEX As String
        '  IEX = LocalProgramFilesFolder & "Internet Explorer\IExplore.exe "  '###x86
        '
        '  If Not FileExists(IEX) Then
        '    MsgBox "We could not find Microsoft Internet Explorer on your machine."
        '  End If
        '
        '  If WaitForExit Then
        '    ShellAndWait IEX & URL
        '  Else
        '    ShellOut IEX & URL
        '  End If
    End Function

    Public Sub RunFile(ByVal tFile As String, Optional ByVal DoMaximized As Boolean = True)
        Const SW_SHOWNORMAL As Integer = 1
        Const SW_SHOWMAXIMIZED As Integer = 3
        Const SW_SHOWMINIMIZED As Integer = 2

        LastProcessID = ShellExecute(0&, "open", tFile, "", GetFilePath(tFile), IIf(DoMaximized, SW_SHOWMAXIMIZED, SW_SHOWNORMAL))
    End Sub

    Public Function RunCmdToOutputWithArgs(ByVal FullPathToCmd As String, Optional ByVal Args As String = "", Optional ByRef ErrStr As String = "") As String
        Dim ShortCmdName As String
        If FullPathToCmd = "" Then
            ErrStr = "WINCDS FAILURE - No Full Command Path Specified in RunCmdToOutputWithArgs"
            Exit Function
        End If

        If Not FileExists(FullPathToCmd) Then
            ErrStr = "WINCDS FAILURE - Command File does not exist:" & vbCrLf & "CMD: " & FullPathToCmd
            Exit Function
        End If

        ShortCmdName = GetShortName(FullPathToCmd)
        If FullPathToCmd <> "" And ShortCmdName = "" Then
            ErrStr = "WINCDS FAILURE - Unable to obtain Short Command Name (GetShortName via RunCmdToOutputWithArgs)" & vbCrLf & "CMD: " & FullPathToCmd
            Exit Function
        End If

        RunCmdToOutputWithArgs = RunCmdToOutput(ShortCmdName & " " & Args, ErrStr)
    End Function

    ' This routine changes the directory temporarily and executes a batch file, returns and sets the current directory back to the previous location.
    Public Sub ShellOut_Shell(tForm As Form, Name As String)
        On Error GoTo HandleErr

        '    WriteFile AppFolder & "shellout.txt", "shellout: " & Name

        With tForm
            Dim WindowStatePrevious As Integer
            WindowStatePrevious = .WindowState
            .WindowState = 1
            .Refresh()
            '    Dim t As Single: t = Timer
            '    Do:    Loop Until (Timer - t) > 4
            ' add delay to allow for tForm window to minimize

            ShellAndWait(Name)
            .WindowState = WindowStatePrevious
            .Select()
        End With
        Exit Sub

HandleErr:
        If Err.Number = 53 Then
            Exit Sub
        End If
    End Sub

End Module
