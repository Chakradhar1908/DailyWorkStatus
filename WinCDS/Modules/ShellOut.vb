
Module ShellOut
    Public Const CREATE_NO_WINDOW = &H8000000
    Public Const NORMAL_PRIORITY_CLASS = &H20&
    Public Const INFINITE = -1&
    Private Const ASW As String = "AppShell.Form1.ShellAndWait: "
    Public LastProcessID as integer
    Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName as integer, ByVal lpCommandLine As String, ByVal lpProcessAttributes as integer, ByVal lpThreadAttributes as integer, ByVal bInheritHandles as integer, ByVal dwCreationFlags as integer, ByVal lpEnvironment as integer, ByVal lpCurrentDirectory as integer, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) as integer
    Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle as integer, ByVal dwMilliseconds as integer) as integer
    Declare Function CloseHandle Lib "kernel32" (hObject as integer) As Boolean
    Enum EnSW
        enSW_HIDE = 0
        enSW_NORMAL = 1
        enSW_MAXIMIZE = 3
        enSW_MINIMIZE = 6
    End Enum
    Public Structure PROCESS_INFORMATION
        Dim hProcess as integer
        Dim hThread as integer
        Dim dwProcessId as integer
        Dim dwThreadId as integer
    End Structure
    Public Structure STARTUPINFO
        Dim Cb as integer
        Dim lpReserved as integer ' !!! must be Long for Unicode string
        Dim lpDesktop as integer  ' !!! must be Long for Unicode string
        Dim lpTitle as integer    ' !!! must be Long for Unicode string
        Dim dwX as integer
        Dim dwY as integer
        Dim dwXSize as integer
        Dim dwYSize as integer
        Dim dwXCountChars as integer
        Dim dwYCountChars as integer
        Dim dwFillAttribute as integer
        Dim dwFlags as integer
        Dim wShowWindow As Integer
        Dim cbReserved2 As Integer
        Dim lpReserved2 as integer
        Dim hStdInput as integer
        Dim hStdOutput as integer
        Dim hStdError as integer
    End Structure
    Private Structure OSVERSIONINFO
        Dim dwOSVersionInfoSize as integer
        Dim dwMajorVersion as integer
        Dim dwMinorVersion as integer
        Dim dwBuildNumber as integer
        Dim dwPlatformId as integer
        <VBFixedString(128)> Dim szCSDVersion As String
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
        Dim rc as integer
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

    Public Function DoShell(ByVal App As String, Optional ByVal WindowStyle As AppWinStyle = vbMinimizedFocus) as integer
        LastProcessID = Shell(App, WindowStyle)
        DoShell = LastProcessID
    End Function

    Public Function RunCmdToOutput(ByVal cmd As String, Optional ByRef ErrStr As String = "", Optional ByVal AsAdmin As Boolean = False) As String
        On Error GoTo RunError
        Dim A As String, B As String, C As String
        Dim tLen as integer, Iter as integer
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
        Const MaxIter as integer = 10
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
    Public Function RunFileAsAdmin(ByVal App As String, Optional ByVal nHwnd as integer = 0, Optional ByVal WindowState as integer = modAPI.SW_SHOWNORMAL) As Boolean
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
        Dim rc as integer

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

End Module
