Module modConnectCMD
    Public Function ConnectCMDUpgrade() As Boolean
        Const CurrentVersion = "2.0"

        If ConnectCMDVersion <> CurrentVersion Then
            WriteFile(ConnectCMDFile, ConnectCMDv2, True)
            If ConnectCMDVersion <> CurrentVersion And Not ServiceMode And Not AutoPatching Then
                If MessageBox.Show("Could not write to Connect.cmd file:" & vbCrLf & ConnectCMDFile() & vbCrLf2 & "Write to desktop?", ProgramName & " Connect Upgrade", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                    WriteFile(LocalDesktopFolder() & "Connect.cmd", ConnectCMDv2, True)
                    MessageBox.Show("Written", ProgramName & " Connect Upgrade")
                End If
            End If
        End If

        ConnectCMDUpgrade = True
    End Function

    Private Function ConnectCMDVersion() As String
        Dim S As String, FN As String, FV As String

        ConnectCMDVersion = "1.0"   ' Default

        S = ReadFile(ConnectCMDFile, 1, 1)
        If Left(Trim(S), 4) <> "REM " Then Exit Function
        S = Mid(Trim(S), 5)

        FN = CSVField(S, 1)
        If UCase(FN) <> "CONNECT.CMD" Then Exit Function ' Who Knows...
        FV = CSVField(S, 2)
        If LCase(Left(FV, 1)) <> "v" Then Exit Function ' Not a version??

        ConnectCMDVersion = Mid(FV, 2)
    End Function

    Private Function ConnectCMDv2() As String
        Dim S As String, N As String, M As String

        N = vbCrLf
        M = ""
        S = ""

        S = S & M & "REM CONNECT.CMD,v2.0  "
        S = S & N & "::::::::::::::::::::::::::::::::::::::::::::  "
        S = S & N & ":: Automatically check & get admin rights V2  "
        S = S & N & "::::::::::::::::::::::::::::::::::::::::::::  "
        S = S & N & "@ECHO OFF "
        S = S & N & "CLS "
        S = S & N & "ECHO. "
        S = S & N & "ECHO =============================  "
        S = S & N & "ECHO Running Admin shell  "
        S = S & N & "ECHO =============================  "
        S = S & N & "  "
        S = S & N & ":init "
        S = S & N & "setlocal DisableDelayedExpansion  "
        S = S & N & "set ""batchPath=%~0"" "
        S = S & N & "for %%k in (%0) do set batchName=%%~nk  "
        S = S & N & "set ""vbsGetPrivileges=%temp%\OEgetPriv_%batchName%.vbs"" "
        S = S & N & "setlocal EnableDelayedExpansion "
        S = S & N & "  "
        S = S & N & ":checkPrivileges  "
        S = S & N & "NET FILE 1>NUL 2>NUL  "
        S = S & N & "if '%errorlevel%' == '0' ( goto gotPrivileges ) else ( goto getPrivileges ) "
        S = S & N & "  "
        S = S & N & ":getPrivileges  "
        S = S & N & "if '%1'=='ELEV' (echo ELEV & shift /1 & goto gotPrivileges) "
        S = S & N & "ECHO. "
        S = S & N & "ECHO ************************************** "
        S = S & N & "ECHO Invoking UAC for Privilege Escalation  "
        S = S & N & "ECHO ************************************** "
        S = S & N & "  "
        S = S & N & "ECHO Set UAC = CreateObject^(""Shell.Application""^) > ""%vbsGetPrivileges%"" "
        S = S & N & "ECHO args = ""ELEV "" >> ""%vbsGetPrivileges%"" "
        S = S & N & "ECHO For Each strArg in WScript.Arguments >> ""%vbsGetPrivileges%"" "
        S = S & N & "ECHO args = args ^& strArg ^& "" ""  >> ""%vbsGetPrivileges%""  "
        S = S & N & "ECHO Next >> ""%vbsGetPrivileges%"" "
        S = S & N & "ECHO UAC.ShellExecute ""!batchPath!"", args, """", ""runas"", 1 >> ""%vbsGetPrivileges%"" "
        S = S & N & "%SystemRoot%\System32\WScript.exe ""%vbsGetPrivileges%"" %* "
        S = S & N & "exit /B "
        S = S & N & "  "
        S = S & N & ":gotPrivileges  "
        S = S & N & "setlocal & pushd .  "
        S = S & N & "cd /d %~dp0 "
        S = S & N & "if '%1'=='ELEV' (del ""%vbsGetPrivileges%"" 1>nul 2>nul  &  shift /1) "
        S = S & N & "  "
        S = S & N & "::::::::::::::::::::::::::::  "
        S = S & N & "::START "
        S = S & N & "::::::::::::::::::::::::::::  "
        S = S & N & "REM ECHO ON "
        S = S & N & "CLS "
        S = S & N & "REM Run shell as admin (example) - put here code as you like  "
        S = S & N & "REM ECHO %batchName% Arguments: %1 %2 %3 %4 %5 %6 %7 %8 %9  "
        S = S & N & "REM cmd /k  "
        S = S & N & "  "
        S = S & N & "for /f ""tokens=2*"" %%a in ('reg query ""HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System"" /v ConsentPromptBehaviorAdmin 2^>^&1^|find ""REG_""') do @set consentAdmin=%%b "
        S = S & N & "REM echo %consentAdmin% "
        S = S & N & "reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v ConsentPromptBehaviorAdmin /t REG_DWORD /d 0x0 /f "
        S = S & N & "  "
        S = S & N & "@ECHO OFF "
        S = S & N & "CLS "
        S = S & N & "ECHO. "
        S = S & N & "ECHO Connecting...  "
        S = S & N & "ECHO. "
        S = S & N & "IF %~d0 == \\ PushD %~d0%~p0  "
        S = S & N & "REM If winvnc4.exe is already running, then don't start it. "
        S = S & N & "TaskList /Fi ""IMAGENAME eq winvnc4.exe"" | Find ""winvnc4.exe"" /i "
        S = S & N & "CLS "
        S = S & N & "ECHO. "
        S = S & N & "ECHO Connecting...  "
        S = S & N & "ECHO. "
        S = S & N & "IF %ErrorLevel% == 1 (  "
        S = S & N & "  Start winvnc4.exe -noconsole"
        S = S & N & "  Sleep -m 500"
        S = S & N & ") "
        S = S & N & "winvnc4.exe -connect Office.SimplifiedPOS.com "
        S = S & N & "IF %~d0 == \\ ( "
        S = S & N & "  Sleep -m 500"
        S = S & N & "  PopD"
        S = S & N & ") "
        S = S & N & "  "
        S = S & N & "  "
        S = S & N & "  "
        S = S & N & "  "
        S = S & N & "CLS "
        S = S & N & "ECHO VNC server is running....  "
        S = S & N & "ECHO. "
        S = S & N & "ECHO Leave this window open while connected.  "
        S = S & N & "ECHO. "
        S = S & N & "ECHO --Only press a key when done-- "
        S = S & N & "ECHO. "
        S = S & N & "PAUSE "
        S = S & N & "  "
        S = S & N & "reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v ConsentPromptBehaviorAdmin /t REG_DWORD /d %consentAdmin% /f  "
        S = S & N & "  "
        S = S & N & "  "
        S = S & N & "  "

        ConnectCMDv2 = S
    End Function

End Module
