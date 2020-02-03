Imports System
Imports System.Threading
Imports VBA
Module modDebugging
    Private mTerm As Integer
    Private mSection As String
    Public Function LogFile(ByVal File As String, ByVal Text As String, Optional ByVal NoDateStamp As Boolean = True, Optional ByVal PreventNL As Boolean = False) As Boolean
        Dim TDesc As String, P As String, Ext As String

        Const LogFileSizeReduceTo As Double = 0.3 ' Reduce to 30% when full

        On Error Resume Next
        If Not NoDateStamp Then
            TDesc = Format(Now, "yyyy-mm-dd hh:mm:ss")
            TDesc = TDesc & " [" & Thread.CurrentThread.ManagedThreadId & "]"
            Text = TDesc & "  " & Text
        End If

        If InStr(File, ":\") > 0 Then
            P = File
        Else
            P = LogFolder() & File
        End If
        Ext = Right(P, 4)
        If Left(Ext, 1) <> "." Then P = P & ".txt"

        If LogFileMaxSize(File) <= 0 Then             ' this can be used to cleanly disable a given log file
            If FileExists(P) Then Kill(P)
            Exit Function
        End If

        WriteFile(P, Text, False, PreventNL)

        If FileLen(P) > LogFileMaxSize(File) Then
            ' keeps log file size managable
            ' if log file is over threshold, read entire file, delete the first SHRINK characters, and re-write.
            ' Should not result in happening every time.  Should only happen "every so often".
            Dim Cont As String
            Cont = ReadEntireFile(P)
            Cont = Mid(P, LogFileMaxSize(File) * LogFileSizeReduceTo)
            WriteFile(P, Cont, True, True)
        End If
        LogFile = True
    End Function

    Public Function LogFolder(Optional ByVal SubFolder As String = "") As String
        ' NOTE: We purposefully do NOT use the standard functions here because this is for debugging.
        ' If we used InventFolder, it would call IsServer, and we wouldn't be able to test many things...
        ' By keeping diagnostic folders separate, we prevent recursion while debugging the deep sub-functions
        '
        ' *** CALL NO SUPPORT functions whatsoever in this function (See Note Above)
        ' *** Use Only native or atomic functions to derive this value.
        Dim A(), L, Cn As String
        Static PathCache As String

        If PathCache <> "" Then
            LogFolder = PathCache
            Exit Function
        End If

        A = {"I:\Invent", "C:\CDSData\Invent", "C:\CDSData"}
        On Error GoTo NextEntry
        For Each L In A
            If Left(L, 2) = "I:" Then
                Cn = UCase(GetLocalComputerName())
                Select Case Cn
                    Case "JERRY-LAPTOP", "INVENTORY" : GoTo NextEntry
                End Select
            End If
            LogFolder = L
            If Dir(LogFolder, vbDirectory) <> "" Then
                LogFolder = LogFolder & "\"
                PathCache = LogFolder
                Exit Function
            End If
        Next
NextEntry:
        LogFolder = AppFolder()
        PathCache = LogFolder
    End Function

    Private Function LogFileMaxSize(Optional ByVal vLogFile As String = "") As Integer
        ' Setting the max size to zero (0) will remove and prevent the log file.
        ' Default max file size is 1MB.
        ' We generally don't need to keep these that long
        Select Case LCase(vLogFile)
'    Case "update":        LogFileMaxSize = 0
'    Case "awslog":        LogFileMaxSize = 0
'    Case "voidsale":      LogFileMaxSize = 0
            Case "backup" : LogFileMaxSize = FileSize_1MB
            Case "shellandwait" : LogFileMaxSize = FileSize_1MB
            Case "activelog" : LogFileMaxSize = 20 * FileSize_1MB
            Case Else : LogFileMaxSize = FileSize_1MB
        End Select
    End Function

    Public Function IsIDE() As Boolean
        'IsIDE = False
        'Exit Function
        Dim a, b As Integer
        ' works on a very simple princicple... debug statements don't get compiled...
        On Error GoTo IDEInUse
        a = 1
        Debug.Print(a \ 0) 'division by zero error

        IsIDE = False
        Exit Function
IDEInUse:
        IsIDE = True
    End Function

    Public Function ErrMsg(ByVal Text As String, Optional ByVal Style As MsgBoxStyle = vbCritical, Optional ByVal AltTitle As String = "") As VbMsgBoxResult
        ErrMsg = ErrorMsg(Text, Style, AltTitle)
    End Function
    Public Function ErrorMsg(ByVal Text As String, Optional ByVal Style As MsgBoxStyle = MsgBoxStyle.DefaultButton1, Optional ByVal AltTitle As String = "") As VbMsgBoxResult
        Dim Num As Integer, Desc As String, M As String
        Num = Err.Number
        Desc = Err.Description

        AltTitle = IIf(AltTitle = "", ProgramErrorTitle, AltTitle)

        If Not IsFormLoaded("frmSplash") Then
            ReportError(AltTitle)
        End If

        M = ""
        M = M & "You've encountered an error in our software." & vbCrLf
        M = M & vbCrLf
        M = M & Text & vbCrLf
        If Num <> 0 Then M = M & "[" & Num & "]: " & Desc & vbCrLf
        M = M & vbCrLf
        M = M & "Please contact " & AdminContactString(2) & vbCrLf
        M = M & vbCrLf
        M = M & "ver=" & SoftwareVersion(False, True, False, True) & vbCrLf
        M = M & AdminContactString(3) & vbCrLf

        '  If Style And vbCritical <> 0 Then
        ''    text = text & vbCrLf2 & "Please contact " & admincontactstring(
        '  ElseIf Style And vbExclamation <> 0 Then
        '  ElseIf Style And vbQuestion <> 0 Then
        '  ElseIf Style And vbInformation <> 0 Then
        '  Else
        '  End If

        'MsgBox(M, Style, AltTitle)
        MessageBox.Show(M, AltTitle, MessageBoxButtons.OK)
    End Function

    Public Function DevErr(ByVal Text As String, Optional ByVal Style As VbMsgBoxStyle = MsgBoxStyle.DefaultButton1, Optional ByVal AltTitle As String = "Developer Error") As VbMsgBoxResult
        ' we purposefully do not provide a timeout on this function
        ' because the principle is that this is only used for an
        ' obvious and glaring error in code--something that should
        ' never happen. Hence, it is intentionally not 'user friendly'.
        ' See Also:  MsgBox, ErrMsg
        Text = "=== Developer Error ===" & vbCrLf & Text
        DevErr = ErrMsg(Text, Style, AltTitle)
    End Function

    Public Sub tPr(Optional ByVal vSection As String = "")
        mSection = vSection
        If mSection = "" Then mSection = "#"
        mTerm = 0
    End Sub

    Public Function Tp(Optional ByVal Msg As String = "", Optional ByVal Target As String = "") As Boolean
        If Target = "" Then Target = "debug"
        mTerm = mTerm + 1
        LogFile(Target, Trim("[" & mTerm & " - " & mSection & "] " & Msg))
        Tp = True
    End Function

    Public Function KillLog(ByVal S As String) As Boolean
        ' normally, the log function above keeps the logs trimmed...
        ' But, we provide an interface to clear it
        On Error Resume Next
        ' If they specified a full path, they can delete it themselves..
        If InStr(S, ":") <> 0 Then
            MsgBox "Unable to delete log:" & S
    Exit Function
        End If
        Kill LogFolder() & S
  Kill LogFolder() & S & ".txt"
  KillLog = True
    End Function

End Module
