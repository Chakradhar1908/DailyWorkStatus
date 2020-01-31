Module modBenchmark
    Private LastLongBenchmark As Date, LastShrtBenchmark As Date
    Private BenchmarkHistory() As String, BenchmarkHistoryCount As Integer
    Private BenchmarkThreshold As Integer
    Private IsVistaAndLater As Boolean

    Public Function RecordWinCDSBenchmark() As Boolean
        Dim R As String, C As String, L As Object, PName As String, pId As String, Mem As String, K As String

        'If CLng(LastShrtBenchmark) = 0 Then LastShrtBenchmark = DateAdd("h", -3, Now)
        If IsNothing(LastShrtBenchmark) Then LastShrtBenchmark = DateAdd("h", -3, Now)

        'If CLng(LastLongBenchmark) = 0 Then LastLongBenchmark = DateAdd("h", -3, Now)
        If IsNothing(LastLongBenchmark) Then LastLongBenchmark = DateAdd("h", -3, Now)

        ' Every 5 minutes
        If DateBefore(Now, DateAdd("n", 5, LastShrtBenchmark), True, "n") Then Exit Function
        BenchmarkMemInfo(True)
        LastShrtBenchmark = Now

        ' Every 60 minutes
        If DateBefore(Now, DateAdd("n", 60, LastLongBenchmark), True, "n") Then Exit Function
        LastLongBenchmark = Now

        R = RunCmdToOutput("tasklist /FO CSV")
        R = Replace(R, vbLf, "")

        C = WinCDSEXEName(True, True)
        For Each L In Split(R, vbCr)
            ParseTasklist(L, PName, pId, , , Mem)
            If LCase(CSVField(L, 1)) = LCase(C) Then
                'If App.ThreadID = PID Then
                K = PName & " " & pId & " " & Mem & " " & L
                LogFile("Benchmark", K, False)
                'End If
            End If
        Next
    End Function

    Public Function BenchmarkMemInfo(Optional ByVal DoRecord As Boolean = False) As Integer
        Dim pId As Integer, R As String, L As Object, C As String, Mem As Integer

        pId = GetCurrentProcessId

        R = RunCmdToOutput("tasklist /FO CSV")
        R = Replace(R, vbLf, "")

        For Each L In Split(R, vbCr)
            C = L
            If Val(CSVField(C, 2)) = pId Then
                Mem = Val(Replace(CSVField(C, 5), ",", ""))
                If IsDevelopment() And Not IsIDE() Then
                    If Mem > NextThreshold Then
                        BenchmarkThreshold = NextThreshold(Mem)
                        'Robert 5/8/2017 commented this out.  Still not sure how it is getting here, because above it is asking 'IsDevelopment, but
                        'I swear I'm not in Development mode.  There may be more problems than just this one
                        '
                        'BFH20170612 - IsDevelopment REPLACED.  You did it wrong.
                        'JERRY-LAPTOP computer is ALWAYS in development mode unless SWITCHED OFF.
                        'If the computer is in DEVELOPMENT MODE, it should show a small DEV in the top right of the main menu.
                        'To shut it off, you need to go into the practice screen (double click on "WinCDS Pro") and double-click
                        'on the lower right "DEV MODE" label and select a new mode.
                        '
                        ' REGARDLESS -- DO NOT SHUT THIS CHECK OFF!!!
                        ' It is here to catch memory leaks
                        ' EITHER:  Leave this notice IN TACT and ignore messages, or raise the limit (which I hav done)!

                        MessageBox.Show("DEVELOPER NOTICE:" & vbCrLf & "Potentially Large Memory Usage" & vbCrLf & "BenchmarkMemInfo: " & DescribeFileSize(Mem * FileSize_1KB))
                    End If
                End If

                BenchmarkMemInfo = Mem
                If DoRecord Then
                    BenchmarkHistoryCount = BenchmarkHistoryCount + 1
                    'ReDim Preserve BenchmarkHistory(1 To BenchmarkHistoryCount)
                    ReDim Preserve BenchmarkHistory(0 To BenchmarkHistoryCount - 1)
                    BenchmarkHistory(BenchmarkHistoryCount - 1) = """" & Now & """" & "," & C
                End If
            End If
        Next
    End Function

    Private ReadOnly Property NextThreshold(Optional ByVal Current As Integer = 0) As Integer
        Get
            'BFH20170126 11 sets threshhold to 12 MB
            Const tMin As Integer = 14   ' will alert when memory is ((tMin + 1) * 10) MEGABYTES
            If Current <> 0 Then BenchmarkThreshold = (Trunc(Current / 10000, 0)) * 10000
            NextThreshold = BenchmarkThreshold / 10000
            If NextThreshold <= tMin Then NextThreshold = tMin
            NextThreshold = (NextThreshold + 1) * 10000
        End Get
    End Property

    Private Function ParseTasklist(ByVal CSV As String, Optional ByRef ProcName As String = "", Optional ByRef pId As String = "", Optional ByRef SessionName As String = "", Optional ByRef SessionNo As String = "", Optional ByRef MemUsage As String = "") As Boolean
        '"OSPPSVC.EXE","3816","Services","0","11,096 K"
        '"SystemSettingsBroker.exe","2976","RDP-Tcp#11","2","18,704 K"
        '"dllhost.exe","2096","Services","0","9,532 K"
        '"NetworkUXBroker.exe","9460","RDP-Tcp#11","2","17,688 K"
        '"VB6.EXE","896","RDP-Tcp#11","2","158,168 K"
        '"APILOAD.EXE","5648","RDP-Tcp#11","2","22,080 K"
        '"cmd.exe","8052","RDP-Tcp#11","2","2,848 K"
        '"conhost.exe","2760","RDP-Tcp#11","2","8,640 K"
        '"SearchProtocolHost.exe","7716","Services","0","7,940 K"
        '"SearchFilterHost.exe","948","Services","0","5,592 K"
        '"tasklist.exe","8376","RDP-Tcp#11","2","7,172 K"

        ProcName = CSVField(CSV, 1)
        pId = CSVField(CSV, 2)
        SessionName = CSVField(CSV, 3)
        SessionNo = CSVField(CSV, 4)
        MemUsage = CSVField(CSV, 5)
        ParseTasklist = True
    End Function

    Public Function BenchmarkMemoryProfile() As String
        Dim hSnap As Integer
        Dim Pe As PROCESSENTRY32
        Dim hProc As Integer
        Dim MI As PROCESS_MEMORY_COUNTERS
        Dim I As Integer
        'Dim Li As ListItem

        Dim pId As Integer, Res As String, N As String, M As String
        pId = GetCurrentProcessId

        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap = INVALID_HANDLE_VALUE Then Exit Function

        Pe.dwSize = Len(Pe)

        If Process32First(hSnap, Pe) Then
            Do
                hProc = OpenProcess(IIf(IsVistaAndLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), False, Pe.th32ProcessID)
                If hProc Then
                    If Pe.th32ProcessID = pId Then
                        MI.Cb = Len(MI)
                        GetProcessMemoryInfo(hProc, MI, Len(MI))

                        Res = ""
                        M = ""
                        N = vbCrLf
                        Res = Res & M & ""
                        Res = Res & M & "Process ID:       " & Pe.th32ProcessID
                        Res = Res & N & "Working Set Size: " & DescribeFileSize(MI.WorkingSetSize, 1)
                        Res = Res & N & "Pagefile Usage:   " & DescribeFileSize(MI.PagefileUsage, 1)
                        Res = Res & N & "Page fault count: " & MI.PageFaultCount
                        Res = Res & N & "Peak Page Usage:  " & DescribeFileSize(MI.PeakPagefileUsage, 1)
                        Res = Res & N & "Peak Working Set: " & DescribeFileSize(MI.PeakWorkingSetSize, 1)
                        Res = Res & N & "QuotaNonPagedPoolUsage:     " & DescribeFileSize(MI.QuotaNonPagedPoolUsage, 1)
                        Res = Res & N & "QuotaPagedPoolUsage:        " & DescribeFileSize(MI.QuotaPagedPoolUsage, 1)
                        Res = Res & N & "QuotaPeakNonPagedPoolUsage: " & DescribeFileSize(MI.QuotaPeakNonPagedPoolUsage, 1)
                        Res = Res & N & "QuotaPeakPagedPoolUsage:    " & DescribeFileSize(MI.QuotaPeakPagedPoolUsage, 1)
                    End If

                    CloseHandle(hProc)
                    I = I + 1
                End If
            Loop While Process32Next(hSnap, Pe)
        End If

        CloseHandle(hSnap)

        BenchmarkMemoryProfile = Res
    End Function
End Module
