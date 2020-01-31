Module modTrackUsage
    Private TakingTooLong As Boolean
    Public Function TrackUsage(ByVal Title As String, Optional ByVal Desc As String = "") As Boolean
        Dim V As String
        V = Title & IIf(Desc = "", "", " - " & Desc)
        SetConfigLastRun(Title)
        PostUsage(Title, Desc)

        LogFile("Usage", Title, False)
        TrackUsage = True
    End Function

    Private Function SetConfigLastRun(ByVal vName As String, Optional ByVal nValue As String = "") As Boolean
        If vName = "" Then Exit Function
        If nValue = "" Then nValue = Today
        SetConfigLastRun = SetConfigTableValue(FeatureFieldName(vName), nValue)
    End Function

    Public Function PostUsage(ByVal vName As String, Optional ByVal Desc As String = "") As Boolean
        Const Host As String = CompanyURL_BARE
        Const Port As String = "80"
        Const vURL As String = "usage/usage.php"
        Const Q As String = "?"
        Const A As String = "&"

        Const BenchmarkThreshold As Integer = 1250

        If TakingTooLong Then Exit Function  ' If their internet is broken or something...  Will persist until program restart.

        On Error GoTo NetFail

        Dim Benchmark As Integer, Duration As Integer
        Benchmark = GetTickCount

        Dim R As String, Z As Boolean
        R = ""
        R = R & CompanyURL
        R = R & vURL
        R = R & Q & "post=1"
        R = R & A & "store=" & URLEncode(StoreSettings(1).Name, False)
        R = R & A & "item=" & URLEncode(vName, True)
        R = R & A & "desc=" & URLEncode(Desc, True)


        'Debug.Print R
        ' Truth is, we just want the request to SEND.  We care 0% about the result...
        ' If even just 50% of these go through, that is enough.. or 20%... We just want a usage slice of SOME kind..
        INETTimeout = 1
        Z = INETRequestOnly(R)

        Duration = GetTickCount - Benchmark
        If Duration > BenchmarkThreshold Then TakingTooLong = True
        '  Debug.Print "Dur=" & Duration

        PostUsage = True
        Exit Function

NetFail:
        TakingTooLong = True      ' Stop using if they have a network failure
        Exit Function
    End Function

    Private Function FeatureFieldName(ByVal vName As String) As String
        Const MRU_TAG As String = "LastRun_"
        FeatureFieldName = MRU_TAG & LCase(vName)
    End Function

End Module
