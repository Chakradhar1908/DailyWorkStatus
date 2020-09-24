Imports Microsoft.VisualBasic.Interaction
Module modTimeDuration
    Const DUR_MS_S As Integer = 1000
    Const DUR_S_M As Integer = 60
    Const DUR_M_H As Integer = 60
    Const DUR_H_D As Integer = 24
    Const DUR_D_W As Integer = 7
    Const DUR_D_T As Integer = 30
    Const DUR_D_Y As Integer = 365

    Public Enum TimeDurationStyles
        tdsty_Text = 0
        tdsty_Short = 1
        tdsty_Long = 2
        tdsty_Clock = 3
    End Enum

    Public Enum TimeSegments
        tseg_MS
        tseg_S
        tseg_MI
        tseg_H
        tseg_D
        tseg_W
        tseg_MO
        tseg_Y
    End Enum

    Public Function DescribeTimeDurationMS(ByVal Ms As Integer, Optional ByVal Style As TimeDurationStyles = TimeDurationStyles.tdsty_Text, Optional ByVal Resolution As TimeSegments = TimeSegments.tseg_Y) As String
        DescribeTimeDurationMS = DescribeTimeDuration(Ms / DUR_MS_S, Style, Resolution)
    End Function

    Public Function DescribeTimeDuration(ByVal tS As Double, Optional ByVal Style As TimeDurationStyles = TimeDurationStyles.tdsty_Text, Optional ByVal Resolution As TimeSegments = TimeSegments.tseg_Y) As String
        Dim Ms As Integer, S As Double, M As Integer, H As Integer, D As Integer
        Dim W As Integer, T As Integer, Y As Integer
        Dim Res As String

        Ms = TimeDurationSegment(tS, TimeSegments.tseg_MS, Resolution)
        S = TimeDurationSegment(tS, TimeSegments.tseg_S, Resolution)
        M = TimeDurationSegment(tS, TimeSegments.tseg_MI, Resolution)
        H = TimeDurationSegment(tS, TimeSegments.tseg_H, Resolution)
        D = TimeDurationSegment(tS, TimeSegments.tseg_D, Resolution)
        W = TimeDurationSegment(tS, TimeSegments.tseg_W, Resolution)
        T = TimeDurationSegment(tS, TimeSegments.tseg_MO, Resolution)
        Y = TimeDurationSegment(tS, TimeSegments.tseg_Y, Resolution)

        Res = ""
        If Style = TimeDurationStyles.tdsty_Clock Then
            If Y > 0 Then Res = Res & Y & DescribeTimeSegment(TimeSegments.tseg_Y, TimeDurationStyles.tdsty_Short) & " "
            If D > 0 Then Res = Res & D & DescribeTimeSegment(TimeSegments.tseg_D, TimeDurationStyles.tdsty_Short) & " "
            Res = Res & Format(H, "00") & ":" & Format(M, "00") & ":" & Format(S, "00")
            DescribeTimeDuration = Res
            Exit Function
        End If

        If S <> 0 Then
            If Y > 0 Or M > 0 Or D > 0 Or H > 0 Or Ms = 0 Then
                Res = Trunc0(S) & " " & DescribeTimeSegment(TimeSegments.tseg_S, Style) & IIf(Res = "", "", ", ") & Res
            ElseIf Y = 0 And M = 0 And D = 0 And H = 0 And M = 0 And S <= 5 And Resolution <= TimeSegments.tseg_S Then
                Res = Format(S * 1000 + Ms) & " " & DescribeTimeSegment(TimeSegments.tseg_MS, Style) & IIf(Res = "", "", ", ") & Res
            Else
                Res = Format(S, "0.000") & " " & DescribeTimeSegment(TimeSegments.tseg_S, Style) & IIf(Res = "", "", ", ") & Res
            End If
        End If
        If M <> 0 Then Res = M & " " & DescribeTimeSegment(TimeSegments.tseg_MI, Style) & IIf(Res = "", "", ", ") & Res
        If H <> 0 Then Res = H & " " & DescribeTimeSegment(TimeSegments.tseg_H, Style) & IIf(Res = "", "", ", ") & Res
        If D <> 0 Then Res = D & " " & DescribeTimeSegment(TimeSegments.tseg_D, Style) & IIf(Res = "", "", ", ") & Res

        ' Only show "weeks" if not showing months or years
        If W <> 0 And T = 0 And Y = 0 Then
            Res = D & " " & DescribeTimeSegment(TimeSegments.tseg_W, Style) & IIf(Res = "", "", ", ") & Res
        Else
            If T <> 0 Then Res = D & " " & DescribeTimeSegment(TimeSegments.tseg_MO, Style) & IIf(Res = "", "", ", ") & Res
            If Y <> 0 Then Res = D & " " & DescribeTimeSegment(TimeSegments.tseg_Y, Style) & IIf(Res = "", "", ", ") & Res
        End If

        If Style = TimeDurationStyles.tdsty_Short Then Res = Replace(Res, ",", "")
        If Style = TimeDurationStyles.tdsty_Short Then Res = Replace(Res, " ", "")

        DescribeTimeDuration = Trim(Res)
    End Function

    Public Function TimeDurationSegment(ByVal tS As Double, Optional ByRef Segment As TimeSegments = 0, Optional ByVal Resolution As TimeSegments = TimeSegments.tseg_Y) As Integer
        Dim Ms As Double, S As Double, M As Integer, H As Integer, D As Integer
        Dim W As Integer, T As Integer, Y As Integer
        Dim Res As String

        Ms = Decimals(tS) * 1000
        If Resolution <= TimeSegments.tseg_MS Then
            Ms = Trunc0(tS * DUR_MS_S)
            GoTo Render
        End If

        S = tS
        If Resolution <= TimeSegments.tseg_S Then GoTo Render

        M = Trunc0(tS / DUR_S_M)
        S = tS - M * DUR_S_M
        If Resolution <= TimeSegments.tseg_MI Then GoTo Render

        H = Trunc0(M / DUR_M_H)
        M = M - H * DUR_M_H
        If Resolution <= TimeSegments.tseg_H Then GoTo Render

        D = Trunc0(H / DUR_H_D)
        H = H - D * DUR_H_D
        If Resolution <= TimeSegments.tseg_D Then GoTo Render

        W = Trunc0(D / DUR_D_W)
        If Resolution <= TimeSegments.tseg_W Then GoTo Render
        T = Trunc0(D / DUR_D_T)
        If Resolution <= TimeSegments.tseg_MO Then GoTo Render
        Y = Trunc0(D / DUR_D_Y)
        If Resolution <= TimeSegments.tseg_Y Then GoTo Render

Render:
        Select Case Segment
            Case TimeSegments.tseg_MS : TimeDurationSegment = Ms
            Case TimeSegments.tseg_S : TimeDurationSegment = S
            Case TimeSegments.tseg_MI : TimeDurationSegment = M
            Case TimeSegments.tseg_H : TimeDurationSegment = H
            Case TimeSegments.tseg_D : TimeDurationSegment = D
            Case TimeSegments.tseg_W : TimeDurationSegment = W
            Case TimeSegments.tseg_MO : TimeDurationSegment = T
            Case TimeSegments.tseg_Y : TimeDurationSegment = Y
        End Select
    End Function

    Public Function DescribeTimeSegment(ByVal Seg As TimeSegments, Optional ByVal Style As TimeDurationStyles = TimeDurationStyles.tdsty_Text) As String
        Select Case Seg
            Case TimeSegments.tseg_MS : DescribeTimeSegment = Switch(Style = TimeDurationStyles.tdsty_Long, "millisecond", Style = TimeDurationStyles.tdsty_Short, "ms", True, "")
            Case TimeSegments.tseg_S : DescribeTimeSegment = Switch(Style = TimeDurationStyles.tdsty_Long, "second", Style = TimeDurationStyles.tdsty_Short, "s", True, "sec")
            Case TimeSegments.tseg_MI : DescribeTimeSegment = Switch(Style = TimeDurationStyles.tdsty_Long, "minute", Style = TimeDurationStyles.tdsty_Short, "m", True, "min")
            Case TimeSegments.tseg_H : DescribeTimeSegment = Switch(Style = TimeDurationStyles.tdsty_Long, "hour", Style = TimeDurationStyles.tdsty_Short, "h", True, "hr")
            Case TimeSegments.tseg_D : DescribeTimeSegment = Switch(Style = TimeDurationStyles.tdsty_Long, "day", Style = TimeDurationStyles.tdsty_Short, "d", True, "day")
            Case TimeSegments.tseg_W : DescribeTimeSegment = Switch(Style = TimeDurationStyles.tdsty_Long, "week", Style = TimeDurationStyles.tdsty_Short, "w", True, "wk")
            Case TimeSegments.tseg_MO : DescribeTimeSegment = Switch(Style = TimeDurationStyles.tdsty_Long, "month", Style = TimeDurationStyles.tdsty_Short, "mo", True, "mo")
            Case TimeSegments.tseg_Y : DescribeTimeSegment = Switch(Style = TimeDurationStyles.tdsty_Long, "year", Style = TimeDurationStyles.tdsty_Short, "y", True, "yr")
        End Select
    End Function

    Public Function TimeRemaining(Optional ByVal ItemsRemaining As Integer = -1, Optional ByVal AverageOver As Integer = 25) As String
        Static BeginningTick As Double, LastTick As Double, LastCount As Integer, Avg As Double, LastReturn As Double, LastCheck As Integer, LastResult As String
        'Debug.Print "TimeRemaining(IR=" & ItemsRemaining & ", AO=" & AverageOver & ")"

        Dim ItemsElapsed As Integer
        Dim Tick As Integer, Dur As Integer, RemainingTicks As Double
        Dim UseCurr As Boolean
        Tick = GetTickCount


        If ItemsRemaining <= 0 Then
            BeginningTick = Tick
            LastTick = Tick
            LastCount = 0
            LastReturn = 0
            LastCheck = 0
            Avg = 0
            Exit Function
        End If

        ItemsElapsed = LastCount - ItemsRemaining
        Dur = Tick - LastTick
        Avg = (Dur + (AverageOver - 1) * IIf(Avg = 0, Dur, Avg)) / AverageOver
        RemainingTicks = ItemsRemaining * Avg / IIf(ItemsElapsed = 0, 1, ItemsElapsed)

        If Dur < 50 Then TimeRemaining = LastResult : Exit Function ' Just a redraw
        'Debug.Print "Items=" & ItemsRemaining & ", Tick=" & Tick & ", Last = " & LastTick & ", Dur = " & Dur & ", Avg = " & FormatQuantity(Avg) & ", Remaining=" & RemainingTicks

        ' This is an anti-flicker approach..  Don't show constant up/down..  Only go back up if a certain period has elapsed.
        If LastReturn = 0 Or RemainingTicks < LastReturn Or Tick - LastCheck > 2500 Then
            LastReturn = RemainingTicks
            LastCheck = Tick
            UseCurr = True
        End If

        LastTick = Tick
        LastCount = ItemsRemaining

        On Error Resume Next ' Sometimes, a few of these were overflowing...  This prevents errors
        TimeRemaining = DescribeTimeDurationMS(IIf(UseCurr, RemainingTicks, LastReturn) + 1000, TimeDurationStyles.tdsty_Clock) '+ 1000 because you NEVER show 0s...  0s means done, not 1s remaining.
        LastResult = TimeRemaining
    End Function
End Module
