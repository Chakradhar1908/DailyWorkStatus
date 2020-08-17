Public Class frmCCAd
    Private WithEvents M As frmMenu

    Public Sub Advertize()
        Exit Sub                              ' ################ DISABLED

        ' NOTE: Below code is not required. Because exit sub is used in the first line.   <----------

        'Const AdvStart = #6/9/2008#
        'Dim Freq as integer
        'Dim X as integer, Y as integer

        ''  If Not IsDevelopment Then Unload Me: Exit Sub
        ''  If IsDevelopment Then Unload Me: Exit Sub
        'If SwipeCards() Then Me.Close() : Exit Sub

        ''If DateBetween(Of Date, AdvStart, DateAdd("d", 7, AdvStart))() Or
        'If DateBetween(Date.Today, AdvStart, DateAdd("d", 7, AdvStart)) Or DateBetween(Date.Today, DateAdd("d", 21, AdvStart), DateAdd("d", 28, AdvStart)) Then
        '    'DateBetween(Of Date, DateAdd("d", 21, AdvStart), DateAdd("d", 28, AdvStart))() Then
        '    Freq = -DateDiff("d", Today, AdvStart)
        '    If Freq >= 8 Then Freq = Freq - 21
        '    X = Random(10)
        '    Y = Random(4)
        '    If IsDevelopmentMANUAL() Then X = 10 : Y = 1
        '    If X > Freq And Y = 1 Then
        '        Show()
        '        tmr.Enabled = False
        '        tmr.Interval = 25000
        '        tmr.Enabled = True
        '        Exit Sub
        '    End If
        'End If

        ''Unload Me
        'Me.Close()
    End Sub

End Class