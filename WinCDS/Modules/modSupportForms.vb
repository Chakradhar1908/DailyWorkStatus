Imports VBA
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Module modSupportForms
    Private PR As frmProgress
    Private PR2 As frmProgress2
    Private PR3 As frmProgress3
    Private PS As frmProgressStatic

    Private SuppressMessagesUntil As Date
    Public Enum ProgressBarStyle
        prgDefault = 0
        prg3DFloat = 1
        prgFlatFloat = 2
        prgStatic = 3
        prgSpin = 4
        prgIndefinite = 5
    End Enum

    Public Function SelectOptionArray(ByRef selTitle As String, ByRef selType As frmSelectOption.ESelOpts, ByRef selOptions() As Object, Optional ByVal SelectButtonCaption As String = "&Print", Optional ByVal PreSelChk As String = "x") 'as integer
        '::::SelectOptionArray
        ':::SUMMARY
        ': Raise a select option dialog
        ':::DESCRIPTION
        ': Present a user with a list of options and require them to select one to continue (modal/blocking).
        ':::PARAMETERS
        ':::RETURN
        SelectOptionArray = frmSelectOption.SelectOptionArray(selTitle, selType, selOptions, SelectButtonCaption, PreSelChk)
    End Function
    'Public Function ProgressForm(Optional ByVal Value as integer = -1, Optional ByVal Max as integer = -1, Optional ByVal Caption As String = "#", Optional ByVal vButtons As VbMsgBoxStyle = 0, Optional ByVal BarColor as integer = vbInactiveBorder, Optional ByVal BackColor as integer = vbInactiveBorder, Optional ByVal Style As ProgressBarStyle = ProgressBarStyle.prgDefault, Optional ByVal Lt as integer = 0, Optional ByVal Tp as integer = 0) As VbMsgBoxResult
    '    '::::ProgressForm
    '    ':::SUMMARY
    '    ': Raise, Control, and Remove a progress form
    '    ':::DESCRIPTIONS
    '    ':: Raise Progress Form
    '    ': - ProgressForm 0, <MAX>, <Caption>
    '    ':: Update Progress Form
    '    ': - ProgressForm <Current Value>
    '    ':: Remove Progress Form
    '    ': - ProgressForm
    '    ':::RETURN
    '    ': - vbMsgBoxResult

    '    On Error Resume Next
    '    Dim s As String = SuppressMessagesUntil
    '    'If CLng(SuppressMessagesUntil) <> 0 Then
    '    If CLng(s) <> 0 Then
    '        If Not DateAfter(Now, SuppressMessagesUntil, , "n") Then
    '            Exit Function
    '        Else
    '            SuppressMessagesUntil = "0"
    '        End If
    '    End If


    '    If Value = -1 Or Max = 0 Then
    '        ClearOtherProgressForms()

    '        Exit Function
    '    End If

    '---> NOTE: COMMENTED THE BELOW CODE TO FIND OUT "ucPBar" control used in the frmProgress form.  <----

    '      If BarColor <> VBRUN.SystemColorConstants.vbInactiveBorder Then PR.prg.FillColor = BarColor
    '      If BackColor <> vbInactiveBorder Then PR.prg.BackColor = BackColor

    '      If IsWin5 And Not IsIDE() And IsIn(Style, prgSpin, prgIndefinite) Then Style = prgStatic ' WinxP is too slow to do this
    '      If Not Gif89Installed And Not IsIn(Style, prgStatic, prg3DFloat, prgDefault) Then Style = prgDefault
    '      If Style = prgDefault Then Style = IIf(Value = 0 And Max = 1, prgStatic, prg3DFloat)
    '      ClearOtherProgressForms Style

    'Select Case Style
    '          Case prg3DFloat
    '              If PR Is Nothing Then Set PR = New frmProgress
    '    PR.Progress Value, Max, Caption, True, True ', vButtons
    '    'If Lt > 0 Then
    '    '  PR.Left = Lt + 200
    '    '  PR.Top = Tp + 100
    '    'End If
    '          Case prgFlatFloat
    '              If PR2 Is Nothing Then Set PR2 = New frmProgress2
    '    PR2.Progress Value, Max, Caption, True, True, vButtons
    '  Case prgIndefinite
    '              If PR3 Is Nothing Then Set PR3 = New frmProgress3
    '    PR3.Progress Caption, True, True, vButtons
    '  Case prgSpin
    '              If PS Is Nothing Then Set PS = New frmProgressStatic
    '    PS.ProgressSpin Caption, True, True
    '  Case prgStatic
    '              If PS Is Nothing Then Set PS = New frmProgressStatic
    '    PS.Progress Caption, True, True
    'End Select

    '      If Value = 0 And Max = 1 Then
    '      Else
    '      End If
    'End Function
    Private Function ClearOtherProgressForms(Optional ByVal Style As ProgressBarStyle = ProgressBarStyle.prgDefault) As Boolean
        If Style <> ProgressBarStyle.prg3DFloat Then DisposeDA(PR)
        If Style <> ProgressBarStyle.prgFlatFloat Then DisposeDA(PR2)
        If Style <> ProgressBarStyle.prgIndefinite Then DisposeDA(PR3)
        If Style <> ProgressBarStyle.prgSpin And Style <> ProgressBarStyle.prgStatic Then DisposeDA(PS)
    End Function

    Public Function PreviewItemByStyle(Optional ByVal Style As String = "", Optional ByRef frm As Form = Nothing) As Boolean
        '::::PreviewItemByStyle
        ':::SUMMARY
        ': Used to preview the Item By using Style.
        ':::DESCRIPTION
        ': Technically, this may fail...
        ': If the calling form is displayed modally, and we try to show a form in the background that isn't modal,VB6 usually errors.  We simply fail silently to maintain program integrity.
        ': It should be up to the developer to make sure this doesn't happen
        ':::PARAMETERS
        ': - Style
        ': - Frm
        ':::RETURN
        ': Boolean - Returns the result whether it is True or False.

        '  If Not IsDevelopment Then Exit Function
        On Error Resume Next
        If Style = "" Then
            'Unload frmItemView
            frmItemView.Close()
            Exit Function
        End If

        ' Technically, this may fail...
        ' If the calling form is displayed modally, and we try to show a form in the background that isn't modal,
        ' VB6 usually errors.  We simply fail silently to maintain program integrity.
        ' It should be up to the developer to make sure this doesn't happen.
        PreviewItemByStyle = frmItemView.PreviewStyle(Style, frm)
        If Not PreviewItemByStyle Then
            'Unload frmItemView
            frmItemView.Close()
        End If
    End Function

    Public Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "", Optional ByVal HelpFile As String = "", Optional ByVal Context As Integer = 0, Optional ByVal MaxDisplay As Integer = 0, Optional ByVal FlashButton As Long = 0, Optional ByVal SecureConfirmation As String = "", Optional ByVal DimBackground As Boolean = True) As VbMsgBoxResult
        '::::MsgBox
        ':::SUMMARY
        ': Overrides native VB6 MsgBox
        ':::DESCRIPTION
        ': This function is used to define the properities to Message Box using parameters
        ':::PARAMETERS
        ':::RETURN
        ': - vbMsgBoxResult

        Dim R As New frmMsg2
        'MsgBox = VBA.MsgBox(Prompt, Buttons, Title, Helpfile, Context)   ' good for comparing original with ours in debugging
        On Error Resume Next
        If SuppressMessagesUntil <> "" Then
            If Not DateAfter(Now, SuppressMessagesUntil, , "n") Then
                LogFile("Suppressed.txt", "(v" & WinCDSBuildNumber() & ") [" & Title & "]: " & Replace(Replace(Prompt, vbCr, "/"), vbLf, ""), False)
                'Debug.Print "Suppressed MsgBox: " & Prompt
                ActiveLog("MsgBox::SUPPRESSED: [" & Title & "]: " & Prompt, 7)
                MsgBox = vbOK
                Exit Function
            Else
                SuppressMessagesUntil = ""
            End If
        End If

        HideSplash()
        ProgressForm ' Not usual to show msg box during progress form

        If DimBackground Then DimAllForms
        MsgBox = R.MsgBox(Prompt, Buttons, Title, HelpFile, Context, MaxDisplay, FlashButton, SecureConfirmation)
        ActiveLog("MsgBox::MsgBox: [" & Title & "]: " & Prompt, 8)
        If DimBackground Then UnDimAllForms
        R = Nothing
    End Function

    Public Function ProgressForm(Optional ByVal Value As Long = -1, Optional ByVal Max As Long = -1, Optional ByVal Caption As String = "#", Optional ByVal vButtons As VbMsgBoxStyle = 0, Optional ByVal BarColor As Long = vbInactiveBorder, Optional ByVal BackColor As Long = vbInactiveBorder, Optional ByVal Style As ProgressBarStyle = ProgressBarStyle.prgDefault, Optional ByVal Lt As Long = 0, Optional ByVal Tp As Long = 0) As VbMsgBoxResult
        '::::ProgressForm
        ':::SUMMARY
        ': Raise, Control, and Remove a progress form
        ':::DESCRIPTIONS
        ':: Raise Progress Form
        ': - ProgressForm 0, <MAX>, <Caption>
        ':: Update Progress Form
        ': - ProgressForm <Current Value>
        ':: Remove Progress Form
        ': - ProgressForm
        ':::RETURN
        ': - vbMsgBoxResult

        '      On Error Resume Next
        '      If SuppressMessagesUntil <> "" Then
        '          If Not DateAfter(Now, SuppressMessagesUntil, , "n") Then
        '              Exit Function
        '          Else
        '              SuppressMessagesUntil = ""
        '          End If
        '      End If


        '      If Value = -1 Or Max = 0 Then
        '          ClearOtherProgressForms()

        '          Exit Function
        '      End If


        '      If BarColor <> vbInactiveBorder Then PR.prg.FillColor = BarColor
        '      If BackColor <> vbInactiveBorder Then PR.prg.BackColor = BackColor

        '      If IsWin5 And Not IsIDE() And IsIn(Style, prgSpin, prgIndefinite) Then Style = prgStatic ' WinxP is too slow to do this
        '      If Not Gif89Installed And Not IsIn(Style, prgStatic, prg3DFloat, prgDefault) Then Style = prgDefault
        '      If Style = prgDefault Then Style = IIf(Value = 0 And Max = 1, prgStatic, prg3DFloat)
        '      ClearOtherProgressForms Style

        'Select Case Style
        '          Case prg3DFloat
        '              If PR Is Nothing Then Set PR = New frmProgress
        '    PR.Progress Value, Max, Caption, True, True ', vButtons
        '    'If Lt > 0 Then
        '    '  PR.Left = Lt + 200
        '    '  PR.Top = Tp + 100
        '    'End If
        '          Case prgFlatFloat
        '              If PR2 Is Nothing Then Set PR2 = New frmProgress2
        '    PR2.Progress Value, Max, Caption, True, True, vButtons
        '  Case prgIndefinite
        '              If PR3 Is Nothing Then Set PR3 = New frmProgress3
        '    PR3.Progress Caption, True, True, vButtons
        '  Case prgSpin
        '              If PS Is Nothing Then Set PS = New frmProgressStatic
        '    PS.ProgressSpin Caption, True, True
        '  Case prgStatic
        '              If PS Is Nothing Then Set PS = New frmProgressStatic
        '    PS.Progress Caption, True, True
        'End Select

        '      If Value = 0 And Max = 1 Then
        '      Else
        '      End If
    End Function


End Module
