Imports VBA
Imports Microsoft.VisualBasic.Interaction
Public Class FrmMsg2
    Private Style As VbMsgBoxStyle
    Private Const MASK_ICONS As Integer = &H70  ' 0001110000
    Private Const MASK_BUTTONS As Integer = &H7   ' 0000000111
    Private Const MASK_DEFAULTS As Integer = &H300 ' 1100000000
    Private FlashButton As Integer
    Private Result As VbMsgBoxResult
    Private Const MESSAGE_RTL_FONTNAME As String = "Courier New"
    Private Const MESSAGE_RTL_FONTSIZE As Integer = 8
    Private Const MESSAGE_FONTNAME As String = "Arial"
    Private Const MESSAGE_FONTSIZE As Integer = 8
    Private Const MSG_GAP_ICON As Integer = 180
    Private Const MSG_GAP As Integer = 120
    Private Const MIN_WIDTH As Integer = 2640 ' 5145
    Private Const MAX_WIDTH As Integer = 8715
    Private Const LABEL_GAP As Integer = 240
    Private Const LABEL_NORMAL_GAP As Integer = 135
    Private Const BUTTON_GAP As Integer = 105

    Public Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "", Optional ByVal HelpFile As String = "", Optional ByVal Context As Integer = 0, Optional ByVal MaxDisplay As Integer = 0, Optional ByVal nFlashButton As Integer = 0, Optional ByVal SecureConfirmation As String = "") As VbMsgBoxResult
        Dim HelpF As String, I As Integer, R As VbMsgBoxStyle

        Style = Buttons ' record it for the duration

        If MaxDisplay > 0 Then
            tmrMax.Tag = MaxDisplay
            tmrMax.Enabled = True
        Else
            On Error Resume Next
            tmrMax.Enabled = False
        End If

        MessageTitle = Title
        MessageIcon = Buttons
        lblMessage.Text = Prompt

        'HelpF = App.HelpFile  'remember current file

        'App.HelpFile = HelpFile
        'HelpContextID = Context

        For I = 1 To 4
            'ButtonText(I) = ""
            ButtonText(I, "")
        Next

        Select Case (Buttons And MASK_BUTTONS)
            'Case vbRetryCancel : ButtonText(1) = "&Retry" : ButtonText(2) = "&Cancel" : ButtonCancel = 2
            Case vbRetryCancel : ButtonText(1, "&Retry") : ButtonText(2, "&Cancel") : ButtonCancel = 2
            'Case vbYesNo : ButtonText(1) = "&Yes" : ButtonText(2) = "&No" : ButtonCancel = 2
            Case vbYesNo : ButtonText(1, "&Yes") : ButtonText(2, "&No") : ButtonCancel = 2
            'Case vbYesNoCancel : ButtonText(1) = "&Yes" : ButtonText(2) = "&No" : ButtonText(3) = "&Cancel" : ButtonCancel = 3
            Case vbYesNoCancel : ButtonText(1, "&Yes") : ButtonText(2, "&No") : ButtonText(3, "&Cancel") : ButtonCancel = 3
            'Case vbAbortRetryIgnore : ButtonText(1) = "&Abort" : ButtonText(2) = "&Retry" : ButtonText(3) = "&Ignore" : ButtonCancel = 3
            Case vbAbortRetryIgnore : ButtonText(1, "&Abort") : ButtonText(2, "&Retry") : ButtonText(3, "&Ignore") : ButtonCancel = 3
            'Case vbOKCancel : ButtonText(1) = "&OK" : ButtonText(2) = "&Cancel" : ButtonCancel = 2
            Case vbOKCancel : ButtonText(1, "&OK") : ButtonText(2, "&Cancel") : ButtonCancel = 2
                'Case Else : ButtonText(1) = "&OK" : ButtonCancel = 1 ' vbOKOnly
            Case Else : ButtonText(1, "&OK") : ButtonCancel = 1 ' vbOKOnly
        End Select
        If (Buttons And VbMsgBoxStyle.vbMsgBoxHelpButton) Then ButtonText(ButtonCount(Buttons), "&Help")

        R = (Buttons And MASK_DEFAULTS)
        ButtonDefault = Switch(R = vbDefaultButton2, 2, R = vbDefaultButton3, 3, R = VbMsgBoxStyle.vbDefaultButton4, 4, True, 1)

        'lblMessage.Alignment = IIf(Buttons And vbMsgBoxRight, vbRightJustify, vbLeftJustify)
        lblMessage.TextAlign = IIf(Buttons And vbMsgBoxRight, ContentAlignment.MiddleRight, ContentAlignment.MiddleLeft)
        TopMost = (Buttons And vbMsgBoxSetForeground)

        FlashButton = nFlashButton
        txtConfirm.Tag = SecureConfirmation

        Rearrange(Buttons)
        '  Dim F As Form
        '  Set F = MainMenu
        'Show vbModal ', F
        ShowDialog()

        'App.HelpFile = HelpF  ' Reset the application helpfile
        MsgBox = Result

        Style = vbDefaultButton1 ' reset the style after completion
        'Unload Me
        Me.Close()
    End Function

    Private WriteOnly Property MessageTitle() As String
        Set(value As String)
            Me.Text = IIf(value = "", ProgramName, value)
        End Set
    End Property

    Private WriteOnly Property MessageIcon() As MsgBoxStyle
        Set(value As MsgBoxStyle)
            Dim R As String
            value = MASK_ICONS And value
            R = Switch(value = vbCritical, "critical", value = vbQuestion, "question", value = vbExclamation, "exclamation", value = vbInformation, "information", True, "")
            If R <> "" Then
                'picIcon.Image = imlStyles.ListImages(R).Picture
                picIcon.Image = imlStyles.Images.Item(R)
                picIcon.Visible = True

            Else
                picIcon.Visible = False
            End If
        End Set
    End Property

    Private Function ButtonText(ByVal N As Integer, Optional ByVal S As String = "")
        Select Case N
            Case 1
                cmdButton1.Text = S
                cmdButton1.Visible = (S <> "")
            Case 2
                cmdButton2.Text = S
                cmdButton2.Visible = (S <> "")
            Case 3
                cmdButton3.Text = S
                cmdButton3.Visible = (S <> "")
            Case 4
                cmdButton4.Text = S
                cmdButton4.Visible = (S <> "")
        End Select
        Return Nothing
    End Function

    Private Property ButtonDefault() As Integer
        Get
            If AcceptButton Is cmdButton1 Then ButtonDefault = 0 : Exit Property
            If AcceptButton Is cmdButton2 Then ButtonDefault = 1 : Exit Property
            If AcceptButton Is cmdButton3 Then ButtonDefault = 2 : Exit Property
            If AcceptButton Is cmdButton4 Then ButtonDefault = 3 : Exit Property
            Return Nothing
        End Get

        Set(value As Integer)
            On Error Resume Next
            Select Case value
                Case 0
                    Me.AcceptButton = cmdButton1
                Case 1
                    Me.AcceptButton = cmdButton2
                Case 2
                    Me.AcceptButton = cmdButton3
                Case 3
                    Me.AcceptButton = cmdButton4
            End Select
        End Set
    End Property

    Private Property ButtonCancel() As Integer
        Get
            If CancelButton Is cmdButton1 Then ButtonCancel = 0 : Exit Property
            If CancelButton Is cmdButton2 Then ButtonCancel = 1 : Exit Property
            If CancelButton Is cmdButton3 Then ButtonCancel = 2 : Exit Property
            If CancelButton Is cmdButton4 Then ButtonCancel = 3 : Exit Property
            Return Nothing
        End Get
        Set(value As Integer)
            On Error Resume Next
            Select Case value
                Case 0
                    Me.CancelButton = cmdButton1
                Case 1
                    Me.CancelButton = cmdButton2
                Case 2
                    Me.CancelButton = cmdButton3
                Case 3
                    Me.CancelButton = cmdButton4
            End Select
        End Set
    End Property

    Private Function ButtonCount(ByVal Buttons As VbMsgBoxStyle) As Integer
        Select Case (Buttons And MASK_BUTTONS)
            Case vbRetryCancel : ButtonCount = 2
            Case vbYesNo : ButtonCount = 2
            Case vbYesNoCancel : ButtonCount = 3
            Case vbAbortRetryIgnore : ButtonCount = 3
            Case vbOKCancel : ButtonCount = 2
            Case Else : ButtonCount = 1
        End Select
        If Buttons And VbMsgBoxStyle.vbMsgBoxHelpButton Then ButtonCount = ButtonCount + 1
    End Function

    Public Sub Rearrange(ByVal Buttons As VbMsgBoxStyle)
        Dim lGap As Integer, lWidth As Integer, lBorder As Integer
        Dim lTextWidth As Integer, lBW As Integer
        Dim HasIcon As Boolean, BCount As Integer
        Dim ConfirmSpace As Integer, ConfirmWidth As Integer

        HasIcon = ((Buttons And MASK_ICONS) <> 0)
        BCount = ButtonCount(Buttons)

        If (Style And vbMsgBoxRtlReading) = vbMsgBoxRtlReading Then
            'FontName = MESSAGE_RTL_FONTNAME
            'FontSize = MESSAGE_RTL_FONTSIZE
            Me.Font = New Font(MESSAGE_RTL_FONTNAME, MESSAGE_RTL_FONTSIZE)
        Else
            'FontName = MESSAGE_FONTNAME
            'FontSize = MESSAGE_FONTSIZE
            Me.Font = New Font(MESSAGE_FONTNAME, MESSAGE_FONTSIZE)
        End If
        'FontName = lblMessage.FontName
        'FontSize = lblMessage.FontSize
        'lTextWidth = TextWidth(lblMessage)
        lTextWidth = CreateGraphics.MeasureString(lblMessage.Text, lblMessage.Font).Width
        'lBorder = ScaleX((GetSystemMetrics(SM_CXBORDER) + GetSystemMetrics(SM_CXDLGFRAME)), vbPixels, vbTwips)

        If HasIcon Then
            lWidth = lTextWidth + (2 * MSG_GAP_ICON) + (picIcon.Left + picIcon.Width)
        Else
            lWidth = lTextWidth + (2 * MSG_GAP) + lBorder
        End If

        Width = FitRange(MIN_WIDTH, lWidth, MAX_WIDTH)
        If BCount >= 3 And Width < 3940 Then Width = 3940
        If BCount >= 4 And Width < 5310 Then Width = 5310


        ' Size the label
        If HasIcon Then
            lblMessage.Left = picIcon.Left + picIcon.Width + MSG_GAP_ICON
            lblMessage.Width = Width - (2 * MSG_GAP_ICON) - (picIcon.Left + picIcon.Width)
        Else
            lblMessage.Left = MSG_GAP
            lblMessage.Width = Width - (2 * MSG_GAP)
        End If

        lGap = picButtons.Top - (lblMessage.Top + lblMessage.Height)

        If txtConfirm.Tag = "" Then
            ConfirmSpace = 0
            txtConfirm.Visible = False
        Else
            ConfirmSpace = txtConfirm.Height + 120
            txtConfirm.Text = ""
            'FontName = txtConfirm.FontName
            'FontSize = txtConfirm.FontSize
            Me.Font = New Font(txtConfirm.Font.Name, txtConfirm.Font.Size)
            'ConfirmWidth = TextWidth(txtConfirm.Tag & "X") + 60
            ConfirmWidth = CreateGraphics.MeasureString(txtConfirm.Tag & "X", txtConfirm.Font).Width
            ConfirmWidth = ConfirmWidth + 60
            If ConfirmWidth > 1500 Then ConfirmWidth = 1500
            'txtConfirm.Move(Width - ConfirmWidth) / 2, lblMessage.Top + lblMessage.Height + 60, ConfirmWidth
            txtConfirm.Location = New Point((Me.Width - ConfirmWidth) / 2, lblMessage.Top + lblMessage.Height + 60)
            txtConfirm.Size = New Size(ConfirmWidth, Height)
            txtConfirm.Visible = True
        End If

        If lGap < LABEL_NORMAL_GAP Then
            Height = Height + (LABEL_NORMAL_GAP - lGap) + ConfirmSpace
        Else
            Height = Height + ConfirmSpace
        End If

        'lBW = cmdButton(1).Width
        lBW = cmdButton1.Width
        Select Case BCount
            Case 1
                'cmdButton(1).Left = (Width - lBW - lBorder) \ 2
                cmdButton1.Left = (Width - lBW - lBorder) \ 2
            Case 2
                'cmdButton(1).Left = ((Width - BUTTON_GAP - lBorder) \ 2) - lBW
                'cmdButton(2).Left = BUTTON_GAP + lBW + cmdButton(1).Left
                cmdButton1.Left = ((Width - BUTTON_GAP - lBorder) \ 2) - lBW
                cmdButton2.Left = BUTTON_GAP + lBW + cmdButton1.Left

            Case 3
                'cmdButton(2).Left = (Width - lBW - lBorder) \ 2
                'cmdButton(1).Left = cmdButton(2).Left - lBW - BUTTON_GAP
                'cmdButton(3).Left = cmdButton(2).Left + lBW + BUTTON_GAP

                cmdButton2.Left = (Width - lBW - lBorder) \ 2
                cmdButton1.Left = cmdButton2.Left - lBW - BUTTON_GAP
                cmdButton3.Left = cmdButton2.Left + lBW + BUTTON_GAP
            Case 4
                'cmdButton(1).Left = ((Width - BUTTON_GAP - lBorder) \ 2) - lBW - lBW - BUTTON_GAP
                'cmdButton(2).Left = BUTTON_GAP + lBW + cmdButton(1).Left
                'cmdButton(3).Left = BUTTON_GAP + lBW + cmdButton(2).Left
                'cmdButton(4).Left = BUTTON_GAP + lBW + cmdButton(3).Left

                cmdButton1.Left = ((Width - BUTTON_GAP - lBorder) \ 2) - lBW - lBW - BUTTON_GAP
                cmdButton2.Left = BUTTON_GAP + lBW + cmdButton1.Left
                cmdButton3.Left = BUTTON_GAP + lBW + cmdButton2.Left
                cmdButton4.Left = BUTTON_GAP + lBW + cmdButton3.Left
        End Select
    End Sub

End Class