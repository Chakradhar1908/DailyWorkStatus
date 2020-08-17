Public Class frmSendMail
    Private Const EMAIL_TIMEOUT as integer = 15000

    'Public WithEvents objSendMail As vbSendMail.clsSendMail
    Public mStatus As String, ErrorMsg As String
    Private Expiration As Date, Sending As Boolean

    Public Function DoSimpleSendMail(ByVal From As String, ByVal FromName As String, ByVal T As String, ByVal TName As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachments As String = "") As String
        Dim Host As String
        DoLog("DoSimpleSendMail(" & From & ", " & FromName & ", " & T & ", " & TName & ", " & Subject & ",...)")

        'On Error GoTo NoSendMail
        ErrorMsg = "Error on CreateObject (clsSendMail)"
        'If objSendMail Is Nothing Then
        '    objSendMail = CreateObject("vbSendMail.clsSendMail")
        'End If
        'If objSendMail Is Nothing Then DoSimpleSendMail = "Could not create SendMail object: No Reason Specified." : Exit Function
        'ErrorMsg = ""

        On Error GoTo 0
        'With objSendMail
        'Debug.Print "Host=" & Host
        'If Val(GetEmailSetting("SOutlook")) = 1 And GetEmailSetting("smtpHost") <> "" And GetEmailSetting("smtpPort") <> "" Then
        '        .SMTPHost = GetEmailSetting("smtpHost")
        '        .SMTPPort = GetEmailSetting("smtpPort")
        '        If GetEmailSetting("smtpUser") <> "" And GetEmailSetting("smtpPass") <> "" Then
        '            .UserName = GetEmailSetting("smtpUser")
        '            .Password = GetEmailSetting("smtpPass")
        '            .UseAuthentication = True
        '        Else
        '            .UseAuthentication = False
        '        End If
        '    Else
        '        Host = .MXQuery(DestString(T))
        '        If Host = "" Then
        '            DoSimpleSendMail = "Could not resolve host MX record: " & DestString(T)
        '            Exit Function
        '        End If
        '        .SMTPHost = Host
        '    End If
        '    .From = From
        '    .FromDisplayName = FromName
        '    .Recipient = T
        '    .RecipientDisplayName = TName
        '    If CC <> "" Then .CcRecipient = CC
        '    If BCC <> "" Then .BccRecipient = BCC
        '    .ReplyToAddress = From
        '    .Subject = Subject
        '    .Message = Body & vbCrLf & "."
        '    .AsHTML = True
        '    .EncodeType = MIME_ENCODE
        '    .ContentBase = ""
        '    .Priority = NORMAL_PRIORITY
        '    .UsePopAuthentication = False
        '        Dim X As String, Y As String
        '            replaceHTMLimages(Body, , X, Y)
        '            Attachments = Attachments & IIf(Len(Attachments) > 0, ";", "") & Y
        '            If Attachments <> "" Then .Attachment = Attachments

        '            ErrorMsg = ""
        '            Sending = True
        '            Expiration = DateAdd("s", 5 + (EMAIL_TIMEOUT / 1000), Now) ' +5 seconds from expiration date
        '            ResetTimeout()
        '            ProgressForm 0, 1, "Sending email...", , , , prgIndefinite
        '    .Send
        '        End With

        '        Do Until Not Sending
        '            If DateAfter(Now, Expiration, True, "s") Then
        '                ErrorMsg = "Extra Timeout"
        '                Exit Do
        '            End If
        '            Application.DoEvents()     ' Highly non-glamorous synchronous use of an asych component...
        '        Loop
        '        ProgressForm()

        'NoSendMail:
        '        objSendMail = Nothing
        '        DoSimpleSendMail = ErrorMsg()

        '        'Unload Me
        '        Me.Close()
    End Function
    Private Sub DoLog(ByVal Msg As String, Optional ByVal Priority as integer = 7)
        Dim T As String
        T = IIf(Microsoft.VisualBasic.Left(Msg, 3) = "...", Msg, mStatus & "::" & Msg)
        LogFile("vbsendmail", T, False)
        ActiveLog("frmSendMail: t, priority")
        Debug.Print("::" & T)
    End Sub
    Public Function DestString(ByVal TAddr As String) As String
        Dim X as integer
        X = InStr(TAddr, "@")
        If X <= 0 Then Exit Function
        DestString = Mid(TAddr, X + 1)
    End Function

    Private Sub ResetTimeout(Optional ByVal StopTimer As Boolean = False)
        tmr.Enabled = False
        If StopTimer Then Exit Sub
        'tmr.Interval = EMAIL_TIMEOUT
        'tmr.Enabled = True
    End Sub

End Class