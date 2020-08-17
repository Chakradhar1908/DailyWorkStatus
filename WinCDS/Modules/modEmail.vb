Imports VBA
Module modEmail
    Private Const CHILKAT_LICENSE As String = "KATZCA.EM11217_NHCieJmcFO92"
    ' Chilkat Docs: https://www.chilkatsoft.com/refdoc/vbnetMailManRef.html
    ' Sample: https://www.example-code.com/vb6/smtp_simpleSend.asp

    Private Const PR_ATTACH_MIME_TAG As String = "http://schemas.microsoft.com/mapi/proptag/0x370E001E"
    Private Const PR_ATTACH_CONTENT_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
    Private Const PR_ATTACHMENT_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
    Private Const olMailItem As Integer = 0

    Public Structure EmailResult
        Dim PoNo As Integer
        Dim VendorAddress As String
        Dim VendorName As String
        Dim SendTime As String
        Dim Success As Boolean
    End Structure

    Public Function SendSimpleEmail(ByVal From As String, ByVal FromName As String, ByVal T As String, ByVal TName As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachments As String = "") As String
        Dim Host As String

        ' First, we allow DEV MODE to skip all emails
        If IsDevelopment() Then
            If MsgBox("DEVELOPER: Skip email send?", vbYesNo, "DEVELOPER SKIP") = vbYes Then
                Dim TF As String
                SendSimpleEmail = "Email Send SKIPPED by developer"
                TF = TempFile(DevOutputFolder, Slug(Subject, 15) & "-", ".htm")
                WriteFile(TF, Body)

                LogFile("email", "SendSimpleEmail: """ & FromName & " <" & From & ">"": " & Subject)
                MsgBox("Email body written to: " & vbCrLf & TF, vbInformation, "DEVELOPER INFORMATION")
                Exit Function
            End If
        End If

        ' If they are using any email client, run this.
        If Val(GetEmailSetting("SOutlook")) = 0 Then
            SendSimpleEmail = SendThroughOutlook(From, FromName, T, TName, Subject, Body, CC, BCC, Attachments)
            Exit Function
        End If

        ' If they have the new DLL, use it.
        If CBool(GetEmailSetting("uchilkat")) Then
            SendSimpleEmail = SendMailSMTP_Chilkat(From, FromName, T, TName, Subject, Body, CC, BCC, Attachments)
            Exit Function
        End If

        ' Original standby...
        SendSimpleEmail = SendMailSMTP_VBAccelerator(From, FromName, T, TName, Subject, Body, CC, BCC, Attachments)
    End Function

    Private Function SendThroughOutlook(ByVal From As String, ByVal FromName As String, ByVal T As String, ByVal TName As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachments As Object = Nothing) As String
        Dim olOutlookApp As Object, olEMail As Object
        Dim NeedQuit As Boolean
        Dim TT As Integer
        Dim L As Object, A As Object, B As Object, N As Integer
        Dim CID As Object, tCID As String, img As String, ImgID As String, ImgNo As Integer, FS As Object
        Dim mCID As Object '@NO-LINT-NTYP
        Dim Res As String
        TT = GetEmailSetting("WOutlook") ' Which Outlook?
        If TT = 0 Then
            MessageBox.Show("No outlook program is specified in setup.  Please either select Outlook or Outlook Express or both if you wish to send through MS Outlook", "No Program Specified", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Function
        End If

        On Error Resume Next
        If (TT And 1) <> 0 Then
            olOutlookApp = GetObject(, "Outlook.Application")
            If Val(Err()) <> 0 Then    '   Outlook not running?
                olOutlookApp = CreateObject("Outlook.Application")
                NeedQuit = True
            End If
        End If

        If olOutlookApp Is Nothing Then  ' try outlook express..?
            SendThroughOutlook = "MS Outlook not available."
            If (TT And 2) <> 0 Then
                ' we have 2 ways to try to send through MAPI....  one uses MSMAPI.MAPISession, the other MAPI.Session......  Hopefully one will work!!
                SendThroughOutlook = SendThroughMAPI(From, FromName, T, TName, Subject, Body, CC, BCC, Attachments)
                If SendThroughOutlook <> "" Then ActiveLog("frmEmail::SendThroughMAPI Failure: " & SendThroughOutlook, 3)

                If SendThroughOutlook <> "" Then SendThroughOutlook = SendThroughMAPI2(From, FromName, T, TName, Subject, Body, CC, BCC, Attachments)
                If SendThroughOutlook <> "" Then ActiveLog("frmEmail::SendThroughMAPI2 Failure: " & SendThroughOutlook, 3)
                If SendThroughOutlook = "" Then
                    Math.Log("Email Sent Through MAPI: " & TName)
                    '        If Mode = emSimple Then
                    '          RaiseEvent SimpleEmailDone(True, "Email Sent.")
                    '        Else
                    '          MsgBox "Message sent.", vbInformation, "Email Sent"
                    '        End If
                Else
                    If (TT And 1) Then GoTo ViaOutlook
                    Math.Log("Email FAILED to Send Through MAPI: " & TName)
                    SendThroughOutlook = "Could not send mail through Outlook Express: " & SendThroughOutlook
                End If
            End If
            Exit Function
        End If

ViaOutlook:
        '   Create E-mail
        olEMail = olOutlookApp.CreateItem(0) ' (olMailItem)
        '    olEMail.Recipients.Add T
        olEMail.To = T
        olEMail.Subject = Subject
        '    olEMail.BodyFormat = 2 ' olFormatHTML
        Dim BP As String
        If IsFormLoaded("MailBookEmail") Then BP = GetFilePath(MailBookEmail.txtBodyFile.Text)
        Body = replaceHTMLimages(Body, BP, tCID, img)
        CID = tCID ' for the attachment to work, the CID has to be untyped.  Byref, above, needs typed
        '    olEMail.Body = "HTML Message Enclosed"
        olEMail.HTMLBody = Body
        If CC <> "" Then olEMail.CC = CC
        If BCC <> "" Then olEMail.BCC = BCC

        If img <> "" Then
            ImgNo = 0
            '      ImgID = olEMail.EntryID
            FS = Split(CID, ";")
            For Each L In Split(img, ";")
                Dim Att As Object, oPA As Object
                ImgNo = ImgNo + 1
                Att = olEMail.Attachments.Add(L)
                oPA = Att.PropertyAccessor
                oPA.SetProperty(PR_ATTACH_MIME_TAG, HTTPImageType(L))
                mCID = FS(ImgNo - 1) ' This variable must be un-typed for the attachment
                oPA.SetProperty(PR_ATTACH_CONTENT_ID, mCID) 'change myident for another other image
            Next
        End If

        If Attachments <> "" Then
            For Each L In Split(Attachments, ";")
                olEMail.Attachments.Add(L)
            Next
        End If

        olEMail.Send
        Math.Log("Email Sent through Outlook: " & TName)
        '    If Mode = emSimple Then
        '      RaiseEvent SimpleEmailDone(True, "Email Sent.")
        '    Else
        '      MsgBox "Message sent.", vbInformation, "Email Sent"
        '    End If


        If NeedQuit Then olOutlookApp.Quit
        olOutlookApp = Nothing
    End Function

    Public Function SendThroughMAPI2(ByVal From As String, ByVal FromName As String, ByVal T As String, ByVal TName As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachments As String = "") As String
        Dim oSession As Object, oNewMessage As Object
        Dim oRecipient As Object, oAttachment As Object
        Dim L As Object, AttNo As Integer
        Const CdoFileData As Integer = 1 ' && ReadFromFile method Attachment is a file (Default value)
        Const CdoFileLink As Integer = 2 ' &&Source property  Attachment is a link to a file
        Const CdoOLE As Integer = 3 ' &&ReadFromFile method  Attachment is an OLE object
        Const CdoEmbeddedMessage As Integer = 4 ' &&ID property of the Message object to be embedded  Attachment is an embedded message
        Const CdoDefaultFolderOutbox As Integer = 2


        On Error GoTo NoMAPI
        oSession = CreateObject("MAPI.Session")
        On Error GoTo NoLogon
        oSession.Logon '"Microsoft Outlook Internet Settings", "" ' <password>

        ' oFolder = oSession.GetDefaultFolder(CdoDefaultFolderOutbox)
        ' oMessage = oFolder.Messages.Add()

        On Error GoTo ErrHandler

        oNewMessage = oSession.Outbox.Messages.Add()
        oNewMessage.Subject = Subject
        oNewMessage.Text = Body
        '  oNewMessage.DeliveryReceipt = .T. 'Optional, forces a receipt send back

        oRecipient = oNewMessage.Recipients.Add()
        oRecipient.Name = T ' && This name must appear in your address book.
        '  oRecipient.Resolve ' && If name does not appear this will cause an error.
        oRecipient.Resolve


        AttNo = 0
        For Each L In Split(Attachments, ",")
            oAttachment = oNewMessage.Attachments.Add()
            oAttachment.Position = AttNo  ' && Important to position the attachement
            oAttachment.Type = CdoFileData
            oAttachment.ReadFromFile(InventFolder() & L)
            oAttachment.Source = InventFolder() & L
            oAttachment.Name = L
            AttNo = AttNo + 1
            oNewMessage.Update
        Next

        oNewMessage.Update
        oNewMessage.Send(1, 0, 0) ' && The 1 will save a copy of the message in your sent items folder of Outlook.
        oSession.LogOff

        oAttachment = Nothing
        oRecipient = Nothing
        oNewMessage = Nothing
        oSession = Nothing
        SendThroughMAPI2 = ""
        Exit Function

NoMAPI:
        SendThroughMAPI2 = "MAPI.Session could not be created: " & Err.Description
        Exit Function
NoLogon:
        SendThroughMAPI2 = "No available Outlook profile was available (or selected)."
        oSession = Nothing
        Exit Function
ErrHandler:
        SendThroughMAPI2 = "Sending through MAPI failed: " & Err.Description
        oAttachment = Nothing
        oRecipient = Nothing
        oNewMessage = Nothing
        oSession = Nothing
        Exit Function
    End Function

    Private Function SendThroughMAPI(ByVal From As String, ByVal FromName As String, ByVal T As String, ByVal TName As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachments As String = "") As String
        Dim objSession As Object 'MAPI.Session
        Dim objMessages As Object 'MAPI.Message
        Dim X As Integer, I As Integer
        Dim img As String, CID As String

        'http://lists.topica.com/lists/VB6Helper/read/message.html?sort=d&mid=812381415
        On Error GoTo NoMAPI
        objSession = CreateObject("MSMAPI.MAPISession")
        objMessages = CreateObject("MSMAPI.MAPIMessages")
        On Error GoTo NoSend

        objSession.DownLoadMail = False
        objSession.NewSession = False
        objSession.UserName = GetEmailSetting("username")
        objSession.Password = GetEmailSetting("password")
        objSession.SignOn

        X = objSession.SessionID
        objSession.SessionID = X

        I = 0
        objMessages.Compose
        objMessages.RecipIndex = I
        objMessages.RecipAddress = T
        objMessages.RecipDisplayName = TName
        objMessages.RecipType = 1

        If CC <> "" Then
            I = I + 1
            objMessages.RecipIndex = I
            objMessages.RecipAddress = CC
            '      .RecipDisplayName = tname
            objMessages.RecipType = 2
        End If

        If BCC <> "" Then
            I = I + 1
            objMessages.RecipIndex = I
            objMessages.RecipAddress = T
            '        .RecipDisplayName = BCC
            objMessages.RecipType = 3
        End If

        Dim FF As String
        FF = LocalDesktopFolder() & "tmp.htm"
        '      Body = replaceHTMLimages(Body, , CID, Img)
        WriteFile(FF, Body, True)

        objMessages.AttachmentIndex = 0
        objMessages.AttachmentType = 0
        objMessages.AttachmentName = "Sale.htm"
        objMessages.AttachmentPosition = 0
        objMessages.AttachmentPathName = FF

        objMessages.MsgSubject = Subject
        objMessages.MsgType = "text/html"
        objMessages.MsgNoteText = "See Attachment."

        Dim L As Object, NNN As Integer
        NNN = 0

        '      If Img <> "" Then
        '      ImgNo = 0
        '      FS = split(CID, ";")
        '      For Each L In split(Img, ";")
        '        NNN = NNN + 1
        '        olEMail.Attachments.Add L
        '        olEMail.Close olSave
        '        ImgID = olEMail.EntryID
        '        For Each L In split(Attachments, ";")
        '          NNN = NNN + 1
        '          .AttachmentIndex = NNN
        '          .AttachmentType = 1
        '          .AttachmentName = L
        '          .AttachmentPosition = NNN
        '          .AttachmentPathName = L
        '  Set oField = colFields.Add(CdoPR_ATTACH_MIME_TAG, "image/bmp")
        '  Set oField = colFields.Add(&H3712001E, CID)
        '  oMsg.Fields.Add "{0820060000000000C000000000000046}0x8514", 11, True
        '        Next
        '      End If
        '
        If Attachments <> "" Then
            For Each L In Split(Attachments, ";")
                NNN = NNN + 1
                objMessages.AttachmentIndex = NNN
                objMessages.AttachmentType = 1
                objMessages.AttachmentName = L
                objMessages.AttachmentPosition = NNN
                objMessages.AttachmentPathName = L
            Next
        End If

        objMessages.AddressResolveUI = False
        objMessages.Send(0)
        Kill(FF)

        objSession.SignOff
        objSession = Nothing
        objMessages = Nothing
        Exit Function
NoSend:
        If Err.Number = 32001 Then
            MessageBox.Show("You have cancelled the email--nothing has been sent.", Err.Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            SendThroughMAPI = "Send through MAPI failed: " & Err.Description
        End If

        objSession = Nothing
        objMessages = Nothing
        Exit Function
NoMAPI:
        objSession = Nothing
        objMessages = Nothing
        SendThroughMAPI = "You do not have MAPI installed."
        Exit Function
    End Function

    Public Function SendMailSMTP_VBAccelerator(ByVal From As String, ByVal FromName As String, ByVal T As String, ByVal TName As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachments As String = "") As String
        Dim F As frmSendMail
        F = New frmSendMail
        SendMailSMTP_VBAccelerator = F.DoSimpleSendMail(From, FromName, T, TName, Subject, Body, CC, BCC, Attachments)
        'Unload F
        F.Close()

        F = Nothing
    End Function

    Public Function SendMailSMTP_Chilkat(ByVal From As String, ByVal FromName As String, ByVal T As String, ByVal TName As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachments As String = "") As String
        Dim Success As Integer
        Dim L As Object
        Dim MM As Chilkat_v9_5_0.ChilkatMailMan
        Dim Mail As Chilkat_v9_5_0.ChilkatEmail
        MM = New Chilkat_v9_5_0.ChilkatMailMan
        Mail = New Chilkat_v9_5_0.ChilkatEmail

        Success = MM.UnlockComponent(CHILKAT_LICENSE)
        If Success <> 1 Then
            DevErr("Chilkat Email Component Licensure failure.")
            SendMailSMTP_Chilkat = "Chilkat Email Component Licensure failure."
            Exit Function
        End If

        MM.SMTPHost = GetEmailSetting("smtphost")
        MM.SMTPPort = GetEmailSetting("smtpport")
        MM.SmtpUsername = GetEmailSetting("smtpuser")
        MM.SmtpPassword = GetEmailSetting("smtppass")
        '    .SmtpPassword = "mawdyweqinofhdke"

        Mail.FromAddress = From
        Mail.FromName = FromName
        Mail.AddTo(TName, T)
        Mail.Subject = Subject

        Mail.Body = Body

        If CC <> "" Then Mail.AddMultipleCC(CC)
        If BCC <> "" Then Mail.AddMultipleBcc(BCC)

        If Attachments <> "" Then
            For Each L In Split(Attachments, ",")
                Mail.AddFileAttachment(PXfile(L))
            Next
        End If

        ProgressForm(0, 1, "Sending Email...", , , , ProgressBarStyle.prgIndefinite)
        Success = MM.SendEmail(Mail)
        ProgressForm()
        If Success <> 1 Then SendMailSMTP_Chilkat = "Sending failed (" & MM.SmtpFailReason & ")."

        DisposeDA(Mail)
    End Function


    '    Public Function SendMailSMTP_Chilkat(ByVal From As String, ByVal FromName As String, ByVal T As String, ByVal TName As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachments As String = "") As String
    '        Dim Success as integer
    '        Dim L As Object
    '        Dim MM As Chilkat_v9_5_0.ChilkatMailMan
    '        Dim Mail As Chilkat_v9_5_0.ChilkatEmail
    '        MM = New Chilkat_v9_5_0.ChilkatMailMan
    '        Mail = New Chilkat_v9_5_0.ChilkatEmail

    '        Success = MM.UnlockComponent(CHILKAT_LICENSE)
    '        If Success <> 1 Then
    '            DevErr("Chilkat Email Component Licensure failure.")
    '            SendMailSMTP_Chilkat = "Chilkat Email Component Licensure failure."
    '            Exit Function
    '        End If

    '        MM.SMTPHost = GetEmailSetting("smtphost")
    '        MM.SMTPPort = GetEmailSetting("smtpport")
    '        MM.SmtpUsername = GetEmailSetting("smtpuser")
    '        MM.SmtpPassword = GetEmailSetting("smtppass")
    '        '    .SmtpPassword = "mawdyweqinofhdke"

    '        Mail.FromAddress = From
    '        Mail.FromName = FromName
    '        Mail.AddTo TName, T
    '  Mail.Subject = Subject

    '        Mail.Body = Body

    '        If CC <> "" Then Mail.AddMultipleCC CC
    '  If BCC <> "" Then Mail.AddMultipleBcc BCC

    '  If Attachments <> "" Then
    '            For Each L In Split(Attachments, ",")
    '                Mail.AddFileAttachment PXfile(L)
    '    Next
    '        End If

    '        ProgressForm 0, 1, "Sending Email...", , , , prgIndefinite
    '  Success = MM.SendEmail(Mail)
    '        ProgressForm()
    '        If Success <> 1 Then SendMailSMTP_Chilkat = "Sending failed (" & MM.SmtpFailReason & ")."

    '        DisposeDA Mail
    'End Function

    Public Function GetEmailSetting(ByVal Key As String) As String
        Dim T As String
        T = EmailSettingAlt(Key)
        If T <> "" Then GetEmailSetting = T : Exit Function   ' allows homogenous interface over multiple data sources
        Select Case Key
            Case "uchilkat"
                '###SENDMAIL
                GetEmailSetting = False
                '      GetEmailSetting = FileExists(GetWindowsSystemDir() & DIR_SEP & "ChilkatAx-9.5.0-win32.dll")
            Case Else : GetEmailSetting = GetConfigTableValue(EmailSettingKey(Key), EmailSettingDef(Key))
        End Select
    End Function

    Public Function EmailSettingAlt(ByRef Key As String) As String
        EmailSettingKey(Key, , EmailSettingAlt)
    End Function

    Public Function EmailSettingDef(ByRef Key As String) As String
        EmailSettingKey(Key, EmailSettingDef)
    End Function

    Public Function EmailSettingKey(ByRef KeyIn As String, Optional ByRef EmailSettingDef As String = "", Optional ByRef EmailSettingAlt As String = "") As String
        Select Case LCase(KeyIn)
            Case "emailcfg" : EmailSettingKey = "EPO_EMailCFG" : EmailSettingDef = ""

            Case "fromname" : EmailSettingKey = "EPO_FromName" : EmailSettingAlt = StoreSettings.Name
            Case "fromaddr" : EmailSettingKey = "EPO_FromAddr" : EmailSettingAlt = StoreSettings.Email

            Case "smtphost" : EmailSettingKey = "EPO_SMTPHost" : EmailSettingDef = ""
            Case "smtpport" : EmailSettingKey = "EPO_SMTPPort" : EmailSettingDef = ""
            Case "smtpsecr" : EmailSettingKey = "EPO_SMTPSecr" : EmailSettingDef = "0"
            Case "smtpuser" : EmailSettingKey = "EPO_SMTPUser" : EmailSettingDef = ""
            Case "smtppass" : EmailSettingKey = "EPO_SMTPPass" : EmailSettingDef = ""
            Case "soutlook" : EmailSettingKey = "EPO_SOutlook" : EmailSettingDef = "0"
            Case "woutlook" : EmailSettingKey = "EPO_WOutlook" : EmailSettingDef = "3"
            Case "username" : EmailSettingKey = "EPO_OutlkUsr" : EmailSettingDef = ""
            Case "password" : EmailSettingKey = "EPO_OutlkPwd" : EmailSettingDef = ""
            Case "uchilkat" : EmailSettingKey = "EPO_UChilKat" : EmailSettingDef = ""
            Case Else : DevErr("Program error: Invalid Email Setting [" & KeyIn & "]", vbCritical, "Program Error")
        End Select
    End Function

    '    Private Function SendThroughOutlook(ByVal From As String, ByVal FromName As String, ByVal T As String, ByVal TName As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachments As Object = Nothing) As String
    '        Dim olOutlookApp As Object, olEMail As Object
    '        Dim NeedQuit As Boolean
    '        Dim TT as integer
    '        Dim L As Object, A As Object, B As Object, N as integer
    '        Dim CID As Object, tCID As String, img As String, ImgID As String, ImgNo as integer, FS As Object
    '        Dim mCID As Object '@NO-LINT-NTYP
    '        Dim Res As String
    '        TT = GetEmailSetting("WOutlook") ' Which Outlook?
    '        If TT = 0 Then
    '            MsgBox("No outlook program is specified in setup.  Please either select Outlook or Outlook Express or both if you wish to send through MS Outlook", vbExclamation, "No Program Specified")
    '            Exit Function
    '        End If

    '        On Error Resume Next
    '        If (TT And 1) <> 0 Then
    '            olOutlookApp = GetObject(, "Outlook.Application")
    '            If Err() <> 0 Then    '   Outlook not running?
    '                olOutlookApp = CreateObject("Outlook.Application")
    '                NeedQuit = True
    '            End If
    '        End If

    '        If olOutlookApp Is Nothing Then  ' try outlook express..?
    '            SendThroughOutlook = "MS Outlook not available."
    '            If (TT And 2) <> 0 Then
    '                ' we have 2 ways to try to send through MAPI....  one uses MSMAPI.MAPISession, the other MAPI.Session......  Hopefully one will work!!
    '                SendThroughOutlook = SendThroughMAPI(From, FromName, T, TName, Subject, Body, CC, BCC, Attachments)
    '                If SendThroughOutlook <> "" Then ActiveLog "frmEmail::SendThroughMAPI Failure: " & SendThroughOutlook, 3

    '      If SendThroughOutlook <> "" Then SendThroughOutlook = SendThroughMAPI2(From, FromName, T, TName, Subject, Body, CC, BCC, Attachments)
    '                If SendThroughOutlook <> "" Then ActiveLog "frmEmail::SendThroughMAPI2 Failure: " & SendThroughOutlook, 3
    '      If SendThroughOutlook = "" Then
    '                    Log("Email Sent Through MAPI: " & TName)
    '                    '        If Mode = emSimple Then
    '                    '          RaiseEvent SimpleEmailDone(True, "Email Sent.")
    '                    '        Else
    '                    '          MsgBox "Message sent.", vbInformation, "Email Sent"
    '                    '        End If
    '                Else
    '                    If (TT And 1) Then GoTo ViaOutlook
    '                    Log("Email FAILED to Send Through MAPI: " & TName)
    '                    SendThroughOutlook = "Could not send mail through Outlook Express: " & SendThroughOutlook
    '                End If
    '            End If
    '            Exit Function
    '        End If

    'ViaOutlook:
    '        '   Create E-mail
    '        olEMail = olOutlookApp.CreateItem(0) ' (olMailItem)
    '        '    olEMail.Recipients.Add T
    '        olEMail.To = T
    '        olEMail.Subject = Subject
    '        '    olEMail.BodyFormat = 2 ' olFormatHTML
    '        Dim BP As String
    '        If IsFormLoaded("MailBookEmail") Then BP = GetFilePath(MailBookEmail.txtBodyFile)
    '        Body = replaceHTMLimages(Body, BP, tCID, img)
    '        CID = tCID ' for the attachment to work, the CID has to be untyped.  Byref, above, needs typed
    '        '    olEMail.Body = "HTML Message Enclosed"
    '        olEMail.HTMLBody = Body
    '        If CC <> "" Then olEMail.CC = CC
    '        If BCC <> "" Then olEMail.BCC = BCC

    '        If img <> "" Then
    '            ImgNo = 0
    '            '      ImgID = olEMail.EntryID
    '            FS = Split(CID, ";")
    '            For Each L In Split(img, ";")
    '                Dim Att As Object, oPA As Object
    '                ImgNo = ImgNo + 1
    '                Att = olEMail.Attachments.Add(L)
    '                oPA = Att.PropertyAccessor
    '                oPA.SetProperty PR_ATTACH_MIME_TAG, HTTPImageType(L)
    '      mCID = FS(ImgNo - 1) ' This variable must be un-typed for the attachment
    '                oPA.SetProperty PR_ATTACH_CONTENT_ID, mCID 'change myident for another other image
    '            Next
    '        End If

    '        If Attachments <> "" Then
    '            For Each L In Split(Attachments, ";")
    '                olEMail.Attachments.Add L
    '    Next
    '        End If

    '        olEMail.Send
    '        Log("Email Sent through Outlook: " & TName)
    '        '    If Mode = emSimple Then
    '        '      RaiseEvent SimpleEmailDone(True, "Email Sent.")
    '        '    Else
    '        '      MsgBox "Message sent.", vbInformation, "Email Sent"
    '        '    End If


    '        If NeedQuit Then olOutlookApp.Quit
    '        olOutlookApp = Nothing
    '    End Function
    Public Function replaceHTMLimages(ByVal Doc As String, Optional ByVal BasePath As String = "", Optional ByRef C As String = "", Optional ByRef I As String = "") As String
        Dim X As Integer, Y As Integer, R As String, S As String, QC As String, CID As String
        Dim Src As String, F As String, Src1 As String
        If BasePath = "" Then BasePath = PXFolder()

        X = 1
        Do While True
            X = InStr(X, Doc, "<img ", vbTextCompare)
            If X = 0 Then Exit Do
            'Y = InStr(X, Doc, ">")   ERROR
            R = Mid(Doc, X, Y - X + 1)

            'Src1 = GetHTMLTagArgument(R, "src", QC)   ERROR
            Src = Replace(Src1, "/", "\\")
            F = CleanPath(Src, BasePath)
            If FileExists(F) Then
                'CID = Right(CreateUniqueID("{-}"), 15)   ERROR
                C = C & IIf(Len(C) > 0, ";", "") & CID
                I = I & IIf(Len(I) > 0, ";", "") & F
                S = Replace(R, QC & Src1 & QC, """" & "cid:" & CID & """")
                '      S = Replace(R, QC & Src & QC, """" & ImageToDataURI(F) & """")  ' could have worked, but IE7 disallows inline data
                Doc = Replace(Doc, R, S, , , vbTextCompare)
            End If

            X = X + Len(R)
        Loop
        replaceHTMLimages = Doc
    End Function

    Public Function EmailFactOrdNotAckBodyHTML() As String
        If FileExists(EmailTemplateFactOrdNotAckFile) Then
            MainMenu.rtbn.RichTextBox.LoadFile(EmailTemplateFactOrdNotAckFile)
            EmailFactOrdNotAckBodyHTML = MainMenu.rtbn.asHtml()
        Else
            EmailFactOrdNotAckBodyHTML = DefaultEmailFactOrdNotAck()
        End If
    End Function

    Public Function EmailOverdueOrdersBodyHTML() As String
        If FileExists(EmailTemplateOverdueOrdersFile) Then
            MainMenu.rtbn.RichTextBox.LoadFile(EmailTemplateOverdueOrdersFile)
            EmailOverdueOrdersBodyHTML = MainMenu.rtbn.asHtml()
        Else
            EmailOverdueOrdersBodyHTML = DefaultEmailOverdueOrders()
        End If
    End Function

    Public Function HasSendMail() As Boolean
        Dim X As Object
        On Error Resume Next
        X = CreateObject("vbSendMail.clsSendMail")
        HasSendMail = Not (X Is Nothing)
        X = Nothing
    End Function

    Public Function EmailTemplateFactOrdNotAckFile() As String
        EmailTemplateFactOrdNotAckFile = FXFolder() & "EmailTemplate-FactOrdNotAck.rtf"
    End Function

    Public Function DefaultEmailFactOrdNotAck(Optional ByVal DoStripHTML As Boolean = False) As String
        Dim S As String
        S = ""
        S = S & "To whom it may concern:<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "We have not received an acknowledgement yet on this order.<br/>" & vbCrLf
        S = S & "Please confirm <b>STATUS</b>, Acknowledgement No, and Anticipated <u>Shipping Date</u>.<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "Thank you.<br/>" & vbCrLf

        If DoStripHTML Then S = StripHTML(S)
        DefaultEmailFactOrdNotAck = S
    End Function

    Public Function EmailTemplateOverdueOrdersFile() As String
        EmailTemplateOverdueOrdersFile = FXFolder() & "EmailTemplate-OverdueOrders.rtf"
    End Function

    Public Function DefaultEmailOverdueOrders(Optional ByVal DoStripHTML As Boolean = False) As String
        Dim S As String
        S = ""
        S = S & "To whom it may concern:<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "The PO indicated above is currently <b>OVERDUE</b>.<br/>" & vbCrLf
        S = S & "<i>Please advise</i> the anticipated <b>Due Date</b> and <u>reason for delay</u>.<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "Thank you.<br/>" & vbCrLf

        If DoStripHTML Then S = StripHTML(S)
        DefaultEmailOverdueOrders = S
    End Function

    Private Function StripHTML(ByVal S As String) As String
        S = Replace(S, "<br/>", "") ' already handled with vbCrLf
        S = Replace(S, "<b>", "")
        S = Replace(S, "</b>", "")
        S = Replace(S, "<u>", "")
        S = Replace(S, "</u>", "")
        S = Replace(S, "<i>", "")
        S = Replace(S, "</i>", "")
        StripHTML = S
    End Function

    Public Function EmailChargeBackBodyHTML() As String
        If FileExists(EmailTemplateChargeBackFile) Then
            MainMenu.rtbn.RichTextBox.LoadFile(EmailTemplateChargeBackFile)
            EmailChargeBackBodyHTML = MainMenu.rtbn.asHtml()
        Else
            EmailChargeBackBodyHTML = DefaultEmailChargeBack()
        End If
    End Function

    Public Function EmailTemplateChargeBackFile() As String
        EmailTemplateChargeBackFile = FXFolder() & "EmailTemplate-ChargeBack.rtf"
    End Function

    Public Function DefaultEmailChargeBack(Optional ByVal DoStripHTML As Boolean = False) As String
        Dim S As String
        S = ""
        S = S & "To whom it may concern:<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "Attention: Accounts Receivable Department<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "Dear Sir:<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf

        If DoStripHTML Then S = StripHTML(S)
        DefaultEmailChargeBack = S
    End Function
End Module
