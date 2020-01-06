Public Class MessageCenter
    Private EDate As Date
    Private CurrentMessage As String

    Public Property EffectiveDate() As Date
        Get
            'EffectiveDate = IIf(CLng(EDate) = 0, Today, EDate)
            EffectiveDate = IIf(IsNothing(EDate), Today, EDate)
            '  If IsDevelopment Then EffectiveDate = #9/10/2016#      ' For Testing
        End Get
        Set(value As Date)
            EDate = value
            CheckMessages
        End Set
    End Property

    Public Sub Reset()
        Const xW As Long = 360
        'X.Move Width - X.Width, 0
        X.Location = New Point(Width - X.Width, 0)
        X.Visible = False
        'txt.Move 0, xW, Width, Height - xW
        txt.Location = New Point(0, xW)
        txt.Size = New Size(Width, Height - xW)
        txt.Visible = False

        'img.Move 0, xW, Width, Height - xW
        img.Location = New Point(0, xW)
        img.Size = New Size(Width, Height - xW)
        img.Visible = False
    End Sub

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        CheckMessages()
    End Sub

    Public Function CheckMessages() As Boolean
        On Error Resume Next
        CheckMessages = True
        Reset()
        X.Visible = True

        ' Give thought to the order.  Only the first will show....
        ' Commissions and EOY take precedence over Ads
        ' TEST is last, or first.
        ' Just consider it..

        If EndOfYearMessage() Then Exit Function
        If CommissionsOnFirstOfMonth() Then Exit Function

        If CashOptReverse() Then Exit Function

        If OldPasswordMessage() Then Exit Function

        If DisplayDispatchTrackAd() Then Exit Function
        If DisplayAWSAd() Then Exit Function
        If IntegratedCCAd() Then Exit Function

        If DisplayInstallmentNotice() Then Exit Function

        If DisplayTEST() Then Exit Function

        If SoftwareOutOfDate Then Exit Function

        X.Visible = False
        CheckMessages = False
    End Function

    Private Function EndOfYearMessage() As Boolean
        Dim S As String
        If Month(EffectiveDate) <> 1 Then Exit Function
        If DateAndTime.Day(EffectiveDate) > 10 Then Exit Function

        If Val(GetConfigTableValue("MessageCenter_EndOfYearMessage", "0")) >= Year(EffectiveDate) Then Exit Function

        S = ""
        '  S = S & "On Jan 1, after you have run your Quarterly Sales Comparison Reports and Best Seller Report, "
        '  S = S & vbCrLf & "you MUST Annually Update Quarterly Sales."
        '  S = S & vbCrLf2 & "This function is located on the FILE MENU."
        '  S = S & vbCrLf & "If you have one store or multiple stores click ONE time only! "
        '  S = S & vbCrLf2 & "This message will stop after January 10th."

        S = S & "On January 1, after you have run your Quarterly Sales Comparison Reports and Best Seller Report, "
        S = S & vbCrLf & "you MUST Annually Update For Quarterly Sales."
        S = S & vbCrLf & "This function is located on the FILE MENU; then click Maintenance."
        S = S & vbCrLf2 & "MAKE A 12/31 BACKUP AND ARCHIVE FOR USE LATER."
        S = S & vbCrLf & "Also please run from Inventory Reports; Year End Cost Report. "
        S = S & vbCrLf2 & "If you have one store or multiple stores click ONE time only! "
        S = S & vbCrLf2 & "This message will stop after January 10."

        txt.Text = S
        txt.Visible = True
        CurrentMessage = "EndOfYearMessage"
        EndOfYearMessage = True
    End Function

    Private Function CommissionsOnFirstOfMonth() As Boolean
        Dim G As Long

        If Not Installment Then Exit Function
        G = StoreSettings.GracePeriod

        ' This method only works for Grace <= 10 (and then not in February).
        '  If Not IsIn(Day(EffectiveDate), 1 + G, 10 + G, 20 + G) Then Exit Function
        'Checking (EffectiveDate - Grace) will allow for any date.
        If Not IsIn(DateAndTime.Day(DateAdd("d", -G, EffectiveDate)), 1, 10, 20) Then Exit Function

        If Val(GetConfigTableValue("MessageCenter_CommissionsOnFirstOfMonth", "0")) >= Val(DateStamp) Then Exit Function

        On Error Resume Next
        'img.Picture = iml.ListImages("comm").Picture
        img.Image = iml.Images(4)
        img.Visible = True
        CurrentMessage = "CommissionsOnFirstOfMonth"
        CommissionsOnFirstOfMonth = True
    End Function

    Private Function CashOptReverse() As Boolean
        Dim S As String

        If Not Installment Then Exit Function
        If Not DateBetween(EffectiveDate, #8/18/2016#, #8/28/2016#) Then Exit Function

        S = S & "NOTICE:  Due to popular demand, the fields for Deferred Payment and Same as Cash"
        S = S & vbCrLf & "have been REVERSED on the Installment Payment Setup."
        S = S & vbCrLf2 & "Please note this change, that the 'Same as Cash' option"
        S = S & vbCrLf & "is now in the main column, and the deferred payment has been moved to the right."
        S = S & vbCrLf2 & "This message will stop after Sept 10."

        txt.Text = S

        txt.Visible = True
        X.Visible = False

        CurrentMessage = "CashOptReverse"
        CashOptReverse = True
    End Function

    Private Function OldPasswordMessage() As Boolean
        OldPasswordMessage = False
        Exit Function

        Dim S As String
        If Not DateBetween(EffectiveDate, #7/6/2012#, #10/15/2012#) Then Exit Function
        '  If Not AllowUseLastEntry Then Exit Function
        '  If IsDevelopment Then Exit Sub
        S = ""
        S = S & "*** Notice -- Important Password System Change"
        S = S & vbCrLf & ""
        S = S & vbCrLf & "The password system has been changed to make it more usable."
        S = S & vbCrLf & "After you enter your password, you will still be able to access"
        S = S & vbCrLf & "the software for three minutes after your last interaction"
        S = S & vbCrLf & "with the software.  If you continue to use the software, you will"
        S = S & vbCrLf & "not need to re-enter your password for every operation you make."
        S = S & vbCrLf & ""
        S = S & vbCrLf & "Additionally, to provide additional security, if you would like to"
        S = S & vbCrLf & "be sure you are logged out so that no one can operate the software"
        S = S & vbCrLf & "without your permission, a logout button is provided in the upper left."
        '  DisplayMainMenuMessage S
    End Function

    Private Function DisplayDispatchTrackAd() As Boolean
        If DDTLicensed Then Exit Function
        If Val(GetConfigTableValue("MessageCenter_DisplayDispatchTrackAd", "0")) >= Val(DateStamp) Then Exit Function
        If Not DateBetween(EffectiveDate, #5/11/2015#, #5/15/2015#) Then Exit Function

        On Error Resume Next
        X.Visible = False
        'img.Picture = iml.ListImages("DT").Picture
        img.Image = iml.Images(3)
        img.Visible = True
        CurrentMessage = "DisplayDispatchTrackAd"
        DisplayDispatchTrackAd = True
    End Function

    Private Function DisplayAWSAd() As Boolean
        If StoreSettings.AmazonCustomerBucket <> "" Then Exit Function
        If Val(GetConfigTableValue("MessageCenter_DisplayAWSAd", "0")) >= Val(DateStamp) Then Exit Function
        If Not DateBetween(EffectiveDate, #6/1/2015#, #6/5/2015#) Then Exit Function

        On Error Resume Next
        X.Visible = False
        'img.Picture = iml.ListImages("aws").Picture
        img.Image = iml.Images(2)
        img.Visible = True
        CurrentMessage = "DisplayAWSAd"
        DisplayAWSAd = True
    End Function

    Private Function IntegratedCCAd() As Boolean
        If IsDevelopment() And DateBefore(Today, #1/20/2017#) Then GoTo ShowNow
        If StoreSettings.CCConfig <> "" Then Exit Function
        If Val(GetConfigTableValue("MessageCenter_IntegratedCCAd", "0")) >= Val(DateStamp) Then Exit Function
        If Not DateBetween(EffectiveDate, #1/15/2017#, #1/20/2017#) Then Exit Function

ShowNow:
        On Error Resume Next
        X.Visible = False
        'img.Picture = iml.ListImages("cc2").Picture
        img.Image = iml.Images(1)
        img.Visible = True
        CurrentMessage = "IntegratedCCAd"
        IntegratedCCAd = True
    End Function

    Private Function DisplayInstallmentNotice()
        On Error Resume Next
        Dim S As String
        If Not Installment Then Exit Function
        If Not DateBetween(EffectiveDate, #7/18/2017#, #7/25/2017#) Then Exit Function
        If Val(GetConfigTableValue("MessageCenter_DisplayInstallmentNotice", "0")) >= Year(EffectiveDate) Then Exit Function

        S = ""
        S = S & "INSTALLMENT CUSTOMERS:"
        S = S & vbCrLf2 & "If you are doing Credit Reporting, we have a new upgrade "
        S = S & "for Add on Sales to better comply with new requirements." & vbCrLf2 & "Most noticeable will be having a "
        S = S & "New Account Number assigned to the add on sale.  This is automatically processed for you."

        txt.Text = S
        txt.Visible = True
        X.Visible = True
        CurrentMessage = "DisplayInstallmentNotice"
        DisplayInstallmentNotice = True
    End Function

    Private Function DisplayTEST() As Boolean
        Exit Function ' Comment this line out to test...
        On Error Resume Next
        X.Visible = False
        'img.Picture = iml.ListImages("test").Picture
        img.Image = iml.Images(5)
        img.Visible = True
        CurrentMessage = "DisplayTEST"
        DisplayTEST = True
    End Function

    Private Function SoftwareOutOfDate()
        On Error Resume Next
        Dim S As String, R As String
        If Not DateAfter(Today, DateAdd("m", 3, BuildDate)) Then Exit Function
        R = GetConfigTableValue("MessageCenter_SoftwareOutOfDate", "0")
        If DateBefore(Today, DateAdd("m", 3, DateStampValue(R))) Then
            Exit Function
        End If

        S = ""
        S = "Software Out of Date"
        S = S & vbCrLf2 & "Your WinCDS software appears to be getting out of date."
        S = S & vbCrLf & "WinCDS is updated regularly via your internet connection, and"
        S = S & vbCrLf & "this copy was created on " & BuildDate() & "."

        txt.Text = S
        txt.Visible = True
        X.Visible = True
        CurrentMessage = "SoftwareOutOfDate"
        SoftwareOutOfDate = True
    End Function

    Private Sub MessageCenter_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        On Error Resume Next
        Reset()
    End Sub

    Private Function ConfigTableClosedField(Optional ByVal Msg As String = "") As String
        ConfigTableClosedField = "MessageCenter_" & IIf(Msg = "", CurrentMessage, Msg)
    End Function

    Private Sub X_Click(sender As Object, e As EventArgs) Handles X.Click
        Select Case CurrentMessage
            Case "OldPasswordMessage" : DevErr("Legacy - OldPasswordMessage")
            Case "DisplayTEST" : MessageBox.Show("Displayed by Code, cannot remove")

            Case "EndOfYearMessage" : SetConfigTableValue(ConfigTableClosedField, Year(EffectiveDate))
            Case "CashOptReverse" : SetConfigTableValue(ConfigTableClosedField, EffectiveDate)

            Case Else : SetConfigTableValue(ConfigTableClosedField, DateStamp)
        End Select
        CheckMessages()
    End Sub
End Class
