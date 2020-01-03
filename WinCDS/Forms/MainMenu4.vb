Public Class MainMenu4
    Private Const FRM_W_MIN As Long = 14355
    Private Const FRM_H_MIN As Long = 9810
    Private Const WM_NCLBUTTONDOWN As Long = &HA1
    Private Const HTCAPTION As Long = 2
    Private Const MINWIDTH As Long = 762
    Private Const MINHEIGHT As Long = 507
    ' BFH20060522 - Info about the controls on this page:
    '   MSComm1             - Used only in Functions::OpenCashDrawer
    '   tmrMaintafin        - Only referenced in MainMenu::tmrMaintain_Timer.. used to restart the program and perform maintenance.  Disables itself if a file named PreventRestart.txt exists in the local Store1\NewOrder\ folder
    '   tmrVoidCatch        - Used to catch Random Void sales.  NOT ACTIVE
    '   tmrPulse            - Used to make the main menu items flash slowly in time.
    '   cdgFile             - Used in MainMenu::ImportDiscrepancyData and frmPictures::imgPicture_DblClick and frmExportMail and frmPriceChangeExcel
    '   rtb                 - Various
    '   rtbn                - Various (workaround)
    '   rtbStorePolicy      - To take the place of the one on the BoS form... It actually had 2
    '   flb                 - Used by modDesignTag
    '                         Because getting a file list is difficult w/o a FileListBox
    '
    '   imlStandardButtons  - Used throughout software for all standard buttons
    '   imlMiniButtons      - Used throughout software for all mini buttons
    '   imlMM               - Used locally
    '
    '   picSplash           - Displays the CDS logo in the background of the main menu
    '   lblStore(0..2)      - Reflects the currently logged in store (name/add1/add2)
    '   lblStoreType        - Static display of the phrase "POS Software"
    '
    '   txtPassword         - Where users can enter a password - This and cmdEnterPassword work in conjunction with modPasswords
    '   cmdEnterPassword    - The 'OK' button that corresponds to txtPassword
    '
    '   picAlpha            - modAPI.DrawRectangleToDC
    '
    '   datPicture
    '
    ' Store Info for the store we're currently logged into.
    Public CurrentMenu As String, ParentMenu As String, CurrentMenuIndex As Long

    Private Initializing As Boolean, Highlighting As String, ActiveForm As Boolean, CurrentHLIndex As Long, ItemHLIndex As Long
    Public MouseX As Single, MouseY As Single
    Public LastMouseMove As Date
    Public ZeroLaunch As String

    Public WithEvents WebServ As frmHTTPServer
    Private WithEvents m_cHotKey As cRegHotKey

    Private Sub MainMenu4_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        On Error Resume Next
        DisplayDevState()

        'lblStore(0) = StoreSettings.Name
        'lblStore(1) = StoreSettings.Address
        'lblStore(2) = StoreSettings.City
        lblStore0.Text = StoreSettings.Name
        lblStore1.Text = StoreSettings.Address
        lblStore2.Text = StoreSettings.City

        imgStoreLogo.Image = Nothing
        imgStoreLogo.Image = StoreLogoPicture()

        'imgStoreLogo.Visible = imgStoreLogo.Visible And (imgStoreLogo.Picture <> 0)
        imgStoreLogo.Visible = imgStoreLogo.Visible And Not IsNothing(imgStoreLogo.Image)
        imgStoreLogoBorder.Visible = imgStoreLogo.Visible
        'imgStoreLogoBorder.BackStyle = 1

        On Error GoTo 0
        cmdLogout.Visible = False
        If Not AllowUseLastEntry Then
            ClearAccess                 ' whenever we come back here, we clear it..
        Else
            If IsIn(SecurityLevel, ComputerSecurityLevels.seclevOfficeComputer, ComputerSecurityLevels.seclevSalesFloor) Then
                If IsLoggedIn Then cmdLogout.Visible = True
            End If
            ResetLastLoginExpiry(True)
        End If

        gblLastDeliveryDate = Today

        QBSM_Reset

        LoadMainMenu()
        ShowMsgs(CurrentMenu = "")
        ActiveForm = True
        CatchKeys()
        'Form_Resize
        MainMenu4_Resize(Me, New EventArgs)
    End Sub

    Private Sub CatchKeys()
        On Error Resume Next
        'Debug.Print "CatchKeys (Active=" & ActiveForm & "), fActiveForm=" & fActiveForm.Caption
        If Not ActiveForm Or Not fActiveForm Is Me Then Exit Sub
        If txtPassword.Visible Then
            txtPassword.Select()
        Else
            KeyCatch.Select()
        End If
    End Sub

    Private Sub ShowMsgs(Optional ByVal Show As Boolean = False)
        msgs.Move 4000, 2550, 8025, 2700
  msgs.Visible = Show And msgs.CheckMessages
    End Sub

    Private Sub MainMenu4_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize

    End Sub

    Private Sub LoadMainMenu()
        Dim I As Long, J As Long, Count As Long, X As Long, Y As Long
        Dim Cap As String, R As Long, M As Long, TPP As Long
        Dim MenuCaptions() As Object

        ShowInfo(False)
        ShowMsgs(False)
        '
        I = 0
        'bvb(I).Tag = "File" : I = I + 1
        'bvb(I).Tag = "Order Entry" : I = I + 1
        'bvb(I).Tag = "Inventory" : I = I + 1
        'bvb(I).Tag = "Accounting" : I = I + 1
        'bvb(I).Tag = "Mailing" : I = I + 1
        'bvb(I).Tag = "Installment" : I = I + 1

        bvb0.Tag = "File" : I = I + 1
        bvb1.Tag = "Order Entry" : I = I + 1
        bvb2.Tag = "Inventory" : I = I + 1
        bvb3.Tag = "Accounting" : I = I + 1
        bvb4.Tag = "Mailing" : I = I + 1
        bvb5.Tag = "Installment" : I = I + 1
        '  bvb(I - 1).Visible = Installment
        ' No longer used
        '  If IsUFO() Then bvb(I).Caption = "&Time Clock": I = I + 1
    End Sub

    Private Sub ShowInfo(Optional ByVal Show As Boolean = False)
        txtInfo.Text = AdminContactCompany & vbCrLf2 & AdminContactString(0, True, False, True, True, True, True, True, True, True)
        'txtInfo.Locked = True
        txtInfo.ReadOnly = True
        'txtInfo.Move 235 * 15, 152 * 15, 535 * 15, 180 * 15
        txtInfo.Location = New Point(235 * 15, 152 * 15)
        txtInfo.Size = New Size(535 * 15, 180 * 15)
        txtInfo.Visible = Show
    End Sub

    Private Sub cmdLogout_Click(sender As Object, e As EventArgs) Handles cmdLogout.Click
        DoLogOut
    End Sub

    Public Sub DoLogOut()
        modPasswords.LogOut
    End Sub

    Private Sub DisplayDevState(Optional ByVal ForceOff As Boolean = False)
        Dim AllowState As Boolean

        '  AllowState = True
        AllowState = IsIDE() Or IsDevelopment() ' Jerry doesn't want these shown
        If ForceOff Then AllowState = False

        lblDEMO.Visible = IsDemo()
        'lblDEMO.ToolTipText = "DEMO EXPIRES: " & DemoExpirationDate()
        ttpMainMenu.SetToolTip(lblDEMO, "DEMO EXPIRES: " & DemoExpirationDate())

        If AllowState Then
            lblCDSComputer.Visible = IsCDSComputer()
            lblIDE.Visible = IsIDE()
            lblDevMode.Visible = IsDevelopment()
            lblBETA.Visible = IsBetaChannel()
            If lblBETA.Visible Then lblBETA.Text = ExeChannelName() : lblBETA.ForeColor = ExeChannelNameColor
            lblELEVATE.Visible = IsElevated
            ' this is the only one a user might see.
            lblRDP.Visible = SessionIsRemote
        Else
            lblCDSComputer.Visible = False
            lblIDE.Visible = False
            lblDevMode.Visible = False
            lblBETA.Visible = False
            lblELEVATE.Visible = False
            lblRDP.Visible = False
        End If
    End Sub


    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Case "newsale"
        If CrippleBug("New Sales") Then Exit Sub
        If Not CheckAccess("Create Sales") Then Exit Sub
        Order = "A"
        'frmSalesList.SafeSalesClear = True
        frmSalesList.SalesCode = ""
        'Unload BillOSale
        BillOSale.Close()
        MainMenu.Hide()
        'BillOSale.HelpContextID = 42000
        'BillOSale.HelpContextID = 42002
        BillOSale.Show()
        'MailCheck.HelpContextID = 42000
        'MailCheck.optTelephone.Value = True
        MailCheck.HidePriorSales = True
        'MailCheck.Show vbModal  ' If this is loaded "vbModal, BillOSale", lockup may occur.
        MailCheck.ShowDialog()
        MailCheck.HidePriorSales = False
        'Unload MailCheck
        MailCheck.Close()
    End Sub

    Private Sub MainMenu4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'NOTE: In vb6, for image control(imgPicture) assigned datasource as datacontrl and datafied as "Picture" column(code is in mod2DataPictures modules ->GetDatabasePicture function).
        'Replacement for it in vb.net is the below line. This code line is not in vb6. Values are directly assigned in the design time properties window of imgPicture image control.

        Try
            'imgPicture.DataBindings.Clear()  NOTE: REMOVE THIS COMMENTE IF imgPicture.DataBindings.Add will expect Clear first before Add.
            imgPicture.DataBindings.Add("Image", datPicture, "Picture")
        Catch ex As ArgumentException
            'ArgumentException will raise because before adodc control(datPicture) connection code to execute in another form, this MainMenu4 form will executes.
        End Try
    End Sub


End Class