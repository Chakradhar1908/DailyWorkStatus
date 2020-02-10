Imports VBA
Imports VBRUN
Imports Microsoft.VisualBasic.Interaction
Imports stdole
Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class MainMenu4
    Private Const FRM_W_MIN As Integer = 14355
    Private Const FRM_H_MIN As Integer = 9810
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTCAPTION As Integer = 2
    Private Const MINWIDTH As Integer = 762
    Private Const MINHEIGHT As Integer = 507
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
    Public CurrentMenu As String, ParentMenu As String, CurrentMenuIndex As Integer

    Private Initializing As Boolean, Highlighting As String, ActiveForm As Boolean, CurrentHLIndex As Integer, ItemHLIndex As Integer
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
            ClearAccess()                 ' whenever we come back here, we clear it..
        Else
            If IsIn(SecurityLevel, ComputerSecurityLevels.seclevOfficeComputer, ComputerSecurityLevels.seclevSalesFloor) Then
                If IsLoggedIn Then cmdLogout.Visible = True
            End If
            ResetLastLoginExpiry(True)
        End If

        gblLastDeliveryDate = Today

        QBSM_Reset()

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
        If Not ActiveForm Or Not fActiveForm() Is Me Then Exit Sub
        If txtPassword.Visible Then
            txtPassword.Select()
        Else
            KeyCatch.Select()
        End If
    End Sub

    Private Sub ShowMsgs(Optional ByVal Show As Boolean = False)
        'msgs.Move(4000, 2550, 8025, 2700)        msgs is a custom active control. Still not developed.
        'msgs.Visible = Show And msgs.CheckMessages
    End Sub

    Private Sub MainMenu4_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        On Error Resume Next

        If Width < FRM_W_MIN Then
            Width = FRM_W_MIN
            Exit Sub
        End If

        If Height < FRM_H_MIN Then
            Height = FRM_H_MIN
            Exit Sub
        End If

        'LockWindowUpdate hWnd
        LockWindowUpdate(Handle)

        'imgInfo.Move ScaleWidth - imgInfo.Width, 0
        imgInfo.Location = New Point(Me.ClientSize.Width - imgInfo.Width, 0)

        Const bW As Integer = 30

        'imgStoreLogo.Move ScaleWidth / 2 - 5415 / 2, ScaleHeight - 1325 - 1755 / 2, 5415, 1755
        imgStoreLogo.Location = New Point(Me.ClientSize.Width / 2 - 5415 / 2, Me.ClientSize.Height - 1325 - 1755 / 2)
        imgStoreLogo.Size = New Size(5415, 1755)
        'imgStoreLogo.Stretch = True
        imgStoreLogo.SizeMode = PictureBoxSizeMode.StretchImage

        ResizeAndCenterPicture(imgStoreLogo, LoadPictureStd(StoreLogoFile()))
        'imgStoreLogoBorder.Move imgStoreLogo.Left - bW, imgStoreLogo.Top - bW, imgStoreLogo.Width + 2 * bW, imgStoreLogo.Height + 2 * bW
        imgStoreLogoBorder.Location = New Point(imgStoreLogo.Left - bW, imgStoreLogo.Top - bW)
        imgStoreLogoBorder.Size = New Size(imgStoreLogo.Width + 2 * bW, imgStoreLogo.Height + 2 * bW)

        'imgBackground.Move 0, 0, ScaleWidth, ScaleHeight
        imgBackground.Location = New Point(0, 0)
        imgBackground.Size = New Size(Me.ClientSize.Width, Me.ClientSize.Height)
        '  PaintPicture imgBackground.Picture, 0, 0, ScaleWidth, ScaleHeight, 0, 0, imgBackground.Picture.Width, imgBackground.Picture.Height

        'lblStore(0).Move 60, ScaleHeight - lblStore(0).Height - 60
        lblStore0.Location = New Point(60, Me.ClientSize.Height - lblStore0.Height - 60)

        'lblStore(1).Move ScaleWidth / 2 - lblStore(1).Width / 2, ScaleHeight - lblStore(1).Height - 60
        lblStore1.Location = New Point(Me.ClientSize.Width / 2 - lblStore1.Width / 2, Me.ClientSize.Height - lblStore1.Height - 60)
        'lblStore(2).Move ScaleWidth - lblStore(2).Width - 60, ScaleHeight - lblStore(2).Height - 60
        lblStore2.Location = New Point(Me.ClientSize.Width - lblStore2.Width - 60, Me.ClientSize.Height - lblStore2.Height - 60)

        LockWindowUpdate(IntPtr.Zero)

    End Sub

    Private Sub MainMenu4_Deactivate(sender As Object, e As EventArgs) Handles MyBase.Deactivate
        'Debug.Print "Form_Deactivate"
        ActiveForm = False
    End Sub

    Private Sub LoadMainMenu()
        Dim I As Integer, J As Integer, Count As Integer, X As Integer, Y As Integer
        Dim Cap As String, R As Integer, M As Integer, TPP As Integer
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

    Private Sub MainMenu4_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'QUUERY UNLOAD EVENT CODE

        'Select Case UnloadMode
        'Select Case e.CloseReason
        'Case vbFormCode, vbFormControlMenu
        If Not EndProgram() Then e.Cancel = True
        'End Select

        'UNLOAD EVENT CODE
        '  RemoveCustomFrame Me
        CleanUpProgram()
        '  Unload Me
        End
    End Sub

    Public Sub InitHotKeysLocal()
        m_cHotKey = New cRegHotKey
        'm_cHotKey.Attach(hWnd)
        m_cHotKey.Attach(Handle)
        InitHotKeys(m_cHotKey)
    End Sub

    Private Sub lblRDP_Click(sender As Object, e As EventArgs) Handles lblRDP.Click
        DisplayDevState(True)
    End Sub

    Private Sub lblBETA_Click(sender As Object, e As EventArgs) Handles lblBETA.Click
        DisplayDevState(True)
    End Sub

    Private Sub lblCDSComputer_Click(sender As Object, e As EventArgs) Handles lblCDSComputer.Click
        DisplayDevState(True)
    End Sub

    Private Sub lblDevMode_Click(sender As Object, e As EventArgs) Handles lblDevMode.Click
        DisplayDevState(True)
    End Sub

    Private Sub lblELEVATE_Click(sender As Object, e As EventArgs) Handles lblELEVATE.Click
        DisplayDevState(True)
    End Sub

    Private Sub lblIDE_Click(sender As Object, e As EventArgs) Handles lblIDE.Click
        DisplayDevState(True)
    End Sub

    Private Sub lblDEMO_Click(sender As Object, e As EventArgs) Handles lblDEMO.Click
        DisplayDevState(True)
    End Sub

    Private Sub lblMenuCaption_DoubleClick(sender As Object, e As EventArgs) Handles lblMenuCaption.DoubleClick
        'frmAbout.Show vbModal
        frmAbout.ShowDialog()
    End Sub

    Private Sub lblMenuItem_MouseMove(sender As Object, e As MouseEventArgs) Handles lblMenuItem.MouseMove
        'If imgMenuItem(Index).Width < 1000 Then MenuItemHighlight Index
        Dim I As Integer
        I = Mid(lblMenuItem.Name, 12)
        If imgMenuItem.Width < 1000 Then MenuItemHighlight(I)
    End Sub

    Private Sub m_cHotKey_HotKeyPress(ByVal sName As String, ByVal eModifiers As cRegHotKey.EHKModifiers, ByVal eKey As KeyCodeConstants)
        ReadHotKeyPress(sName)
    End Sub

    Private ReadOnly Property ScreenDX() As Double
        Get
            ScreenDX = 1 ' Screen.TwipsPerPixelX / 15
        End Get
    End Property

    Private ReadOnly Property ScreenDY() As Double
        Get
            ScreenDY = 1 ' Screen.TwipsPerPixelY / 15
        End Get
    End Property

    Private Sub StoreLogIn()
        On Error Resume Next
        ' BFH20150711 - Fixed the underlying, had to comment this out to make it not work again
        StoresSld = 1 ' DefaultLoginStore
    End Sub

    Public Sub ShutDown(Optional ByVal vQuick As Boolean = False)
        modMainMenu.DoShutDown(vQuick)
    End Sub

    Private Function EndProgram() As Boolean  ' called by Form_QueryUnload
        Dim mRes As VbMsgBoxResult
        If QuickQuit Or IsDevelopment() Then
            mRes = vbOK
        Else
            If IsWinXP() Then
                mRes = MessageBox.Show("Leave WinCDS running each night to get automatic updates!", "Exit WinCDS", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation)
            End If
            mRes = vbOK
        End If
        EndProgram = (mRes = vbOK)
    End Function

    Private Sub CleanUpProgram()
        Dim El As Form

        'For Each El In Forms
        For Each El In My.Application.OpenForms
            'If El.Name <> "MainMenu" Then Unload El
            If El.Name <> "MainMenu" Then El.Close()
        Next
        Domain_exit()
        Close()  ' This closed all open files in the old DOS days.
        Reset()  ' Close all files opened on the disk.
    End Sub

    Private Sub lblWinCDS_DoubleClick(sender As Object, e As EventArgs) Handles lblWinCDS.DoubleClick
        DoPractice()
    End Sub

    Public Sub StartVoidCheck(ByVal LeaseNo As String)
        '  tmrVoidCatch.Tag = LeaseNo
        '  RenewVoidCheck
    End Sub

    Public Sub RenewVoidCheck()
        '  tmrVoidCatch.Enabled = False
        '  tmrVoidCatch.Interval = 7000   ' 7 seconds
        '  tmrVoidCatch.Enabled = True
    End Sub

    Public Function StopVoidCheck() As String
        '  StopVoidCheck = tmrVoidCatch.Tag
        '  tmrVoidCatch.Enabled = False
        '  tmrVoidCatch.Tag = ""
    End Function

    ' backup GM lines to csv file to help catch void bug
    Public Function BackupSale(ByVal LeaseNo As String, Optional ByVal RemoveFile As Boolean = False) As Boolean
        Dim F As String, I As Integer
        F = InventFolder() & "Sale-" & LeaseNo & ".csv"

        On Error Resume Next
        Kill(F)
        If RemoveFile Then BackupSale = True : Exit Function

        For I = 1 To BillOSale.UGridIO1.LastRowUsed
            WriteFile(F, CSVLine(BillOSale.QueryStyle(I), BillOSale.QueryMfg(I), BillOSale.QueryMfgNo(I), BillOSale.QueryLoc(I), BillOSale.QueryStatus(I), BillOSale.QueryQuan(I), BillOSale.QueryDesc(I), BillOSale.QueryPrice(I)))
        Next
        BackupSale = True
    End Function

    Public Sub DetectBrokenSales()
        Dim T As Object, L As Object

        T = AllFiles(InventFolder() & "Sale-*.csv")
        For Each L In T
            If L <> "" Then
                On Error Resume Next
                MessageBox.Show("Failed To Save Sale Correctly: " & L & vbCrLf & "Please contact " & AdminContactName & " as soon as possible with this error message")
                'Name InventFolder() & L As InventFolder & "br" & L
                My.Computer.FileSystem.MoveFile(InventFolder() & L, InventFolder() & "br" & L)
            End If
        Next
    End Sub

    Private Sub mnuHelpLicenseAgreement_Click()
        ShowLicenseAgreement(True)
    End Sub

    Private Sub mnuHelpScreenShare_Click()
        Dim sbc As Object
        sbc = AddControlToForm("WinCDs.ScreenBroadcast", MainMenu)
        sbc.OpenConn("localhost", 5001)
    End Sub

    Private Sub txtPassword_Enter(sender As Object, e As EventArgs) Handles txtPassword.Enter
        SelectContents(txtPassword)
    End Sub

    Private Sub cmdEnterPassword_Click(sender As Object, e As EventArgs) Handles cmdEnterPassword.Click
        cmdEnterPassword.Visible = False
        txtPassword.Visible = False
    End Sub

    Public Sub GetSpeechInputMode(ByRef Result As Boolean, ByVal Func As String, ByVal cName As String)
        If Func = "command" Then Result = True
        If Func = "mainmenu" Then Result = True
        If IsIn(Func, "numbers", "letters", "styles") Then Result = False
    End Sub

    Private Sub tmrMaintain_Tick(sender As Object, e As EventArgs) Handles tmrMaintain.Tick
        MainMenu_Maintain_Timer()
    End Sub

    Private Sub KeyCatch_KeyDown(sender As Object, e As KeyEventArgs) Handles KeyCatch.KeyDown
        On Error Resume Next

        'MainMenu_NumberKeys KeyCode
        MainMenu_NumberKeys(e.KeyCode)
        ' what to do on the zero-key-press.  Used by speech interface as an anti-modal workaround.
        'If KeyCode = 48 And ZeroLaunch <> "" Then
        If e.KeyCode = Keys.D0 And ZeroLaunch <> "" Then
            Dim R As Object
            R = Split(ZeroLaunch, "|")
            ZeroLaunch = ""
            SelectMenuItem(, R(0), R(1))
            Exit Sub
        End If

        'Form_KeyDown KeyCode, Shift
        MainMenu4_KeyDown(Me, New KeyEventArgs(Keys.KeyCode And Keys.Shift))
    End Sub

    Private Sub KeyCatch_Leave(sender As Object, e As EventArgs) Handles KeyCatch.Leave
        CatchKeys()
    End Sub

    Private Sub MainMenu4_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        'MainMenu_KeyDown(KeyCode, Shift)
        MainMenu_KeyDown(e.KeyValue, e.KeyData)
    End Sub

    Public Sub SetWindowState(ByVal X As FormWindowStateConstants)
        If IsDevelopment() Then
            If X = FormWindowStateConstants.vbMaximized Then
                If IsFormLoaded("frmSplash") Then Exit Sub ' Developers really don't like this turned on..
            End If
        End If
        ActiveLog("MainMenu::SetWindowState(" & Switch(X = FormWindowStateConstants.vbMinimized, "vbMinimized", X = FormWindowStateConstants.vbMaximized, "vbMaximized", X = FormWindowStateConstants.vbNormal, "vbNormal", True, "" & X) & ")", 8)
        WindowState = X
    End Sub

    Private Sub txtInfo_DoubleClick(sender As Object, e As EventArgs) Handles txtInfo.DoubleClick
        txtInfo.Visible = False
    End Sub

    Private Sub imgInfo_DoubleClick(sender As Object, e As EventArgs) Handles imgInfo.DoubleClick
        LoadMenuToForm("")
        ShowMsgs(False)
        ShowInfo(Not txtInfo.Visible)
    End Sub

    Private Sub lblHelp_Click()
        LaunchHelp()
    End Sub

    Private Sub lblLogin_Click()
        FormClickLogin()
    End Sub

    Private Sub imgLogin_Click()
        FormClickLogin()
    End Sub

    Private Sub FormClickLogin()
        If Not CheckAccess("Log In To Other Stores") Then Exit Sub
        LaunchLogin()
    End Sub

    Private Sub LaunchLogin()
        'LogIn.Show vbModal, Me
        LogIn.ShowDialog(Me)
    End Sub

    Public Sub ReloadMenus()
        ResetMenus()
        LoadMenuToForm(CurrentMenu)
    End Sub

    Private Sub imgBackground_MouseMove(sender As Object, e As MouseEventArgs) Handles imgBackground.MouseMove
        LastMouseMove = Now
        MenuItemHighlight(-1, True)
        CatchKeys()
    End Sub

    Private Sub MainMenu4_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp
        'If Button = vbRightButton Then SystemMenuOnMouseUp Me
        If e.Button = MouseButtons.Right Then SystemMenuOnMouseUp(Me)
    End Sub

    Private Function tCaption() As String
        tCaption = ProgramCaption
        '  tCaption = tCaption & " v" & SoftwareVersion(True, False)
        tCaption = tCaption & " - " & GetLocalComputerName()
        tCaption = tCaption & " - " & IIf(IsServer, "SERVER", "WORKSTATION")
    End Function

    Private Sub InitForm()
        Initializing = True

        'HelpContextID = 10000
        SetAppIcon()
        'Caption = tCaption()
        Text = tCaption()

        LoadMainMenu()

        LoadMenuToForm("")

        Initializing = False
    End Sub

    'Private Sub bvb_MouseEnter(Index As Integer) : MainMenuPulse(Index) : End Sub
    'Private Sub bvb_MouseExit(Index As Integer) : MainMenuPulse Index, True: End Sub
    Private Sub bvb_MouseEnter(sender As Object, e As EventArgs) Handles bvb0.MouseEnter, bvb1.MouseEnter, bvb2.MouseEnter, bvb3.MouseEnter, bvb4.MouseEnter, bvb5.MouseEnter
        Dim p As PictureBox

        p = CType(sender, PictureBox)
        MainMenuPulse(Microsoft.VisualBasic.Right(p.Name, 1))
    End Sub

    Private Sub bvb_MouseLeave(sender As Object, e As EventArgs) Handles bvb0.MouseLeave, bvb1.MouseLeave, bvb2.MouseLeave, bvb3.MouseLeave, bvb4.MouseLeave, bvb5.MouseLeave
        Dim p As PictureBox

        p = CType(sender, PictureBox)
        MainMenuPulse(Microsoft.VisualBasic.Right(p.Name, 1), True)
    End Sub

    Private Sub MainMenuPulse(ByVal Index As Integer, Optional ByVal StopIt As Boolean = False)
        tmrPulse.Enabled = True
        'bvb(Index).LightnessPct = 0  -> Commented this line. Because in vb6, bvb is an "alphaimage" control. But .Net is not supporting it. 
        '--> So replaced it with picturebox control. And picturebox does not have "LightnessPct" property.
        tmrPulse.Tag = "-1"
        If StopIt Then Exit Sub
        tmrPulse.Interval = 10
        tmrPulse.Tag = Index
        tmrPulse.Enabled = True
    End Sub

    Private Sub tmrPulse_Tick(sender As Object, e As EventArgs) Handles tmrPulse.Tick
        Const C As Integer = 6
        Const T As Integer = 1200
        Const Q As Double = 300
        Dim R As Double

        If tmrPulse.Tag = "-1" Then
            tmrPulse.Enabled = False
            Exit Sub
        End If
        R = (GetTickCount Mod T) - Q
        R = IIf(R < 500, R, T - R)
        'bvb(Val(tmrPulse.Tag)).LightnessPct = R * C / T + 25 --> Commented two lines,because bvb is third party control called "alphaimage".
        'bvb(Val(tmrPulse.Tag)).Refresh                           Replaced with picture box. 
    End Sub

    'Private Sub bvb_Click(Index As Integer) : MainMenuClick Index: End Sub
    Private Sub bvb_Click(sender As Object, e As EventArgs) Handles bvb0.Click, bvb1.Click, bvb2.Click, bvb3.Click, bvb4.Click, bvb5.Click
        Dim p As PictureBox

        p = CType(sender, PictureBox)
        MainMenuClick(Microsoft.VisualBasic.Right(p.Name, 1))
    End Sub

    Public Sub MainMenuClick(ByVal Index As Integer)
        On Error Resume Next
        Dim I As Integer, K As String
        Dim C As Integer, W As Integer
        Dim T As Boolean

        If Index = 5 And Not Installment Then
            MessageBox.Show("Installment module not enabled." & vbCrLf & "Please contact " & AdminContactName & " at " & AdminContactPhone2 & " for pricing.", "Error - Not Installed", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        'For I = bvb.LBound To bvb.UBound
        For I = 0 To 5
            'K = LCase(Replace(bvb(I).Tag, " ", "")) & "U"
            Select Case I
                Case 0
                    K = LCase(Replace(bvb0.Tag, " ", "")) & "U"
                    bvb0.Image = MainMenu4_Images.MenuImage("mm", K)
                Case 1
                    K = LCase(Replace(bvb1.Tag, " ", "")) & "U"
                    bvb1.Image = MainMenu4_Images.MenuImage("mm", K)
                Case 2
                    K = LCase(Replace(bvb2.Tag, " ", "")) & "U"
                    bvb2.Image = MainMenu4_Images.MenuImage("mm", K)
                Case 3
                    K = LCase(Replace(bvb3.Tag, " ", "")) & "U"
                    bvb3.Image = MainMenu4_Images.MenuImage("mm", K)
                Case 4
                    K = LCase(Replace(bvb4.Tag, " ", "")) & "U"
                    bvb4.Image = MainMenu4_Images.MenuImage("mm", K)
                Case 5
                    K = LCase(Replace(bvb5.Tag, " ", "")) & "U"
                    bvb5.Image = MainMenu4_Images.MenuImage("mm", K)
            End Select
            'Debug.Print K
            'Set bvb(I).Picture = MainMenu4_Images.MenuImage("mm", K)
        Next

        C = 300
        'W = bvb(0).Width
        W = bvb0.Width

        Dim Tk As Integer, Mk As Integer
        Tk = GetTickCount
        Mk = Tk + C
        Do While GetTickCount < Mk
            I = (GetTickCount - Tk)
            If Index <> -1 Then
                'bvb(Index).Left = -I / C * W
                Select Case Index
                    Case 0
                        bvb0.Left = -I / C * W
                    Case 1
                        bvb1.Left = -I / C * W
                    Case 2
                        bvb2.Left = -I / C * W
                    Case 3
                        bvb3.Left = -I / C * W
                    Case 4
                        bvb4.Left = -I / C * W
                    Case 5
                        bvb5.Left = -I / C * W
                End Select

                If I > (C / 2) And Not T Then
                    T = True
                    'K = LCase(Replace(bvb(Index).Tag, " ", "")) & "D"
                    Select Case Index
                        Case 0
                            K = LCase(Replace(bvb0.Tag, " ", "")) & "D"
                            bvb0.Image = MainMenu4_Images.MenuImage("mm", K)
                            bvb0.Refresh()
                        Case 1
                            K = LCase(Replace(bvb1.Tag, " ", "")) & "D"
                            bvb1.Image = MainMenu4_Images.MenuImage("mm", K)
                            bvb1.Refresh()
                        Case 2
                            K = LCase(Replace(bvb2.Tag, " ", "")) & "D"
                            bvb2.Image = MainMenu4_Images.MenuImage("mm", K)
                            bvb2.Refresh()
                        Case 3
                            K = LCase(Replace(bvb3.Tag, " ", "")) & "D"
                            bvb3.Image = MainMenu4_Images.MenuImage("mm", K)
                            bvb3.Refresh()
                        Case 4
                            K = LCase(Replace(bvb4.Tag, " ", "")) & "D"
                            bvb4.Image = MainMenu4_Images.MenuImage("mm", K)
                            bvb4.Refresh()
                        Case 5
                            K = LCase(Replace(bvb5.Tag, " ", "")) & "D"
                            bvb5.Image = MainMenu4_Images.MenuImage("mm", K)
                            bvb5.Refresh()
                    End Select
                    'Debug.Print K
                    'Set bvb(Index).Picture = MainMenu4_Images.MenuImage("mm", K)
                End If
                'bvb(Index).Refresh
            End If

            If CurrentMenuIndex <> -1 Then
                'bvb(CurrentMenuIndex).Left = -(W - I / C * W)
                'bvb(CurrentMenuIndex).Refresh
                Select Case CurrentMenuIndex
                    Case 0
                        bvb0.Left = -(W - I / C * W)
                        bvb0.Refresh()
                    Case 1
                        bvb1.Left = -(W - I / C * W)
                        bvb1.Refresh()
                    Case 2
                        bvb2.Left = -(W - I / C * W)
                        bvb2.Refresh()
                    Case 3
                        bvb3.Left = -(W - I / C * W)
                        bvb3.Refresh()
                    Case 4
                        bvb4.Left = -(W - I / C * W)
                        bvb4.Refresh()
                    Case 5
                        bvb5.Left = -(W - I / C * W)
                        bvb5.Refresh()
                End Select
            End If
        Loop

        If Index <> -1 Then
            'bvb(Index).Left = 0
            'bvb(Index).Refresh
            Select Case Index
                Case 0
                    bvb0.Left = 0
                    bvb0.Refresh()
                Case 1
                    bvb1.Left = 0
                    bvb1.Refresh()
                Case 2
                    bvb2.Left = 0
                    bvb2.Refresh()
                Case 3
                    bvb3.Left = 0
                    bvb3.Refresh()
                Case 4
                    bvb4.Left = 0
                    bvb4.Refresh()
                Case 5
                    bvb5.Left = 0
                    bvb5.Refresh()
            End Select
        End If

        If CurrentMenuIndex <> -1 Then
            'bvb(CurrentMenuIndex).Left = 0
            'bvb(CurrentMenuIndex).Refresh
            Select Case CurrentMenuIndex
                Case 0
                    bvb0.Left = 0
                    bvb0.Refresh()
                Case 1
                    bvb1.Left = 0
                    bvb1.Refresh()
                Case 2
                    bvb2.Left = 0
                    bvb2.Refresh()
                Case 3
                    bvb3.Left = 0
                    bvb3.Refresh()
                Case 4
                    bvb4.Left = 0
                    bvb4.Refresh()
                Case 5
                    bvb5.Left = 0
                    bvb5.Refresh()
            End Select
        End If

        CurrentMenuIndex = Index

        If Index <> -1 Then
            'LoadMenuToForm bvb(Index).Tag
            Select Case Index
                Case 0
                    LoadMenuToForm(bvb0.Tag)
                Case 1
                    LoadMenuToForm(bvb1.Tag)
                Case 2
                    LoadMenuToForm(bvb2.Tag)
                Case 3
                    LoadMenuToForm(bvb3.Tag)
                Case 4
                    LoadMenuToForm(bvb4.Tag)
                Case 5
                    LoadMenuToForm(bvb5.Tag)
            End Select
        End If
    End Sub

    'Private Sub imgMenuItem_MouseEnter(Index As Integer) : MenuItemHighlight Index: End Sub
    'Private Sub imgMenuItem_MouseExit(Index As Integer) : MenuItemHighlight Index, True: End Sub

    Private Sub imgMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles imgMenuItem.MouseEnter
        Dim I As Integer
        I = Mid(imgMenuItem.Name, 12)
        MenuItemHighlight(I)
    End Sub

    Private Sub imgMenuItem_MouseLeave(sender As Object, e As EventArgs) Handles imgMenuItem.MouseLeave
        'MenuItemHighlight(Index, True)
        Dim I As Integer
        I = Mid(imgMenuItem.Name, 12)
        MenuItemHighlight(I, True)
    End Sub

    Public Sub MenuItemHighlight(ByVal Index As Integer, Optional ByVal StopIt As Boolean = False)
        Const X As Double = 0.15
        Dim I As Integer, C As Integer
        Dim L As Integer, T As Integer
        Dim D As Integer

        If Index < 0 Then
            imgSelected.Visible = False
            imgSubSelected.Visible = False
            Exit Sub
        End If


        'If imgMenuItem(Index).Width > 1000 Then
        If imgMenuItem.Width > 1000 Then
            'L = imgMenuItem(Index).Left
            L = imgMenuItem.Left
            'T = imgMenuItem(Index).Top
            T = imgMenuItem.Top

            If StopIt Then
                'imgMenuItem(Index).Top = imgMenuItem(Index).Top + 1920 * X
                imgMenuItem.Top = imgMenuItem.Top + 1920 * X
                'imgMenuItem(Index).Left = imgMenuItem(Index).Left + 1920 * X
                imgMenuItem.Left = imgMenuItem.Left + 1920 * X
                'imgMenuItem(Index).Width = 1920
                imgMenuItem.Width = 1920
                'imgMenuItem(Index).Height = 1920
                imgMenuItem.Height = 1920
                'imgMenuItem(Index).Effect = lvicNoEffects
                'imgMenuItem.Effect = lvicNoEffects --------> Commented 2 lines because this property if of Alphaimage control.
                'imgMenuItem(Index).Effects.GrayScale = lvicNoGrayScale
                imgSelected.Visible = False
            Else
                Dim Tk As Integer, Mk As Integer
                C = 150
                Tk = GetTickCount
                Mk = Tk + C
                Do While GetTickCount < Mk
                    I = GetTickCount - Tk
                    D = CLng(1920.0# * X * CDbl(I) / CDbl(C))
                    'imgMenuItem(Index).Move L - D, T - D, 1920 + 2 * D, 1920 + 2 * D
                    'imgMenuItem.Move L - D, T - D, 1920 + 2 * D, 1920 + 2 * D
                    imgMenuItem.Location = New Point(L - D, T - D)
                    imgMenuItem.Size = New Size(1920 + 2 * D, 1920 + 2 * D)
                    'imgMenuItem(Index).Refresh()
                    imgMenuItem.Refresh()
                    'MoveControlTo imgSelected, imgMenuItem(Index), True
                    MoveControlTo(imgSelected, imgMenuItem, True)
                    imgSelected.Refresh()
                Loop
            End If
        Else
            If StopIt Then
                imgSubSelected.Visible = False
                'lblMenuItem(Index).BackStyle = cc2BackstyleTransparent
            Else
                MoveControl(imgSubSelected, -10000, , , , True, True)
                'lblMenuItem(Index).ZOrder 0
                lblMenuItem.BringToFront()
                '        lblMenuItem(Index).AutoSize = True

                '        lblMenuItem(Index).BackStyle = cc2BackstyleOpaque
                '        lblMenuItem(Index).BackColor = vbRed
                'imgSubSelected.Move lblMenuItem(Index).Left - 100, imgMenuItem(Index).Top + 20, 3800, lblMenuItem(Index).Height + 60
                imgSubSelected.Location = New Point(lblMenuItem.Left - 100, imgMenuItem.Top + 20)
                imgSubSelected.Size = New Size(3800, lblMenuItem.Height + 60)
                'MoveControlTo imgSubSelected, lblMenuItem(Index)
            End If
        End If
    End Sub

    'Private Sub imgMenuItem_Click(Index As Integer) : SelectMenuItem Index: End Sub
    'Private Sub lblMenuItem_Click(Index As Integer) : SelectMenuItem Index: End Sub

    Private Sub imgMenuItem_Click(sender As Object, e As EventArgs) Handles imgMenuItem.Click
        'SelectMenuItem(Index)
        Dim I As Integer
        I = Mid(imgMenuItem.Name, 12)
        SelectMenuItem(I)
    End Sub

    Private Sub lblMenuItem_Click(sender As Object, e As EventArgs) Handles lblMenuItem.Click
        'SelectMenuItem(Index)
        Dim I As Integer
        I = Mid(lblMenuItem.Name, 12)
        SelectMenuItem(I)
    End Sub

    Public Sub LoadMenuToForm(ByVal Menu As String)
        Dim I As Integer, ArtPic As String, Art As StdPicture
        ActiveLog("MainMenu::LoadMenuToForm(" & Menu & ")", 4)

        ShowMsgs(Menu = "")

        Menu = Replace(LCase(Menu), "&", "")

        Select Case LCase(Menu)
            Case "payables", "payroll", "banking", "time c"
                LaunchProgram(LCase(Menu))
            Case Else
                On Error Resume Next
                'For I = 1 To imgMenuItem.UBound
                'Unload imgMenuItem(I)
                'Unload lblMenuItem(I)
                'Next
                imgMenuItem.Hide()
                lblMenuItem.Hide()

                'If imgStoreLogo.Picture = 0 Then Set imgStoreLogo.Picture = LoadPictureStd(StoreLogoFile())
                If IsNothing(imgStoreLogo.Image) Then imgStoreLogo.Image = LoadPictureStd(StoreLogoFile())
                imgStoreLogo.Visible = (Menu = "" And IsNothing(imgStoreLogo.Image))
                imgStoreLogoBorder.Visible = imgStoreLogo.Visible
                'imgStoreLogoBorder.BackStyle = 1
                lblMenuCaption.Tag = LCase(Menu)

                On Error GoTo 0
                GenericLoader(Menu)
                '      MsgBox "Program referenced a non-existant menu: " & Menu, vbCritical, "Menu Error"
        End Select
    End Sub

    Private Function ItemOptionString(ByVal Caption As String, ByVal Menu As String, ByVal Operation As String, ByVal Src As String, ByVal HotKeys As String) As String
        ItemOptionString = CSVLine(Caption, Menu, Operation, Src, HotKeys)
    End Function

    Private Function ItemOptionCaption(ByVal S As String) As String
        ItemOptionCaption = CSVField(S, 1)
    End Function
    Private Function ItemOptionMenu(ByVal S As String) As String
        ItemOptionMenu = CSVField(S, 2)
    End Function
    Private Function ItemOptionOp(ByVal S As String) As String
        ItemOptionOp = CSVField(S, 3)
    End Function
    Private Function ItemOptionSrc(ByVal S As String) As String
        ItemOptionSrc = CSVField(S, 4)
    End Function
    Public Function ItemOptionHotKeys(ByVal S As String) As String
        ItemOptionHotKeys = CSVField(S, 5)
    End Function

    Private Sub SetMenuItemImage(ByVal MI As PictureBox, ByVal Menu As String, ByVal Operation As String)
        MI.Image = MainMenu4_Images.MenuImage(Menu, Operation)
    End Sub

    Private Function UnloadHRs() As Boolean
        Dim I As Integer
        'For I = imgHR.UBound To 1 Step -1
        'Unload imgHR(I)
        'Unload lblHR(I)
        'Next
        imgHR.Hide()
        UnloadHRs = True
    End Function

    Private Function LoadHR(ByVal Caption As String, ByVal X As Integer, ByVal Y As Integer, Optional ByVal W As Integer = 4500) As Boolean
        Dim N As Integer
        'N = imgHR.UBound + 1
        'Load imgHR(N)
        'Set imgHR(N).Picture = imgHR(0).Picture

        'MoveControl imgHR(N), X, Y, W, 300, True, True
        MoveControl(imgHR, X, Y, W, 380, True, True)
        'Load lblHR(N)
        'lblHR(N).Caption = Caption
        lblHR.Text = Caption
        'MoveControl lblHR(N), imgHR(N).Left + 120, imgHR(N).Top, , , True, True
        MoveControl(lblHR, imgHR.Left + 120, imgHR.Top, , , True, True)
        LoadHR = True
    End Function

    Private Sub GenericLoader(ByVal MenuName As String)
        Dim MM As MyMenu, Idx As Integer, Li As Integer, Lh As Integer
        Dim I As Integer, Src As String, Ctrl As StdPicture, MICap As String, TTT As String
        Dim MI As MyMenuItem, HR As MyMenuHR
        Dim R As Integer, TPP As Integer, TPP2 As Integer

        'TPP = Screen.TwipsPerPixelX
        'TPP2 = Screen.TwipsPerPixelY

        On Error Resume Next
        'For I = 1 To imgMenuItem.UBound
        'Unload imgMenuItem(I)
        'Unload lblMenuItem(I)
        'Next
        imgMenuItem.Hide()
        lblMenuItem.Hide()

        lblMenuCaption.Tag = MenuName

        MM = GetMyMenu(MenuName, Idx)
        If Idx = -1 Then MessageBox.Show("Unknown menu: " & MenuName, "Ooops!", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Exit Sub
        CurrentMenu = MenuName
        ParentMenu = MM.ParentMenu
        'HelpContextID = MM.HCID

        Li = -1
        Li = UBound(MM.Items)
        Lh = -1
        Lh = UBound(MM.HRs)
        On Error GoTo 0

        If MM.Caption = "WinCDS" Then
            lblMenuCaption.Text = ""
        Else
            lblMenuCaption.Text = MM.Caption
        End If

        UnloadHRs()

        If MM.Layout = eMyMenuLayouts.eMML_4x8x8 Or MM.Layout = eMyMenuLayouts.eMML_3x8x8 Then
            If MM.SubTitle1 <> "" Then LoadHR(MM.SubTitle1, 3200, 4000)
            If MM.SubTitle2 <> "" Then LoadHR(MM.SubTitle2, 8700, 4000)
        ElseIf MM.Layout = eMyMenuLayouts.eMML_3x2x4x4 Or MM.Layout = eMyMenuLayouts.eMML_4x2x4x4 Or MM.Layout = eMyMenuLayouts.eMML_4x2x5x5 Then
            If MM.SubTitle1 <> "" Then LoadHR(MM.SubTitle1, 3200, 5700)
            If MM.SubTitle2 <> "" Then LoadHR(MM.SubTitle2, 8700, 5700)
        End If

        If Li >= 0 Then
            For I = LBound(MM.Items) + 1 To UBound(MM.Items) + 1
                MI = MM.Items(I - 1)
                TTT = IIf(MI.ControlCode = "", MI.ToolTipText, "[" & MI.ControlCode & "] " & MI.ToolTipText)
                'Load imgMenuItem(I)
                'Load lblMenuItem(I)
                'imgMenuItem(I).Tag = ItemOptionString(MI.Caption, MenuName, MI.Operation, MI.ImageKey, MI.HotKeys)
                imgMenuItem.Tag = ItemOptionString(MI.Caption, MenuName, MI.Operation, MI.ImageKey, MI.HotKeys)
                'lblMenuItem(I).Caption = Replace(MI.Caption, "/", vbCrLf)
                lblMenuItem.Text = Replace(MI.Caption, "/", vbCrLf)
                'imgMenuItem(I).ToolTipText = TTT
                'imgMenuItem.ToolTipText = TTT
                'SetMenuItemImage imgMenuItem(I), MenuName, MI.ImageKey
                SetMenuItemImage(imgMenuItem, MenuName, MI.ImageKey)
                If Not MI.IsSubItem Then
                    'imgMenuItem(I).Move(MI.Left + MM.ImageW + 10), MI.Top, 1650, 1650
                    'imgMenuItem(I).Move(MI.Left + MM.ImageW + 10), MI.Top, 1650, 1650
                    imgMenuItem.Location = New Point((MI.Left + MM.ImageW + 10), MI.Top)
                    imgMenuItem.Size = New Size(1650, 1650)
                    'lblMenuItem(I).Move imgMenuItem(I).Left + imgMenuItem(I).Width / 2 - lblMenuItem(I).Width / 2, imgMenuItem(I).Top + imgMenuItem(I).Height + 60
                    lblMenuItem.Location = New Point(imgMenuItem.Left + imgMenuItem.Width / 2 - lblMenuItem.Width / 2, imgMenuItem.Top + imgMenuItem.Height + 60)
                Else
                    'imgMenuItem(I).Move MI.Left, MI.Top, 500, 500
                    imgMenuItem.Location = New Point(MI.Left, MI.Top)
                    imgMenuItem.Size = New Size(500, 500)
                    'lblMenuItem(I).Move imgMenuItem(I).Left + imgMenuItem(I).Width, imgMenuItem(I).Top + 60, 4000
                    lblMenuItem.Location = New Point(imgMenuItem.Left + imgMenuItem.Width, imgMenuItem.Top + 60)
                    lblMenuItem.Size = New Size(4000, lblMenuItem.Height)
                    'lblMenuItem(I).Alignment = 0
                    'lblMenuItem.Alignment = 0
                    'lblMenuItem(I).Caption = Replace(lblMenuItem(I).Caption, vbCrLf, " ")
                    lblMenuItem.Text = Replace(lblMenuItem.Text, vbCrLf, " ")
                End If
                'lblMenuItem(I).ToolTipText = TTT
                'Debug.Print imgMenuItem(I).Left & "x" & imgMenuItem(I).Top & "..." & imgMenuItem(I).Width & "x" & imgMenuItem(I).Height & " - " & ScaleWidth & "x" & ScaleHeight

                'imgMenuItem(I).Visible = True
                imgMenuItem.Visible = True
                'imgMenuItem(I).ZOrder 0
                imgMenuItem.BringToFront()
                'lblMenuItem(I).Visible = True
                lblMenuItem.Visible = True
                'lblMenuItem(I).ZOrder 0
                lblMenuItem.BringToFront()

            Next
        End If

    End Sub

    Public Function SelectMenuItem(Optional ByVal Index As Integer = -1, Optional ByVal ExplicitMenu As String = "", Optional ByVal ExplicitOperation As String = "") As Boolean
        Dim X As Integer, F As String
        Dim Operation As String, Source As String, MI As MyMenuItem
        ActiveLog("MainMenu::SelectMenuItem(" & Index & ", " & ExplicitMenu & ", " & ExplicitOperation & ")", 5)

        Dim Fail As Boolean, FailMsg As String, FailTitle As String ' not really planning to fail
        If Index = -1 And ExplicitMenu <> "" And ExplicitOperation <> "" Then
            Source = ExplicitMenu
            Operation = ExplicitOperation
        Else
            Source = lblMenuCaption.Tag
            'On Error Resume Next
            'Operation = ItemOptionOp(imgMenuItem(Index).Tag)
            Operation = ItemOptionOp(imgMenuItem.Tag)
        End If
        FailMsg = "You have encountered a program error or the resource has moved." & vbCrLf & "Please contact " & AdminContactCompany & " at " & AdminContactPhone2 & " immediately." & vbCrLf & "Thank-you, and sorry for the inconvenience." & vbCrLf & "Source=" & Source & vbCrLf & "Operation=" & Operation
        FailTitle = "Unknown Menu Function"

        If Microsoft.VisualBasic.Left(Operation, 1) = "#" Then
            GenericLoader(Mid(Operation, 2))
            Exit Function
        End If

        SelectMenuItem = modMainMenu.MainMenu_Dispatch(Source, Operation)
    End Function

    Public Function DeveloperEx() As String
        Dim S As String
        S = ""
        S = MainMenu_NumberKeys_DeveloperEx()
        DeveloperEx = S
    End Function

    Public Function NOOP() As Boolean
    End Function

    Public Function CalcBase64() As String
        CalcBase64 = SoftwareVersionForLog()
    End Function

    Private Sub WebServ_HandleGET(FileName As String, Result As String, StatusCode As String, Headers As String)
        MessageBox.Show("WebServ Get FILENAME" & vbCrLf & FileName)
    End Sub

    Private Sub lblWinCDS_Click(sender As Object, e As EventArgs) Handles lblWinCDS.Click
        ReloadMenus()
    End Sub

    Private Sub mnuFileSettings_Click()
        SelectMenuItem(, "file", "systemsetup")
    End Sub
    Private Sub mnuFileUpdate_Click()
        SelectMenuItem(, "file:maintenance", "webupdates")
    End Sub
    Private Sub mnuFileExit_Click()
        SelectMenuItem(, "file", "exit")
    End Sub

    Private Sub mnuStore_Click()
        SelectMenuItem(, "file", "login")
    End Sub

    Private Sub mnuHelpSupport_Click()
        LaunchAutoVNC()
    End Sub
    Private Sub mnuHelpUploadLogs_Click()
        DiagnosticDataUpload(Logs:=True)
    End Sub
    Private Sub mnuHelpUploadData_Click()
        DiagnosticDataUpload()
    End Sub
    Private Sub mnuHelpContact_Click()
        MessageBox.Show(AdminContactCompany & vbCrLf2 & AdminContactString(0, True, False, True, True, True, True, True, True, True), "Company Contact Information")
    End Sub
    Private Sub mnuHelpContents_Click()
        ShowHelp()
    End Sub
    Private Sub mnuHelpAbout_Click()
        'frmVersionControl.Show 1
        frmVersionControl.ShowDialog()
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
        DoLogOut()
    End Sub

    Public Sub DoLogOut()
        modPasswords.LogOut()
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
            If lblBETA.Visible Then lblBETA.Text = ExeChannelName() : lblBETA.ForeColor = ExeChannelNameColor()
            lblELEVATE.Visible = IsElevated()
            ' this is the only one a user might see.
            lblRDP.Visible = SessionIsRemote()
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
        Main()
        Try
            If modKillBug.KillBug Then End
            AdjustFormForLargeFonts()
            InitForm()
            StoreLogIn()  ' Default store 1
            modSetup.CheckEXEUpdate()
            mUpdateInstance = Random(120)

            CurrentMenuIndex = 1

            'Top = Screen.Height - Height / 2
            Top = Screen.PrimaryScreen.Bounds.Height - Height / 2
            'Left = Screen.Width - Width / 2
            Left = Screen.PrimaryScreen.Bounds.Width - Width / 2

            'mnuHelpScreenShare.Visible = IsCDSComputer("prototype")
            '  If IsIDE And IsDevelopment Then InitHotKeysLocal
            '  SetCustomFrame Me, ncMacLook

            'NOTE: In vb6, for image control(imgPicture) assigned datasource as datacontrl and datafied as "Picture" column(code is in mod2DataPictures modules ->GetDatabasePicture function).
            'Replacement for it in vb.net is the below line. This code line is not in vb6. In vb6 values are directly assigned in the design time properties window of imgPicture image control.

            '---------      NOTE: BELOW CODE LINE IS NOT IN VB6.  ------------
            'imgPicture.DataBindings.Clear()  NOTE: REMOVE THIS COMMENTE IF imgPicture.DataBindings.Add will expect Clear first before Add.
            imgPicture.DataBindings.Add("Image", datPicture, "Picture")
        Catch ex As ArgumentException
            'ArgumentException will raise because before adodc control(datPicture) connection code to execute in another form, this MainMenu4 form will executes.
        End Try
    End Sub

    Private Sub AdjustFormForLargeFonts()
        Dim dX As Double, dY As Double
        Dim L As Object, Name As String, TName As String

        'If Screen.TwipsPerPixelX = 15 Then Exit Sub

        Dim g As Graphics
        g = CreateGraphics()
        Dim r As Single
        r = 15 / 1440 * g.DpiX
        'dX = Screen.TwipsPerPixelX / 15
        'dY = Screen.TwipsPerPixelY / 15

        ActiveLog("MainMenu::AdjustFormForLargeFonts - Adjusting...  dX=" & dX & ", dY=" & dY)

        On Error GoTo BadMove
        For Each L In Controls
            Name = L.Name
            TName = TypeName(L)
            If IsIn(TName, "ImageList", "MSComm", "Timer", "CommonDialog", "Inet", "Menu") Then GoTo SkipControl

            If IsIn(Name, "imgBackground", "imgStoreLogo", "imgStoreLogoBorder") Then GoTo SkipControl
            '    If IsIn(Name, "imgStoreLogo") Then GoTo SkipControl

            If TName = "Line" Then
                L.X1 = L.X1 * dX
                L.X2 = L.X2 * dX
                L.Y1 = L.Y1 * dY
                L.Y2 = L.Y2 * dY
            Else
                L.Left = L.Left * dX
                L.Top = L.Top * dY

                Select Case TName
                    Case "Label", "TextBox"
                        L.FontSize = L.FontSize * dX
                        L.Width = L.Width * dX
                        L.Height = L.Height * dY
                    Case "Image"
                        L.Width = L.Width * dX
                        L.Height = L.Height * dY
                End Select
            End If
SkipControl:
        Next
        Exit Sub

BadMove:
        Debug.Print("Bad Move: " & Err.Description)
        Debug.Print("Control: " & Name)
        Debug.Print("Type: " & TName)
        Err.Clear()
        Resume Next
    End Sub
End Class