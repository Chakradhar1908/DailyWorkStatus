Imports stdole
Imports VBA
Imports Microsoft.VisualBasic.Interaction
Module modMainMenu
    Public mUpdateInstance As Integer ' To keep everyone from hitting the server at the same time...
    Public Structure MyMenu
        Dim Name As String
        Dim ParentMenu As String
        Dim Caption As String
        Dim Visible As Boolean
        Dim Layout As eMyMenuLayouts
        Dim HCID As Integer

        Dim ImageW As Integer
        Dim ImageH As Integer
        Dim vSP As Integer


        Dim ImageSource As Object 'ImageList
        Dim CaptionStyle As eCaptionStyles
        Dim CaptionMargin As Integer
        Dim MaskColor As Integer

        Dim Items() As MyMenuItem
        Dim HRs() As MyMenuHR

        Dim SubTitle1 As String
        Dim SubTitle2 As String
    End Structure
    Public Enum eMyMenuLayouts
        eMML_Manual = 0
        eMML_2x3
        eMML_2x4
        eMML_3x3
        eMML_3x4
        eMML_4x2
        eMML_4x3
        eMML_4x4

        eMML_2x3Across
        eMML_2x4Across
        eMML_3x3Across
        eMML_3x4Across
        eMML_4x2Across
        eMML_4x3Across
        eMML_4x4Across

        eMML_3x8x8
        eMML_4x8x8
        eMML_3x2x4x4
        eMML_4x2x4x4
        eMML_4x2x5x5
    End Enum
    Public Enum eCaptionStyles
        eCS_None = 0
        eCS_RightCenter
        eCS_RightBottom
        eCS_Below
    End Enum
    Public Structure MyMenuItem
        Dim ImageKey As String
        Dim Caption As String
        Dim Top As Integer
        Dim Left As Integer

        Dim ToolTipText As String
        Dim ControlCode As String

        Dim HotKeys As String
        Dim Operation As String

        Dim Visible As String
        Dim Image As StdPicture

        Dim IsSubItem As Boolean
    End Structure

    Public Structure MyMenuHR
        Dim Top As Integer
        Dim Left As Integer
        Dim Width As Integer
    End Structure

    Private frmSplas As Form = frmSplash2
    Public Const frmSplashType As String = "frmSplash2"
    Public MyMenusInitialized As Boolean
    Private MyMenus() As MyMenu
    Private LastPayPalCheck As Date
    Private DidUpdate As Boolean     ' Used to prevent more automatic updates after first call
    Private DidAutoBackup As Boolean   ' Used to record whether the software has automatically backed itself up
    Private mExportTaskList As Boolean
    Public MenuItemCount As Integer
    Public Const MainMenuType As String = "MainMenu4"

    'Public ReadOnly Property frmSplash As frmSplash2
    '    Get
    '        Return frmSplas
    '    End Get
    'End Property

    Public ReadOnly Property frmSplashIsLoaded As Boolean
        Get
            frmSplashIsLoaded = IsFormLoaded(frmSplashType)
        End Get
    End Property

    Public Sub SetButtonImage(ByRef cmd As Button, Optional ByVal ImageIndex As Integer = -1, Optional ByVal MiniButton As Boolean = False)
        '::::SetButtonImage
        ':::SUMMARY
        ': Set image on CmdButton control
        ':::DESCRIPTION
        ': Initializes a command button image.  The button must already be set to type Graphical, as these cannot be set in code.
        ':
        '::Available Image Keys:
        ': - calc,gear,config,notes,none,calendar,cancel
        ': - rStop,rDelete,rInfo,rNext,rAdd,rPrefs,rReload,rSearch
        ': - ok,clear,map,import,print,menu,back,forward
        ': - zoom,preview,next,previous,next1,previous1,delete,plus,minus,refresh
        ': - south,west,east,north
        ':::PARAMETERS
        ': - cmd - Indicates the Command Button.
        ': - ImageName - Indicates the Image Name.
        ': - MiniButton - Indicates whether it is true or false.
        'Dim T As String
        'If cmd.Style <> vbButtonGraphical Then
        'If cmd.Image Is Nothing Then
        '    Debug.Print("Bad button")
        '    If IsDevelopment() Then
        '        Err.Raise(-1, "Development Code", "Not a graphical button: " & cmd.Name)
        '        Stop
        '    End If
        'End If
        'cmd.UseMaskColor = True
        'cmd.MaskColor = vbWhite
        'If ImageName = "" Then
        '    T = LCase(cmd.Name)
        '    If LCase(Left(T, 3)) = "cmd" Then T = Mid(T, 4)
        '    If IsIn(T, "ok", "apply", "post", "done", "close", "process") Then
        '        ImageName = "ok"
        '    ElseIf T Like "*preview" Then
        '        ImageName = "preview"
        '    ElseIf T = "cancel" Then
        '        ImageName = "cancel"
        '    ElseIf T = "clear" Then
        '        ImageName = "clear"
        '    ElseIf IsIn(T, "config", "settings", "setup", "options", "save") Then
        '        ImageName = "config"
        '    ElseIf T = "print" Then
        '        ImageName = "print"
        '    ElseIf T Like "*menu*" Then
        '        ImageName = "menu"
        '    ElseIf T Like "*next*" Then
        '        ImageName = "next"
        '    ElseIf T Like "*prev*" Then
        '        ImageName = "previous"
        '    ElseIf T Like "*del*" Then
        '        ImageName = "delete"
        '    ElseIf T Like "*calendar*" Then
        '        ImageName = "calendar"
        '    ElseIf T Like "*refresh*" Then
        '        ImageName = "refresh"
        '    ElseIf T Like "*down*" Then
        '        ImageName = "south"
        '    ElseIf T Like "*up*" Then
        '        ImageName = "north"
        '    Else
        '        ImageName = "ok"
        '    End If
        'End If

        If MiniButton Then
            'cmd.Picture = MiniButtonImage(LCase(ImageName))
            'cmd.Image = MiniButtonImage(LCase(ImageName))
            cmd.Image = MiniButtonImage(ImageIndex)
        Else
            'cmd.Picture = StandardButtonImage(LCase(ImageName))
            'cmd.Image = StandardButtonImage(LCase(ImageName))
            'cmd.Image = StandardButtonImage(ImageIndex)
            cmd.Image = MainMenu.imlStandardButtons.Images(ImageIndex)
            cmd.ImageAlign = ContentAlignment.MiddleCenter
            cmd.TextAlign = ContentAlignment.BottomCenter
            cmd.TextImageRelation = TextImageRelation.ImageAboveText
        End If
    End Sub

    Public Sub SetButtonImageSmall(ByRef cmd As Button, ByVal ImageIndex As Integer)
        cmd.Image = MainMenu.imlSmallButtons.Images(ImageIndex)
    End Sub

    Public Function MiniButtonImage(ByVal ImageName As String) As StdPicture
        '::::MiniButtonImage
        ':::SUMMARY
        ': Returns a MiniButton image
        ':::DESCRIPTION
        ': Pulls a specified image from the MiniButtonImageList control
        ':::PARAMETERS
        ': - ImageName
        ':::RETURN
        ': StdPicture

        On Error Resume Next
        ImageName = LCase(ImageName)
        'MiniButtonImage = MiniButtonImageList.ListImages(ImageName).Picture
        MiniButtonImage = MiniButtonImageList.Images(ImageName)

        If MiniButtonImage Is Nothing Then
            If IsDevelopment() Then MsgBox("Not a valid mini image name: " & ImageName, vbCritical, "Development Error")
            'MiniButtonImage = MiniButtonImageList.ListImages("none").Picture
            MiniButtonImage = MiniButtonImageList.Images("none")
        End If
    End Function

    Public Function StandardButtonImage(ByVal ImageIndex As Integer) As StdPicture
        '::::StandardButtonImage
        ':::SUMMARY
        ': Used to check whether the StandardButtonImage is Nothing or not.
        ':::DESCRIPTION
        ': This function is used to display the Standard Button Image using ImageName and to check whether the StandardButtonImage is Nothing or not and print the respective message.
        ':::PARAMETERS
        ': - ImageName - Indicates the Image Name.
        ':::RETURN
        ': StdPicture - Returns the result as StdPicture object.

        On Error Resume Next
        'ImageName = LCase(ImageName)
        'StandardButtonImage = StandardButtonImageList.ListImages(ImageName).Picture
        'StandardButtonImage = StandardButtonImageList.Images(ImageName)
        StandardButtonImage = StandardButtonImageList.Images.Item(ImageIndex)

        'If StandardButtonImage Is Nothing Then
        '    If IsDevelopment() Then MsgBox("Not a valid standard image name: " & ImageName, vbCritical, "Development Error")
        '    'StandardButtonImage = StandardButtonImageList.ListImages("none").Picture
        '    StandardButtonImage = StandardButtonImageList.Images("none")
        'End If
    End Function

    Public Function MiniButtonImageList() As ImageList
        '::::MiniButtonImageList
        ':::SUMMARY
        ': Return the ImageList of MiniButton images
        ':::DESCRIPTION
        ': Returns the ImageList control with the MiniButton images for display throughout the software.
        ':::RETURN
        ': ImageList

        MiniButtonImageList = MainMenu.imlMiniButtons
    End Function

    Public Function StandardButtonImageList() As ImageList
        '::::StandardButtonImageList
        ':::SUMMARY
        ': Returns the ImageList of the standard images
        ':::DESCRIPTION
        ': Returns the ImageList control with the StandardImages for display throughout software.
        ':::RETURN
        ': ImageList
        ':::SEE ALSO
        ': MiniButtonImageList
        StandardButtonImageList = MainMenu.imlStandardButtons
    End Function
    '    Public Property Get MainMenu() As MainMenu4
    '        '  If IsCDSComputer Then Set MainMenu = MainMenu4: Exit Property
    '        Set MainMenu = MainMenu4
    'End Property
    Public ReadOnly Property MainMenu() As MainMenu4
        Get
            '  If IsCDSComputer Then Set MainMenu = MainMenu4: Exit Property
            MainMenu = MainMenu4
        End Get
    End Property

    Public Sub InitHotKeys(ByRef CHK As cRegHotKey)
        '::::InitHotKeys
        ':::SUMMARY
        ': Initialize global hot keys, if enabled.
        ':::DESCRIPTION
        ': Hook for main menu to initialize the hot keys.
        ':::PARAMETERS
        ': - CHK - ByRef
        On Error Resume Next
        '  m_cHotKey.RegisterKey "Activate", vbKeyUp, MOD_ALT + MOD_CONTROL
        CHK.RegisterKey("Security Monitor", VBRUN.KeyCodeConstants.vbKeyF9, cRegHotKey.EHKModifiers.MOD_ALT + cRegHotKey.EHKModifiers.MOD_CONTROL)
        '  CHK.RegisterKey "Printers", vbKeyF5, MOD_ALT + MOD_CONTROL
        '  CHK.RegisterKey "Calculator", vbKeyF2, MOD_ALT + MOD_CONTROL
        CHK.RegisterKey("ErrorReport", VBRUN.KeyCodeConstants.vbKeySnapshot, cRegHotKey.EHKModifiers.MOD_WIN)
    End Sub

    Public Sub DoShutDown(Optional ByVal vQuick As Boolean = False)
        '::::DoShutDown
        ':::SUMMARY
        ': Causes the software to shut down (by closing the main menu).
        ':::DESCRIPTION
        ': Shuts down the software immediately from anywhere in the software.
        ':::PARAMETER
        ': - vQuick - If set, performs the operation without confirmation.
        QuickQuit = vQuick
        'Unload MainMenu
        MainMenu.Close()
    End Sub

    Public Sub DoPractice()
        '::::DoPractice
        ':::SUMMARY
        ': Opens developer panel
        ':::DESCRIPTION
        ': Opens the Developer Panel.
        Dim CHK As VbMsgBoxResult, SS As String

        CHK = vbRetry

        If Not CheckAccess("Store Setup") Then Exit Sub

        If (IsDevelopment() Or IsIDE() Or IsCDSComputer()) Then
            CHK = vbOK
        ElseIf SecurityLevel = ComputerSecurityLevels.seclevNoPasswords Then
            SS = InputBox("Enter Password: ", "Developer's Panel", , "*")
            CHK = IIf(Backdoor(SS), vbOK, vbCancel)
        End If

        If CHK = vbRetry Then
            CHK = IIf(IsDevelopment, vbOK, MessageBox.Show("You have entered the Developer's Area.  Please cancel!", "Warning: Entering Developer's Area", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
        End If

        If CHK = vbOK Then
            MainMenu.Hide()
            Practice.Show()
        End If
    End Sub

    Public Sub ShowLicenseAgreement(Optional ByVal ReShow As Boolean = False)
        '::::ShowLicenseAgreement
        ':::SUMMARY
        ': Opens License Agreement
        ': - ReShow - When not set, only shows if not seen before.  If set, will show regardless.
        ':::RETURN
        frmLicenseAgreement.LicenseAgreement(ReShow)
    End Sub

    Public Sub MainMenu_NumberKeys(ByVal KeyCode As Integer)
        '::::MainMenu_NumberKeys
        ':::SUMMARY
        ': Handles the NumberKeys event from the main menu
        ':::DECSRIPTION
        ': Handles an imaginary 'NumberKeys' event from the main menu
        ':
        ': The Main menu should merely call this function and return
        ':::PARAMETERS
        ': KeyCode - Returns the Integer value.
        On Error Resume Next
        If Not IsDevelopment() Then Exit Sub
        ' BFH20090130
        ' This is our development, main-menu quick-launch section.
        ' Put parts of the program you wish to access quickly that you are working on, so you don't have to navigate through sub-menus.
        ' It may be a little redundant (just specify a key-board shortcut in the menu item creation?), but it works.
        Select Case KeyCode

' Do not use "0" as it is 'special' and handled below.
            Case 49 '1
                MainMenu.WebServ = New frmHTTPServer
                MainMenu.WebServ.HTTPPort = 8080
                MainMenu.WebServ.StartHTTP()
                MainMenu.WebServ.Show()               ' is a form
            Case 50 '2
                frmSupportHost.Listen()
            Case 51 '3
                PracticeCommandPromptFunctions.Show()
            Case 52 '4
                MainMenu.Hide()
                frmAshleyEDI888.Show()
            Case 53 '5
                Order = "TABLE-VIEWER"
                frmTableView.Show() 'vbModal
'        Order = ""
            Case 54 '6
                MainMenu.Hide()
                frmAWSAdmin.Show()
            Case 55 '7
                MainMenu.Hide()
                PracticeDiagnostics.Show()
            Case 56 '8
                PracticeCommandPrompt.Show()
            Case 57 '9
                PermissionMonitor(0)
        End Select
    End Sub

    Public Function MainMenu_KeyDown(KeyCode As Integer, Shift As Integer)
        '::::MainMenu_KeyDown
        ':::SUMMARY
        ': Key down handler for main menu
        ':::DESCRIPTION
        ': Contains the handler code for the Key Down event in the main menu.  (Only) The Main Menu should (only) call this function on KeyDown event
        Dim M As String, L As String, VT As Integer
        Dim A As String, B As String

        If FindControlCode(Shift, KeyCode, M, L) Then
            If Not (M = "file" And L = "login") Then     ' we don't zoom to the contextual menu for store login (available from the top always)
                MainMenu.LoadMenuToForm(M)
            End If
            MainMenu.SelectMenuItem(, M, L)
            Exit Function
        End If

        VT = Val(MainMenu.Tag)

        Dim T As String, I As Integer
        T = Format(Shift, "00") & Format(KeyCode, "0000")
        If IsIn(T, "040018", "020017") Then Exit Function
        '  Debug.Print T
        Select Case T
            Case "000027"                                          ' ESC
                MainMenu.MenuItemHighlight(-1, True)
                If MainMenu.ParentMenu = "" Then MainMenu.MainMenuClick(-1)
                MainMenu.LoadMenuToForm(MainMenu.ParentMenu)
                MainMenu.DoLogOut()
            Case "040065" : MainMenu.MainMenuClick(3)                         ' Alt-A   'LoadMenuToForm "file:maintenance"
            Case "040066"                                          ' Alt-B
            Case "040067" : MainMenu.LoadMenuToForm("inventory:po")           ' Alt-C
            Case "040068" : MainMenu.LoadMenuToForm("inventory:deliveries")   ' Alt-D
            Case "040069" : MainMenu.LoadMenuToForm("inventory:ashley")       ' Alt-E
            Case "040070" : MainMenu.MainMenuClick(0)                         ' Alt-F   'LoadMenuToForm "file"
            Case "040071" : MainMenu.LoadMenuToForm("general ledger")         ' Alt-G
            Case "040072"                                          ' Alt-H
            Case "040073" : MainMenu.MainMenuClick(2)                         ' Alt-I   'LoadMenuToForm "inventory"
            Case "040074"                                          ' Alt-J
            Case "040075" : MainMenu.LoadMenuToForm("file:backup")            ' Alt-K
            Case "040076"                                          ' Alt-L
            Case "040077" : MainMenu.MainMenuClick(4)                         ' Alt-M   'LoadMenuToForm "mailing"
            Case "040078" : MainMenu.MainMenuClick(5)                         ' Alt-N   'LoadMenuToForm "installment"
            Case "040079" : MainMenu.MainMenuClick(1)                         ' Alt-O   LoadMenuToForm "order entry"
            Case "040080"                                          ' Alt-P
            Case "040081"                                          ' Alt-Q
            Case "040082" : MainMenu.LoadMenuToForm("file:restore")           ' Alt-R
            Case "040083"                                          ' Alt-S
            Case "040084"                                          ' Alt-T
            Case "040085"                                          ' Alt-U
            Case "040086"                                          ' Alt-V
            Case "040087" : MainMenu.LoadMenuToForm("file:web")               ' Alt-W
            Case "040088"                                          ' Alt-X
            Case "040089"                                          ' Alt-Y
            Case "040090"                                          ' Alt-Z

            Case "000107", "010187"
                Select Case VT
                    Case 0 : MainMenu.Tag = "2"
                    Case 2 : MainMenu.Tag = "4"
                    Case 4
                        VersionControlDialog()
                        MainMenu.Tag = 0
                    Case Else : MainMenu.Tag = ""
                End Select
            Case "000106"
                If VT = 2 Then
                    MainMenu.Tag = ""
                    DoPractice()

                    MainMenu.Hide()
                    Practice.Show()
                End If
            Case "000109"
                If VT = 2 Then
                    MainMenu.Tag = ""
                    If Not CheckAccess("Store Setup") Then Exit Function
                    ProgressForm(0, 1, "Preparing Settings...", , , , ProgressBarStyle.prgSpin)
                    MainMenu.Hide()
                    frmSetup.Show()
                    ProgressForm()
                End If
            Case Else
                MainMenu.Tag = ""
                On Error Resume Next
                If Val(T) >= 65 And Val(T) < 90 Then
                    '        Debug.Print "t=" & T & ", chr=" & LCase(Chr(Val(T)))
                    'NOTE: THIS FOR LOOP REPLACEMENT IS AFTER THESE COMMENTED LINES.
                    'For I = 1 To MainMenu.imgMenuItem.UBound
                    '    A = LCase(Chr(Val(T)))
                    '    B = LCase(MainMenu.ItemOptionHotKeys(MainMenu.imgMenuItem(I).Tag))
                    '    If B <> "" And A = B Then MainMenu.SelectMenuItem I: Exit For
                    '    '            Debug.Print "i=" & I & ", tag=" & LCase(imgMenuItem(I).Tag)
                    '    '            If InStr(LCase(Chr(Val(T))), LCase(.ItemOptionHotKeys(.imgMenuItem(i).Tag))) > 0 Then .SelectMenuItem i: Exit For
                    'Next

                    Dim C As Control
                    For Each C In MainMenu.Controls
                        If Left(C.Name, 11) = "imgMenuItem" Then
                            I = Mid(C.Name, 12)
                            A = LCase(Chr(Val(T)))
                            B = LCase(MainMenu.ItemOptionHotKeys(C.Tag))
                            If B <> "" And A = B Then MainMenu.SelectMenuItem(I) : Exit For
                        End If
                    Next
                End If
        End Select
    End Function

    Public Sub ResetMenus()
        '::::ResetMenus
        ':::SUMMARY
        ': Used to Reset the Menus.
        ':::DESCRIPTION
        ': This function is used to Reset the main menu data, causing it to be re-loaded at next use (clears the cache)
        MyMenusInitialized = False
    End Sub

    Public Sub LaunchProgram(ByVal Which As String)
        '::::LaunchProgram
        ':::SUMMMARY
        ': Used to launch external programs.
        ':::DESCRIPTION
        ': Launches the desired program (froma  list of possibilities).
        ':::PARAMETERS
        ': - Which - Indicates the String value.
        MainMenu.Hide()
        '  ActiveLog "MainMenu.LaunchProgram: " & FileBanking
        Select Case LCase(Which)
            Case "payables" : ShellOut_Shell(MainMenu, FileAccountPayable)
            Case "payroll" : ShellOut_Shell(MainMenu, FilePayroll)
            Case "banking" : ShellOut_Shell(MainMenu, FileBanking)
            Case "general ledger" : ShellOut_Shell(MainMenu, FileGenLedger)
                '   Case "time clock":     ShellOut_Shell mainmenu, FileTimeClock
            Case Else : MessageBox.Show("Could not launch " & Which & vbCrLf & "Please contact " & AdminContactCompany & " at " & AdminContactPhone2 & ".", "Unknown Program", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Select
        MainMenu.Show()
    End Sub

    Public Function GetMyMenu(ByVal Name As String, Optional ByRef Index As Integer = 0) As MyMenu
        '::::GetMyMenu
        ':::SUMMARY
        ': Return menu object by name
        ':::DESCRIPTION
        ': Returns menu object by name.
        ':::PARAMETERS
        ': - Name - Indiactes the Name of MyMenu.
        ': - Index - Byref. Indiactes the Index of MyMenu.
        ':::RETURN
        ': MyMenu
        Dim I As Integer
        InitializeMenus()  ' this will only run once.  Immediately exits on subseq calls

        For I = LBound(MyMenus) To UBound(MyMenus)
            If LCase(MyMenus(I).Name) = LCase(Name) Then GetMyMenu = MyMenus(I) : Index = I : Exit Function
        Next

        Index = -1
    End Function

    Public Function MainMenu_Dispatch(ByVal Source As String, ByVal Operation As String) As Boolean
        '::::MainMenu_Dispatch
        ':::SUMMARY
        ': Dispatches a main menu action.
        ':::DESCRIPTION
        ': When a main menu item is selected (clicking, keyboad, shortcut, etc), this function performs the desired operation in
        ': the software.
        ':::PARAMETERS
        ': - Source - Indicates the name of Source file.
        ': - Operation - Indicates the functionality.
        ':::RETURN
        ':': Boolean - Returns true on success.
        Dim Fail As Boolean, FailMsg As String, FailTitle As String
        Dim UsageStr As String, UsageDsc As String

        FailMsg = "You have encountered a program error or the resource has moved." & vbCrLf & "Please contact " & AdminContactCompany & " at " & AdminContactPhone2 & " immediately." & vbCrLf & "Thank-you, and sorry for the inconvenience." & vbCrLf & "Source=" & Source & vbCrLf & "Operation=" & Operation
        FailTitle = "Unknown Menu Function"

        ClearProgramState()
        UsageStr = "MM - " & Source & " - " & Operation
        UsageDsc = GetOperationCaption(Source, Operation)

        TrackUsage(UsageStr, UsageDsc)


        Select Case Source
            Case "file", "file:system", "file:utilities", "file:maintenance"
                Select Case Operation
                    Case "systemsetup"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        frmSetup.Show()
                        ProgressForm()
''''''''''''''''
                    Case "password"
                        If modStores.SecurityLevel = ComputerSecurityLevels.seclevNoPasswords Then
                            'MsgBox "You cannot set up the password until you have taken the computer out of No Passwords Mode in Store Setup." & vbCrLf &
                            '"Click F1 for Help.", vbExclamation + vbMsgBoxHelpButton, ProgramMessageTitle, App.HelpFile, 31100 '34000
                            MessageBox.Show("You cannot set up the password until you have taken the computer out of No Passwords Mode in Store Setup." & vbCrLf &
                             "Click F1 for Help.", ProgramMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Else
                            PassWord.ChangePassword()
                        End If
                    Case "configw"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        frmScanner.Show()
                    Case "download"
                        MainMenu.Hide()
                        frmScannerDownload.Show()
                    Case "creditcardmanager"
                        If Not CheckAccess("Credit Administration") Then Exit Function
                        MainMenu.Hide()
                        'Load frmCCAdmin
                        'frmCCAdmin.HelpContextID = 47500
                        frmCCAdmin.Show()
                    Case "webupdates"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        frmUpgrade.Show()
                    Case "email"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        frmEmailSetup.Show()

''''''''''''''''
                    Case "quarterly"
                        If Not CheckAccess("Annual Maintenance") Then Exit Function
                        If MessageBox.Show("Caution:  This should only be done at CALENDAR year end!  When you click Ok, Quarterly Sales on your Inventory Data Base gets updated. Unit sales for the current year will be moved to the prior year.  The current year unit sales will be empty!", "", MessageBoxButtons.OKCancel) = DialogResult.Cancel Then Exit Function
                        If MessageBox.Show(" Are You Sure You want To Update Quarterly Inventory? ", "", MessageBoxButtons.YesNo) = DialogResult.No Then Exit Function

                        'MainMenu.MousePointer = vbHourglass
                        MainMenu.Cursor = Cursors.WaitCursor
                        ExecuteRecordsetBySQL("UPDATE [2Data] SET PSales1=Sales1, PSales2=Sales2, " &
                        "PSales3=Sales3, PSales4=Sales4, Sales1=0, Sales2=0, Sales3=0, Sales4=0", , GetDatabaseInventory)
                        'MainMenu.MousePointer = vbDefault
                        MainMenu.Cursor = Cursors.Default

                        MessageBox.Show("Units sales are transferred!", "Quarterly Update Complete!", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Case "restoredel"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        Inven = "R"
                        InvenA.DoInvRestore()
                    Case "racklabels"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        frmRackLabel.Show()
                    Case "loadorig-manual"
                        Inven = "L"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        MainMenu.Hide()
                        InvenA.Show()
                        InvenA.Text = "Loading Original Inventory Quantities"
                        'InvenA.HelpContextID = 37000
                        'Load InvAutoMan
                        'InvAutoMan.HelpContextID = 37000
                        InvAutoMan.Show()
                    Case "loadorig-import"
                        Inven = "L"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        MainMenu.Hide()
                        Reports = "LoadOriginalInvByBarcodes"
                        frmPhysicalInventoryMainMenu.Show()
                        'frmPhysicalInventoryMainMenu.HelpContextID = 37000
                    Case "tags"
                        Inven = "TAGS"
                        'If Not CheckAccess("") Then Exit Sub
                        MainMenu.Hide()
                        frmPrintAllTickets.Show()
                    Case "speech"
                        On Error Resume Next
                        'Load frmSpeech
                        If IsFormLoaded("frmSpeech") Then frmSpeech.Show()
                        MainMenu.Select()
                    Case "ashley"
                        frmAshleyEDI888.Show()
                        MainMenu.Hide()
                    Case "s3"
                        If Not IsServer() Then
                            MessageBox.Show("The Cloud backup/restore feature should only be run from the server.", "You are on a Workstation!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            Exit Function
                        End If
                        frmAWS.Show()
                        MainMenu.Hide()
                    Case "exportitems"
                        frmExport.ShowExport()
                        MainMenu.Hide()
                    Case "importitems"
                        frmExport.ShowImport()
                        MainMenu.Hide()
                    Case "old-reports"
                        frmPDFReports.Show()
                        MainMenu.Hide()
                    Case "login"
                        If Not CheckAccess("Log In To Other Stores") Then Exit Function
                        LogIn.ShowDialog()
                    Case "exit"
                        'Unload MainMenu
                        MainMenu.Close()
                    Case Else : Fail = True
                End Select
            Case "file:backup", "file:restore"
                ''''''''''''''''
                If Not CheckAccess("Backup/Restore") Then Exit Function

                If Not IsServer() Then 'Check if running from server
                    MessageBox.Show("Backups MUST be made from main (Inventory) computer only!", ProgramMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Function
                End If

                If Left(Operation, 7) = "restore" Then
                    If MessageBox.Show("CAUTION: Restoring from Backup will wipe out all transactions after date of files on this disk!", ProgramMessageTitle, MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) = DialogResult.Cancel Then
                        Exit Function
                    End If
                End If

                Inven = "A"

                Select Case Operation
                    Case "backuppos" : frmBackUpGeneric.Display(BackupMode.bkBackup, BackupType.bkPS, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "backuppayables" : frmBackUpGeneric.Display(BackupMode.bkBackup, BackupType.bkAP, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "backuppayroll" : frmBackUpGeneric.Display(BackupMode.bkBackup, BackupType.bkPR, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "backupbanking" : frmBackUpGeneric.Display(BackupMode.bkBackup, BackupType.bkBK, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "backupgl" : frmBackUpGeneric.Display(BackupMode.bkBackup, BackupType.bkGL, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "backuppx" : frmBackUpGeneric.Display(BackupMode.bkBackup, BackupType.bkpx, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "backupall" : frmBackUpGeneric.Display(BackupMode.bkBackup, BackupType.bkAll, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "restorepos" : frmBackUpGeneric.Display(BackupMode.bkRestore, BackupType.bkPS, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "restorepayables" : frmBackUpGeneric.Display(BackupMode.bkRestore, BackupType.bkAP, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "restorepayroll" : frmBackUpGeneric.Display(BackupMode.bkRestore, BackupType.bkPR, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "restorebanking" : frmBackUpGeneric.Display(BackupMode.bkRestore, BackupType.bkBK, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "restoregl" : frmBackUpGeneric.Display(BackupMode.bkRestore, BackupType.bkGL, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "restoress" : frmBackUpGeneric.Display(BackupMode.bkRestore, BackupType.bkSS, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case "restorepx" : frmBackUpGeneric.Display(BackupMode.bkRestore, BackupType.bkpx, MainMenu, VBRUN.FormShowConstants.vbModal)
                    Case Else : Fail = True
                End Select

            Case "file:web"
                If CrippleBug("Online Sales") Then Exit Function
                If Not CheckAccess("Store Setup") Then Exit Function
                Select Case Operation
                    Case "webgen"
                        MainMenu.Hide()
                        'Load frmAutoWeb
                        frmAutoWeb.FOwner = MainMenu
                        frmAutoWeb.Show()
                        'frmAutoWeb.HelpContextID = 37060
                    Case "webcsv"
                        'Load frmAutoWeb
                        frmAutoWeb.FOwner = MainMenu
                        Dim F As Object
                        F = frmAutoWeb.BuildCSV
                        'MsgBox "CSV Update Complete!" & vbCrLf & "File written to " & F & ".", vbExclamation, "Update CSV", App.HelpFile, 37070
                        MessageBox.Show("CSV Update Complete!" & vbCrLf & "File written to " & F & ".", "Update CSV", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        If Not FormIsLoaded("frmAutoWeb") Is Nothing Then frmAutoWeb.Close()
                        'frmAutoWeb.HelpContextID = 37070
                    Case "webopenmonitor"
                        MainMenu.Hide()
                        'Load frmAutoInv
                        frmAutoInv.FOwner = MainMenu
                        frmAutoInv.Show()
                        'frmAutoInv.HelpContextID = 37080
                    Case "webopensite"
                        On Error Resume Next
                        ShellOut_URL(GetConfigTableValue("Website", WebDemoURL))
                    Case Else : Fail = True
                End Select

            Case "order entry", "order entry:reports"
                Select Case Operation
                    Case "login"
                        If Not CheckAccess("Log In To Other Stores") Then Exit Function
                        LogIn.ShowDialog()
                    Case "newsale"
                        If CrippleBug("New Sales") Then Exit Function
                        If Not CheckAccess("Create Sales") Then Exit Function
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
                        MailCheck.optTelephone.Checked = True
                        MailCheck.HidePriorSales = True
                        MailCheck.ShowDialog()  ' If this is loaded "vbModal, BillOSale", lockup may occur.
                        MailCheck.HidePriorSales = False
                        'Unload MailCheck
                        MailCheck.Close()
                    Case "cashreg"
                        If CrippleBug("New Sales") Then Exit Function
                        If StoreSettings.bManualBillofSaleNo Then
                            MessageBox.Show("You have selected to use Manually entered Bill of Sale numbers." & vbCrLf & "To use the cash register, you must unselect this feature in the store setup.", "Cannot use Cash Register", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            Exit Function
                        End If
                        '          MsgBox "cashreg: -1"
                        If Not CheckAccess("Create Sales") Then Exit Function
                        '          MsgBox "cashreg: 0"
                        MainMenu.Hide()
                        '          MsgBox "cashreg: 1"
                        'Load frmCashRegister
                        '          MsgBox "cashreg: 2"
                        'frmCashRegister.HelpContextID = 42500
                        '          MsgBox "cashreg: 3"
                        frmCashRegister.BeginSale()
                        'frmCashRegister.HelpContextID = 42500
                        Order = "CashRegister"
                    Case "deliver"
                        If CrippleBug("Delivering Sales") Then Exit Function
                        If Not CheckAccess("Deliver Sales") Then Exit Function
                        MainMenu.Hide()
                        Order = "B"
                        BillOSale.Show()
                        'BillOSale.HelpContextID = 43000
                        BillOSale.BillOSale2_Show()
                        'MailCheck.HelpContextID = 43000
                        'MailCheck.optSaleNo.Value = True
                        MailCheck.optSaleNo.Checked = True
                        MailCheck.ShowDialog()  ' If this is loaded "vbModal, BillOSale", lockup may occur.
                        'Unload MailCheck
                        MailCheck.Close()
                    Case "payment"
                        If CrippleBug("Paying on Sales") Then Exit Function
                        If Not CheckAccess("Accept Payments") Then Exit Function
                        MainMenu.Hide()
                        Order = "D"
                        BillOSale.Show()
                        'BillOSale.HelpContextID = 44000
                        BillOSale.BillOSale2_Show()
                        'MailCheck.HelpContextID = 44000
                        'MailCheck.optSaleNo.Value = True
                        MailCheck.optSaleNo.Checked = True
                        MailCheck.ShowDialog()  ' If this is loaded "vbModal, BillOSale", lockup may occur.
                        'Unload MailCheck
                        MailCheck.Close()
                    Case "viewsale"
                        If Not CheckAccess("View Sales") Then Exit Function
                        MainMenu.Hide()
                        Order = "E"
                        'BillOSale.HelpContextID = 45000
                        BillOSale.Show()
                        BillOSale.BillOSale2_Show()
                        'MailCheck.HelpContextID = 45000
                        'MailCheck.optSaleNo.Value = True
                        MailCheck.optSaleNo.Checked = True
                        MailCheck.ShowDialog()  ' If this is loaded "vbModal, BillOSale", lockup may occur.
                        'Unload MailCheck
                        MailCheck.Close()
                    Case "voidsale"
                        If CrippleBug("Voiding Sales") Then Exit Function
                        If Not CheckAccess("Void Sales") Then Exit Function
                        MainMenu.Hide()
                        Order = "C"
                        BillOSale.Show()
                        'BillOSale.HelpContextID = 46000
                        BillOSale.BillOSale2_Show()
                        'MailCheck.HelpContextID = 46000
                        'MailCheck.optSaleNo.Value = True
                        MailCheck.optSaleNo.Checked = True
                        MailCheck.ShowDialog()  ' If this is loaded "vbModal, bos2", lockup may occur.  However, alt-tab can MainMenu.Hide Bos2 if we don't.
                        'Unload MailCheck
                        MailCheck.Close()
                    Case "cashdrawer"
                        If Not CheckAccess("Cash Drawer") Then Exit Function
                        MainMenu.Hide()
                        OrdCashflow.Show()
                    Case "preview"
                        If Not CheckAccess("View Stock Quantities") Then Exit Function
                        MainMenu.Hide()
                        Order = "F"
                        OrdPreview.Show()
                        'InvCkStyle.HelpContextID = 48000
                    Case "adjustments"
                        If CrippleBug("Adjusting Sales") Then Exit Function
                        If Not CheckAccess("Adjust Sales") Then Exit Function
                        MainMenu.Hide()
                        Order = "Credit"
                        'OnScreenReport.HelpContextID = 49000
                        OnScreenReport.CustomerAdjustment()
                    Case "dailyaudit"
                        If Not CheckAccess("Daily Audit Report") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "Audit"
                        OrdAudit2.Show()
                        'OrdAudit2.HelpContextID = 49620
                    Case "undelivered"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "R"
                        ReportPrint.Show()
                        'ReportPrint.HelpContextID = 49630
                    Case "layaway"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "L"
                        ReportPrint.Show()
                        'ReportPrint.HelpContextID = 49640
                    Case "backorder"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "B"
                        ReportPrint.Show()
                        'ReportPrint.HelpContextID = 49650
                    Case "creditsales"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "C"
                        ReportPrint.Show()
                        'ReportPrint.HelpContextID = 49660
                    Case "customerhistory"
                        If Not CheckAccess("View Sales") Then Exit Function
                        MainMenu.Hide()
                        Reports = "H"
                        'MailCheck.HelpContextID = 49670
                        OnScreenReport.CustomerHistory()
                        'OnScreenReport.HelpContextID = 49670
                    Case "salestax"
                        If Not CheckAccess("Store Finances") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "ST"
                        DateForm.Show()
                        'DateForm.HelpContextID = 49680
                    Case "advertising"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "ATR"
                        'DateForm.HelpContextID = 49700
                        DateForm.Show()
                        'DateForm.HelpContextID = 49700
                    Case Else : Fail = True
                End Select

            Case "service"
                If Not CheckAccess("Service Orders") Then Exit Function
                Select Case Operation
                    Case "servicecalls"
                        Order = "S"
                        'MailCheck.HelpContextID = 49500
                        MailCheck.HidePriorSales = True
                        'MailCheck.optTelephone.Value = True
                        MailCheck.optTelephone.Checked = True
                        MailCheck.ShowDialog()
                        MailCheck.HidePriorSales = False
                    Case "damagedstock"
                        Order = "SDam"
                        MainMenu.Hide()
                        'ServiceParts.HelpContextID = 49502
                        ServiceParts.SelectMode(ServiceParts.ServiceForMode.ServiceMode_ForStock, True, True)
                        ServiceParts.Show()
                    Case "partsorders"
                        Order = "SParts"
                        MainMenu.Hide()
                        'ServiceParts.HelpContextID = 49503
                        ServiceParts.SelectMode(ServiceParts.ServiceForMode.ServiceMode_ForStock, True, True)
                        ServiceParts.Text = "Parts Orders Form"
                        ServiceParts.Show()
                    Case "openservicecalls"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Order = "SCR"
                        'ServiceReports.HelpContextID = 49504
                        ServiceReports.Show()
                    Case "openpartsorders"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Order = "SPR"
                        'ServiceReports.HelpContextID = 49505
                        ServiceReports.Show()
                    Case "partsorderbilling"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Order = "SBR"
                        'ServiceReports.HelpContextID = 49506
                        ServiceReports.Show()
                    Case "unpaidbilling"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Order = "SBU"
                        'ServiceReports.HelpContextID = 49507
                        ServiceReports.Show()
                    Case Else : Fail = True
                End Select

            Case "inventory"
                Select Case Operation
                    Case "newitems"
                        If CrippleBug("New Items") Then Exit Function
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        MainMenu.Hide()
                        Inven = "A"
                        On Error Resume Next
                        'Unload InvenA
                        InvenA.Close()
                        InvenA.Show()
                        InvenA.Text = "Adding New Items"
                        InvenA.Desc.Text = "Enter Description Here"
                        'InvenA.HelpContextID = 51000
                        'Load InvCkStyle
                        'InvCkStyle.HelpContextID = 51000
                        InvCkStyle.Show() 'vbModal, InvenA
'        Case "pricechanges"
'          If CrippleBug("Price Changes") Then Exit Sub
'          If Not CheckAccess("Change Item Prices") Then Exit Sub
'          If Not CheckAccess("View Cost and Gross Margin", False, True, False) Then Exit Sub
'          MainMenu.Hide
'          Inven = "B"
'          InvenA.Show
'          InvenA.Caption = "Changing Price Structure"
'          InvenA.HelpContextID = 52000
'          Load InvAutoMan
'          InvAutoMan.HelpContextID = 52000
'          InvAutoMan.Show
                    Case "factoryshipments"
                        If CrippleBug("Factory Shipments") Then Exit Function
                        If Not CheckAccess("Factory Shipments") Then Exit Function
                        If Not CheckAccess("Schedule Deliveries", , True) Then Exit Function
                        MainMenu.Hide()
                        Inven = "D"
                        InvenA.Show()
                        InvenA.Text = "Processing Factory Shipments"
                        'InvenA.HelpContextID = 53000
                        'Load InvCkStyle
                        'InvCkStyle.HelpContextID = 53000
                        InvCkStyle.Show()
                    Case "storetransfers"
                        If CrippleBug("Store Transfers") Then Exit Function
                        If Not CheckAccess("Store Transfers") Then Exit Function
                        If Not CheckAccess("Schedule Deliveries", , True) Then Exit Function
                        MainMenu.Hide()
                        Inven = "T"
                        InvenA.Show()
                        InvenA.Text = "Processing Store Transfers"
                        'InvenA.HelpContextID = 54000
                        'Load InvCkStyle
                        'InvCkStyle.HelpContextID = 54000
                        InvCkStyle.Show()
                    Case "viewstock"
                        MainMenu.Hide()
                        Inven = "E"
                        InvenA.Show()
                        InvenA.Text = "View Any Item"
                        'InvenA.HelpContextID = 55000
                        'Load InvCkStyle
                        'InvCkStyle.HelpContextID = 55000
                        InvCkStyle.Show()
                    Case "changecontents"
                        If CrippleBug("Change Contents") Then Exit Function
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        MainMenu.Hide()
                        Inven = "H"
                        InvenA.Show()
                        InvenA.Text = "Inventory Maintenance"
                        'InvenA.HelpContextID = 56000
                        'Load InvAutoMan
                        'InvAutoMan.HelpContextID = 56000
                        InvAutoMan.Show()
                    Case "orderstatus"
                        If CrippleBug("Sales Adjustments") Then Exit Function
                        If Not CheckAccess("Order Status") Then Exit Function
                        InvOrdStatus.ShowCost = CheckAccess("View Cost and Gross Margin", False, True, False)
                        MainMenu.Hide()
                        InvOrdStatus.CheckDeliveryStatus()
                    Case "comm"
                        If Not CheckAccess("Commissions") Then Exit Function
                        MainMenu.Hide()
                        'InvPayComm.HelpContextID = 59700
                        InvPayComm.Show()
                    Case "ss"
                        MainMenu.Hide()
                        Reports = "VSS"
                        'frmViewSSLoc.HelpContextID = 59990
                        frmViewSSLoc.Show()
                    Case Else : Fail = True
                End Select
            Case "inventory:transfers"
                If CrippleBug("Store Transfers") Then Exit Function
                If Not CheckAccess("Store Transfers") Then Exit Function
                If Not CheckAccess("Schedule Deliveries", , True) Then Exit Function
                Select Case Operation
                    Case "schedule"
                        Inven = "T"
                        frmTransferSelect.Show()
                        MainMenu.Hide()
                    Case "show"
                        Inven = "View Transfer"
                        frmTransferLookup.Show()
                        MainMenu.Hide()
                    Case "reportopen"
                        frmTransferReports.Mode = "Pending"
                        frmTransferReports.Show()
                        'frmTransferReports.HelpContextID = 54030
                        MainMenu.Hide()
                    Case "reportclosed"
                        frmTransferReports.Mode = "Previous"
                        frmTransferReports.Show()
                        'frmTransferReports.HelpContextID = 54040
                        MainMenu.Hide()
                    Case "invdelmulti-trans"
                        InvPull.Pull = 4
                        InvPull.Show()
                        'InvPull.HelpContextID = 59400
                        MainMenu.Hide()
                        '        Case "invdelmulti-lists"
                        '          InvPull.Pull = 5
                        '          InvPull.Show
                        '          InvPull.HelpContextID = 59500
                        '          MainMenu.Hide
                        '        Case "invdelmulti-cross"
                        '          InvPull.Pull = 3
                        '          InvPull.Show
                        '          InvPull.HelpContextID = 59300
                        '          MainMenu.Hide
                    Case Else : Fail = True
                End Select
            Case "inventory:po", "inventory:potrack", "inventory:poorder", "inventory:ashley"
                If Not CheckAccess("Manage Purchase Orders") Then Exit Function
                If Not CheckAccess("View Cost and Gross Margin", False, True, False) Then Exit Function
                Inven = "P"
                MainMenu.Hide()

                Select Case Operation
                    Case "poorder"
                    Case "poeditview"
                        If CrippleBug("PO Changes") Then Exit Function
                        PurchaseOrder = "EDIT"
                        EditPO.Show()
                    Case "poquickprint"
                        Inven = "P"
                        'InvPoPrint.HelpContextID = 57300
                        InvPoPrint.Show()
                    Case "pofaxprint"
                        Inven = "FPO"
                        'FaxPo.HelpContextID = 57400
                        FaxPo.Show()
                    Case "poemail"
                        On Error Resume Next
                        Inven = "EPO"
                        'frmEmail.HelpContextID = 57450
                        frmEmail.Mode = frmEmail.EmailMode.emPO
                        frmEmail.Show()
                    Case "porec"
                        If CrippleBug("PO Changes") Then Exit Function
                        PurchaseOrder = "REC"
                        On Error Resume Next
                        'Load EditPO
                        If PurchaseOrder <> "" Then EditPO.Show()
                    Case "povoid"
                        PurchaseOrder = "Void"
                        On Error Resume Next
                        'Load EditPO
                        If PurchaseOrder <> "" Then EditPO.Show()
                    Case "porepreceiving"
                        Inven = "PRec"
                        'InvPoPrint.HelpContextID = 57700
                        InvPoPrint.Show()
                    Case "porepnotack"
                        Inven = "AK"
                        'InvPoPrint.HelpContextID = 57800
                        InvPoPrint.Show()
                    Case "porepoverdue"
                        Inven = "OverdueOrders"
                        'InvPoPrint.HelpContextID = 57850
                        InvPoPrint.Show()
                    Case "porepnotackE"
                        Inven = "AK-E"
                        'frmPOEmails.HelpContextID = 57805
                        frmPOEmails.sType = 0
                        frmPOEmails.Show()
                    Case "porepoverdueE"
                        Inven = "OverdueOrders-E"
                        'frmPOEmails.HelpContextID = 57855
                        frmPOEmails.sType = 1
                        frmPOEmails.Show()
                    Case "porepopen"
                        PurchaseOrder = "ReOpen"
                        'InvPoPrint.HelpContextID = 57900
                        InvPoPrint.OpenPOReport()
                    Case "poordermanual"
                        Inven = "P"
                        'InvPoSelect.HelpContextID = 57100
                        InvPoSelect.Show()
                    Case "poorderminimum"
                        PurchaseOrder = "OrderMinimum"
                        'InvAutoReOrder.HelpContextID = 57100
                        InvAutoReOrder.Show()
                        InvAutoReOrder.OrderAutomatic(CheckAccess("View Cost and Gross Margin", False, True, False))
                    Case "poorderdemand"
                        PurchaseOrder = "OrderDemand"
                        'InvAutoReOrder.HelpContextID = 57100
                        InvAutoReOrder.Show()
                        InvAutoReOrder.OrderByDemand(StoresSld, Today, DateAdd("d", 7, Today), CheckAccess("View Cost and Gross Margin", False, True, False))
                    Case "pocombine"
                        PurchaseOrder = "POCombine"
                        frmCombinePOs.Show()
                    Case "ashley"
                        Order = "AshleyEDI"
                        frmAshleyEDI.Show()
                    Case "ashleyasn"
                        '          frmAshleyEDIReceive.Mode = "ASN"
                        Order = "AshleyASN"
                        frmAshleyEDIReceive.Show()
                    Case "ashleyopenpo"
                        'InvPoPrint.HelpContextID = 57900
                        Reports = "Ashley"
                        InvPoPrint.OpenPOReport()
                    Case Else : Fail = True
                End Select

            Case "inventory:deliveries"
                If Not CheckAccess("Schedule Deliveries") Then Exit Function
                MainMenu.Hide()

                Select Case Operation
                    Case "invdelpullloads"
                        InvPull.Pull = 1
                        InvPull.Show()
                        'InvPull.HelpContextID = 59100
                    Case "invdeltickets"
                        InvPull.Pull = 2
                        InvPull.Show()
                        'InvPull.HelpContextID = 59200
                    Case "invdelmulti-cross"
                        If Not CheckAccess("View Cost and Gross Margin", False, True, False) Then MainMenu.Show() : 
                        Exit Function
                        InvPull.Pull = 3
                        InvPull.Show()
                        'InvPull.HelpContextID = 59300
'        Case "invdelmulti-trans"
'          InvPull.Pull = 4
'          InvPull.Show
'          InvPull.HelpContextID = 59400
                    Case "invdelmulti-lists"
                        InvPull.Pull = 5
                        InvPull.Show()
                        'InvPull.HelpContextID = 59500
                    Case "invdelcalendar"
                        Inven = "Calendar"
                        Calendar.Show()
                        'Calendar.HelpContextID = 59600
                    Case "invdelpastdeliveries"
                        InvPull.Pull = 6
                        InvPull.Show()
                        'invpull.HelpContextID =0
                    Case Else : Fail = True
                End Select

            Case "inventory:package"
                Select Case Operation
                    Case "invpackmake"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        If Not CheckAccess("View Cost and Gross Margin", False, True, False) Then Exit Function
                        MainMenu.Hide()
                        Reports = "MT"
                        PackagePrice.Show()
                        'PackagePrice.HelpContextID = 59730
                    Case "invpackedit"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        If Not CheckAccess("View Cost and Gross Margin", False, True, False) Then Exit Function
                        MainMenu.Hide()
                        Reports = "ET"
                        PackagePrice.EditPackages()
                        'PackagePrice.HelpContextID = 59740
                    Case "invpacklist"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        If Not CheckAccess("View Cost and Gross Margin", False, True, False) Then Exit Function
                        MainMenu.Hide()
                        Reports = "RT"
                        InvKitReport.Show()
                        'InvKitReport.HelpContextID = 59750
                    Case "invpacklookup"
                        If Not CheckAccess("View Stock Quantities") Then Exit Function
                        MainMenu.Hide()
                        Reports = "CS"
                        'InvKitStock.HelpContextID = 59760
                        InvKitStock.ShowPackages()
                        'InvKitStock.HelpContextID = 59760
                    Case Else : Fail = True
                End Select

            Case "inventory:reports"
                If Not CheckAccess("View Inventory Reports") Then Exit Function
                MainMenu.Hide()

                Select Case Operation
                    Case "invrepinven"
                        Reports = "I"
                        Inven = "IRep"
                        InvReports.Show()
                    Case "invrepmargin"
                        Reports = "M"
                        Inven = "MRep"
                        InvReports.Show()
                    Case "invrepmanuf"
                        Reports = "I"
                        Inven = "ML"
                        'InvPoPrint.HelpContextID = 59930
                        InvPoPrint.Show()
                    Case "invrepbest"
                        Inven = "BS"
                        Reports = "I"
                        'ReportPrint2.HelpContextID = 59940
                        ReportPrint2.Show()
                    Case "invrepdog"
                        Reports = "I"
                        Inven = "DG"
                        'ReportPrint3.HelpContextID = 59970
                        ReportPrint3.Show()
                    Case "invrepss"
                        Reports = "I"
                        Inven = "SS"
                        'ReportPrint3.HelpContextID = 59960
                        ReportPrint3.Show()
                    Case "invrepserialno"
                        Reports = "I"
                        Inven = "SNO"
                        'ReportPrint3.HelpContextID = 59980
                        ReportPrint3.Show()
                    Case "invrepbarcode"
                        Reports = "I"
                        Inven = "PhysicalInventory"
                        frmPhysicalInventoryMainMenu.Show()
                        'frmPhysicalInventoryMainMenu.HelpContextID = 59950
                    Case "designtag"
                        Reports = "DESIGNTAG"
                        frmDesignTag.Show()
                    Case "invrepreturn"
                    Case "storecatalog"
                        MainMenu.Hide()
                        Reports = "Store Catalog"
                        frmCatalog.Show()
                    Case "special"
                        If Not CheckAccess("Change Item Prices") Then Exit Function
                        MainMenu.Hide()
                        Ticket.Show()
                    Case "mini"
                        If CrippleBug("Mini Scanners") Then Exit Function
                        If Not CheckAccess("View Stock Quantities") Then Exit Function
                        Reports = "Mini-Barcode Scanner"
                        BarcodeInventoryCheck()
                    Case Else : Fail = True
                End Select


            Case "accounting"
                Select Case Operation
                    Case "ap"
                        LaunchProgram("payables")
                    Case "pr"
                        LaunchProgram("payroll")
                    Case "bk"
                        LaunchProgram("banking")
                    Case "gl"
                        LaunchProgram("general ledger")
                    Case "qb"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        'Load frmQuickBooksAccountSetup
                        'frmQuickBooksAccountSetup.HelpContextID = 67000
                        frmQuickBooksAccountSetup.Show()
                    Case Else : Fail = True
                End Select

            Case "mailing"
                Select Case Operation
                    Case "add"
                        If Not CheckAccess("Create Sales") Then Exit Function
                        Mail = "ADD/Edit"
                        MainMenu.Hide()
                        BillOSale.Show()
                        'BillOSale.HelpContextID = 71000
                        'MailCheck.HelpContextID = 71000
                        'MailCheck.optTelephone.Value = True
                        MailCheck.optTelephone.Checked = True
                        MailCheck.ShowDialog() ' If this is loaded "vbModal, BillOSale", lockup may occur.
                    Case "print"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        'MailPrint.HelpContextID = 72000
                        MailPrint.Show()
                    Case "export"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        frmMailExport.Show()
                    Case "advert"   ' advertising by Zip
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        CustomerTypes.Show()
                        'CustomerTypes.HelpContextID = 73400
                        MainMenu.Hide()
                    Case "book"
                        If Not CheckAccess("Create Sales") Then Exit Function
                        Mail = "Book"
                        MainMenu.Hide()
                        MailBook.Show()
                    Case "merge"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        Mail = "Fix"
                        MailFix.Show()
                    Case Else : Fail = True
                End Select
            Case "installment", "installment:applications"
                If Not Installment Then Exit Function 'Check for installment module
                AddOnAcc.Typee = ArAddOn_Nil
                ARPaySetUp.AccountFound = ""
                Select Case Operation
                    Case "estimator" 'Payment Estimator
                        If Not CheckAccess("Create Sales") Then Exit Function
                        MainMenu.Hide()
                        ArSelect = "E"
                        ARPaySetUp.Show()
                        'ARPaySetUp.HelpContextID = 81000
                    Case "payview" 'Payment and View
                        If Not CheckAccess("Accept Payments") Then Exit Function
                        MainMenu.Hide()
                        ArSelect = "P"
                        'Unload ArCard  'BFH20061102 - just to be sure
                        'ArCard.HelpContextID = 82000
                        ArCard.GetCustomerAccount()
                    Case "edit" 'Edit Accounts
                        If Not CheckAccess("Credit Administration") Then Exit Function
                        MainMenu.Hide()
                        ArSelect = "Edit"
                        'Load ArCard
                        'ArCard.HelpContextID = 83000
                        ArCard.GetCustomerAccount()
                    Case "void" 'Void Contract
                        If Not CheckAccess("Credit Administration") Then Exit Function
                        MainMenu.Hide()
                        ArSelect = "V"
                        'Load ArCard
                        'ArCard.HelpContextID = 84000
                        ArCard.GetCustomerAccount()
                        ArCard.VoidAccount()
                    Case "oldapps" 'Old Account Setup
                        If Not CheckAccess("Credit Administration") Then Exit Function
                        MainMenu.Hide()
                        ArSelect = "S"
                        BillOSale.Show()
                        'BillOSale.HelpContextID = 85000
                        'Load MailCheck
                        'MailCheck.HelpContextID = 85000
                        'MailCheck.optTelephone.Value = True
                        MailCheck.optTelephone.Checked = True
                        MailCheck.ShowDialog()  ' If this is loaded "vbModal, BillOSale", lockup may occur.
                    Case "creditapps-new" 'Add New Appications
                        ArSelect = "A"
                        MainMenu.Hide()
                        BillOSale.Show()
                        'MailCheck.optTelephone.Value = True
                        MailCheck.optTelephone.Checked = True
                        MailCheck.ShowDialog()
                        'ArApp.HelpContextID = 86100
                    Case "creditapps-old" 'Edit Old Appications
                        ArSelect = "EA"
                        MainMenu.Hide()
                        ArApp.Show()
                        'ArApp.HelpContextID = 86200
                        ArApp.NextEntry()
                    Case "revolving"
                        ' Add new form here
                        'MsgBox "This is where the revolving account management screen goes.", , "Placeholder"
                        ArSelect = "Revolving"
                        MainMenu.Hide()
                        frmRevolving.Show()
                    Case Else : Fail = True
                End Select
            Case "installment:reports"
                If Not Installment Then Exit Function 'Check for installment module
                MainMenu.Hide()

                Select Case Operation
                    Case "monthly" 'Monthly Billing
                        ArSelect = "B"
                        'Load ArReports
                        'ArReports.HelpContextID = 87100
                        ArReports.Show()
                    Case "latecharges" 'Late Charge And Notices
                        ArSelect = "L"
                        'Load ArReports
                        'ArReports.HelpContextID = 87200
                        ArReports.Show()
                    Case "aging" 'A/R Ageing Reports
                        ArSelect = "A"
                        'Load ArReports
                        'ArReports.HelpContextID = 87300
                        ArReports.Show()
                    Case "wholate"    ' Whos Late Report
                        ArSelect = "WHOLATE"
                        'Load ArReports
                        'ArReports.HelpContextID = 87350
                        ArReports.Show()
                    Case "delinquent" 'A/R Delinquent Accounts
                        ArSelect = "D"
                        'Load ArReports
                        'ArReports.HelpContextID = 87400
                        ArReports.Show() 'vbModal
                    Case "trial" 'A/R Trial Balance
                        ArSelect = "T"
                        'Load ArReports
                        'ArReports.HelpContextID = 87500
                        ArReports.Show() 'vbModal
                    Case "newaccounts" 'New Account Report
                        ArSelect = "N"
                        'Load ArReports
                        'ArReports.HelpContextID = 87600
                        ArReports.Show() 'vbModal
                    Case "closedaccounts" 'Closed Account Report
                        ArSelect = "O"
                        'Load ArReports
                        'ArReports.HelpContextID = 87700
                        ArReports.Show() 'vbModal
                    Case "writeoff" 'write off report
                        ArSelect = "W"
                        'Load ArReports
                        'ArReports.HelpContextID = 87710
                        ArReports.Show() 'vbModal
                    Case "repo"  ' Repo Report
                        ArSelect = "R"
                        'Load ArReports
                        'ArReports.HelpContextID = 87720
                        ArReports.Show()
                    Case "legal" ' Legal Report
                        ArSelect = "LG"
                        'Load ArReports
                        'ArReports.HelpContextID = 87730
                        ArReports.Show()
                    Case "losscombo" 'Loss Combo = Write Off, Repo, Legal, Bankruptcy
                        ArSelect = "LC"
                        'Load ArReports
                        'ArReports.HelpContextID = 87710
                        ArReports.Show()
                    Case "nonpayment"
                        ArSelect = "NP"
                        'Load ArReports
                        'ArReports.HelpContextID = 87740
                        ArReports.Show()
                    Case "export" 'Export to Credit Bureau
                        ArSelect = "X"
                        'Load ArReports
                        'ArReports.HelpContextID = 87800
                        ArReports.Show() 'vbModal
                    Case "restore"
                        ArSelect = "RestoreAR"
                        Dim Xx As String, R As ADODB.Recordset
                        Xx = InputBox("Enter AR Account to be restored:", "Restore Voided Account")
                        If Xx <> "" Then
                            R = GetRecordsetBySQL("SELECT * FROM InstallmentInfo WHERE arno='" & ProtectSQL(Xx) & "'")
                            If R.RecordCount = 0 Then
                                MessageBox.Show("Account No [" & Xx & "] does not exist.", "No Such Account", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            Else
                                If R("Status").Value <> "V" Then
                                    MessageBox.Show("Account No [" & Xx & "] is not void.", "Invalid Account", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                Else
                                    ExecuteRecordsetBySQL("UPDATE InstallmentInfo SET Status='O' WHERE ArNo='" & ProtectSQL(Xx) & "'")
                                    AddNewARTransactionExisting(StoresSld, Xx, , arPT_stReO, 0, 0, "Restored Voided Account: " & GetCashierName)
                                    MessageBox.Show("Account [" & Xx & "] restored.  Status set to 'Open'.", "Account Restored", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                End If
                            End If
                        End If
                        DisposeDA(R)
                        MainMenu.Show()
                    Case Else : Fail = True
                End Select
            Case Else : Fail = True
        End Select

        If Fail Then
            MessageBox.Show(FailMsg, FailTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            MainMenu.Show()
        End If

        TrackUsage(UsageStr)
    End Function

    Public Function MainMenu_NumberKeys_DeveloperEx() As String
        '::::MainMenu_NumberKeys_DeveloperEx
        ':::SUMMARY
        ': DeveloperEx function for Number Keys
        Dim S As String
        S = ""
        S = S & "1 - frmHTTPServ"
        S = S & "2 - MainMenu2"
        S = S & "3 - MainMenu3"
        S = S & "4 - frmAshleyEDI888"
        S = S & "5 - frmTableViewer"
        S = S & "6 - AWS Admin"
        S = S & "7 - frmAWS (Amazon)"
        S = S & "8 - CDS CMD"
        S = S & "9 - PermissionMonitor 0"

        MainMenu_NumberKeys_DeveloperEx = S
    End Function

    Public Sub MainMenu_Maintain_Timer()
        '::::MainMenu_Maintain_Timer
        ':::SUMMARY
        ': Main Menu scheduled event timer event.
        ':::DESCRIPTION
        ': Should be called by the timer running on the Main Menu.
        ':
        ': Currently hit every 10 seconds.
        ':
        ': 20060303, 20060710
        ': As it is currently, the nightly maintenance schedule is as following:
        ': Between 1:00a-2:30a
        ': Backup to Cloud Storage, if available.
        ': Between 3:00a-4:00a
        ': Reboot the Software...  Clears up possibility of memory leaks, etc.  Fresh start every day.
        ': Between 4:05a-6:30a
        ': Look for Software updates.  This could, of course, restart the software again.
        Dim Msg As String, FileList As String
        On Error Resume Next

        'NOTE: for all of these TRUE is normal run, FALSE is debugging

        ''''''''''''''''''''''''' AUTO-SHUTDOWN (Part of update)
#If True Then
        If ShutdownSemaforeFile(ItExists:=True) Then DoShutDown(True) : Exit Sub
#End If

        ''''''''''''''''''''''''' AUTO-SHUTDOWN (Part of update)
#If True Then
        If BackupSemaforeFile(ItExists:=True) Then
            Domain_exit()
        End If
#End If

        ''''''''''''''''''''''''' DEVELOPMENT (To make IDE not jump into here while in Debug Mode)

        ' Robert: Haven't figured out this bypass part yet 4/11/2017

        ' True bypasses this "timer kill", keeping the maintenance check running
        ' False allows IsDevelopment() to break this
#If Not False Then
        If IsIDE() Then   ' nice saftey in case we leave this turned off
            MainMenu.tmrMaintain.Enabled = False
            Exit Sub
        End If
#End If

        ''''''''''''''''''''''''' PAYPAL ORDER CHECKS... EVERY N MINUTES
#If True Then
        If DateAfter(Now, DateAdd("n", 7, LastPayPalCheck), , "n") Then
            modPayPal.PayPalCheckSales()  ' exits immediately if there is no setup info
            LastPayPalCheck = Now
        End If
#End If

        ''''''''''''''''''''''''' AUTO-UPDATE
        '  If IsServer Then
        If Not DidUpdate Then
#If True Then
            If DateBetween(TimeValue(Now), DateAdd("n", mUpdateInstance, #4:05:00 AM#), #6:30:00 AM#, , "s") Then
#Else
      ' this one is for development purposes...
    If DateBetween(TimeValue(Now), #1:00:00 PM#, #9:00:00 PM#, , "s") Then
#End If
                DidUpdate = True
                'Load frmUpgrade
                frmUpgrade.DoSilentUpdate(Msg, FileList)
                'Unload frmUpgrade
                frmUpgrade.Close()
                Exit Sub      ' exit here in case executable is being reset
            End If
        End If
        '  End If


        ''''''''''''''''''''''''' AUTO-MENU-CLEAR
#If True Then
        If DateDiff("s", MainMenu.LastMouseMove, Now) > 300 Then   ' 5 minutes to clear
#Else
  If DateDiff("s", LastMouseMove, Now) > 15 Then    ' 15 seconds to clear
#End If
            If Not IsYourFurnitureStore() Then
                MainMenu.LastMouseMove = Now
                MainMenu.LoadMenuToForm("")
            End If
        End If

#If True Then
        If LastLoginExpired And MainMenu.cmdLogout.Visible Then modPasswords.LogOut()
#End If


        ''''''''''''''''''''''''' AUTO-BACKUP
        If AWS_AutoBackup() Then
            'Debug.Print "Checking on Auto-Backup at " & FormatDateTime(Now, vbShortTime)
#If True Then
            If DateBetween(TimeValue(Now), #1:05:00 AM#, #2:00:00 AM#, , "s") Then
#Else
      ' this one is for development purposes...
    If DateBetween(TimeValue(Now), #1:00:00 PM#, #5:00:00 PM#, , "s") Then
#End If
                If Not DidAutoBackup Then
                    '        Debug.Print "Attempting backup at " & FormatDateTime(Now, vbShortTime)
                    DidAutoBackup = True
                    DoAmazonAutoBackup()
                Else
                    '        Debug.Print "Already backed up at " & FormatDateTime(Now, vbShortTime)
                End If
            Else
                ' Reset the check after the backup window closes
                '     This is technically not necessary, but is left in for good measure.
                '     The nightly reboot between 3a and 4a would accomplish this otherwise.
                DidAutoBackup = False
                '      Debug.Print "Not backing up at " & FormatDateTime(Now, vbShortTime)
            End If
        End If

#If True Then
        If LastLoginExpired And MainMenu.cmdLogout.Visible Then modPasswords.LogOut()
#End If

        '''''''''''''''''''''''''  AUTO-BENCHMARK
#If True Then
        RecordWinCDSBenchmark()
#End If

        '''''''''''''''''''''''''  INTERNET MONITOR
#If True Then
        MonitorInternet()
#End If

        ''''''''''''''''''''''''' AUTO-RESTART
#If True Then   ' this will only run between 3am and 4am
        If DateBetween(TimeValue(Now), #3:00:00 AM#, #4:00:00 AM#, , "s") Then
#Else           ' this one could be used for debugging, if you comment out the If DateEqual(... below
  If DateBetween(TimeValue(Now), #1:00:00 PM#, #5:00:00 PM#, , "s") Then
#End If
            If DateEqual(Today, DateValue(ProgramStart)) Then Exit Sub      ' if it's already marked as run today, then exit now
            NightlyCleanup()
            'BFH20170427 - Nightly restart disabled upon Jerry's request
#If False Then
    RestartProgram
#End If
        End If
    End Sub

    Public Sub LaunchHelp()
        '::::LaunchHelp
        ':::SUMMARY
        ': Launches the help file.
        ':::DESCRIPTION
        ': Opens whatever CHM is set as the help file on the App object.  Uses whatever method works.

        OpenCHM()

        '  Not reliable...
        '  SendKeys_safe "{F1}"
        '  RunShellExecute "open", App.HelpFile, 0&, 0&, SW_SHOWDEFAULT
        '  OpenHelp Me.KeyCatch.hWnd
    End Sub

    Private Function FindControlCode(ByVal Shift As Integer, ByVal KeyCode As Integer, Optional ByRef Menu As String = "", Optional ByRef Item As String = "") As Boolean
        Dim I As Integer, J As Integer, A As Integer, B As Integer, LL As MyMenu, MM As MyMenuItem
        Dim S As String
        S = ""
        Select Case Shift
            Case 0 : S = S & ""
            Case 1 : S = S & "Shift-"
            Case 2 : S = S & "Ctrl-"
            Case 4 : S = S & "Alt-"
            Case Else : Exit Function
        End Select

        Select Case KeyCode
            Case 8 : S = S & "BkSp"
            Case 45 : S = S & "Insert"
            Case 46 : S = S & "Delete"
            Case 65 To 90 : S = S & Chr(KeyCode)
            Case 112 To 123 : S = S & "F" & (KeyCode - 111)
            Case Else : Exit Function
        End Select

        '  Debug.Print "Key Pressed: " & S

        On Error Resume Next
        A = -1
        A = UBound(MyMenus)
        On Error GoTo 0

        If A >= 0 Then
            For I = 0 To A
                LL = MyMenus(I)
                On Error Resume Next
                B = -1
                B = UBound(LL.Items)
                On Error GoTo 0
                If B >= 0 Then
                    For J = 0 To B
                        MM = LL.Items(J)
                        If MM.ControlCode = S Then
                            Menu = LL.Name
                            Item = MM.Operation
                            FindControlCode = True
                            Exit Function
                        End If
                    Next
                End If
            Next
        End If
    End Function

    Public Sub InitializeMenus(Optional ByVal ExportTaskList As Boolean = False)
        '::::InitializeMenus
        ':::SUMMARY
        ': Initializes the Main Menu
        ':::DESCRIPTION
        ': This function is used to Initialize the main menu items.  Handing is smart, and unless reset, is only run once in the software.

        Dim R() As MyMenu
        Dim X As String
        Dim F As Form

        If MyMenusInitialized Then Exit Sub

        MyMenusInitialized = True
        '  MyMenusInitialized4 = True
        F = MainMenu

        MyMenus = R

        If ExportTaskList Then DeleteFileIfExists(DevOutputFolder() & "MainMenuOptions.txt") : mExportTaskList = ExportTaskList

        On Error Resume Next

        '''''''''''''''''''''  OPENING MENU  '''''''''''''''''''''
        X = AddMyMenu(Name:="", Caption:=ProgramName, ImageSource:=Nothing, HCID:=10000)

        '''''''''''''''''''''  FILE MENU  '''''''''''''''''''''
        X = AddMyMenu(Name:="file", Caption:="File", MenuLayout:=eMyMenuLayouts.eMML_4x3Across, HCID:=30000)
        MyMenuAddItem(X, "system", "S&ystem...", , , 1, "#file:system", "Open this menu to perform system functions such as store setup, backup, and restore.")
        MyMenuAddItem(X, "utilities", "&Utilities...", , , 2, "#file:utilities", "Open this menu to access some of the utilities of WinCDS.")
        MyMenuAddItem(X, "maintenance", "&Maintenance...", , , 3, "#file:maintenance", "Open this menu for some of the maintenance functions of WinCDS.")
        MyMenuAddItem(X, "web", "&Web Development...", , , 4, "#file:web", "Open this menu to access WinCDs's built-in web development system.")
        MyMenuAddItem(X, "setup", "&Store Setup", , , 5, "systemsetup", "This allows you to configure all your store options.", "F12")
        MyMenuAddItem(X, "login", "Store Lo&gin", , , 11, , "Login to other stores", "F2")
        MyMenuAddItem(X, "exit", "E&xit", , , 12, , "Click here to exit WinCDS.", "Alt-X")

        X = AddMyMenu(Name:="file:system", Caption:="File - System", MenuLayout:=eMyMenuLayouts.eMML_4x3Across, ParentMenu:="file", HCID:=30000)
        MyMenuAddItem(X, "systemH", "S&ystem...", , , 1, "#file:system", "Open this menu to perform system functions such as store setup, backup, and restore.")
        MyMenuAddItem(X, "utilities", "&Utilities...", , , 2, "#file:utilities", "Open this menu to access some of the utilities of WinCDS.")
        MyMenuAddItem(X, "maintenance", "&Maintenance...", , , 3, "#file:maintenance", "Open this menu for some of the maintenance functions of WinCDS.")
        MyMenuAddItem(X, "web", "&Web Development...", , , 4, "#file:web", "Open this menu to access WinCDs's built-in web development system.")
        MyMenuAddItem(X, "login", "Store Lo&gin", , , 11, , "Login to other stores", "F2")
        MyMenuAddItem(X, "exit", "E&xit", , , 12, , "Click here to exit WinCDS.", "Alt-X")
        MyMenuAddItem(X, "setup", "&Store Setup", , , 5, "systemsetup", "This allows you to configure all your store options.", "F12")
        MyMenuAddItem(X, "backup", "&Backup...", , , 7, "#file:backup", "Make backups of your database every day!")
        MyMenuAddItem(X, "restore", "&Restore...", , , 8, "#file:restore", "This will wipe out your existing data and replace it with a backup that you specify.")

        X = AddMyMenu(Name:="file:backup", Caption:="File - System", MenuLayout:=eMyMenuLayouts.eMML_4x2x5x5, ParentMenu:="file:system", HCID:=30000, SubTitle1:="Backup Databases", SubTitle2:="Other Files")
        MyMenuAddItem(X, "systemH", "S&ystem...", , , 1, "#file:system", "Open this menu to perform system functions such as store setup, backup, and restore.")
        MyMenuAddItem(X, "utilities", "&Utilities...", , , 2, "#file:utilities", "Open this menu to access some of the utilities of WinCDS.")
        MyMenuAddItem(X, "maintenance", "&Maintenance...", , , 3, "#file:maintenance", "Open this menu for some of the maintenance functions of WinCDS.")
        MyMenuAddItem(X, "web", "&Web Development...", , , 4, "#file:web", "Open this menu to access WinCDs's built-in web development system.")
        '  MyMenuAddItem X, "login", "Store Lo&gin", , , 12, , "Login to other stores", "F2"
        '  MyMenuAddItem X, "exit", "E&xit", , , 16, , "Click here to exit WinCDS.", "Alt-X"
        '  MyMenuAddItem X, "setup", "&Store Setup", , , 2, "systemsetup"
        MyMenuAddItem(X, "backupH", "&Backup...", , , 7, "#file:backup", "Make backups of your database every day!")
        MyMenuAddItem(X, "restore", "&Restore...", , , 8, "#file:restore", "This will wipe out your existing data and replace it with a backup that you specify.")
        MyMenuAddItem(X, "backupitem", "&POS", , , 9, "backuppos", "This backs up your POS database.", "Shift-Insert")
        MyMenuAddItem(X, "backupitem", "&Pictures", , , 10, "backuppx", "This backs up your InventPX folder.")
        MyMenuAddItem(X, "backupitem", "&Everything", , , 11, "backupall", "This backs up all of yoru databases.", "Ctrl-Insert")
        MyMenuAddItem(X, "backupitem", "P&ayables", , , 14, "backuppayables", "This backs up your AP database.")
        MyMenuAddItem(X, "backupitem", "Pa&yroll", , , 15, "backuppayroll", "This backs up your Payroll database.")
        MyMenuAddItem(X, "backupitem", "Ban&king", , , 16, "backupbanking", "This backs up your bank database.")
        MyMenuAddItem(X, "backupitem", "&GL", , , 17, "backupgl", "This backs up your General Ledger database.")


        X = AddMyMenu(Name:="file:restore", Caption:="File - System", MenuLayout:=eMyMenuLayouts.eMML_4x2x5x5, ParentMenu:="file:system", HCID:=30000, SubTitle1:="Restore Databases", SubTitle2:="Other Files")
        MyMenuAddItem(X, "systemH", "S&ystem...", , , 1, "#file:system", "Open this menu to perform system functions such as store setup, backup, and restore.")
        MyMenuAddItem(X, "utilities", "&Utilities...", , , 2, "#file:utilities", "Open this menu to access some of the utilities of WinCDS.")
        MyMenuAddItem(X, "maintenance", "&Maintenance...", , , 3, "#file:maintenance", "Open this menu for some of the maintenance functions of WinCDS.")
        MyMenuAddItem(X, "web", "&Web Development...", , , 4, "#file:web", "Open this menu to access WinCDs's built-in web development system.")
        '  MyMenuAddItem X, "login", "Store Lo&gin", , , 12, , "Login to other stores", "F2"
        '  MyMenuAddItem X, "exit", "E&xit", , , 16, , "Click here to exit WinCDS.", "Alt-X"
        '  MyMenuAddItem X, "setup", "&Store Setup", , , 2, "systemsetup"
        MyMenuAddItem(X, "backup", "&Backup...", , , 7, "#file:backup", "Make backups of your database every day!")
        MyMenuAddItem(X, "restoreH", "&Restore...", , , 8, "#file:restore", "This will wipe out your existing data and replace it with a backup that you specify.")
        MyMenuAddItem(X, "restoreitem", "&POS", , , 9, "restorepos", "This restores your POS database.", "Shift-Delete")
        MyMenuAddItem(X, "restoreitem", "&Store Setup", , , 10, "restoress", "This restores your store settings.", "Alt-BkSp")
        MyMenuAddItem(X, "restoreitem", "&Pictures", , , 11, "restorepx", "This restores your InventPX Folder.")
        MyMenuAddItem(X, "restoreitem", "P&ayables", , , 14, "restorepayables", "This restores your AP database.")
        MyMenuAddItem(X, "restoreitem", "Pa&yroll", , , 15, "restorepayroll", "This restores your Payroll database.")
        MyMenuAddItem(X, "restoreitem", "Ban&king", , , 16, "restorebanking", "This restores your bank database.")
        MyMenuAddItem(X, "restoreitem", "&GL", , , 17, "restoregl", "This restores your General Ledger database.")

        X = AddMyMenu(Name:="file:utilities", Caption:="File - Utilities", MenuLayout:=eMyMenuLayouts.eMML_4x8x8, ParentMenu:="file", HCID:=30000, SubTitle1:="Utilities", SubTitle2:="Services")
        MyMenuAddItem(X, "system", "S&ystem...", , , 1, "#file:system", "Open this menu to perform system functions such as store setup, backup, and restore.")
        MyMenuAddItem(X, "utilitiesH", "&Utilities...", , , 2, "#file:utilities", "Open this menu to access some of the utilities of WinCDS.")
        MyMenuAddItem(X, "maintenance", "&Maintenance...", , , 3, "#file:maintenance", "Open this menu for some of the maintenance functions of WinCDS.")
        MyMenuAddItem(X, "web", "&Web Development...", , , 4, "#file:web", "Open this menu to access WinCDs's built-in web development system.")
        '  MyMenuAddItem X, "login", "Store Lo&gin", , , 12, , "Login to other stores", "F2"
        '  MyMenuAddItem X, "exit", "E&xit", , , 16, , "Click here to exit WinCDS.", "Alt-X"
        ' First sub-set begins @ 5
        MyMenuAddItem(X, "webupdates", "Manually C&heck for Updates", , , 5, , "The program should periodically check for updates automatically.  Click here to manually check.")
        MyMenuAddItem(X, "password", "&Password" & vbCrLf & "Setup", , , 6, , "This will set up your password system.")
        MyMenuAddItem(X, "configw", "Config. W&ireless" & vbCrLf & "Scanner", , , 7, , "This will open a form that will allow you to configure your CipherLAB scanner.")
        MyMenuAddItem(X, "download", "&Download to" & vbCrLf & "Scanner", , , 8, , "This will download your current item catalog to your CipherLAB scanner.")
        MyMenuAddItem(X, "speech", "&Speech Recognition", , , 9, , "Enable Speech Recognition.")
        ' Second sub-set begins @ 13
        MyMenuAddItem(X, "old-reports", "Report Archi&ve", , , 13, , "View historical reports.")
        MyMenuAddItem(X, "email", "Email Config/Pane&l", , , 14, , "This will open the email settings panel for sending email from WinCDS.")
        MyMenuAddItem(X, "creditcardmanager", "C&redit Card/Manager", , , 15, , "Click here to enter the credit card admin panel.", "Ctrl-L")
        MyMenuAddItem(X, "s3", "Ama&zon AWS", , , 16, , "Amazon AWS Simple Storage Solution Automatic Cloud-based Backup and Restore.")
        MyMenuAddItem(X, "ashley", "&Ashley Maintenance", , , 17, , "Opens the Ashley Vendor Price and Item Maintenance Panel for 888 Item Alignment.")
        MyMenuAddItem(X, "loadorig-manual", "Export Invent&ory", , , 10, "exportitems", "This will export the inventory to a program such as Excel.")
        MyMenuAddItem(X, "racklabels", "Import I&nventory", , , 11, "importitems", "This will import the inventory items from a program such as Excel.")

        X = AddMyMenu(Name:="file:maintenance", Caption:="File - Maintenance", MenuLayout:=eMyMenuLayouts.eMML_4x8x8, ParentMenu:="file", HCID:=30000, SubTitle1:="Maintenance", SubTitle2:="Maintenance")
        MyMenuAddItem(X, "system", "S&ystem...", , , 1, "#file:system", "Open this menu to perform system functions such as store setup, backup, and restore.")
        MyMenuAddItem(X, "utilities", "&Utilities...", , , 2, "#file:utilities", "Open this menu to access some of the utilities of WinCDS.")
        MyMenuAddItem(X, "maintenanceH", "&Maintenance...", , , 3, "#file:maintenance", "Open this menu for some of the maintenance functions of WinCDS.")
        MyMenuAddItem(X, "web", "&Web Development...", , , 4, "#file:web", "Open this menu to access WinCDs's built-in web development system.")
        '  MyMenuAddItem X, "login", "Store Lo&gin", , , 12, , "Login to other stores", "F2"
        '  MyMenuAddItem X, "exit", "E&xit", , , 16, , "Click here to exit WinCDS.", "Alt-X"
        MyMenuAddItem(X, "quarterly", "Annual U&pdate For/Quarterly Sales", , , 5, , "Run this report to clear your quarterly sales totals at the beginning of the year.")
        MyMenuAddItem(X, "restoredel", "Restore &Deleted/Records", , , 6, , "This tool allows you to restore an item you deleted from your item list but would like to restore.")
        MyMenuAddItem(X, "racklabels", "Pr&int Rack" & vbCrLf & "Labels", , , 7, , "This will allow you to print rack labels.")
        MyMenuAddItem(X, "loadorig-manual", "&Enter Quantities/Manually", , , 8, , "This will let you load your original inventory manually.")
        MyMenuAddItem(X, "loadorig-import", "&Import Styles From/Barcode Scanner", , , 13, , "This will let you load your original inventory via your barcode scanner.")
        MyMenuAddItem(X, "tags", "Print Al&l Tags", , , 14, , "Use this to print tags for all items in your database (such as when you first load your inventory).")

        X = AddMyMenu(Name:="file:web", Caption:="File - Web Development", MenuLayout:=eMyMenuLayouts.eMML_4x8x8, ParentMenu:="file", HCID:=30000, SubTitle1:="Web Development")
        MyMenuAddItem(X, "system", "S&ystem...", , , 1, "#file:system", "Open this menu to perform system functions such as store setup, backup, and restore.")
        MyMenuAddItem(X, "utilities", "&Utilities...", , , 2, "#file:utilities", "Open this menu to access some of the utilities of WinCDS.")
        MyMenuAddItem(X, "maintenance", "&Maintenance...", , , 3, "#file:maintenance", "Open this menu for some of the maintenance functions of WinCDS.")
        MyMenuAddItem(X, "webH", "&Web Development...", , , 4, "#file:web", "Open this menu to access WinCDs's built-in web development system.")
        '  MyMenuAddItem X, "login", "Store Lo&gin", , , 12, , "Login to other stores", "F2"
        '  MyMenuAddItem X, "exit", "E&xit", , , 16, , "Click here to exit WinCDS.", "Alt-X"
        MyMenuAddItem(X, "webgen", "&Generate Web Site", , , 5, , "This will open a form to guide you through building or updating your WinCDS generated website.")
        MyMenuAddItem(X, "webcsv", "&Update CSV", , , 6, , "This will generate a new copy of the CSV file which you can use to update your items on your WinCDS generated website.")
        MyMenuAddItem(X, "webopenmonitor", "&Real-time" & vbCrLf & "Monitor", , , 7, , "This will open a monitoring window to let you automatically process orders if you are using Yahoo for your WinCDS generated website.")
        MyMenuAddItem(X, "webopensite", "&Open Website", , , 8, , "This will attempt to open Internet Explorer to display your current (or the demonstration) website.")

        '''''''''''''''''''''  ORDER ENTRY MENU  '''''''''''''''''''''
        X = AddMyMenu(Name:="order entry", Caption:="Order Entry", MenuLayout:=eMyMenuLayouts.eMML_4x3Across, HCID:=40000)
        MyMenuAddItem(X, "newsale", "&New Sale", , , 1, , "Click here to create a new sale.", "Ctrl-N")
        MyMenuAddItem(X, "adjustments", "&Adjustments", , , 2, , "Click here to adjust an existing sale.", "Ctrl-A")
        MyMenuAddItem(X, "viewsale", "&View Sale", , , 3, , "Click here to view an existing sale.", "Ctrl-V")
        MyMenuAddItem(X, "payment", "&Payment on Account", , , 5, , "Click here to make a payment on a sale.", "Ctrl-Y")
        MyMenuAddItem(X, "deliver", "&Deliver Sales", , , 6, , "Click here to deliver a sale.", "Ctrl-D")
        MyMenuAddItem(X, "voidsale", "Vo&id Sale", , , 7, , "Click here to void an existing sale.", "Ctrl-X")
        MyMenuAddItem(X, "cashdrawer", "&Cash Drawer Entries", , , 8, , "Click here to make entries in your cash drawer.", "Ctrl-H")
        MyMenuAddItem(X, "cashreg", "Cas&h Register", , , 9, , "Open the cash-register panel.", "Ctrl-R")
        MyMenuAddItem(X, "servicemodule", "Service &Module...", , , 10, "#service", "This goes to the Service Module.")
        MyMenuAddItem(X, "preview", "Customer &Stock/Preview", , , 11, , "Open the customer preview window for your inventory.", "Ctrl-W")
        MyMenuAddItem(X, "reports", "&Reports...", , , 12, "#order entry:reports", "This goes to the Order Entry Reports menu.")

        X = AddMyMenu(Name:="order entry:reports", Caption:="Order Entry Reports", MenuLayout:=eMyMenuLayouts.eMML_4x8x8, ParentMenu:="order entry", HCID:=49600, SubTitle1:="Other Reports")
        MyMenuAddItem(X, "oereport", "&Daily Audit Report", , , 1, "dailyaudit", "Run the daily audit report.", "F5")
        MyMenuAddItem(X, "oereport", "&Undelivered Sales", , , 2, "undelivered", "Run the Undelivered Sales report.", "F6")
        MyMenuAddItem(X, "oereport", "&Lay-A-Way Report", , , 3, "layaway", "Run the Lay-A-Way report.", "F7")
        MyMenuAddItem(X, "oereport", "&Backorder Sales/Receivables Report", , , 4, "backorder", "Run the Backorder Report.", "F8")
        MyMenuAddItem(X, "oereport", "&Credit Sales", , , 5, "creditsales", "Run the credit sales report.", "Shift-F1")
        MyMenuAddItem(X, "oereport", "Customer &History", , , 6, "customerhistory", "Run the Customer History report.", "Shift-F2")
        MyMenuAddItem(X, "oereport", "Sales Ta&x Report", , , 7, "salestax", "Run the Sales Tax report.", "Shift-F3")
        MyMenuAddItem(X, "oereport", "Adverti&sing Report", , , 8, "advertising", "Run the advertising report.", "Shift-F4")

        '''''''''''''''''''''  SERVICE MODULE MENU  '''''''''''''''''''''
        X = AddMyMenu(Name:="service", Caption:="Service", MenuLayout:=eMyMenuLayouts.eMML_3x8x8, ParentMenu:="order entry", HCID:=49500, SubTitle1:="Service Reports")
        MyMenuAddItem(X, "servicecalls", "&Service Calls", , , 1, , "Create or View Existing Service Calls.", "Ctrl-C")
        MyMenuAddItem(X, "damagedstock", "&Damaged Stock", , , 2, , "Enter the Damaged Stock window.", "Ctrl-I")
        MyMenuAddItem(X, "partsorders", "&Parts Order", , , 3, , "Create or View parts orders.", "Ctrl-B")
        MyMenuAddItem(X, "servicereport", "Open S&ervice/Calls Report", , , 4, "openservicecalls", "Run the Open Service Calls report.")
        MyMenuAddItem(X, "servicereport", "Open P&arts/Orders", , , 5, "openpartsorders", "Run the Open Parts Orders report.")
        MyMenuAddItem(X, "servicereport", "Parts O&rder/Billing Report", , , 6, "partsorderbilling", "Run the Parts Order Billing report.")
        MyMenuAddItem(X, "servicereport", "Unpaid Service Orders", , , 7, "unpaidbilling", "Run the Parts Order Billing report.")

        '''''''''''''''''''''  INVENTORY MENU  '''''''''''''''''''''
        X = AddMyMenu(Name:="inventory", Caption:="Inventory", MenuLayout:=eMyMenuLayouts.eMML_4x3Across, HCID:=50000)
        MyMenuAddItem(X, "newitems", "&New Items", , , 1, , "Create a new inventory item.")
        '  MyMenuAddItem X, "pricechanges", "Price &Changes", , , 2, , "Change prices of items."
        MyMenuAddItem(X, "changecontents", "In&ventory/Maintenance", , , 2, , "This allows you to edit an inventory item.", "Ctrl-G")
        MyMenuAddItem(X, "factoryshipments", "&Factory Shipments", , , 3, , "Enter factory shipments.")
        MyMenuAddItem(X, "storetransfers", "&Store Transfers...", , , 4, "#inventory:transfers", "Enter store transfers.")
        '  MyMenuAddItem X, "storetransfers", "&Store Transfers", , , 4, , "Enter store transfers."
        MyMenuAddItem(X, "viewstock", "&View Stock/Items", , , 5, , "Displays items, showing pricing (if authorized), stock and on order levels, and an item picture.", "Ctrl-K")
        MyMenuAddItem(X, "deliveries", "Customer &Deliveries...", , , 6, "#inventory:deliveries", "This opens the Customer Deliveries menu.")
        '  MyMenuAddItem X, "changecontents", "Chan&ge Contents/of Item", , , 6, , "This allows you to edit an inventory item.", "Ctrl-G"
        MyMenuAddItem(X, "po", "&Purchase Orders...", , , 7, "#inventory:po", "This opens the Purchase Orders Menu.")
        MyMenuAddItem(X, "orderstatus", "Check Order S&tatus", , , 8, , "This lets you check the status of an order.", "Ctrl-T")
        MyMenuAddItem(X, "package", "Pac&kage Ticket Builder...", , , 9, "#inventory:package", "Opens the Package Ticket Builder Menu...")
        MyMenuAddItem(X, "comm", "Co&mmissions", , , 10, , "Opens the commissions report window.")
        MyMenuAddItem(X, "reports", "&Reports...", , , 11, "#inventory:reports", "Opens the Inventory Reports menu.")


        X = AddMyMenu(Name:="inventory:transfers", Caption:="Inventory - Transfers", MenuLayout:=eMyMenuLayouts.eMML_3x8x8, ParentMenu:="inventory", HCID:=54000, SubTitle1:="Transfer Reports")
        MyMenuAddItem(X, "storetransfers", "Schedule a &Transfer", , , 1, "schedule", "Create a Store Transfer.")
        MyMenuAddItem(X, "viewstock", "T&ransfer And Void", , , 2, "show", "Allows you to complete or cancel scheduled transfers.")
        MyMenuAddItem(X, "reports", "Pen&ding Transfers", , , 4, "reportopen", "Shows open store transfers.")
        MyMenuAddItem(X, "reports", "Pre&vious Transfers", , , 5, "reportclosed", "Shows closed store transfers.")
        MyMenuAddItem(X, "invdelmulti-trans", "Delivery Tickets + Transfer/Stock Bi&lling", , , 6, , "Runs the Transfer Billing report.  This is the billing report for all items transfered through Store Transfers.")
        '  MyMenuAddItem X, "reports", "Pen&ding Transfers", , , 4, "reportopen", "Shows open store transfers."
        '  MyMenuAddItem X, "reports", "Pre&vious Transfers", , , 5, "reportclosed", "Shows closed store transfers."
        '  MyMenuAddItem X, "invdelmulti-trans", "Delivery Tickets + Transfer/Stock Bi&lling", , , 6, , "Runs the Transfer Billing report.  This is the billing report for all items transfered through Store Transfers."

        '  MyMenuAddItem X, "invdelmulti-cross", "Multi-Store C&ross/Selling Billing", , , 7, , "Runs the Cross-Selling report.  This reports on all the sales order items sold from a store other than your current store."
        '  MyMenuAddItem X, "invdelmulti-lists", "Multi-Store Trans&fer/Lists", , , 9, , "Prints the Transfer Lists.  This is a list of all store transfers."

        X = AddMyMenu(Name:="inventory:po", Caption:="Inventory - POs", MenuLayout:=eMyMenuLayouts.eMML_4x2x4x4, ParentMenu:="inventory", HCID:=57000, SubTitle1:="Utilities", SubTitle2:="More Options")
        MyMenuAddItem(X, "poeditview", "&Edit View POs", , , 1, , "Edits or views a Purchase Order.", "Ctrl-P")
        MyMenuAddItem(X, "porec", "Receive &Shipments", , , 2, , "Receives a Purchase Order.", "Ctrl-S")
        MyMenuAddItem(X, "povoid", "&Void Stock POs", , , 4, , "Voids a Purchase Order.", "Ctrl-O")

        MyMenuAddItem(X, "poorder", "&Order for Stock...", , , 5, "#inventory:poorder", "Opens the Order for Stock menu.")
        MyMenuAddItem(X, "poreport", "Order Trac&king...", , , 6, "#inventory:potrack", "Runs the Purchase Order 'Orders Not Acknowledged' report.")
        MyMenuAddItem(X, "ashley", "&Ashley Direct...", , , 8, "#inventory:ashley", "Send POs to Ashley via EDI.")

        MyMenuAddItem(X, "pocombine", "&Combine POs", , , 9, , "Gives you the ability to combine vendor POs.")
        MyMenuAddItem(X, "poreport", "&Receiving Report", , , 10, "porepreceiving", "Runs the Purchase Order receiving report.")

        MyMenuAddItem(X, "poquickprint", "&Quick Print POs", , , 13, , "Prints unprinted POs.")
        '  MyMenuAddItem X, "pofaxprint", "&Fax Print POs", , , 5, , "Faxes or prints POs."
        MyMenuAddItem(X, "poemail", "e&Mail POs", , , 14, , "Emails POs.")


        X = AddMyMenu(Name:="inventory:Ashley", Caption:="Inventory - Ashley POs", MenuLayout:=eMyMenuLayouts.eMML_3x2x4x4, ParentMenu:="inventory:po", HCID:=57000, SubTitle1:="Other")
        MyMenuAddItem(X, "ashley", "&Ashley Direct", , , 1, "ashley", "Send POs to Ashley via Electronic Data Interface (EDI).")
        MyMenuAddItem(X, "ashley", "Ashley Ad&vance/Ship Notice", , , 2, "ashleyasn", "Check for Ashley Advance Shipment Notification")

        MyMenuAddItem(X, "poreport", "&Open Ashley POs", , , 7, "ashleyopenpo", "Display all Open Ashley POs.")
        MyMenuAddItem(X, "return", "Retur&n", , , 6, "#inventory:po", "Return to the Purchase Orders Menu.")


        X = AddMyMenu(Name:="inventory:potrack", Caption:="POs - PO Tracking", MenuLayout:=eMyMenuLayouts.eMML_3x3Across, ParentMenu:="inventory:po", HCID:=57000)
        MyMenuAddItem(X, "poreport", "Orders Not &Acknowledged", , , 1, "porepnotack", "Runs the Purchase Order 'Orders Not Acknowledged' report.")
        MyMenuAddItem(X, "poreport", "Overdue Or&ders", , , 2, "porepoverdue", "Runs the Purchase Order 'Overdue Orders' report.")
        MyMenuAddItem(X, "poreport", "&Open Purchase Orders", , , 3, "porepopen", "Runs the Purchase Order 'Open Orders' report.")
        MyMenuAddItem(X, "poemail", "Email Factor&y/Not Ack", , , 4, "porepnotackE", "Creates POs using the 'Order on Demand' system.")
        MyMenuAddItem(X, "poemail", "Email Facto&ry/Overdue", , , 5, "porepoverdueE", "Return to the Purchase Orders Menu.")
        MyMenuAddItem(X, "return", "Retur&n", , , 9, "#inventory:po", "Return to the Purchase Orders Menu.")


        X = AddMyMenu(Name:="inventory:poorder", Caption:="POs - Order for Stock", MenuLayout:=eMyMenuLayouts.eMML_3x3Across, ParentMenu:="inventory:po", HCID:=57000)
        MyMenuAddItem(X, "poorder", "M&anually Entry", , , 1, "poordermanual", "Creates POs manually.")
        MyMenuAddItem(X, "poorder", "M&inimun Stk Lvl", , , 2, "poorderminimum", "Create POs based on Minimum Stock Level.")
        MyMenuAddItem(X, "poorder", "&Order By Demand", , , 3, "poorderdemand", "Creates POs using the 'Order on Demand' system.")
        MyMenuAddItem(X, "return", "Retur&n", , , 9, "#inventory:po", "Return to the Purchase Orders Menu.")


        X = AddMyMenu(Name:="inventory:deliveries", Caption:="Customer Deliveries ", MenuLayout:=eMyMenuLayouts.eMML_3x3, ParentMenu:="inventory", HCID:=59000)
        MyMenuAddItem(X, "invdelpullloads", "&Pull Loads", , , 1, , "Prints Pull Loads.", "Ctrl-F7")
        MyMenuAddItem(X, "invdeltickets", "Delivery &Tickets", , , 2, , "Prints Delivery Tickets.", "Ctrl-F8")
        MyMenuAddItem(X, "invdelcalendar", "Delivery &Calendar", , , 3, , "Opens the Delivery Calendar.", "Ctrl-F11")
        MyMenuAddItem(X, "invdelmulti-cross", "Multi-Store C&ross/Selling Billing", , , 7, , "Runs the Cross-Selling report.  This reports on all the sales order items sold from a store other than your current store.")
        '  MyMenuAddItem X, "invdelmulti-trans", "Multi-Store Tra&nsfer/Billing", , , 8, , "Runs the Transfer Billing report.  This is the billing report for all items transfered through Store Transfers."
        MyMenuAddItem(X, "invdelmulti-lists", "Past Deliveries", , , 8, "invdelpastdeliveries", "Past Deliveries report.")
        MyMenuAddItem(X, "invdelmulti-lists", "Multi-Store Tran&sfer/Summary", , , 9, , "Prints the Transfer Lists on Delivered Sales.  This is a Summary List of all store transfers.")


        X = AddMyMenu(Name:="inventory:package", Caption:="Package Ticket Builder", MenuLayout:=eMyMenuLayouts.eMML_3x3Across, ParentMenu:="inventory", HCID:=59730)
        MyMenuAddItem(X, "invpackmake", "M&ake Packages", , , 1, , "Creates a new package.")
        MyMenuAddItem(X, "invpackedit", "&Edit Packages", , , 2, , "Edits a package.")
        MyMenuAddItem(X, "invpacklist", "Master Kit &List", , , 3, , "Views the master kit list.")
        MyMenuAddItem(X, "invpacklookup", "Kits &Inventory/Look Up", , , 4, , "Looks up kits.")
        MyMenuAddItem(X, "return", "Retur&n", , , 9, "#inventory", "Return to the Inventory Menu.")

        X = AddMyMenu(Name:="inventory:reports", Caption:="Inventory - Reports", MenuLayout:=eMyMenuLayouts.eMML_3x8x8, ParentMenu:="inventory", HCID:=59900, SubTitle1:="Other Reports", SubTitle2:="Utilities")
        MyMenuAddItem(X, "inventoryreports", "&Inventory Reports", , , 1, "invrepinven", "Opens the inventory reports panel.", "F3")
        MyMenuAddItem(X, "marginreports", "&Margin Reports", , , 2, "invrepmargin", "Opens the margin reports panel.", "F4")
        MyMenuAddItem(X, "designtag", "Custom Ta&g Designer", , , 3, , "Enters the Custom Tag Designer")
        MyMenuAddItem(X, "invreport", "Manufacturers &List", , , 4, "invrepmanuf", "Prints the manufacturers list.", "Ctrl-F1")
        MyMenuAddItem(X, "invreport", "&Best Sellers/List (20%)", , , 5, "invrepbest", "Prints the Best-Sellers list (top 20%).", "Ctrl-F2")
        MyMenuAddItem(X, "invreport", "&Dog List", , , 6, "invrepdog", "Prints the Worst-Sellers list.", "Ctrl-F3")
        MyMenuAddItem(X, "invreport", "S&pecial-Special Report", , , 7, "invrepss", "Prints the Special-Special report.", "Ctrl-F4")
        MyMenuAddItem(X, "invreport", "Barcode D&ata/Collector", , , 8, "invrepbarcode", "Opens the Barcode Data Collector panel.", "Ctrl-F5")
        MyMenuAddItem(X, "invreport", "Serial No Trac&king", , , 9, "invrepserialno", "Prints the Serial Number Tracking report.", "Ctrl-F6")
        MyMenuAddItem(X, "ss", "SS Loc&ations", , , 12, , "Runs the Special-Special Locations report.")
        MyMenuAddItem(X, "special", "Special &Event/Price Tags", , , 13, , "This lets you generate special price tags.")
        MyMenuAddItem(X, "mini", "M&ini Scanner/Inventory Check", , , 14, , "This runs the Mini Scanner Inventory Check.")
        MyMenuAddItem(X, "storecatalog", "&Store Catalog", , , 15, , "Lets you view or print a store catalog.")
        MyMenuAddItem(X, "return", "Retur&n", , , 19, "#inventory", "Return to the Inventory Menu.")


        '''''''''''''''''''''  ACCOUNTING MENU  '''''''''''''''''''''
        X = AddMyMenu(Name:="accounting", Caption:="Accounting", MenuLayout:=eMyMenuLayouts.eMML_4x2x4x4, HCID:=50000, SubTitle1:="")
        MyMenuAddItem(X, "qb", "&QuickBooks", , , 1, , "Opens the Quick Books Interface Panel.")
        MyMenuAddItem(X, "gl", "&GenLedgr", , , 5, , "Open the General Ledger Module")
        MyMenuAddItem(X, "ap", "&Payables", , , 6, , "Open the Accounts Payable Module")
        MyMenuAddItem(X, "pr", "Pa&yroll", , , 7, , "Open the Payroll Module")
        MyMenuAddItem(X, "bk", "&Banking", , , 8, , "Open the Banking Module")

        '''''''''''''''''''''  MAILING MENU  '''''''''''''''''''''
        X = AddMyMenu(Name:="mailing", Caption:="Mailing", MenuLayout:=eMyMenuLayouts.eMML_3x3Across, HCID:=70000)
        MyMenuAddItem(X, "add", "&Add Edit Names", , , 1, , "Add or Edit Names from your mailing list.")
        MyMenuAddItem(X, "merge", "&Merge Addresses", , , 2, , "Search the mailing database for duplicates and errors.")
        MyMenuAddItem(X, "book", "Address Boo&k", , , 3, , "Opens the address book.", "Ctrl-Q")
        MyMenuAddItem(X, "export", "&Export Mailing List", , , 4, , "Export your mailing list to one of the available popular formats.")
        MyMenuAddItem(X, "print", "&Print Mailing Labels", , , 5, , "Print mailing labels.")
        MyMenuAddItem(X, "advert", "Advertising by &Zip", , , 6, , "Prints the advertising report by zip code.")

        '''''''''''''''''''''  INSTALLMENT MENU  '''''''''''''''''''''
        X = AddMyMenu(Name:="installment", Caption:="Installment", MenuLayout:=eMyMenuLayouts.eMML_3x4Across, HCID:=80000)
        MyMenuAddItem(X, "estimator", "Payment &Estimator", , , 1, , "Opens the Payment Estimator.")
        MyMenuAddItem(X, "payview", "&Payment and View", , , 2, , "Opens an account to view or make a payment.")
        MyMenuAddItem(X, "edit", "E&dit Accounts", , , 3, , "Edit an existing account.")
        MyMenuAddItem(X, "void", "&Void Accounts", , , 4, , "Void an existing account.")
        MyMenuAddItem(X, "oldapps", "&Old Account Setup", , , 5, , "Set up an old account.")
        MyMenuAddItem(X, "creditapps-old", "Credit Apps...", , , 7, "#installment:applications", "Open the Credit Applications Menu.")
        '  MyMenuAddItem X, "creditapps-new", "&Add New Credit App", , , 6, , "Enter a new credit application."
        '  MyMenuAddItem X, "creditapps-old", "Edi&t Old Credit App", , , 7, , "Edit an existing credit application."
        MyMenuAddItem(X, "revolving", "&Interest Charges and Notices", , , 8, , "Manage revolving accounts.", , ModifiedRevolvingChargeEnabled)
        MyMenuAddItem(X, "reports", "Installment &Reports...", , , 9, "#installment:reports", "Goes to the Installment Reports menu.")

        X = AddMyMenu(Name:="installment:applications", Caption:="Installment - Applications", MenuLayout:=eMyMenuLayouts.eMML_3x4Across, ParentMenu:="installment", HCID:=80000)
        MyMenuAddItem(X, "creditapps-new", "&Add New Credit App", , , 1, , "Enter a new credit application.")
        MyMenuAddItem(X, "creditapps-old", "Edi&t Old Credit App", , , 2, , "Edit an existing credit application.")
        MyMenuAddItem(X, "return", "Retur&n", , , 9, "#installment", "Return to the Installment Menu.")

        X = AddMyMenu(Name:="installment:reports", Caption:="Installment Reports", MenuLayout:=eMyMenuLayouts.eMML_4x8x8, ParentMenu:="installment", HCID:=87000, SubTitle1:="Installment Reports", SubTitle2:="Installment Reports")
        MyMenuAddItem(X, "report-monthly", "Monthly Sta&tements", , , 1, "monthly", "Runs the Installment Monthly Billing report.")
        MyMenuAddItem(X, "report-latecharges", "&Late Charges/and Notices", , , 2, "latecharges", "Generates the Installment late charges and notices.")
        MyMenuAddItem(X, "report-aging", "AR &Aging Report", , , 3, "aging", "Runs the Installment Aging report.")
        MyMenuAddItem(X, "report-delinquent", "AR &Delinquent/Accounts", , , 4, "delinquent", "Runs the Installent Delinquent Accounts report.")
        MyMenuAddItem(X, "report-new", "&New Account/Report", , , 5, "newaccounts", "Runs the Installment New Accounts report.")
        MyMenuAddItem(X, "report-wholate", "Who&s Late", , , 6, "wholate", "Runs the report to show who is late.")
        MyMenuAddItem(X, "report-trial", "AR &Trial Balance", , , 7, "trial", "Runs the Installment Trial Balance report.")
        MyMenuAddItem(X, "report-legal", "Status Report", , , 8, "losscombo", "Runs the Installment Write Off, Repo, Legal, and Bankruptcy report.")
        'MyMenuAddItem X, "report-writeoff", "Write Of&f/Report", , , 10, "writeoff", "Runs the Installment Write Off report."
        'MyMenuAddItem X, "report-repo", "&Repo Report", , , 11, "repo", "Runs the Installment Repo report."
        'MyMenuAddItem X, "report-legal", "Le&gal Report", , , 13, "legal", "Runs the Installment Legal report."
        MyMenuAddItem(X, "report-closed", "&Closed Account/Report", , , 13, "closedaccounts", "Runs the Installment Closed Accounts report.")
        MyMenuAddItem(X, "report-nonpayment", "Non Pa&yment/Report", , , 14, "nonpayment", "Shows No Payments since date selected.")
        MyMenuAddItem(X, "export", "E&xport to Credit/Bureau", , , 15, "export", "Exports the Installment data to the Credit Bureau.")
        '  MyMenuAddItem X, "report-revolving", "Revolving Accounts", , , 16, "revolving", "Runs the Revolving Accounts reports.", , ModifiedRevolvingChargeEnabled
        MyMenuAddItem(X, "restore", "Restore Vo&ided/Account", , , 16, "restore", "Restore an account that has been previously voided.")

    End Sub

    Private Function GetOperationCaption(ByVal Source As String, ByVal Operation As String) As String
        Dim TF As MyMenu, N As Integer
        TF = GetMyMenu(Source)
        For N = LBound(TF.Items) To UBound(TF.Items)
            If Operation = TF.Items(N).Operation Then GetOperationCaption = TF.Items(N).Caption : Exit Function
        Next
    End Function

    Private Function AddMyMenu(ByVal Name As String, ByRef Caption As String, Optional ByRef ImageSource As Object = Nothing, Optional ByVal CaptionStyle As eCaptionStyles = eCaptionStyles.eCS_Below, Optional ByVal ImageW As Integer = 50, Optional ByVal ImageH As Integer = 50, Optional ByVal vSP As Integer = 70, Optional ByRef CaptionMargin As Integer = 10, Optional ByVal MenuLayout As eMyMenuLayouts = eMyMenuLayouts.eMML_Manual, Optional ByVal MaskColor As Integer = VBRUN.ColorConstants.vbRed, Optional ByVal Visible As Boolean = True, Optional ByVal ParentMenu As String = "", Optional ByVal HCID As Integer = 0, Optional ByVal SubTitle1 As String = "", Optional ByVal SubTitle2 As String = "") As String
        Dim X As MyMenu
        X.Name = Name
        X.ParentMenu = ParentMenu
        X.Caption = Caption
        X.Layout = MenuLayout
        X.HCID = HCID
        X.Visible = Visible
        X.ImageSource = ImageSource

        X.CaptionStyle = CaptionStyle
        X.CaptionMargin = CaptionMargin
        X.MaskColor = MaskColor

        X.ImageW = ImageW
        X.ImageH = ImageH
        X.vSP = vSP

        X.SubTitle1 = SubTitle1
        X.SubTitle2 = SubTitle2

        Err.Clear()
        On Error Resume Next
        ReDim Preserve MyMenus(UBound(MyMenus) + 1)
        MyMenus(UBound(MyMenus)) = X

        If Err.Number <> 0 Then
            ReDim MyMenus(0)
            MyMenus(0) = X
        End If

        AddMyMenu = Name
    End Function

    Private Sub MyMenuAddItem(ByRef MM As String, ByVal ImageKey As String, ByVal Caption As String, Optional ByVal Left As Integer = -1, Optional ByVal Top As Integer = -1, Optional ByVal Position As Integer = 0, Optional ByVal Operation As String = "", Optional ByVal ToolTipText As String = "", Optional ByVal ControlCode As String = "", Optional ByVal Visible As Boolean = True)
        Dim X As MyMenuItem, IX As Integer, T As MyMenu
        Dim F As StdPicture

        T = GetMyMenu(MM, IX)
        If IX < 0 Then Err.Raise(-1, , "Invalid Menu: " & MM)
        If mExportTaskList Then
            Dim UsageStr As String
            If Operation = "" Then
                UsageStr = UCase("MM - " & T.Name & " - " & ImageKey)
            Else
                UsageStr = UCase("MM - " & T.Name & " - " & Operation)
            End If

            If Mid(Operation, 1, 1) = "#" Then UsageStr = ""
            If UsageStr <> "" Then
                WriteFile(DevOutputFolder() & "MainMenuOptions.txt", " <item name='" & UsageStr & "' usage=0 />")
            End If
        End If

        MenuItemCount = MenuItemCount + 1
        '  SplashProgress 15 + (MenuItemCount * 60 / MenuItemMax)
        '  Debug.Print "MenuItemCount=" & MenuItemCount

        CalculatePosition(T.Layout, Position, T.vSP, Left, Top)


        On Error Resume Next
        If Not (T.ImageSource Is Nothing) Then
            X.Image = T.ImageSource.ListImages(ImageKey).Picture
        End If
        X.HotKeys = GetHotKeys(Caption) ' byref, remove's ampersands (&)
        X.ImageKey = ImageKey
        X.Caption = Caption
        X.Left = Left
        X.Top = Top
        X.Visible = Visible
        X.Operation = IIf(Operation <> "", Operation, ImageKey)
        X.ToolTipText = ToolTipText
        Dim LL As Object
        For Each LL In Split(X.HotKeys, "")
            X.ToolTipText = Replace(X.ToolTipText, LL, "[" & LL & "]", , 1, vbTextCompare)
        Next
        X.ControlCode = ControlCode
        X.IsSubItem = T.Layout = eMyMenuLayouts.eMML_3x8x8 And Position > 3 Or
               T.Layout = eMyMenuLayouts.eMML_4x8x8 And Position > 4 Or
               T.Layout = eMyMenuLayouts.eMML_3x2x4x4 And Position > 6 Or
               T.Layout = eMyMenuLayouts.eMML_4x2x4x4 And Position > 8 Or
               T.Layout = eMyMenuLayouts.eMML_4x2x5x5 And Position > 8

        Err.Clear()
        On Error Resume Next
        ReDim Preserve MyMenus(IX).Items(UBound(MyMenus(IX).Items) + 1)
        MyMenus(IX).Items(UBound(MyMenus(IX).Items)) = X

        If Err.Number <> 0 Then
            ReDim MyMenus(IX).Items(0)
            MyMenus(IX).Items(0) = X
        End If
    End Sub

    Private Sub CalculatePosition(ByVal Strategy As eMyMenuLayouts, ByVal Position As Integer, ByVal Sp As Integer, ByRef Left As Integer, ByRef Top As Integer)
        Dim R As Integer, Across As Boolean

        If Position = 0 Then Exit Sub
        'If MM4 Then Sp = 2400
        If MM4 Then Sp = 150

        Select Case Strategy
            Case eMyMenuLayouts.eMML_2x3
                Select Case Position
                    Case 1 To 3 : Left = MenuItemLeft1of2
                    Case Else : Left = MenuItemLeft2of2
                End Select
                R = 3 : Do While Position > R : Position = Position - R : Loop
                Top = MenuItemTop + Sp * (Position - 1)
            Case eMyMenuLayouts.eMML_2x4
                Select Case Position
                    Case 1 To 4 : Left = MenuItemLeft1of2
                    Case Else : Left = MenuItemLeft2of2
                End Select
                R = 4 : Do While Position > R : Position = Position - R : Loop
                Top = MenuItemTop + Sp * (Position - 1)
            Case eMyMenuLayouts.eMML_2x3Across, eMyMenuLayouts.eMML_2x4Across
                If (Position - 1) Mod 2 = 0 Then Left = MenuItemLeft1of2
                If (Position - 1) Mod 2 = 1 Then Left = MenuItemLeft2of2
                Top = MenuItemTop + Sp * ((Position - 1) \ 2)
            Case eMyMenuLayouts.eMML_3x3
                Select Case Position
                    Case 1 To 3 : Left = MenuItemLeft1of3
                    Case 4 To 6 : Left = MenuItemLeft2of3
                    Case Else : Left = MenuItemLeft3of3
                End Select
                R = 3 : Do While Position > R : Position = Position - R : Loop
                Top = MenuItemTop + Sp * (Position - 1)
            Case eMyMenuLayouts.eMML_3x4
                Select Case Position
                    Case 1 To 4 : Left = MenuItemLeft1of3
                    Case 5 To 8 : Left = MenuItemLeft2of3
                    Case Else : Left = MenuItemLeft3of3
                End Select
                R = 4 : Do While Position > R : Position = Position - R : Loop
                Top = MenuItemTop + Sp * (Position - 1)
            Case eMyMenuLayouts.eMML_3x3Across, eMyMenuLayouts.eMML_3x4Across
                If (Position - 1) Mod 3 = 0 Then Left = MenuItemLeft1of3
                If (Position - 1) Mod 3 = 1 Then Left = MenuItemLeft2of3
                If (Position - 1) Mod 3 = 2 Then Left = MenuItemLeft3of3
                Top = MenuItemTop + Sp * ((Position - 1) \ 3)
            Case eMyMenuLayouts.eMML_4x2
                Select Case Position
                    Case 1 To 2 : Left = MenuItemLeft1of4
                    Case 3 To 4 : Left = MenuItemLeft2of4
                    Case 5 To 6 : Left = MenuItemLeft3of4
                    Case Else : Left = MenuItemLeft4of4
                End Select
                R = 2 : Do While Position > R : Position = Position - R : Loop
                Top = MenuItemTop + Sp * (Position - 1)
            Case eMyMenuLayouts.eMML_4x3
                Select Case Position
                    Case 1 To 3 : Left = MenuItemLeft1of4
                    Case 4 To 6 : Left = MenuItemLeft2of4
                    Case 7 To 9 : Left = MenuItemLeft3of4
                    Case Else : Left = MenuItemLeft4of4
                End Select
                R = 3 : Do While Position > R : Position = Position - R : Loop
                Top = MenuItemTop + Sp * (Position - 1)
            Case eMyMenuLayouts.eMML_4x4
                Select Case Position
                    Case 1 To 4 : Left = MenuItemLeft1of4
                    Case 5 To 8 : Left = MenuItemLeft2of4
                    Case 9 To 12 : Left = MenuItemLeft3of4
                    Case Else : Left = MenuItemLeft4of4
                End Select
                R = 4 : Do While Position > R : Position = Position - R : Loop
                Top = MenuItemTop + Sp * (Position - 1)
            Case eMyMenuLayouts.eMML_4x2Across, eMyMenuLayouts.eMML_4x3Across, eMyMenuLayouts.eMML_4x4Across
                If (Position - 1) Mod 4 = 0 Then Left = MenuItemLeft1of4
                If (Position - 1) Mod 4 = 1 Then Left = MenuItemLeft2of4
                If (Position - 1) Mod 4 = 2 Then Left = MenuItemLeft3of4
                If (Position - 1) Mod 4 = 3 Then Left = MenuItemLeft4of4
                Top = MenuItemTop + Sp * ((Position - 1) \ 4)
            Case eMyMenuLayouts.eMML_3x8x8
                If Position <= 3 Then
                    Top = MenuItemTop
                    Left = Switch(Position = 1, MenuItemLeft1of3, Position = 2, MenuItemLeft2of3, True, MenuItemLeft3of3)
                Else
                    Top = MenuSubItemTop1 + MenuSubItemHeight * ((Position - 4) Mod 8)
                    If Position <= (3 + 8) Then Left = MenuSubItemLeft1of2 Else Left = MenuSubItemLeft2of2
                End If
            Case eMyMenuLayouts.eMML_4x8x8
                If Position <= 4 Then
                    Top = MenuItemTop
                    Left = Switch(Position = 1, MenuItemLeft1of4, Position = 2, MenuItemLeft2of4, Position = 3, MenuItemLeft3of4, True, MenuItemLeft4of4)
                Else
                    Top = MenuSubItemTop1 + MenuSubItemHeight * ((Position - 5) Mod 8)
                    If Position <= (4 + 8) Then Left = MenuSubItemLeft1of2 Else Left = MenuSubItemLeft2of2
                End If
            Case eMyMenuLayouts.eMML_3x2x4x4
                If Position <= 3 Then
                    Top = MenuItemTop
                    Left = Switch(Position = 1, MenuItemLeft1of3, Position = 2, MenuItemLeft2of3, True, MenuItemLeft3of3)
                ElseIf Position <= 6 Then
                    Top = MenuItemTop + Sp
                    Left = Switch(Position = 4, MenuItemLeft1of3, Position = 5, MenuItemLeft2of3, True, MenuItemLeft3of3)
                Else
                    Top = MenuSubItemTop2 + MenuSubItemHeight * ((Position - 7) Mod 4)
                    If Position <= (6 + 4) Then Left = MenuSubItemLeft1of2 Else Left = MenuSubItemLeft2of2
                End If
            Case eMyMenuLayouts.eMML_4x2x4x4
                If Position <= 4 Then
                    Top = MenuItemTop
                    Left = Switch(Position = 1, MenuItemLeft1of4, Position = 2, MenuItemLeft2of4, Position = 3, MenuItemLeft3of4, True, MenuItemLeft4of4)
                ElseIf Position <= 8 Then
                    Top = MenuItemTop + Sp
                    Left = Switch(Position = 5, MenuItemLeft1of4, Position = 6, MenuItemLeft2of4, Position = 7, MenuItemLeft3of4, True, MenuItemLeft4of4)
                Else
                    Top = MenuSubItemTop2 + MenuSubItemHeight * ((Position - 9) Mod 4)
                    If Position <= (8 + 4) Then Left = MenuSubItemLeft1of2 Else Left = MenuSubItemLeft2of2
                End If
            Case eMyMenuLayouts.eMML_4x2x5x5
                If Position <= 4 Then
                    Top = MenuItemTop
                    Left = Switch(Position = 1, MenuItemLeft1of4, Position = 2, MenuItemLeft2of4, Position = 3, MenuItemLeft3of4, True, MenuItemLeft4of4)
                ElseIf Position <= 8 Then
                    Top = MenuItemTop + Sp
                    Left = Switch(Position = 5, MenuItemLeft1of4, Position = 6, MenuItemLeft2of4, Position = 7, MenuItemLeft3of4, True, MenuItemLeft4of4)
                Else
                    Top = MenuSubItemTop2 + MenuSubItemHeight * ((Position - 9) Mod 5)
                    If Position <= (8 + 5) Then Left = MenuSubItemLeft1of2 Else Left = MenuSubItemLeft2of2
                End If
            Case Else 'manual
                ' no change
        End Select
    End Sub

    Private Function GetHotKeys(ByRef Caption As String) As String
        Dim N As Integer
Again:
        N = InStr(Caption, "&")
        If N > 0 Then
            GetHotKeys = GetHotKeys & Mid(Caption, N + 1, 1)
            Caption = Replace(Expression:=Caption, Find:="&", Replacement:="", Count:=1)
            'Caption = Left(Caption, n - 1) & "[" & Mid(Caption, n + 1, 1) & "]" & Mid(Caption, n + 2)
            GoTo Again
        End If
        Caption = Replace(Caption, "&", "")
    End Function

    Private ReadOnly Property MM4() As Boolean
        Get
            MM4 = IsFormLoaded(MainMenuType)
        End Get
    End Property

    Private ReadOnly Property MenuItemLeft1of2() As Integer
        Get
            MenuItemLeft1of2 = Switch(MM4, X2()(0), True, 242)
        End Get
    End Property

    Private Function X2() As Integer()
        Dim A() As Integer
        ReDim A(0 To 1)
        A(0) = 3300
        A(1) = 8000
        X2 = A
    End Function

    Private ReadOnly Property MenuItemLeft2of2() As Integer
        Get
            MenuItemLeft2of2 = Switch(MM4, X2()(1), True, 480)
        End Get
    End Property

    Private ReadOnly Property MenuItemTop() As Integer
        Get
            'MenuItemTop = Switch(MM4, 1000, True, 152)   ---------- NOTE: Replaced 1000 with 70 ------------------
            MenuItemTop = Switch(MM4, 70, True, 152)
        End Get
    End Property

    Private ReadOnly Property MenuItemLeft1of3() As Integer
        Get
            MenuItemLeft1of3 = Switch(MM4, X3()(0), True, 225)
        End Get
    End Property

    Private Function X3() As Integer()
        Dim A() As Integer
        ReDim A(0 To 2)
        A(0) = 4000
        A(1) = 7600
        A(2) = 11200
        X3 = A
    End Function

    Private ReadOnly Property MenuItemLeft2of3() As Integer
        Get
            MenuItemLeft2of3 = Switch(MM4, X3()(1), True, 400)
        End Get
    End Property

    Private ReadOnly Property MenuItemLeft3of3() As Integer
        Get
            MenuItemLeft3of3 = Switch(MM4, X3()(2), True, 575)
        End Get
    End Property

    Private ReadOnly Property MenuItemLeft1of4() As Integer
        Get
            MenuItemLeft1of4 = Switch(MM4, X4()(0), True, 210)
        End Get
    End Property

    Private Function X4() As Integer()
        Dim A() As Integer

        ReDim A(0 To 3)
        'A(0) = 3200
        'A(1) = 6000
        'A(2) = 8700
        'A(3) = 11600

        'NOTE: REPLACED ABOVE VALUES WITH THE BELOW ONE. BECAUSE VB.NET MEASUREMENTS ARE IN PIXELS. NOT IN TWIPS.
        A(0) = 200
        A(1) = 400
        A(2) = 600
        A(3) = 800

        X4 = A
    End Function

    Private ReadOnly Property MenuItemLeft2of4() As Integer
        Get
            MenuItemLeft2of4 = Switch(MM4, X4()(1), True, 350)
        End Get
    End Property

    Private ReadOnly Property MenuItemLeft3of4() As Integer
        Get
            MenuItemLeft3of4 = Switch(MM4, X4()(2), True, 490)
        End Get
    End Property

    Private ReadOnly Property MenuItemLeft4of4() As Integer
        Get
            MenuItemLeft4of4 = Switch(MM4, X4()(3), True, 630)
        End Get
    End Property

    Private ReadOnly Property MenuSubItemTop1() As Integer
        Get
            MenuSubItemTop1 = 4400
        End Get
    End Property

    Private ReadOnly Property MenuSubItemHeight() As Integer
        Get
            MenuSubItemHeight = 465
        End Get
    End Property

    Private ReadOnly Property MenuSubItemLeft1of2() As Integer
        Get
            MenuSubItemLeft1of2 = 3400
        End Get
    End Property

    Private ReadOnly Property MenuSubItemLeft2of2() As Integer
        Get
            MenuSubItemLeft2of2 = 8700
        End Get
    End Property

    Private ReadOnly Property MenuSubItemTop2() As Integer
        Get
            MenuSubItemTop2 = 6200
        End Get
    End Property

    Public Sub ReadHotKeyPress(ByVal sName As String)
        '  m_cHotKey.RestoreAndActivate Me.hWnd
        '::::ReadHotKeyPress
        ':::SUMMARY
        ': Reads a pressed Hot Key
        ':::DESCRIPTION
        ': Reads and parses a pressed hot key, performing whatever operation is currently set up.
        ':::PARAMETERS
        ': - sName - The key associated with the hot key pressed.

        Select Case sName
            Case "Security Monitor"
                frmPermissionMonitor.Show()
            Case "Printers"
                ViewPrinters
            Case "Calculator"
                OpenCalculator
        End Select
    End Sub

    Public ReadOnly Property frmSplash() As frmSplash2
        Get
            '  If IsCDSComputer Then Set frmSplash = frmSplash2: Exit Property
            frmSplash = frmSplash2
        End Get
    End Property


End Module
