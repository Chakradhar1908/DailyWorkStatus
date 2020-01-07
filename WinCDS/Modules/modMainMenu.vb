Imports stdole
Imports VBA

Module modMainMenu
    Public mUpdateInstance As Long ' To keep everyone from hitting the server at the same time...
    Public Structure MyMenu
        Dim Name As String
        Dim ParentMenu As String
        Dim Caption As String
        Dim Visible As Boolean
        Dim Layout As eMyMenuLayouts
        Dim HCID As Long

        Dim ImageW As Long
        Dim ImageH As Long
        Dim vSP As Long


        Dim ImageSource As Object 'ImageList
        Dim CaptionStyle As eCaptionStyles
        Dim CaptionMargin As Long
        Dim MaskColor As Long

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
        Dim Top As Long
        Dim Left As Long

        Dim ToolTipText As String
        Dim ControlCode As String

        Dim HotKeys As String
        Dim Operation As String

        Dim Visible As String
        Dim Image As StdPicture

        Dim IsSubItem As Boolean
    End Structure

    Public Structure MyMenuHR
        Dim Top As Long
        Dim Left As Long
        Dim Width As Long
    End Structure

    Private frmSplas As Form = frmSplash2
    Public Const frmSplashType As String = "frmSplash2"

    Public ReadOnly Property frmSplash As frmSplash2
        Get
            Return frmSplas
        End Get
    End Property

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
                MainMenu.WebServ.StartHTTP
                MainMenu.WebServ.Show()               ' is a form
            Case 50 '2
                frmSupportHost.Listen
            Case 51 '3
                PracticeCommandPromptFunctions.Show
            Case 52 '4
                MainMenu.Hide()
                frmAshleyEDI888.Show
            Case 53 '5
                Order = "TABLE-VIEWER"
                frmTableView.Show 'vbModal
'        Order = ""
            Case 54 '6
                MainMenu.Hide()
                frmAWSAdmin.Show
            Case 55 '7
                MainMenu.Hide()
                PracticeDiagnostics.Show
            Case 56 '8
                PracticeCommandPrompt.Show
            Case 57 '9
                PermissionMonitor 0
  End Select
    End Sub

    Public Function MainMenu_KeyDown(KeyCode As Integer, Shift As Integer)
        '::::MainMenu_KeyDown
        ':::SUMMARY
        ': Key down handler for main menu
        ':::DESCRIPTION
        ': Contains the handler code for the Key Down event in the main menu.  (Only) The Main Menu should (only) call this function on KeyDown event
        Dim M As String, L As String, VT As Long
        Dim A As String, B As String

        If FindControlCode(Shift, KeyCode, M, L) Then
            If Not (M = "file" And L = "login") Then     ' we don't zoom to the contextual menu for store login (available from the top always)
                MainMenu.LoadMenuToForm M
    End If
            MainMenu.SelectMenuItem , M, L
    Exit Function
        End If

        VT = Val(MainMenu.Tag)

        Dim T As String, I As Long
        T = Format(Shift, "00") & Format(KeyCode, "0000")
        If IsIn(T, "040018", "020017") Then Exit Function
        '  Debug.Print T
        Select Case T
            Case "000027"                                          ' ESC
                MainMenu.MenuItemHighlight -1, True
      If MainMenu.ParentMenu = "" Then MainMenu.MainMenuClick -1
      MainMenu.LoadMenuToForm MainMenu.ParentMenu
      MainMenu.DoLogOut()
            Case "040065" : MainMenu.MainMenuClick 3                         ' Alt-A   'LoadMenuToForm "file:maintenance"
            Case "040066"                                          ' Alt-B
            Case "040067" : MainMenu.LoadMenuToForm "inventory:po"           ' Alt-C
            Case "040068" : MainMenu.LoadMenuToForm "inventory:deliveries"   ' Alt-D
            Case "040069" : MainMenu.LoadMenuToForm "inventory:ashley"       ' Alt-E
            Case "040070" : MainMenu.MainMenuClick 0                         ' Alt-F   'LoadMenuToForm "file"
            Case "040071" : MainMenu.LoadMenuToForm "general ledger"         ' Alt-G
            Case "040072"                                          ' Alt-H
            Case "040073" : MainMenu.MainMenuClick 2                         ' Alt-I   'LoadMenuToForm "inventory"
            Case "040074"                                          ' Alt-J
            Case "040075" : MainMenu.LoadMenuToForm "file:backup"            ' Alt-K
            Case "040076"                                          ' Alt-L
            Case "040077" : MainMenu.MainMenuClick 4                         ' Alt-M   'LoadMenuToForm "mailing"
            Case "040078" : MainMenu.MainMenuClick 5                         ' Alt-N   'LoadMenuToForm "installment"
            Case "040079" : MainMenu.MainMenuClick 1                         ' Alt-O   LoadMenuToForm "order entry"
            Case "040080"                                          ' Alt-P
            Case "040081"                                          ' Alt-Q
            Case "040082" : MainMenu.LoadMenuToForm "file:restore"           ' Alt-R
            Case "040083"                                          ' Alt-S
            Case "040084"                                          ' Alt-T
            Case "040085"                                          ' Alt-U
            Case "040086"                                          ' Alt-V
            Case "040087" : MainMenu.LoadMenuToForm "file:web"               ' Alt-W
            Case "040088"                                          ' Alt-X
            Case "040089"                                          ' Alt-Y
            Case "040090"                                          ' Alt-Z

            Case "000107", "010187"
                Select Case VT
                    Case 0 : MainMenu.Tag = "2"
                    Case 2 : MainMenu.Tag = "4"
                    Case 4
                        VersionControlDialog
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
                    ProgressForm 0, 1, "Preparing Settings...", , , , prgSpin
        MainMenu.Hide()
                    frmSetup.Show
                    ProgressForm()
                End If
            Case Else
                MainMenu.Tag = ""
                On Error Resume Next
                If Val(T) >= 65 And Val(T) < 90 Then
                    '        Debug.Print "t=" & T & ", chr=" & LCase(Chr(Val(T)))
                    For I = 1 To MainMenu.imgMenuItem.UBound
                        A = LCase(Chr(Val(T)))
                        B = LCase(MainMenu.ItemOptionHotKeys(MainMenu.imgMenuItem(I).Tag))
                        If B <> "" And A = B Then MainMenu.SelectMenuItem I: Exit For
                        '            Debug.Print "i=" & I & ", tag=" & LCase(imgMenuItem(I).Tag)
                        '            If InStr(LCase(Chr(Val(T))), LCase(.ItemOptionHotKeys(.imgMenuItem(i).Tag))) > 0 Then .SelectMenuItem i: Exit For
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
            Case "payables" : ShellOut_Shell MainMenu, FileAccountPayable
    Case "payroll" : ShellOut_Shell MainMenu, FilePayroll
    Case "banking" : ShellOut_Shell MainMenu, FileBanking
    Case "general ledger" : ShellOut_Shell MainMenu, FileGenLedger
'   Case "time clock":     ShellOut_Shell mainmenu, FileTimeClock
            Case Else : MsgBox "Could not launch " & Which & vbCrLf & "Please contact " & AdminContactCompany & " at " & AdminContactPhone2 & ".", vbCritical, "Unknown Program"
  End Select
        MainMenu.Show()
    End Sub

    Public Function GetMyMenu(ByVal Name As String, Optional ByRef Index As Long = 0) As MyMenu
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
        Dim I As Long
        InitializeMenus  ' this will only run once.  Immediately exits on subseq calls

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

        TrackUsage UsageStr, UsageDsc

  Select Case Source
            Case "file", "file:system", "file:utilities", "file:maintenance"
                Select Case Operation
                    Case "systemsetup"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        frmSetup.Show
                        ProgressForm()
''''''''''''''''
                    Case "password"
                        If modStores.SecurityLevel = seclevNoPasswords Then
                            MsgBox "You cannot set up the password until you have taken the computer out of No Passwords Mode in Store Setup." & vbCrLf &
                   "Click F1 for Help.", vbExclamation + vbMsgBoxHelpButton, ProgramMessageTitle, App.HelpFile, 31100 '34000
                        Else
                            PassWord.ChangePassword
                        End If
                    Case "configw"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        frmScanner.Show
                    Case "download"
                        MainMenu.Hide()
                        frmScannerDownload.Show
                    Case "creditcardmanager"
                        If Not CheckAccess("Credit Administration") Then Exit Function
                        MainMenu.Hide()
                        Load frmCCAdmin
          frmCCAdmin.HelpContextID = 47500
                        frmCCAdmin.Show
                    Case "webupdates"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        frmUpgrade.Show
                    Case "email"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        frmEmailSetup.Show

''''''''''''''''
                    Case "quarterly"
                        If Not CheckAccess("Annual Maintenance") Then Exit Function
                        If MsgBox("Caution:  This should only be done at CALENDAR year end!  When you click Ok, Quarterly Sales on your Inventory Data Base gets updated. Unit sales for the current year will be moved to the prior year.  The current year unit sales will be empty!", vbOKCancel + vbCritical) = vbCancel Then Exit Function
                        If MsgBox(" Are You Sure You want To Update Quarterly Inventory? ", vbYesNo + vbCritical) = vbNo Then Exit Function

                        MainMenu.MousePointer = vbHourglass
                        ExecuteRecordsetBySQL "UPDATE [2Data] SET PSales1=Sales1, PSales2=Sales2, " &
            "PSales3=Sales3, PSales4=Sales4, Sales1=0, Sales2=0, Sales3=0, Sales4=0", , GetDatabaseInventory
          MainMenu.MousePointer = vbDefault

                        MsgBox "Units sales are transferred!", vbInformation, "Quarterly Update Complete!"
        Case "restoredel"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        Inven = "R"
                        InvenA.DoInvRestore()
                    Case "racklabels"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        frmRackLabel.Show
                    Case "loadorig-manual"
                        Inven = "L"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        MainMenu.Hide()
                        InvenA.Show()
                        InvenA.Caption = "Loading Original Inventory Quantities"
                        InvenA.HelpContextID = 37000
                        Load InvAutoMan
          InvAutoMan.HelpContextID = 37000
                        InvAutoMan.Show
                    Case "loadorig-import"
                        Inven = "L"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        MainMenu.Hide()
                        Reports = "LoadOriginalInvByBarcodes"
                        frmPhysicalInventoryMainMenu.Show
                        frmPhysicalInventoryMainMenu.HelpContextID = 37000
                    Case "tags"
                        Inven = "TAGS"
                        'If Not CheckAccess("") Then Exit Sub
                        MainMenu.Hide()
                        frmPrintAllTickets.Show
                    Case "speech"
                        On Error Resume Next
                        Load frmSpeech
          If IsFormLoaded("frmSpeech") Then frmSpeech.Show()
                        MainMenu.SetFocus
                    Case "ashley"
                        frmAshleyEDI888.Show
                        MainMenu.Hide()
                    Case "s3"
                        If Not IsServer() Then
                            MsgBox "The Cloud backup/restore feature should only be run from the server.", vbExclamation, "You are on a Workstation!"
            Exit Function
                        End If
                        frmAWS.Show
                        MainMenu.Hide()
                    Case "exportitems"
                        frmExport.ShowExport
                        MainMenu.Hide()
                    Case "importitems"
                        frmExport.ShowImport
                        MainMenu.Hide()
                    Case "old-reports"
                        frmPDFReports.Show
                        MainMenu.Hide()
                    Case "login"
                        If Not CheckAccess("Log In To Other Stores") Then Exit Function
                        LogIn.Show vbModal
        Case "exit"
                        Unload MainMenu
        Case Else : Fail = True
                End Select
            Case "file:backup", "file:restore"
                ''''''''''''''''
                If Not CheckAccess("Backup/Restore") Then Exit Function

                If Not IsServer() Then 'Check if running from server
                    MsgBox "Backups MUST be made from main (Inventory) computer only!", vbExclamation, ProgramMessageTitle
        Exit Function
                End If

                If Left(Operation, 7) = "restore" Then
                    If MsgBox("CAUTION: Restoring from Backup will wipe out all transactions after date of files on this disk!", vbExclamation + vbOKCancel, ProgramMessageTitle) = vbCancel Then
                        Exit Function
                    End If
                End If

                Inven = "A"

                Select Case Operation
                    Case "backuppos" : frmBackUpGeneric.Display BackupMode.bkBackup, BackupType.bkPS, MainMenu, vbModal
        Case "backuppayables" : frmBackUpGeneric.Display BackupMode.bkBackup, BackupType.bkAP, MainMenu, vbModal
        Case "backuppayroll" : frmBackUpGeneric.Display BackupMode.bkBackup, BackupType.bkPR, MainMenu, vbModal
        Case "backupbanking" : frmBackUpGeneric.Display BackupMode.bkBackup, BackupType.bkBK, MainMenu, vbModal
        Case "backupgl" : frmBackUpGeneric.Display BackupMode.bkBackup, BackupType.bkGL, MainMenu, vbModal
        Case "backuppx" : frmBackUpGeneric.Display BackupMode.bkBackup, BackupType.bkpx, MainMenu, vbModal
        Case "backupall" : frmBackUpGeneric.Display BackupMode.bkBackup, BackupType.bkAll, MainMenu, vbModal
        Case "restorepos" : frmBackUpGeneric.Display BackupMode.bkRestore, BackupType.bkPS, MainMenu, vbModal
        Case "restorepayables" : frmBackUpGeneric.Display BackupMode.bkRestore, BackupType.bkAP, MainMenu, vbModal
        Case "restorepayroll" : frmBackUpGeneric.Display BackupMode.bkRestore, BackupType.bkPR, MainMenu, vbModal
        Case "restorebanking" : frmBackUpGeneric.Display BackupMode.bkRestore, BackupType.bkBK, MainMenu, vbModal
        Case "restoregl" : frmBackUpGeneric.Display BackupMode.bkRestore, BackupType.bkGL, MainMenu, vbModal
        Case "restoress" : frmBackUpGeneric.Display BackupMode.bkRestore, BackupType.bkSS, MainMenu, vbModal
        Case "restorepx" : frmBackUpGeneric.Display BackupMode.bkRestore, BackupType.bkpx, MainMenu, vbModal
        Case Else : Fail = True
                End Select

            Case "file:web"
                If CrippleBug("Online Sales") Then Exit Function
                If Not CheckAccess("Store Setup") Then Exit Function
                Select Case Operation
                    Case "webgen"
                        MainMenu.Hide()
                        Load frmAutoWeb
          Set frmAutoWeb.FOwner = MainMenu
          frmAutoWeb.Show
                        frmAutoWeb.HelpContextID = 37060
                    Case "webcsv"
                        Load frmAutoWeb
          Set frmAutoWeb.FOwner = MainMenu
          Dim F As Object
                        F = frmAutoWeb.BuildCSV
                        MsgBox "CSV Update Complete!" & vbCrLf & "File written to " & F & ".", vbExclamation, "Update CSV", App.HelpFile, 37070
          If Not FormIsLoaded("frmAutoWeb") Is Nothing Then Unload frmAutoWeb
          frmAutoWeb.HelpContextID = 37070
                    Case "webopenmonitor"
                        MainMenu.Hide()
                        Load frmAutoInv
          Set frmAutoInv.FOwner = MainMenu
          frmAutoInv.Show
                        frmAutoInv.HelpContextID = 37080
                    Case "webopensite"
                        On Error Resume Next
                        ShellOut_URL GetConfigTableValue("Website", WebDemoURL)
        Case Else : Fail = True
                End Select

            Case "order entry", "order entry:reports"
                Select Case Operation
                    Case "login"
                        If Not CheckAccess("Log In To Other Stores") Then Exit Function
                        LogIn.Show vbModal
        Case "newsale"
                        If CrippleBug("New Sales") Then Exit Function
                        If Not CheckAccess("Create Sales") Then Exit Function
                        Order = "A"
                        'frmSalesList.SafeSalesClear = True
                        frmSalesList.SalesCode = ""
                        Unload BillOSale
          MainMenu.Hide()
                        BillOSale.HelpContextID = 42000
                        BillOSale.HelpContextID = 42002
                        BillOSale.Show()
                        MailCheck.HelpContextID = 42000
                        MailCheck.optTelephone.Value = True
                        MailCheck.HidePriorSales = True
                        MailCheck.Show vbModal  ' If this is loaded "vbModal, BillOSale", lockup may occur.
                        MailCheck.HidePriorSales = False
                        Unload MailCheck
        Case "cashreg"
                        If CrippleBug("New Sales") Then Exit Function
                        If StoreSettings.bManualBillofSaleNo Then
                            MsgBox "You have selected to use Manually entered Bill of Sale numbers." & vbCrLf & "To use the cash register, you must unselect this feature in the store setup.", vbExclamation, "Cannot use Cash Register"
            Exit Function
                        End If
                        '          MsgBox "cashreg: -1"
                        If Not CheckAccess("Create Sales") Then Exit Function
                        '          MsgBox "cashreg: 0"
                        MainMenu.Hide()
                        '          MsgBox "cashreg: 1"
                        Load frmCashRegister
'          MsgBox "cashreg: 2"
                        frmCashRegister.HelpContextID = 42500
                        '          MsgBox "cashreg: 3"
                        frmCashRegister.BeginSale
                        frmCashRegister.HelpContextID = 42500
                        Order = "CashRegister"
                    Case "deliver"
                        If CrippleBug("Delivering Sales") Then Exit Function
                        If Not CheckAccess("Deliver Sales") Then Exit Function
                        MainMenu.Hide()
                        Order = "B"
                        BillOSale.Show()
                        BillOSale.HelpContextID = 43000
                        BillOSale.BillOSale2_Show()
                        MailCheck.HelpContextID = 43000
                        MailCheck.optSaleNo.Value = True
                        MailCheck.Show vbModal  ' If this is loaded "vbModal, BillOSale", lockup may occur.
                        Unload MailCheck
        Case "payment"
                        If CrippleBug("Paying on Sales") Then Exit Function
                        If Not CheckAccess("Accept Payments") Then Exit Function
                        MainMenu.Hide()
                        Order = "D"
                        BillOSale.Show()
                        BillOSale.HelpContextID = 44000
                        BillOSale.BillOSale2_Show()
                        MailCheck.HelpContextID = 44000
                        MailCheck.optSaleNo.Value = True
                        MailCheck.Show vbModal  ' If this is loaded "vbModal, BillOSale", lockup may occur.
                        Unload MailCheck
        Case "viewsale"
                        If Not CheckAccess("View Sales") Then Exit Function
                        MainMenu.Hide()
                        Order = "E"
                        BillOSale.HelpContextID = 45000
                        BillOSale.Show()
                        BillOSale.BillOSale2_Show()
                        MailCheck.HelpContextID = 45000
                        MailCheck.optSaleNo.Value = True
                        MailCheck.Show vbModal  ' If this is loaded "vbModal, BillOSale", lockup may occur.
                        Unload MailCheck
        Case "voidsale"
                        If CrippleBug("Voiding Sales") Then Exit Function
                        If Not CheckAccess("Void Sales") Then Exit Function
                        MainMenu.Hide()
                        Order = "C"
                        BillOSale.Show()
                        BillOSale.HelpContextID = 46000
                        BillOSale.BillOSale2_Show()
                        MailCheck.HelpContextID = 46000
                        MailCheck.optSaleNo.Value = True
                        MailCheck.Show vbModal  ' If this is loaded "vbModal, bos2", lockup may occur.  However, alt-tab can MainMenu.Hide Bos2 if we don't.
                        Unload MailCheck
        Case "cashdrawer"
                        If Not CheckAccess("Cash Drawer") Then Exit Function
                        MainMenu.Hide()
                        OrdCashflow.Show
                    Case "preview"
                        If Not CheckAccess("View Stock Quantities") Then Exit Function
                        MainMenu.Hide()
                        Order = "F"
                        OrdPreview.Show()
                        InvCkStyle.HelpContextID = 48000
                    Case "adjustments"
                        If CrippleBug("Adjusting Sales") Then Exit Function
                        If Not CheckAccess("Adjust Sales") Then Exit Function
                        MainMenu.Hide()
                        Order = "Credit"
                        OnScreenReport.HelpContextID = 49000
                        OnScreenReport.CustomerAdjustment
                    Case "dailyaudit"
                        If Not CheckAccess("Daily Audit Report") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "Audit"
                        OrdAudit2.Show
                        OrdAudit2.HelpContextID = 49620
                    Case "undelivered"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "R"
                        ReportPrint.Show
                        ReportPrint.HelpContextID = 49630
                    Case "layaway"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "L"
                        ReportPrint.Show
                        ReportPrint.HelpContextID = 49640
                    Case "backorder"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "B"
                        ReportPrint.Show
                        ReportPrint.HelpContextID = 49650
                    Case "creditsales"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "C"
                        ReportPrint.Show
                        ReportPrint.HelpContextID = 49660
                    Case "customerhistory"
                        If Not CheckAccess("View Sales") Then Exit Function
                        MainMenu.Hide()
                        Reports = "H"
                        MailCheck.HelpContextID = 49670
                        OnScreenReport.CustomerHistory
                        OnScreenReport.HelpContextID = 49670
                    Case "salestax"
                        If Not CheckAccess("Store Finances") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "ST"
                        DateForm.Show
                        DateForm.HelpContextID = 49680
                    Case "advertising"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Reports = "O"
                        Order = "ATR"
                        DateForm.HelpContextID = 49700
                        DateForm.Show
                        DateForm.HelpContextID = 49700
                    Case Else : Fail = True
                End Select

            Case "service"
                If Not CheckAccess("Service Orders") Then Exit Function
                Select Case Operation
                    Case "servicecalls"
                        Order = "S"
                        MailCheck.HelpContextID = 49500
                        MailCheck.HidePriorSales = True
                        MailCheck.optTelephone.Value = True
                        MailCheck.Show vbModal
          MailCheck.HidePriorSales = False
                    Case "damagedstock"
                        Order = "SDam"
                        MainMenu.Hide()
                        ServiceParts.HelpContextID = 49502
                        ServiceParts.SelectMode ServiceMode_ForStock, True, True
          ServiceParts.Show
                    Case "partsorders"
                        Order = "SParts"
                        MainMenu.Hide()
                        ServiceParts.HelpContextID = 49503
                        ServiceParts.SelectMode ServiceMode_ForStock, True, True
          ServiceParts.Caption = "Parts Orders Form"
                        ServiceParts.Show
                    Case "openservicecalls"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Order = "SCR"
                        ServiceReports.HelpContextID = 49504
                        ServiceReports.Show
                    Case "openpartsorders"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Order = "SPR"
                        ServiceReports.HelpContextID = 49505
                        ServiceReports.Show
                    Case "partsorderbilling"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Order = "SBR"
                        ServiceReports.HelpContextID = 49506
                        ServiceReports.Show
                    Case "unpaidbilling"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        Order = "SBU"
                        ServiceReports.HelpContextID = 49507
                        ServiceReports.Show
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
                        Unload InvenA
          InvenA.Show()
                        InvenA.Caption = "Adding New Items"
                        InvenA.Desc = "Enter Description Here"
                        InvenA.HelpContextID = 51000
                        Load InvCkStyle
          InvCkStyle.HelpContextID = 51000
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
                        InvenA.Caption = "Processing Factory Shipments"
                        InvenA.HelpContextID = 53000
                        Load InvCkStyle
          InvCkStyle.HelpContextID = 53000
                        InvCkStyle.Show()
                    Case "storetransfers"
                        If CrippleBug("Store Transfers") Then Exit Function
                        If Not CheckAccess("Store Transfers") Then Exit Function
                        If Not CheckAccess("Schedule Deliveries", , True) Then Exit Function
                        MainMenu.Hide()
                        Inven = "T"
                        InvenA.Show()
                        InvenA.Caption = "Processing Store Transfers"
                        InvenA.HelpContextID = 54000
                        Load InvCkStyle
          InvCkStyle.HelpContextID = 54000
                        InvCkStyle.Show()
                    Case "viewstock"
                        MainMenu.Hide()
                        Inven = "E"
                        InvenA.Show()
                        InvenA.Caption = "View Any Item"
                        InvenA.HelpContextID = 55000
                        Load InvCkStyle
          InvCkStyle.HelpContextID = 55000
                        InvCkStyle.Show()
                    Case "changecontents"
                        If CrippleBug("Change Contents") Then Exit Function
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        MainMenu.Hide()
                        Inven = "H"
                        InvenA.Show()
                        InvenA.Caption = "Inventory Maintenance"
                        InvenA.HelpContextID = 56000
                        Load InvAutoMan
          InvAutoMan.HelpContextID = 56000
                        InvAutoMan.Show
                    Case "orderstatus"
                        If CrippleBug("Sales Adjustments") Then Exit Function
                        If Not CheckAccess("Order Status") Then Exit Function
                        InvOrdStatus.ShowCost = CheckAccess("View Cost and Gross Margin", False, True, False)
                        MainMenu.Hide()
                        InvOrdStatus.CheckDeliveryStatus
                    Case "comm"
                        If Not CheckAccess("Commissions") Then Exit Function
                        MainMenu.Hide()
                        InvPayComm.HelpContextID = 59700
                        InvPayComm.Show
                    Case "ss"
                        MainMenu.Hide()
                        Reports = "VSS"
                        frmViewSSLoc.HelpContextID = 59990
                        frmViewSSLoc.Show
                    Case Else : Fail = True
                End Select
            Case "inventory:transfers"
                If CrippleBug("Store Transfers") Then Exit Function
                If Not CheckAccess("Store Transfers") Then Exit Function
                If Not CheckAccess("Schedule Deliveries", , True) Then Exit Function
                Select Case Operation
                    Case "schedule"
                        Inven = "T"
                        frmTransferSelect.Show
                        MainMenu.Hide()
                    Case "show"
                        Inven = "View Transfer"
                        frmTransferLookup.Show
                        MainMenu.Hide()
                    Case "reportopen"
                        frmTransferReports.Mode = "Pending"
                        frmTransferReports.Show
                        frmTransferReports.HelpContextID = 54030
                        MainMenu.Hide()
                    Case "reportclosed"
                        frmTransferReports.Mode = "Previous"
                        frmTransferReports.Show
                        frmTransferReports.HelpContextID = 54040
                        MainMenu.Hide()
                    Case "invdelmulti-trans"
                        InvPull.Pull = 4
                        InvPull.Show()
                        InvPull.HelpContextID = 59400
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
                        InvPoPrint.HelpContextID = 57300
                        InvPoPrint.Show()
                    Case "pofaxprint"
                        Inven = "FPO"
                        FaxPo.HelpContextID = 57400
                        FaxPo.Show
                    Case "poemail"
                        On Error Resume Next
                        Inven = "EPO"
                        frmEmail.HelpContextID = 57450
                        frmEmail.Mode = emPO
                        frmEmail.Show()
                    Case "porec"
                        If CrippleBug("PO Changes") Then Exit Function
                        PurchaseOrder = "REC"
                        On Error Resume Next
                        Load EditPO
          If PurchaseOrder <> "" Then EditPO.Show()
                    Case "povoid"
                        PurchaseOrder = "Void"
                        On Error Resume Next
                        Load EditPO
          If PurchaseOrder <> "" Then EditPO.Show()
                    Case "porepreceiving"
                        Inven = "PRec"
                        InvPoPrint.HelpContextID = 57700
                        InvPoPrint.Show()
                    Case "porepnotack"
                        Inven = "AK"
                        InvPoPrint.HelpContextID = 57800
                        InvPoPrint.Show()
                    Case "porepoverdue"
                        Inven = "OverdueOrders"
                        InvPoPrint.HelpContextID = 57850
                        InvPoPrint.Show()
                    Case "porepnotackE"
                        Inven = "AK-E"
                        frmPOEmails.HelpContextID = 57805
                        frmPOEmails.sType = 0
                        frmPOEmails.Show
                    Case "porepoverdueE"
                        Inven = "OverdueOrders-E"
                        frmPOEmails.HelpContextID = 57855
                        frmPOEmails.sType = 1
                        frmPOEmails.Show
                    Case "porepopen"
                        PurchaseOrder = "ReOpen"
                        InvPoPrint.HelpContextID = 57900
                        InvPoPrint.OpenPOReport
                    Case "poordermanual"
                        Inven = "P"
                        InvPoSelect.HelpContextID = 57100
                        InvPoSelect.Show
                    Case "poorderminimum"
                        PurchaseOrder = "OrderMinimum"
                        InvAutoReOrder.HelpContextID = 57100
                        InvAutoReOrder.Show
                        InvAutoReOrder.OrderAutomatic CheckAccess("View Cost and Gross Margin", False, True, False)
        Case "poorderdemand"
                        PurchaseOrder = "OrderDemand"
                        InvAutoReOrder.HelpContextID = 57100
                        InvAutoReOrder.Show
                        InvAutoReOrder.OrderByDemand StoresSld, Date, DateAdd("d", 7, Date), CheckAccess("View Cost and Gross Margin", False, True, False)
        Case "pocombine"
                        PurchaseOrder = "POCombine"
                        frmCombinePOs.Show
                    Case "ashley"
                        Order = "AshleyEDI"
                        frmAshleyEDI.Show
                    Case "ashleyasn"
                        '          frmAshleyEDIReceive.Mode = "ASN"
                        Order = "AshleyASN"
                        frmAshleyEDIReceive.Show
                    Case "ashleyopenpo"
                        InvPoPrint.HelpContextID = 57900
                        Reports = "Ashley"
                        InvPoPrint.OpenPOReport
                    Case Else : Fail = True
                End Select

            Case "inventory:deliveries"
                If Not CheckAccess("Schedule Deliveries") Then Exit Function
                MainMenu.Hide()

                Select Case Operation
                    Case "invdelpullloads"
                        InvPull.Pull = 1
                        InvPull.Show()
                        InvPull.HelpContextID = 59100
                    Case "invdeltickets"
                        InvPull.Pull = 2
                        InvPull.Show()
                        InvPull.HelpContextID = 59200
                    Case "invdelmulti-cross"
                        If Not CheckAccess("View Cost and Gross Margin", False, True, False) Then MainMenu.Show() : 
                        Exit Function
                        InvPull.Pull = 3
                        InvPull.Show()
                        InvPull.HelpContextID = 59300
'        Case "invdelmulti-trans"
'          InvPull.Pull = 4
'          InvPull.Show
'          InvPull.HelpContextID = 59400
                    Case "invdelmulti-lists"
                        InvPull.Pull = 5
                        InvPull.Show()
                        InvPull.HelpContextID = 59500
                    Case "invdelcalendar"
                        Inven = "Calendar"
                        Calendar.Show()
                        Calendar.HelpContextID = 59600
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
                        PackagePrice.HelpContextID = 59730
                    Case "invpackedit"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        If Not CheckAccess("View Cost and Gross Margin", False, True, False) Then Exit Function
                        MainMenu.Hide()
                        Reports = "ET"
                        PackagePrice.EditPackages
                        PackagePrice.HelpContextID = 59740
                    Case "invpacklist"
                        If Not CheckAccess("Create and Edit Items") Then Exit Function
                        If Not CheckAccess("View Cost and Gross Margin", False, True, False) Then Exit Function
                        MainMenu.Hide()
                        Reports = "RT"
                        InvKitReport.Show
                        InvKitReport.HelpContextID = 59750
                    Case "invpacklookup"
                        If Not CheckAccess("View Stock Quantities") Then Exit Function
                        MainMenu.Hide()
                        Reports = "CS"
                        InvKitStock.HelpContextID = 59760
                        InvKitStock.ShowPackages
                        InvKitStock.HelpContextID = 59760
                    Case Else : Fail = True
                End Select

            Case "inventory:reports"
                If Not CheckAccess("View Inventory Reports") Then Exit Function
                MainMenu.Hide()

                Select Case Operation
                    Case "invrepinven"
                        Reports = "I"
                        Inven = "IRep"
                        InvReports.Show
                    Case "invrepmargin"
                        Reports = "M"
                        Inven = "MRep"
                        InvReports.Show
                    Case "invrepmanuf"
                        Reports = "I"
                        Inven = "ML"
                        InvPoPrint.HelpContextID = 59930
                        InvPoPrint.Show()
                    Case "invrepbest"
                        Inven = "BS"
                        Reports = "I"
                        ReportPrint2.HelpContextID = 59940
                        ReportPrint2.Show
                    Case "invrepdog"
                        Reports = "I"
                        Inven = "DG"
                        ReportPrint3.HelpContextID = 59970
                        ReportPrint3.Show
                    Case "invrepss"
                        Reports = "I"
                        Inven = "SS"
                        ReportPrint3.HelpContextID = 59960
                        ReportPrint3.Show
                    Case "invrepserialno"
                        Reports = "I"
                        Inven = "SNO"
                        ReportPrint3.HelpContextID = 59980
                        ReportPrint3.Show
                    Case "invrepbarcode"
                        Reports = "I"
                        Inven = "PhysicalInventory"
                        frmPhysicalInventoryMainMenu.Show
                        frmPhysicalInventoryMainMenu.HelpContextID = 59950
                    Case "designtag"
                        Reports = "DESIGNTAG"
                        frmDesignTag.Show
                    Case "invrepreturn"
                    Case "storecatalog"
                        MainMenu.Hide()
                        Reports = "Store Catalog"
                        frmCatalog.Show
                    Case "special"
                        If Not CheckAccess("Change Item Prices") Then Exit Function
                        MainMenu.Hide()
                        Ticket.Show
                    Case "mini"
                        If CrippleBug("Mini Scanners") Then Exit Function
                        If Not CheckAccess("View Stock Quantities") Then Exit Function
                        Reports = "Mini-Barcode Scanner"
                        BarcodeInventoryCheck
                    Case Else : Fail = True
                End Select



            Case "accounting"
                Select Case Operation
                    Case "ap"
                        LaunchProgram "payables"
        Case "pr"
                        LaunchProgram "payroll"
        Case "bk"
                        LaunchProgram "banking"
        Case "gl"
                        LaunchProgram "general ledger"
        Case "qb"
                        If Not CheckAccess("Store Setup") Then Exit Function
                        MainMenu.Hide()
                        Load frmQuickBooksAccountSetup
          frmQuickBooksAccountSetup.HelpContextID = 67000
                        frmQuickBooksAccountSetup.Show
                    Case Else : Fail = True
                End Select

            Case "mailing"
                Select Case Operation
                    Case "add"
                        If Not CheckAccess("Create Sales") Then Exit Function
                        Mail = "ADD/Edit"
                        MainMenu.Hide()
                        BillOSale.Show()
                        BillOSale.HelpContextID = 71000
                        MailCheck.HelpContextID = 71000
                        MailCheck.optTelephone.Value = True
                        MailCheck.Show vbModal ' If this is loaded "vbModal, BillOSale", lockup may occur.
                    Case "print"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        MailPrint.HelpContextID = 72000
                        MailPrint.Show
                    Case "export"
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        MainMenu.Hide()
                        frmMailExport.Show
                    Case "advert"   ' advertising by Zip
                        If Not CheckAccess("Sales Reports") Then Exit Function
                        CustomerTypes.Show
                        CustomerTypes.HelpContextID = 73400
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
                        MailFix.Show
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
                        ARPaySetUp.HelpContextID = 81000
                    Case "payview" 'Payment and View
                        If Not CheckAccess("Accept Payments") Then Exit Function
                        MainMenu.Hide()
                        ArSelect = "P"
                        Unload ArCard  'BFH20061102 - just to be sure
                        ArCard.HelpContextID = 82000
                        ArCard.GetCustomerAccount
                    Case "edit" 'Edit Accounts
                        If Not CheckAccess("Credit Administration") Then Exit Function
                        MainMenu.Hide()
                        ArSelect = "Edit"
                        Load ArCard
          ArCard.HelpContextID = 83000
                        ArCard.GetCustomerAccount
                    Case "void" 'Void Contract
                        If Not CheckAccess("Credit Administration") Then Exit Function
                        MainMenu.Hide()
                        ArSelect = "V"
                        Load ArCard
          ArCard.HelpContextID = 84000
                        ArCard.GetCustomerAccount
                        ArCard.VoidAccount
                    Case "oldapps" 'Old Account Setup
                        If Not CheckAccess("Credit Administration") Then Exit Function
                        MainMenu.Hide()
                        ArSelect = "S"
                        BillOSale.Show()
                        BillOSale.HelpContextID = 85000
                        Load MailCheck
          MailCheck.HelpContextID = 85000
                        MailCheck.optTelephone.Value = True
                        MailCheck.Show vbModal  ' If this is loaded "vbModal, BillOSale", lockup may occur.
                    Case "creditapps-new" 'Add New Appications
                        ArSelect = "A"
                        MainMenu.Hide()
                        BillOSale.Show()
                        MailCheck.optTelephone.Value = True
                        MailCheck.Show vbModal
          ArApp.HelpContextID = 86100
                    Case "creditapps-old" 'Edit Old Appications
                        ArSelect = "EA"
                        MainMenu.Hide()
                        ArApp.Show()
                        ArApp.HelpContextID = 86200
                        ArApp.NextEntry
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
                        Load ArReports
          ArReports.HelpContextID = 87100
                        ArReports.Show
                    Case "latecharges" 'Late Charge And Notices
                        ArSelect = "L"
                        Load ArReports
          ArReports.HelpContextID = 87200
                        ArReports.Show
                    Case "aging" 'A/R Ageing Reports
                        ArSelect = "A"
                        Load ArReports
          ArReports.HelpContextID = 87300
                        ArReports.Show
                    Case "wholate"    ' Whos Late Report
                        ArSelect = "WHOLATE"
                        Load ArReports
          ArReports.HelpContextID = 87350
                        ArReports.Show
                    Case "delinquent" 'A/R Delinquent Accounts
                        ArSelect = "D"
                        Load ArReports
          ArReports.HelpContextID = 87400
                        ArReports.Show 'vbModal
                    Case "trial" 'A/R Trial Balance
                        ArSelect = "T"
                        Load ArReports
          ArReports.HelpContextID = 87500
                        ArReports.Show 'vbModal
                    Case "newaccounts" 'New Account Report
                        ArSelect = "N"
                        Load ArReports
          ArReports.HelpContextID = 87600
                        ArReports.Show 'vbModal
                    Case "closedaccounts" 'Closed Account Report
                        ArSelect = "O"
                        Load ArReports
          ArReports.HelpContextID = 87700
                        ArReports.Show 'vbModal
                    Case "writeoff" 'write off report
                        ArSelect = "W"
                        Load ArReports
          ArReports.HelpContextID = 87710
                        ArReports.Show 'vbModal
                    Case "repo"  ' Repo Report
                        ArSelect = "R"
                        Load ArReports
          ArReports.HelpContextID = 87720
                        ArReports.Show
                    Case "legal" ' Legal Report
                        ArSelect = "LG"
                        Load ArReports
          ArReports.HelpContextID = 87730
                        ArReports.Show
                    Case "losscombo" 'Loss Combo = Write Off, Repo, Legal, Bankruptcy
                        ArSelect = "LC"
                        Load ArReports
          ArReports.HelpContextID = 87710
                        ArReports.Show
                    Case "nonpayment"
                        ArSelect = "NP"
                        Load ArReports
          ArReports.HelpContextID = 87740
                        ArReports.Show
                    Case "export" 'Export to Credit Bureau
                        ArSelect = "X"
                        Load ArReports
          ArReports.HelpContextID = 87800
                        ArReports.Show 'vbModal
                    Case "restore"
                        ArSelect = "RestoreAR"
                        Dim Xx As String, R As Recordset
                        Xx = InputBox("Enter AR Account to be restored:", "Restore Voided Account")
                        If Xx <> "" Then
            Set R = GetRecordsetBySQL("SELECT * FROM InstallmentInfo WHERE arno='" & ProtectSQL(Xx) & "'")
            If R.RecordCount = 0 Then
                                MsgBox "Account No [" & Xx & "] does not exist.", vbExclamation, "No Such Account"
            Else
                                If R("Status") <> "V" Then
                                    MsgBox "Account No [" & Xx & "] is not void.", vbExclamation, "Invalid Account"
              Else
                                    ExecuteRecordsetBySQL "UPDATE InstallmentInfo SET Status='O' WHERE ArNo='" & ProtectSQL(Xx) & "'"
                AddNewARTransactionExisting StoresSld, Xx, , arPT_stReO, 0, 0, "Restored Voided Account: " & GetCashierName
                MsgBox "Account [" & Xx & "] restored.  Status set to 'Open'.", vbInformation, "Account Restored"
              End If
                            End If
                        End If
                        DisposeDA R
          MainMenu.Show()
                    Case Else : Fail = True
                End Select
            Case Else : Fail = True
        End Select

        If Fail Then
            MsgBox FailMsg, vbCritical, FailTitle
    MainMenu.Show()
        End If

        TrackUsage UsageStr
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
        If ShutdownSemaforeFile(ItExists:=True) Then DoShutDown True: Exit Sub
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
            modPayPal.PayPalCheckSales  ' exits immediately if there is no setup info
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
                Load frmUpgrade
      frmUpgrade.DoSilentUpdate Msg, FileList
      Unload frmUpgrade
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
                MainMenu.LoadMenuToForm ""
    End If
        End If

#If True Then
        If LastLoginExpired And MainMenu.cmdLogout.Visible Then modPasswords.LogOut()
#End If


        ''''''''''''''''''''''''' AUTO-BACKUP
        If AWS_AutoBackup Then
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
                    DoAmazonAutoBackup
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
        RecordWinCDSBenchmark
#End If

        '''''''''''''''''''''''''  INTERNET MONITOR
#If True Then
        MonitorInternet
#End If

        ''''''''''''''''''''''''' AUTO-RESTART
#If True Then   ' this will only run between 3am and 4am
        If DateBetween(TimeValue(Now), #3:00:00 AM#, #4:00:00 AM#, , "s") Then
#Else           ' this one could be used for debugging, if you comment out the If DateEqual(... below
  If DateBetween(TimeValue(Now), #1:00:00 PM#, #5:00:00 PM#, , "s") Then
#End If
            If DateEqual(Of Date, DateValue(ProgramStart))() Then Exit Sub      ' if it's already marked as run today, then exit now
            NightlyCleanup
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

        OpenCHM

        '  Not reliable...
        '  SendKeys_safe "{F1}"
        '  RunShellExecute "open", App.HelpFile, 0&, 0&, SW_SHOWDEFAULT
        '  OpenHelp Me.KeyCatch.hWnd
    End Sub

End Module
