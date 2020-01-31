Module modPasswords
    Private LastLoginName As String  ' Last account name to be authenticated.
    Private LastLoginPriv As String

    Private LastLoginExpiry As Date
    Private Const SECURITY_PASSWORD_TIMEOUT as integer = 180

    Private Const BACKDOOR_PROGRAMMER As String = "QkFDS0RPT1I="
    Private Const BACKDOOR_USERNAME As String = "BACKDOOR"
    Private Const BACKDOOR_0 as integer = 2048  ' Shut off all but programmer back door
    Private Const BACKDOOR_1 as integer = 1
    Private Const BACKDOOR_2 as integer = 2
    Private Const BACKDOOR_3 as integer = 4
    Private Const BACKDOOR_4 as integer = 8
    Private Const BACKDOOR_5 as integer = 16
    Private Const BACKDOOR_6 as integer = 32
    Private Const BACKDOOR_7 as integer = 64
    Private Const BACKDOOR_8 as integer = 128
    Private Const BACKDOOR_9 as integer = 256
    Private Const BACKDOOR_A as integer = 512
    Private Const BACKDOOR_B as integer = 1024
    Private Const BACKDOOR_C as integer = 2048
    Private Const BACKDOOR_D as integer = 4096
    Private Const BACKDOOR_E as integer = 8192
    Private Const BACKDOOR_F as integer = 16384
    Private Const BACKDOOR_P as integer = 32767 '  FFF, 2048

    Private Const SECURITY_NAME_BACKDOOR As String = BACKDOOR_USERNAME
    Private Const SECURITY_NAME_DEMO As String = "DEMO"
    Private Const SECURITY_NAME_EVERYBODY As String = "EVERYBODY"
    Private PermsLoaded As Boolean
    Public PermOptions() As PermissionOption
    ' Public array of option possibilities.
    ' This is loaded by MainModule.main_StoreInitialize,
    ' which is called when the program starts.

    Public Const WINCDS_ONLY_PASSWORD As String = "CDS5150"
    Public Const WINCDS_ONLY_PASSWORD_2 As String = "CUSTOMD"
    Public Const PRACTICE_SECURE_PASSWORD As String = "COMMERCE21"
    Private Const BACKDOOR_PASSWORD1 As String = "WinCDS123"
    ' Currently, we have curtailed release of the backdoor passwords, so these are simply random UUIDs.
    ' If we need to give them out again in the future, we can define these as we want to.
    Private Const BACKDOOR_PASSWORD2 As String = "{45297203-1805-CC4F-A787-7EBB2BA59FA7}"
    Private Const BACKDOOR_PASSWORD3 As String = "{FEA66D76-3B6A-EA42-A8F8-A9ED8363E75E}"
    Private Const BACKDOOR_PASSWORD4 As String = "{16D14616-F525-C543-B4DD-AD35F824EFF0}"
    Private Const BACKDOOR_PASSWORD5 As String = "{016EB39B-4520-744C-BB3B-5BA7985CC3EB}"
    Private Const BACKDOOR_PASSWORD6 As String = "{0FC06AB2-FC76-1440-967C-ED4A4133A6D9}"
    Private Const BACKDOOR_PASSWORD7 As String = "{CF8D0BAB-41FE-2F49-A9A1-2894C537AD99}"
    Private Const BACKDOOR_PASSWORD8 As String = "{0570E689-C395-8F43-A315-96F76387E24A}"

    Private Const BACKDOOR_PASSWORD9 As String = "{9513559B-5527-384F-ACC9-3BE705A258E3}"
    Private Const BACKDOOR_PASSWORDA As String = "{81D93CBE-095F-E44A-8C74-44DD28F7B0F2}"
    Private Const BACKDOOR_PASSWORDB As String = "{518A102D-D07B-6348-A3FE-01EF7081BB36}"
    Private Const BACKDOOR_PASSWORDC As String = "{70BAB22D-296D-4640-8951-DC62E97DDEDD}"
    Private Const BACKDOOR_PASSWORDD As String = "{BD1B1CE9-B88C-5544-AF47-4F0D2460F0F2}"
    Private Const BACKDOOR_PASSWORDE As String = "{DC7221D5-770F-B547-B2C5-9D4D4080B1AD}"
    Private Const BACKDOOR_PASSWORDF As String = "{BE86C4B0-0AB2-0149-89C6-F5EF46F6FA56}"

    Private Const BACKDOOR_PASSWORD_MAX as integer = 15

    Public Enum ComputerSecurityLevels
        seclevOfficeComputer = 0
        seclevSalesFloor = 1
        seclevNoPasswords = 2
    End Enum
    Public Structure PermissionOption
        Dim OptionDesc As String
        Dim OptionID as integer
    End Structure

    Public ReadOnly Property LastLoginExpired() As Boolean
        Get
            If IsFormLoaded("BillOSale") Then LastLoginExpired = False : Exit Property
            If Not IsDate(GetLastLoginExpiry) Then LastLoginExpiry = DateAdd("s", -1, Now)
            LastLoginExpired = DateAfter2(Now, GetLastLoginExpiry, , "s")
        End Get

    End Property
    Public ReadOnly Property GetLastLoginExpiry() As Date
        Get
            GetLastLoginExpiry = LastLoginExpiry
        End Get
    End Property

    Public Function ResetLastLoginExpiry(Optional ByVal NewEntry As Boolean = False) As Boolean
        '::::ResetLastLoginExpiry
        ':::SUMMARY
        ': Reset the Last Login Expiry time.
        ':::DESCRIPTION
        ': This function is used to reset the Last Login Expiry time.  Call this
        ': on every major operation so that the user remains logged in for any normal
        ': procedure of operations.  Prevents the user from being reported as "idle" and logging out.
        ':::PARAMETERS
        ': - NewEntry - Indicates whether it is New Entry or not.
        ':::RETURN
        ': Boolean - Returns True.
        If Not NewEntry And LastLoginExpired Then Exit Function
        LastLoginExpiry = DateAdd("s", SECURITY_PASSWORD_TIMEOUT, Now)
        ResetLastLoginExpiry = True
    End Function

    Public Function CheckAccess(ByRef Zone As String, Optional ByRef Popup As Boolean = False, Optional ByRef UseLastEntry As Boolean = False, Optional ByRef AllowNewEntry As Boolean = False, Optional ByRef Reason As String = "") As Boolean
        '::::CheckAccess
        ':::SUMMARY
        ': Check Access for current operator (possibly raising a password / login prompt)
        ':::DESCRIPTION
        ': This function is used to check permissions, passwords given by Customers to access software.
        ':
        ': - Returns True if the specified account validates and has access to Zone, else false.
        ': - First, check if the Zone requires a password.
        ':::PARAMETERS
        ': - Zone - Password Zone description as defined in modGetSales.
        ': - Popup - Allow raising a pop-up window for password entry.
        ': - UseLastEntry - Allow the use of the previously entered permissions.
        ': - AllowNewEntry - If previously selected permissions are insufficient, prompt again?
        ': - Reason - Indicates the Reason to Access.
        ':::SEE ALSO
        ': - Encrypt, Decrypt, Backdoor
        ':::RETURN
        ': Boolean - Returns True if the specified account validates and has access to Zone, else false.
        Dim CheckAllStores As Boolean
        Dim TargetZone as integer, Security As ComputerSecurityLevels
        Dim PwdOK As Boolean

        TargetZone = QueryPrivZone(Zone)
        If TargetZone = 0 Then Exit Function

        If Not LastLoginExpired Then
            If CheckPrivLevel(QueryPrivZone(Zone), LastLoginPriv) Then
                UseLastEntry = AllowUseLastEntry
            End If
        End If

        ' BFH20060824
        ' For some things, because employees are specific to individual stores,
        ' we need to check the passwords and permissions from all employees
        ' in all stores.  Mainly this is true of Logining into other stores,
        ' because otherwise, the security check for this option becomes immediately
        ' prohibitive when the software starts up in store 1 and you're only defined in
        ' store 2...  This should search them all and check for the appropriate
        ' permissions in all stores.  Groups are defined in the inventory DB so
        ' they are global throughout all stores.
        ' Defined as IsIn() to facilitate adding later in case other functions should
        ' need or want this.
        CheckAllStores = IsIn(CStr(TargetZone), "33")

        UpdatePermissionMonitor(Zone)

        Security = modStores.SecurityLevel
        If Security = ComputerSecurityLevels.seclevNoPasswords Then
            ' Demo level - free access for all.
            LastLoginName = SECURITY_NAME_DEMO
            LastLoginPriv = ""
            UpdatePermissionMonitor(Zone)
            CheckAccess = True
            Exit Function
        End If

        If (Security = ComputerSecurityLevels.seclevOfficeComputer) And CheckPrivLevel(TargetZone, "E") Then
            ' Office mode + Everybody - no password required.
            If Not (UseLastEntry And Not AllowNewEntry) Then
                LastLoginName = SECURITY_NAME_EVERYBODY
                LastLoginPriv = "E"
            End If
            UpdatePermissionMonitor(Zone)
            CheckAccess = True
            Exit Function
        End If

        If UseLastEntry Then
            ' Retain the last set of privs we used; don't ask for a password.
            ' Unless we want to restrict subzones, with popup passwords..
            ' Example: Everybody can run sales reports, only admins can see cost.
            ' We don't want to ask for a password to get into the reports,
            ' only if cost is requested.. In that case, we have to check the
            ' last privs, and if they don't match, ask for new ones.
            ' It's also possible that we only want to check previous privs,
            ' so we need a new option.
            CheckAccess = CheckPrivLevel(QueryPrivZone(Zone), LastLoginPriv)
            If LastLoginName <> SECURITY_NAME_EVERYBODY And Not AllowNewEntry Then
                ResetLastLoginExpiry(True)
                Exit Function
            End If
        End If

        ' New priv check.. clear the record of last logins.
        LastLoginName = ""
        LastLoginPriv = ""
        CheckAccess = False
        UpdatePermissionMonitor(Zone)

        Dim Pwd As String
        ' Get a (username and) password.
        ' For View Cost/GM, it is desirable to use the Last Login privs rather than re-prompt.
        If UCase(Left(Zone, 7)) <> "PREVENT" Then
            If Popup Or Not MainMenu.Visible Then
                Pwd = Password.GetPassword(, Reason, Zone)
            Else
                'PasswordTextbox.ToolTipText = "Enter Password For: " & Zone
                PasswordTextbox.Text = ""
                PasswordTextbox.Visible = True
                PasswordTextbox.Select()
                PasswordCommandButton.Visible = True
                Do While PasswordTextbox.Visible = True
                    Application.DoEvents() ' Yield to other processes.
                Loop
                PasswordCommandButton.Visible = False
                PasswordTextbox.Visible = False
                Pwd = PasswordTextbox.Text
            End If
        End If
        If Pwd = "" Then Exit Function                ' No access with no password.

        If Backdoor(Pwd) Then
            LastLoginName = BACKDOOR_USERNAME
            LastLoginPriv = "A"
            LastLoginExpiry = DateAdd("s", SECURITY_PASSWORD_TIMEOUT, Now)
            UpdatePermissionMonitor(Zone)
            CheckAccess = True
            Exit Function
        End If

        Dim Emp As clsEmployee, FromStore as integer, ToStore as integer, I as integer

        If CheckAllStores Then
            FromStore = 1
            ToStore = NoOfActiveLocations
        Else
            FromStore = StoresSld
            ToStore = StoresSld
        End If

        For I = FromStore To ToStore
            Emp = New clsEmployee
            If FileExists(GetDatabaseAtLocation(I)) Then
                Emp.DataAccess.DataBase = GetDatabaseAtLocation(I)
                Emp.DataAccess.Records_Open()

                Do While Emp.DataAccess.Records_Available
                    ' For each matching account, check group memberships.
                    If Emp.Active Then  ' Inactive accounts never work!

                        PwdOK = False
                        If Emp.Password = Pwd Then PwdOK = True
                        If IsDevelopment() And Left(Pwd, 3) = "*&*" And Mid(Pwd, 4) = Emp.SalesID Then PwdOK = True
                        If IsDevelopment() And Left(Pwd, 3) = "*&[" And LCase(Mid(Pwd, 4)) = LCase(Emp.LastName) Then PwdOK = True

                        If PwdOK Then
                            If CheckPrivLevel(TargetZone, Emp.Privs) Then
                                LastLoginName = Emp.LastName
                                LastLoginPriv = Emp.Privs
                                LastLoginExpiry = DateAdd("s", SECURITY_PASSWORD_TIMEOUT, Now)
                                UpdatePermissionMonitor(Zone)
                                CheckAccess = True
                                Exit Function
                            End If
                        End If

                    End If
                Loop
                DisposeDA(Emp)
            End If
        Next
    End Function

    Private Function QueryPrivZone(ByRef Zone As String) as integer
        Dim I as integer
        LoadPermOptions
        For I = LBound(PermOptions) To UBound(PermOptions)
            If Trim(PermOptions(I).OptionDesc) = Trim(Zone) Then
                QueryPrivZone = PermOptions(I).OptionID    ' ASCII Code for the zone.
                Exit Function
            End If
        Next
    End Function

    Private Function CheckPrivLevel(ByRef PrivZone as integer, ByRef PrivHas As String) As Boolean
        Dim I as integer, Priv As String
        If PrivHas = "" Then Exit Function                ' No privs supplied - fail.
        If InStr(PrivHas, "A") > 0 Then                   ' Admin - ok.
            If Left(Trim(QueryPrivDesc(PrivZone)), 7) = "Prevent" Then           ' But prevent nothing
                CheckPrivLevel = False
                Exit Function
            Else                                            ' Allow everything else
                CheckPrivLevel = True
                Exit Function
            End If
        End If
        For I = 1 To Len(PrivHas)
            Priv = Mid(PrivHas, I, 1)
            If InStr(QueryUserGroupPrivString(StoresSld, Priv), Chr(PrivZone)) > 0 Then
                CheckPrivLevel = True
                Exit Function
            End If
        Next
    End Function

    Public ReadOnly Property AllowUseLastEntry() As Boolean
        Get
            ' BFH20120531 - Enable this to allow 3-minutes timeout for passwords
            AllowUseLastEntry = True
            '  AllowUseLastEntry = IsBFMyer Or IsLapeer Or IsStateLine Or IsCranes
            '  If IsDevelopment Then AllowUseLastEntry = True
        End Get
    End Property

    Public Function RequestManagerApproval(ByRef TargetZone As String, Optional ByVal UseLastLoginName As Boolean = True) As Boolean
        '::::RequestManagerApproval
        ':::SUMMARY
        ': Requesting Manager Approval
        ':::DESCRIPTION
        ': Raise a prompt to request managerial approval for an operation that the current user does not have.
        ':::PARAMETERS
        ': - TargetZone
        ': - UseLastLoginName
        ':::CHANGES
        ': - BFH20050202 - Added RequestManagerApproval instead of CheckAccess for things like 'discounts'
        ':::RETURN
        ': Boolean

        Dim OldLogin As String, OldPriv As String ' , OldZone As String  -- Old zone wasn't stored
        Dim NewLogin As String, NewPriv As String
        RequestManagerApproval = False

        If UseLastLoginName And CheckAccess(TargetZone, False, True, False) Then  ' first check if they have access...
            RequestManagerApproval = True
        Else  ' then check if anyone else does...
            OldLogin = LastLoginName  ' store old login info because check access in this mode changes it..
            OldPriv = LastLoginPriv
            CheckAccess(TargetZone, True, True, True)  ' try to log in as someone else..
            RequestManagerApproval = CheckAccess(TargetZone, False, True, False)  'store their ability to access that zone
            NewLogin = LastLoginName
            NewPriv = LastLoginPriv
            LastLoginName = OldLogin  ' restore the original login's info
            LastLoginPriv = OldPriv
            UpdatePermissionMonitor(TargetZone, NewLogin, NewPriv)
        End If
    End Function

    Public ReadOnly Property GetCashierName() As String
        Get
            '::::GetCashierName
            ':::SUMMARY
            ': Return Cashier Name
            ':::DESCRIPTION
            ': Returns the current Cashier name based on user logged in.
            ': If Cashier Name is empty, then it returns Local Computer Name.
            ':::RETURN
            ': String
            GetCashierName = GetLastLoginName
            If IsIn(GetCashierName, SECURITY_NAME_BACKDOOR, SECURITY_NAME_DEMO, SECURITY_NAME_EVERYBODY) Then GetCashierName = ""
            If GetCashierName = "" Then
                'GetCashierName = GetLocalComputerName() -> GetLocalComputerName() throwing memory error. So replace this line with the below line to store local computer name as cashier name.
                GetCashierName = GetCDSSetting("Terminal", GetLocalComputerName)
                If GetCashierName <> "" And SecurityLevel = ComputerSecurityLevels.seclevOfficeComputer Then GetCashierName = "[Office PC]"
                If GetCashierName <> "" Then GetCashierName = "[" & GetCashierName & "]"
            End If
        End Get
    End Property

    Private Sub UpdatePermissionMonitor(ByVal Zone As String, Optional ByVal ManagerName As String = "", Optional ByVal ManagerPerms As String = "")
        If IsFormLoaded("frmPermissionMonitor") Then
            frmPermissionMonitor.txtUser.Text = LastLoginName
            frmPermissionMonitor.txtGroups.Text = LastLoginPriv
            frmPermissionMonitor.txtLastZone.Text = Zone
            If ManagerName <> "" Or ManagerPerms <> "" Then
                frmPermissionMonitor.txtManagerName.Text = ManagerName
                frmPermissionMonitor.txtManagerGroups.Text = ManagerPerms
                frmPermissionMonitor.txtManagerName.Visible = True
                frmPermissionMonitor.lblManagerName.Visible = True
                frmPermissionMonitor.txtManagerGroups.Visible = True
                frmPermissionMonitor.lblManagerGroups.Visible = True
            Else
                frmPermissionMonitor.txtManagerName.Visible = False
                frmPermissionMonitor.lblManagerName.Visible = False
                frmPermissionMonitor.txtManagerGroups.Visible = False
                frmPermissionMonitor.lblManagerGroups.Visible = False
            End If
        End If
    End Sub
    Private Function PasswordTextbox() As TextBox
        PasswordTextbox = MainMenu.txtPassword
    End Function
    Private Function PasswordCommandButton() As Button
        PasswordCommandButton = MainMenu.cmdEnterPassword
    End Function
    Public Function Backdoor(ByVal Pwd As String) As Boolean
        '::::Backdoor
        ':::SUMMARY
        ': Test Is Backdoor Password
        ':::DESCRIPTION
        ': Tests password for match against backdoor password.
        ': Backdoor password is now customizable based on store due to regular release of these universal credentials to store owners, who then gave them out to employees.
        ':::PARAMETERS
        ': - Pwd
        ':::RETURN
        ': Boolean
        Dim Mode as integer
        Mode = BACKDOOR_1   ' default is to only use password #1
        Pwd = UCase(Pwd)

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''           CUSTOMIZATIONS             '''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '  If IsAdams Then Mode = BACKDOOR_1 OR BACKDOOR_2       ' ALLOWS BOTH
        '  If IsAdams Then Mode = BACKDOOR_2                     ' ALLOWS ONLY NEW #2

        '  If IsAdams Then Mode = BACKDOOR_2
        '  If IsFurnitureDepot Then Mode = BACKDOOR_2            ' 20111217BFH
        '  If IsWarehouseFurniture Then Mode = BACKDOOR_2        ' 20120112BFH
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If Not Backdoor And (Mode And BACKDOOR_1) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORD1))
        If Not Backdoor And (Mode And BACKDOOR_2) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORD2))
        If Not Backdoor And (Mode And BACKDOOR_3) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORD3))
        If Not Backdoor And (Mode And BACKDOOR_4) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORD4))
        If Not Backdoor And (Mode And BACKDOOR_5) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORD5))
        If Not Backdoor And (Mode And BACKDOOR_6) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORD6))
        If Not Backdoor And (Mode And BACKDOOR_7) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORD7))
        If Not Backdoor And (Mode And BACKDOOR_8) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORD8))
        If Not Backdoor And (Mode And BACKDOOR_9) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORD9))
        If Not Backdoor And (Mode And BACKDOOR_A) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORDA))
        If Not Backdoor And (Mode And BACKDOOR_B) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORDB))
        If Not Backdoor And (Mode And BACKDOOR_C) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORDC))
        If Not Backdoor And (Mode And BACKDOOR_D) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORDD))
        If Not Backdoor And (Mode And BACKDOOR_E) Then Backdoor = (Pwd = UCase(BACKDOOR_PASSWORDE))
        If Not Backdoor And (Mode And BACKDOOR_F) Then Backdoor = (Pwd = BACKDOOR_PASSWORDF)
        If Not Backdoor And (Mode And BACKDOOR_P) Then Backdoor = (Pwd = DecodeBase64String(BACKDOOR_PROGRAMMER))
    End Function
    Public Sub LoadPermOptions()
        '::::LoadPermOptions
        ':::SUMMARY
        ': Loads Permission Options
        ':::DESCRIPTION
        ': Loads permission zones.  Called from initial program startup.
        If PermsLoaded Then Exit Sub
        PermsLoaded = True
        ' Current Max Option ID: 38
        ' Option IDs can never be reused!
        ' Option IDs are limited to the range 1-255.
        AddPermOption(1, "Program Administration")
        AddPermOption(2, "  Store Setup")
        AddPermOption(3, "  Backup/Restore")
        AddPermOption(4, "  Annual Maintenance")
        AddPermOption(33, "  Log In To Other Stores")
        AddPermOption(5, "Customer Management")
        AddPermOption(6, "  Create Sales")
        AddPermOption(7, "  View Sales")
        AddPermOption(8, "  Adjust Sales")
        AddPermOption(37, "  Order Status")
        AddPermOption(30, "  Give Discounts")
        AddPermOption(36, "  Prevent Price Adjust")
        AddPermOption(9, "  Deliver Sales")
        AddPermOption(10, "  Void Sales")
        AddPermOption(25, "  Service Orders")
        AddPermOption(11, "  Sales Reports")
        AddPermOption(12, "Inventory Management")
        AddPermOption(13, "  Create and Edit Items")
        AddPermOption(14, "  Change Item Prices")
        AddPermOption(31, "  Factory Shipments")             ' BFH20060403
        AddPermOption(32, "  Store Transfers")               ' BFH20060403
        AddPermOption(15, "  Change Stock Quantities")
        AddPermOption(16, "  View Stock Quantities")
        AddPermOption(17, "  View Cost and Gross Margin")
        AddPermOption(38, "  View Landed Cost")
        AddPermOption(18, "  Manage Purchase Orders")
        AddPermOption(24, "  Schedule Deliveries")
        AddPermOption(19, "  View Inventory Reports")
        AddPermOption(20, "Financial Management")
        AddPermOption(21, "  Accept Payments")
        AddPermOption(28, "  Cash Drawer")
        AddPermOption(26, "  Change Payment Dates")
        AddPermOption(34, "  Change Sale Date")
        AddPermOption(27, "  Forfeit Deposits")
        AddPermOption(22, "  Credit Administration")
        AddPermOption(23, "  Store Finances")
        AddPermOption(35, "  Commissions")
        AddPermOption(29, "  Daily Audit Report")
    End Sub
    Private Function QueryPrivDesc(ByVal ZoneID as integer) As String
        Dim I as integer
        For I = LBound(PermOptions) To UBound(PermOptions)
            If Trim(PermOptions(I).OptionID) = ZoneID Then
                QueryPrivDesc = PermOptions(I).OptionDesc
                Exit For
            End If
        Next
    End Function
    Public ReadOnly Property GetLastLoginName() As String
        Get
            GetLastLoginName = LastLoginName
        End Get
    End Property
    Private Sub AddPermOption(ByRef OptNum as integer, ByRef OptDesc As String)
        On Error Resume Next
        'ReDim Preserve PermOptions(LBound(PermOptions) To UBound(PermOptions) + 1)
        ReDim Preserve PermOptions(0 To UBound(PermOptions) + 1)
        If Err.Number <> 0 Then ReDim PermOptions(0 To 0)
        PermOptions(UBound(PermOptions)).OptionDesc = OptDesc
        PermOptions(UBound(PermOptions)).OptionID = OptNum
    End Sub

    Public Function Encrypt(ByVal UserName As String, ByVal Plaintext As String) As String
        '::::Encrypt
        ':::SUMMARY
        ': Encrypt Password
        ':::DESCRIPTION
        ': Returns an encrypted password, salted by username.
        ':::EXAMPLE
        ': - ?decryptpassword("kroll",encryptpassword("kroll","elvish parsley"))
        ':    - "elvish parsley"
        ':::PARAMETERS
        ': - UserName
        ': - Plaintext
        ':::SEE ALSO
        ': - Decrypt
        ':::RETURN
        ': String

        ' Takes a username and plaintext password.
        ' Returns the password, encrypted.
        ' To encrypt a password, randomize based on the sum of asc(UserName).
        ' This is vulnerable to names which add to the same value!
        On Error GoTo ErrCrypt
        Rnd(-1)                            ' Prepare the randomizer for a repeatable sequence.
        Randomize(StringValue(UserName))   ' Initialize the repeatable sequence by username.
        Dim I As Integer
        For I = 1 To 16 : Rnd()
        Next        ' Discard the first 16 random numbers.
        For I = 1 To Len(Plaintext)       ' Encrypt each character in the password.
            Encrypt = Encrypt & CryptCharacter(Mid(Plaintext, I, 1), Rnd, True)
        Next
        Exit Function

ErrCrypt:
        MsgBox("Error encrypting " & Plaintext & " for " & UserName & "." & vbCrLf &
    "Please contact " & AdminContactString(Format:=1, Phone:=False) & " immediately!", vbCritical, ProgramErrorTitle)
    End Function

    Private Function StringValue(ByVal Inp As String) As Integer
        Dim I As Integer
        For I = 1 To Len(Inp)
            StringValue = StringValue + Asc(Mid(Inp, I, 1))
        Next
    End Function

    Private Function CryptCharacter(ByRef Charr As String, ByRef Seed As Double, ByRef En As Boolean) As String
        '  Debug.Print IIf(En, "+", "-"), Asc(Char), Char, Seed,
        Dim CC As Integer
        If En Then
            CC = (Asc(Charr) + Seed * 255) Mod 255
            '    Debug.Print CC, Chr(CC)
            CryptCharacter = Chr(CC)
        Else
            CC = (Asc(Charr) - Seed * 255 + 255) Mod 255
            '    Debug.Print CC, Chr(CC)
            CryptCharacter = Chr(CC)
        End If
    End Function

    Public Function Decrypt(ByVal UserName As String, ByVal EncText As String) As String
        '::::Decrypt
        ':::SUMMARY
        ': Decrypt Password
        ':::DESCRIPTION
        ': Decodes a password (salted by username)
        ': - UserName
        ': - EncText
        ':::PARAMETERS
        ': - UserName
        ': - EncText
        ':::SEE ALSO
        ': - Encrypt
        ':::RETURN
        ': String

        ' Takes a username and encrypted password.
        ' Returns the password, decrypted.
        If Len(EncText) = 0 Then Exit Function
        On Error GoTo ErrCrypt
        Rnd(-1)                            ' Prepare the randomizer for a repeatable sequence.
        Randomize(StringValue(UserName))   ' Initialize the repeatable sequence by username.
        Dim I As Integer
        For I = 1 To 16 : Rnd() :
        Next        ' Discard the first 16 random numbers.
        For I = 1 To Len(EncText)         ' Encrypt each character in the password.
            Decrypt = Decrypt & CryptCharacter(Mid(EncText, I, 1), Rnd, False)
        Next
        Exit Function

ErrCrypt:
        MsgBox("Error decrypting " & EncText & " for " & UserName & "." & vbCrLf &
    "Please contact " & AdminContactString(Format:=1, Phone:=False) & " immediately!", vbCritical, ProgramErrorTitle)
    End Function

    Public Sub LogOut()
        '::::LogOut
        ':::SUMMARY
        ': Log Out current user (if any)
        ':::DESCRIPTION
        ': This function is used to log out any current user
        'LastLoginExpiry = 0
        LastLoginExpiry = Nothing
        LastLoginPriv = ""
        LastLoginName = ""
        UpdatePermissionMonitor("", "", "")
        MainMenu.cmdLogout.Visible = False
    End Sub

    Public Sub ClearAccess()
        '::::ClearAccess
        ':::SUMMARY
        ': Clears last login
        ':::DESCRIPTION
        ': This function is used to clear last login user Access details like name etc.
        LastLoginName = ""
        LastLoginPriv = ""
        UpdatePermissionMonitor("")
    End Sub

    Public ReadOnly Property IsLoggedIn() As Boolean
        Get
            IsLoggedIn = Not LastLoginExpired
        End Get
    End Property

End Module
