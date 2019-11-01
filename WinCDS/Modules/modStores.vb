Imports stdole

Module modStores
    Public Const COPY_FILE As String = "File Copy"
    Public Const COPY_CUSTOMER As String = "Customer Copy"
    Public Const COPY_SALESMAN As String = "Salesman Copy"
    Public Const COPY_DELIVERY As String = "Delivery Copy"
    Public Const SS_RecLab_NONE As String = "None"
    Public Const SS_RecLab_ALL As String = "All"
    Public Const SS_RecLab_JUSTSTOCK As String = "Just Stock"
    Public Const PO_SPECINSTR_DEFAULT_1 As String = "If order is less than $300.00, HOLD AND SHIP With Other goods!"
    Public Const PO_SPECINSTR_DEFAULT_2 As String = "Ship UPS, PP, or with other goods!"
    Public Const PO_SPECINSTR_DEFAULT_3 As String = "Sold Orders must be shipped complete!"
    Public Const PO_SPECINSTR_DEFAULT_4 As String = ""

    Public ReadOnly Property ssMaxStore() As Integer
        Get
            ssMaxStore = LicensedNoOfStores()
        End Get
    End Property

    Public Function StoreLogoPicture(Optional ByVal StoreNum As Integer = 0) As IPictureDisp
        '::::StoreLogoPicture
        ':::SUMMARY
        ': Store Logo Picture
        ':::DESCRIPTION
        ': Returns the Picture object of Store Logo.  Returns Nothing if no logo.
        ':::PARAMETERS
        ': - StoreNum
        ':::RETURN
        ': IPictureDisp

        On Error Resume Next
        If StoreNum = 0 Then StoreNum = StoresSld
        StoreNum = FitRange(1, StoreNum, Setup_MaxStores)
        StoreLogoPicture = LoadPictureStd(StoreLogoFile(StoreNum))
    End Function
    Public Function DefaultMailingLabelType() As String
        '::::DefaultMailingLabelType
        ':::SUMMARY
        ': Default Mail Label Type
        ':::DESCRIPTION
        ': Returns Default Mail Type
        ':::RETURN
        ': - String

        If IsUFO() Then DefaultMailingLabelType = "30252" : Exit Function
        DefaultMailingLabelType = "30252"
        '  DefaultMailingLabelType = "30323"
    End Function
    Public Property SecurityLevel() As ComputerSecurityLevels
        Get
            Dim T As String
            T = GetCDSSetting("Location", ComputerSecurityLevels.seclevNoPasswords)
            SecurityLevel = IIf(T = "", ComputerSecurityLevels.seclevNoPasswords, Val(T))
            SecurityLevel = Val(GetCDSSetting("Location", ComputerSecurityLevels.seclevNoPasswords))
        End Get
        Set(value As ComputerSecurityLevels)
            SaveCDSSetting("Location", value)
        End Set
    End Property
    Public Function StoreLogoFile(Optional ByVal StoreNum As Integer = 0) As String
        '::::StoreLogoFile
        ':::SUMMARY
        ': Filename of store logo file
        ':::DESCRIPTION
        ': Returns filename of store logo file.  Will return even if file does not exist.
        ':::PARAMETERS
        ': - StoreNum
        ':::RETURN
        ': String
        If StoreNum = 0 Then StoreNum = StoresSld
        If StoreNum <= 0 Or StoreNum >= Setup_MaxStores Then StoreNum = 1
        StoreLogoFile = FXFile("Store" & StoreNum & "Logo")
    End Function
    Public Function BOSFile(Optional ByVal StoreNum As Integer = 0) As String
        '::::BOSFile
        ':::SUMMARY
        ': Bill of Sale File
        ':::DESCRIPTION
        ': Filename of Sale Number Autonumber file.
        ':::PARAMETERS
        ': - StoreNum
        ':::RETUTN
        ': String
        If StoreNum = 0 Then StoreNum = StoresSld
        If StoreNum <= 0 Or StoreNum >= Setup_MaxStores Then StoreNum = 1
        BOSFile = NewOrderFolder(StoreNum) & "BillSale.Dat"
    End Function
    Public ReadOnly Property PasswordProtectedDatabase() As Boolean
        Get
            PasswordProtectedDatabase = DatabasePassword <> ""
        End Get
    End Property
    Public ReadOnly Property PasswordProtectedDatabaseString() As String
        Get
            If Not PasswordProtectedDatabase Then Exit Property
            PasswordProtectedDatabaseString = ";Jet OLEDB:Database Password=" & DatabasePassword
        End Get
    End Property
    Public Property DatabasePassword() As String
        Get
            DatabasePassword = mDatabasePassword
        End Get
        Set(value As String)
            Dim OldPassword As String
            OldPassword = DatabasePassword
            'UpdateAllDBPasswords(value, OldPassword)
            mDatabasePassword = value
        End Set
    End Property
    Public Property mDatabasePassword() As String
        Get
            mDatabasePassword = DecodeBase64String(GetCDSSetting("DB-Password"))
        End Get
        Set(value As String)
            SaveCDSSetting("DB-Password", EncodeBase64String(value))
            WriteStoreSetting(-1, IniSections_StoreSettings.iniSection_StoreSettings, "DB-Password", EncodeBase64String(value))

        End Set

    End Property

    Public Function ArNoFile(Optional ByVal StoreNum As Integer = 0) As String
        '::::ArNoFile
        ':::SUMMARY
        ': Installment Autonumber File
        ':::DESCRIPTION
        ': Returns autonumber filename of next ArNo (Installment) Contract
        ':::PARAMETERS
        ': - StoreNum
        ':::RETURN
        ': String
        If StoreNum = 0 Then StoreNum = StoresSld
        If StoreNum <= 0 Or StoreNum >= Setup_MaxStores Then StoreNum = 1
        ArNoFile = NewOrderFolder(StoreNum) & "ArNo.Dat"
    End Function

    Public Function CustRecFile(Optional ByVal StoreNum As Integer = 0) As String
        '::::CustRecFile
        ':::SUMMARY
        ': Customer Record File
        ':::DESCRIPTION
        ': File Autonumber for Mail Table
        ':::PARAMETERS
        ': - StoreNum
        ':::RETURN
        ': String
        If StoreNum = 0 Then StoreNum = StoresSld
        If StoreNum <= 0 Or StoreNum >= Setup_MaxStores Then StoreNum = 1
        CustRecFile = NewOrderFolder(StoreNum) & "CustRec.Dat"
    End Function

    '    Public Property Get CashDrawerCOMPort() as integer
    '' If you are using our Steel Cash Drawer, you need to select the COM Port to make the Drawer open when you tender cash, or indicate it is a USB Cash Drawer if so equipped..
    '' If you do not know your COM port, you can usually find it quickest just by trying them until you find which one works. Start with COM Port 1 and move up.
    '  CashDrawerCOMPort = Val(GetCDSSetting("Cash Drawer COM Port", 0))
    '  If CashDrawerCOMPort <> 0 Then Exit Property

    '  CashDrawerCOMPort = Val(GetCDSSetting("COM Port", 0, "Cash Drawer"))
    '  If CashDrawerCOMPort <> 0 Then
    '    CashDrawerCOMPort = CashDrawerCOMPort

    '    DeleteSystemSetting RegistrySection, RegistryAppName & "\Cash Drawer", "COM Port"
    '    DeleteSystemKey RegistrySection, RegistryAppName, "Cash Drawer"
    '  End If
    '    End Property
    '    Public Property Let CashDrawerCOMPort(ByVal vData as integer)
    '  SaveCDSSetting "Cash Drawer COM Port", vData
    'End Property

    Public Property CashDrawerCOMPort() As Integer
        Get
            ' If you are using our Steel Cash Drawer, you need to select the COM Port to make the Drawer open when you tender cash, or indicate it is a USB Cash Drawer if so equipped..
            ' If you do not know your COM port, you can usually find it quickest just by trying them until you find which one works. Start with COM Port 1 and move up.
            CashDrawerCOMPort = Val(GetCDSSetting("Cash Drawer COM Port", 0))
            If CashDrawerCOMPort <> 0 Then Exit Property

            CashDrawerCOMPort = Val(GetCDSSetting("COM Port", 0, "Cash Drawer"))
            If CashDrawerCOMPort <> 0 Then
                CashDrawerCOMPort = CashDrawerCOMPort

                DeleteSystemSetting(RegistrySection, RegistryAppName & "\Cash Drawer", "COM Port")
                DeleteSystemKey(RegistrySection, RegistryAppName, "Cash Drawer")
            End If
        End Get
        Set(value As Integer)
            SaveCDSSetting("Cash Drawer COM Port", value)
        End Set
    End Property

    Public Function BOSDiscountFile(Optional ByVal StoreNum As Integer = 0) As String
        '::::BOSDiscountFile
        ':::SUMMARY
        ': Bill of Sale Discount File
        ':::DESCRIPTION
        ': Filename of current Bill of Sale Discounts
        ':::PARAMETERS
        ': - StoreNum
        ':::RETURN
        ': String
        If StoreNum = 0 Then StoreNum = StoresSld
        If StoreNum <= 0 Or StoreNum >= Setup_MaxStores Then StoreNum = 1
        BOSDiscountFile = NewOrderFolder(StoreNum) & "BoSDscnt.Dat"
    End Function
End Module
