Module modStoreSettings
    Public Enum IniSections_StoreSettings
        iniSection_UNKNOWN = 0
        iniSection_StoreSettings = 1
        iniSection_SaleOptions = 2
        iniSection_POOptions = 3
        iniSection_Commissions = 4
        iniSection_TagOptions = 5
        iniSection_Installment = 6
        iniSection_Equifax = 7
        iniSection_PayPal = 8
        iniSection_Ashley = 9
        iniSection_Amazon = 10
        iniSection_DispatchTrack = 11
        iniSection_TRAX = 12
        iniSection_Program = 13
        iniSection_Custom = 14

        ' Virtual INI Sections (used for Installing an INI to alter WinCDS)
        iniSection_InstallSQL = 250
    End Enum
    Public Structure StoreInfo
        Dim refLoaded As Boolean
        Dim refFileDate As Date
        Dim refFileSize As Integer
        Dim refNextCheck As Date

        Dim Name As String
        Dim Address As String
        Dim City As String
        Dim Phone As String
        Dim CommCode As Integer
        Dim FabSeal As Double
        Dim SalesTax As Double

        Dim StoreShipToName As String
        Dim StoreShipToAddr As String
        Dim StoreShipToCity As String
        Dim StoreShipToTele As String

        Dim bPOInvoicesLocation As Boolean
        Dim bDeliveryTaxable As Boolean
        Dim bLaborTaxable As Boolean
        Dim bPicturesOnTags As Boolean

        Dim TagJustify As String

        Dim bNoMerchandisePrice As Boolean
        Dim bNoListPriceOnTags As Boolean

        Dim CalculateList As Integer
        Dim PrintCopies As Integer

        Dim loadedLicense As String

        Dim bPaymentBooksMonthly As Boolean

        Dim SalesCopyID() As String      ' 25-28

        Dim bDelIsCommissionable As Boolean
        Dim bLabIsCommissionable As Boolean
        Dim bNotIsCommissionable As Boolean ' 31

        ' 32 is blank
        ' 33 duplicates 50

        Dim GracePeriod As Integer
        Dim ReceivingLabels As String

        Dim bAPPost As Boolean            ' 36
        Dim bBankManagerPost As Boolean

        Dim bPrintPoNoCost As Boolean

        Dim UseCost As String

        Dim bPOSpecialInstr As Boolean

        Dim loadedInstallmentLicense As String ' 41

        Dim SimpleInterestRate As Double
        Dim DocFee As Decimal
        Dim LateChargePer As Double
        Dim MaxLateCharge As Decimal

        Dim bPrintPaymentBooks As Boolean ' 46
        Dim bManualBillofSaleNo As Boolean
        Dim bShowRegularPrice As Boolean
        Dim bShowManufacturer As Boolean
        Dim bUseStoreCode As Boolean
        Dim bStyleNoInCode As Boolean
        Dim bCostInCode As Boolean

        Dim bShowAvailableStock As Boolean
        Dim bPrintBarCode As Boolean
        Dim bSellFromLoginLocation As Boolean

        Dim MinLateCharge As Decimal

        Dim bPostToLoc1 As Boolean        ' 57

        Dim bRequireAdvertising As Boolean
        Dim bTagIncommingDistinct As Boolean
        Dim bUseCCMachine As Boolean
        Dim bStartMaximized As Boolean
        Dim bUseBTScanner As Boolean


        Dim PoSpecInstr1 As String        ' 63
        Dim PoSpecInstr2 As String
        Dim PoSpecInstr3 As String
        Dim PoSpecInstr4 As String
        Dim DelPercent As String

        Dim bUseCashRegisterAddress As Boolean
        Dim CashRegisterReceiptTailMessage As String

        Dim Email As String               ' 69
        Dim bUseQB As Boolean

        Dim bSeparateCommTables As Boolean
        Dim bOneCalendar As Boolean       ' 72
        Dim EquifaxAccountNo As String    ' 73

        Dim bInstallmentInterestIsTaxable As Boolean '74
        Dim EquifaxSecurityCode As String ' 75
        Dim bAPR As Boolean               ' 76
        Dim TransUnionAcctNo As String    ' 77

        Dim bUseTimeWindows As Boolean

        Dim ExperianAcctNo As String      ' 79
        Dim CompanyIdent As String        ' 80

        Dim bJointLife As Boolean         ' 81

        Dim AshleyID As String            ' 82
        Dim AshleyUName As String         ' 83
        Dim AshleyPWord As String         ' 84
        Dim AshleyPath As String          ' 85

        Dim AshleyATPExternalID As String
        Dim AshleyATPKeyCode As String
        Dim AshleyATPUserName As String
        Dim AshleyATPPassword As String
        Dim AshleyATPPICode As String
        Dim AshleyATPShipToID As String

        Dim bEmailLateChargeNotices As Boolean    ' 86
        Dim bEmailMonthlyStatements As Boolean    ' 87

        Dim bRequestEmail As Boolean      ' 88

        Dim EquifaxSiteID As String       ' 89
        Dim EquifaxPassword As String     ' 90

        Dim CCProcessor As String         ' 91
        Dim CCConfig As String            ' 92

        Dim PayPalUsername As String      ' 93
        Dim PayPalPassword As String      ' 94
        Dim PayPalSignature As String     ' 95
        Dim PayPalAuthKey As String       ' 96

        '  bNoListPriceOnTags As Boolean        ' 97
        Dim bShowPackageItemPrices As Boolean     ' 98

        Dim CashDrawerConfig As String            ' 99

        Dim ServerLock As Boolean                 ' 100
        Dim ModifiedRevolvingCharge As Boolean    ' 101
        Dim ModifiedRevolvingRate As Double       ' 102
        Dim ModifiedRevolvingAPR As Boolean       ' 103
        Dim ModifiedRevolvingSameAsCash As Integer   ' 104
        Dim ModifiedRevolvingMinPmt As Double     ' 105

        ' We used to use a sequential file (.DAT), so we had an Index for each field above.
        ' Now, we have a proper structure, and store the values in an INI file, so we no longer keep count.
        ' The count above is still somewhat useful for upgrading old versions, however (.DAT -> .INI)
        ' But, for everything beyond this line, order no longer matters.

        Dim AmazonKeyID As String
        Dim AmazonSecretKey As String
        Dim AmazonUserName As String
        Dim AmazonCustomerBucket As String
        Dim AmazonPassword As String
        Dim AmazonAWSPanelConfig As String
        Dim AmazonQBPath As String
        Dim AmazonMisc As String

        Dim DispatchTrackLicense As String
        Dim DispatchTrackServiceCode As String
        Dim DispatchTrackServiceAPI As String
        Dim DispatchTrackServiceURL As String

        Dim EquifaxVendorIDCode As String

        Dim TRAXLicense As String
        Dim TRAXID As String

        Dim bInstallmentRoundUp As Boolean

        Dim bPostTermInterest As Boolean
        Public PostTermInterestRate As Double

        ' +++ This is the end of the StoreInfo type..  If you don't know where else to add the new setting, add it here.  Then, find the 5 other places in the code and add it there.

    End Structure
    Private mStoreSettings(Setup_MaxStores) As StoreInfo
    Public Const STORE_INI_SECTION_STORE_SETTINGS As String = "Store Settings"
    Public Const STORE_INI_SECTION_SALE_OPTIONS As String = "Sale Options"
    Public Const STORE_INI_SECTION_PO_OPTIONS As String = "PO Options"
    Public Const STORE_INI_SECTION_COMMISSIONS_OPTIONS As String = "Commissions Options"
    Public Const STORE_INI_SECTION_TAG_OPTIONS As String = "Price Tag Options"
    Public Const STORE_INI_SECTION_INSTALLMENT_OPTIONS As String = "Installment Options"
    Public Const STORE_INI_SECTION_PROGRAM_STATE As String = "Program"
    Public Const STORE_INI_SECTION_CUSTOM_VALUES As String = "Custom"
    Public Const STORE_INI_SECTION_EQUIFAX_OPTIONS As String = "Service-Equifax"
    Public Const STORE_INI_SECTION_PAYPAL_OPTIONS As String = "Service-PayPal"
    Public Const STORE_INI_SECTION_ASHLEY_OPTIONS As String = "Service-Ashley"
    Public Const STORE_INI_SECTION_AMAZON_OPTIONS As String = "Service-Amazon"
    Public Const STORE_INI_SECTION_DISPATCHTRACK_OPTIONS As String = "Service-DispatchTrack"
    Public Const STORE_INI_SECTION_TRAX_OPTIONS As String = "Service-TRAX"

    Public Const STORE_INI_SECTION_INSTALL_SQL As String = "Install-SQL"


    Private Function CreateStoreInformationFile(Optional ByVal nStoreNo As Integer = 0, Optional ByVal StoreFileName As String = "") As String
        If StoreFileName = "" Then StoreFileName = StoreFile(nStoreNo)

        If Not DirExists(GetFilePath(StoreFileName)) Then
            ' If the actual directory doesn't exist, then what?
        End If

        If Not FileExists(StoreFileName) Then
            WriteFile(StoreFileName, "[" & STORE_INI_SECTION_STORE_SETTINGS & "]" & vbCrLf & "Name=Store #" & nStoreNo & vbCrLf2, True, True)
        End If
        CreateStoreInformationFile = IIf(FileExists(StoreFileName), StoreFileName, "")
    End Function

    Public Function StoreSettings(Optional ByVal nStoreNo As Integer = 0, Optional ByVal doReset As Boolean = False, Optional ByVal Optimize As Boolean = False) As StoreInfo

        'StoreSettings = Nothing
        Dim FN As String, DoCreate As Boolean
        'Dim Stale As Boolean. variable not used anywhere in this function.
        Dim Succeeded As Boolean, Tries As Integer, NextTry As Date

        If nStoreNo <= 0 Then nStoreNo = StoresSld
        If nStoreNo <= 0 Then nStoreNo = 1
        If nStoreNo > Setup_MaxStores Then Exit Function

        If doReset Or Not mStoreSettings(nStoreNo).refLoaded Then Optimize = False

        If Not Optimize Then
            FN = StoreFile(nStoreNo)
            If Not FileExists(FN) Then
                doReset = True
                If LicensedNoOfStores() >= nStoreNo Then DoCreate = True Else Exit Function
                CreateStoreInformationFile(nStoreNo)
                '    WriteFile Fn, "", True
            End If

            'If doReset Or DateAfter(Now, mStoreSettings(nStoreNo).refNextCheck, False, "n") Then ' reduces disk access...  1 check every 2 minutes
            If doReset Or DateAfter(Now, mStoreSettings(nStoreNo).refNextCheck, False, DateInterval.Minute) Then ' reduces disk access...  1 check every 2 minutes
                If doReset Or (Not DateEqual(FileDateTime(FN), mStoreSettings(nStoreNo).refFileDate) Or FileLen(FN) <> mStoreSettings(nStoreNo).refFileSize) Then
                    mStoreSettings(nStoreNo) = GetStoreInformation(nStoreNo, , Succeeded)
                    If Not Succeeded Then
                        NextTry = DateAdd("s", 1, Now)
                        Do While True
                            Application.DoEvents()

                            If DateAfter2(Now, NextTry, True, "s") Then
                                Tries = Tries + 1
                                If Tries > 3 Then
                                    MsgBox("Could not open store settings file:" & vbCrLf & StoreFile(nStoreNo), vbCritical, "Error Reading Settings")
                                    Exit Do
                                End If
                                mStoreSettings(nStoreNo) = GetStoreInformation(nStoreNo, , Succeeded)
                                If Succeeded Then Exit Do
                                NextTry = DateAdd("s", 1, Now)
                            End If
                        Loop
                    End If
                    mStoreSettings(nStoreNo).refFileDate = FileDateTime(FN)
                    mStoreSettings(nStoreNo).refFileSize = FileLen(FN)
                    mStoreSettings(nStoreNo).refNextCheck = DateAdd("n", 3, Now)
                    mStoreSettings(nStoreNo).refLoaded = True
                End If
            End If

            If DoCreate Then
                If mStoreSettings(nStoreNo).Name = "" Then mStoreSettings(nStoreNo).Name = "Store #" & nStoreNo
                SaveStoreInformationINI(mStoreSettings(nStoreNo), nStoreNo)
                '    CreateStoreInformationfile nStoreNo
            End If
        End If

        StoreSettings = mStoreSettings(nStoreNo)

    End Function

    Public Function StoreFile(Optional ByVal StoreNum As Integer = 0) As String
        StoreFile = StoreINIFile(StoreNum)
    End Function
    Public Function StoreINIFile(Optional ByVal StoreNum As Integer = 0) As String
        Dim F As String
        If StoreNum <= 0 Then StoreNum = StoresSld

        '  LogStartup "StoreINIFile(" & StoreNum & ")"
        '  LogStartup "StoreINIFile(" & StoreNum & ") - StoreFolder: " & StoreFolder(1)
        '  LogStartup "StoreINIFile(" & StoreNum & ") - E(StoreFolder): " & DirExists(StoreFolder(1))

        On Error Resume Next
        If Not DirExists(StoreFolder(1)) Then ' Always Loc 1
            MsgBox("Directory " & StoreFolder(1) & " must exist." & vbCrLf & "Many other things will most likely fail because this directory cannot be found.", vbCritical, ProgramName & " Critical Error")
        End If

        ' 20060623 - Moving to server for setup files..
        F = "Store" & StoreNum & ".ini"
        StoreINIFile = StoreFolder(1) & F

        '  LogStartup "StoreINIFile: Result=" & StoreINIFile
        If Not FileExists(StoreINIFile) Then
            If StoreNum > LicensedNoOfStores() Then Exit Function
            CreateStoreInformationFile(StoreNum, StoreINIFile)
        End If
    End Function
    Public Function StoreSettingSectionKey(ByVal nSection As IniSections_StoreSettings) As String
        Select Case nSection
'    case iniSection_UNKNOWN        ' DEFAULT TO Store Settings Ection

            Case IniSections_StoreSettings.iniSection_StoreSettings : StoreSettingSectionKey = STORE_INI_SECTION_STORE_SETTINGS
            Case IniSections_StoreSettings.iniSection_SaleOptions : StoreSettingSectionKey = STORE_INI_SECTION_SALE_OPTIONS
            Case IniSections_StoreSettings.iniSection_POOptions : StoreSettingSectionKey = STORE_INI_SECTION_PO_OPTIONS
            Case IniSections_StoreSettings.iniSection_Commissions : StoreSettingSectionKey = STORE_INI_SECTION_COMMISSIONS_OPTIONS
            Case IniSections_StoreSettings.iniSection_TagOptions : StoreSettingSectionKey = STORE_INI_SECTION_TAG_OPTIONS
            Case IniSections_StoreSettings.iniSection_Installment : StoreSettingSectionKey = STORE_INI_SECTION_INSTALLMENT_OPTIONS
            Case IniSections_StoreSettings.iniSection_Equifax : StoreSettingSectionKey = STORE_INI_SECTION_EQUIFAX_OPTIONS
            Case IniSections_StoreSettings.iniSection_PayPal : StoreSettingSectionKey = STORE_INI_SECTION_PAYPAL_OPTIONS
            Case IniSections_StoreSettings.iniSection_Ashley : StoreSettingSectionKey = STORE_INI_SECTION_ASHLEY_OPTIONS
            Case IniSections_StoreSettings.iniSection_Amazon : StoreSettingSectionKey = STORE_INI_SECTION_AMAZON_OPTIONS
            Case IniSections_StoreSettings.iniSection_DispatchTrack : StoreSettingSectionKey = STORE_INI_SECTION_DISPATCHTRACK_OPTIONS
            Case IniSections_StoreSettings.iniSection_TRAX : StoreSettingSectionKey = STORE_INI_SECTION_TRAX_OPTIONS
            Case IniSections_StoreSettings.iniSection_Program : StoreSettingSectionKey = STORE_INI_SECTION_PROGRAM_STATE
            Case IniSections_StoreSettings.iniSection_Custom : StoreSettingSectionKey = STORE_INI_SECTION_CUSTOM_VALUES

            Case IniSections_StoreSettings.iniSection_InstallSQL : StoreSettingSectionKey = STORE_INI_SECTION_INSTALL_SQL

            Case Else : StoreSettingSectionKey = STORE_INI_SECTION_STORE_SETTINGS
        End Select
    End Function
    Public Function ResetStoreSettings() As Boolean
        Dim Discard As Integer, I As Integer
        ResetStoreSettings = False
        For I = 1 To NoOfActiveLocations
            ' the field being read is irrelevant..  Just accessing with ...(I, True) resets the stored value.
            Discard = StoreSettings(I, True).GracePeriod
        Next
    End Function
    Public Function GetStoreInformation(Optional ByVal StoreNum As Integer = 0, Optional ByVal Quiet As Boolean = False, Optional ByRef Success As Boolean = False, Optional ByVal AltFileName As String = "") As StoreInfo
        GetStoreInformation = GetStoreInformationINI(StoreNum, Quiet, Success, AltFileName)
        '  If ReadIniValue(StoreINIFile(StoreNum), STORE_INI_SECTION_STORE_SETTINGS, "Name") = "" Then
        '    GetStoreInformation = GetStoreInformationOLD(StoreNum, Quiet, Success, AltFileName)
        '  Else
        '    GetStoreInformation = GetStoreInformationINI(StoreNum, Quiet, Success, AltFileName)
        '  End If
    End Function
    Private Function SaveStoreInformationINI(ByRef SI As StoreInfo, Optional ByVal StoreNum As Integer = 0, Optional ByVal AltFileName As String = "")
        Dim F As String, FF As Integer, I As Integer
        Dim Section As String

        If StoreNum = 0 Then StoreNum = StoresSld
        If AltFileName = "" Then F = StoreINIFile(StoreNum) Else F = AltFileName
        FF = FreeFile()

        With SI
            On Error Resume Next
            Section = STORE_INI_SECTION_STORE_SETTINGS
            WriteIniValue(F, Section, "License", License, True)


            WriteIniValue(F, Section, "Name", .Name, True)
            WriteIniValue(F, Section, "Address", .Address, True)
            WriteIniValue(F, Section, "City", .City, True)
            WriteIniValue(F, Section, "Phone", .Phone, True)

            WriteIniValue(F, Section, "Email", .Email, True)

            WriteIniValue(F, Section, "ShipToName", .StoreShipToName, True)
            WriteIniValue(F, Section, "ShipToAddr", .StoreShipToAddr, True)
            WriteIniValue(F, Section, "ShipToCity", .StoreShipToCity, True)
            WriteIniValue(F, Section, "ShipToTele", .StoreShipToTele, True)

            WriteIniValue(F, Section, "StartMaximized", .bStartMaximized, True)
            WriteIniValue(F, Section, "UseBTScanner", .bUseBTScanner, True)

            WriteIniValue(F, Section, "UseQB", .bUseQB, True)
            WriteIniValue(F, Section, "OneCalendar", .bOneCalendar, True)
            WriteIniValue(F, Section, "UseTimeWindows", .bUseTimeWindows, True)

            WriteIniValue(F, Section, "CCProcessor", .CCProcessor, True)
            WriteIniValue(F, Section, "CCConfig", .CCConfig, True)

            WriteIniValue(F, Section, "CashDrawerConfig", .CashDrawerConfig, True)
            WriteIniValue(F, Section, "ServerLock", .ServerLock, True)



            Section = STORE_INI_SECTION_SALE_OPTIONS
            WriteIniValue(F, Section, "FabSeal", .FabSeal, True)
            WriteIniValue(F, Section, "SalesTax", .SalesTax, True)

            WriteIniValue(F, Section, "DeliveryTaxable", .bDeliveryTaxable, True)
            WriteIniValue(F, Section, "LaborTaxable", .bLaborTaxable, True)

            WriteIniValue(F, Section, "NoMerchandisePrice", .bNoMerchandisePrice, True)
            WriteIniValue(F, Section, "CalculateList", .CalculateList, True)

            WriteIniValue(F, Section, "PrintCopies", .PrintCopies, True)

            WriteIniValue(F, Section, "SalesCopyID1", .SalesCopyID(0), True)
            WriteIniValue(F, Section, "SalesCopyID2", .SalesCopyID(1), True)
            WriteIniValue(F, Section, "SalesCopyID3", .SalesCopyID(2), True)
            WriteIniValue(F, Section, "SalesCopyID4", .SalesCopyID(3), True)



            WriteIniValue(F, Section, "ManualBillofSaleNo", .bManualBillofSaleNo, True)
            WriteIniValue(F, Section, "SellFromLoginLocation", .bSellFromLoginLocation, True)

            WriteIniValue(F, Section, "RequireAdvertising", .bRequireAdvertising, True)
            WriteIniValue(F, Section, "RequestEmail", .bRequestEmail, True)

            WriteIniValue(F, Section, "DelPercent", .DelPercent, True)
            WriteIniValue(F, Section, "UseCashRegisterAddress", .bUseCashRegisterAddress, True)
            WriteIniValue(F, Section, "CashRegisterReceiptTailMessage", .CashRegisterReceiptTailMessage, True)



            Section = STORE_INI_SECTION_PO_OPTIONS
            WriteIniValue(F, Section, "POInvoicesLocation", .bPOInvoicesLocation, True)
            WriteIniValue(F, Section, "ReceivingLabels", .ReceivingLabels, True)

            WriteIniValue(F, Section, "APPost", .bAPPost, True)
            WriteIniValue(F, Section, "BankManagerPost", .bBankManagerPost, True)

            WriteIniValue(F, Section, "POSpecialInstr", .bPOSpecialInstr, True)

            WriteIniValue(F, Section, "PostToLoc1", .bPostToLoc1, True)
            WriteIniValue(F, Section, "TagIncommingDistinct", .bTagIncommingDistinct, True)

            WriteIniValue(F, Section, "PoSpecInstr1", .PoSpecInstr1, True)
            WriteIniValue(F, Section, "PoSpecInstr2", .PoSpecInstr2, True)
            WriteIniValue(F, Section, "PoSpecInstr3", .PoSpecInstr3, True)
            WriteIniValue(F, Section, "PoSpecInstr4", .PoSpecInstr4, True)



            Section = STORE_INI_SECTION_COMMISSIONS_OPTIONS
            WriteIniValue(F, Section, "CommCode", .CommCode, True)
            WriteIniValue(F, Section, "SeparateCommTables", .bSeparateCommTables, True)

            WriteIniValue(F, Section, "DelIsCommissionable", .bDelIsCommissionable, True)
            WriteIniValue(F, Section, "LabIsCommissionable", .bLabIsCommissionable, True)
            WriteIniValue(F, Section, "NotIsCommissionable", .bNotIsCommissionable, True)



            Section = STORE_INI_SECTION_TAG_OPTIONS
            WriteIniValue(F, Section, "TagJustify", .TagJustify, True)
            WriteIniValue(F, Section, "PicturesOnTags", .bPicturesOnTags, True)
            WriteIniValue(F, Section, "UseStoreCode", .bUseStoreCode, True)
            WriteIniValue(F, Section, "PrintPoNoCost", .bPrintPoNoCost, True)

            WriteIniValue(F, Section, "UseCost", .UseCost, True)

            WriteIniValue(F, Section, "ShowRegularPrice", .bShowRegularPrice, True)
            WriteIniValue(F, Section, "ShowManufacturer", .bShowManufacturer, True)
            WriteIniValue(F, Section, "UseStoreCode", .bUseStoreCode, True)
            WriteIniValue(F, Section, "StyleNoInCode", .bStyleNoInCode, True)
            WriteIniValue(F, Section, "CostInCode", .bCostInCode, True)
            WriteIniValue(F, Section, "ShowAvailableStock", .bShowAvailableStock, True)
            WriteIniValue(F, Section, "PrintBarCode", .bPrintBarCode, True)

            WriteIniValue(F, Section, "NoListPriceOnTags", .bNoListPriceOnTags, True)
            WriteIniValue(F, Section, "ShowPackageItemPrices", .bShowPackageItemPrices, True)



            Section = STORE_INI_SECTION_INSTALLMENT_OPTIONS
            WriteIniValue(F, Section, "Installment License", InstallmentLicense, True)

            WriteIniValue(F, Section, "PaymentBooksMonthly", .bPaymentBooksMonthly, True)
            WriteIniValue(F, Section, "GracePeriod", .GracePeriod, True)
            WriteIniValue(F, Section, "LateChargePer", .LateChargePer, True)

            WriteIniValue(F, Section, "SimpleInterestRate", .SimpleInterestRate, True)
            WriteIniValue(F, Section, "DocFee", .DocFee, True)

            WriteIniValue(F, Section, "MaxLateCharge", .MaxLateCharge, True)
            WriteIniValue(F, Section, "PrintPaymentBooks", .bPrintPaymentBooks, True)

            WriteIniValue(F, Section, "MinLateCharge", .MinLateCharge, True)

            WriteIniValue(F, Section, "InstallmentInterestIsTaxable", .bInstallmentInterestIsTaxable, True)
            WriteIniValue(F, Section, "APR", .bAPR, True)
            WriteIniValue(F, Section, "JointLife", .bJointLife, True)

            WriteIniValue(F, Section, "EmailLateChargeNotices", .bEmailLateChargeNotices, True)
            WriteIniValue(F, Section, "EmailMonthlyStatements", .bEmailMonthlyStatements, True)

            WriteIniValue(F, Section, "ModifiedRevolvingCharge", .ModifiedRevolvingCharge, True)
            WriteIniValue(F, Section, "ModifiedRevolvingRate", .ModifiedRevolvingRate, True)
            WriteIniValue(F, Section, "ModifiedRevolvingAPR", .ModifiedRevolvingAPR, True)
            WriteIniValue(F, Section, "ModifiedRevolvingSameAsCash", .ModifiedRevolvingSameAsCash, True)
            WriteIniValue(F, Section, "ModifiedRevolvingMinPmt", .ModifiedRevolvingMinPmt, True)

            WriteIniValue(F, Section, "RoundUpByDefault", .bInstallmentRoundUp, True)
            WriteIniValue(F, Section, "PostTermInterest", .bPostTermInterest, True)
            WriteIniValue(F, Section, "PostTermInterestRate", .PostTermInterestRate, True)



            Section = STORE_INI_SECTION_EQUIFAX_OPTIONS
            WriteIniValue(F, Section, "CompanyIdent", .CompanyIdent, True)
            WriteIniValue(F, Section, "EquifaxAccountNo", .EquifaxAccountNo, True)
            WriteIniValue(F, Section, "EquifaxSecurityCode", .EquifaxSecurityCode, True)

            WriteIniValue(F, Section, "TransUnionAcctNo", .TransUnionAcctNo, True)
            WriteIniValue(F, Section, "ExperianAcctNo", .ExperianAcctNo, True)

            WriteIniValue(F, Section, "EquifaxSiteID", .EquifaxSiteID, True)
            WriteIniValue(F, Section, "EquifaxPassword", .EquifaxPassword, True)

            WriteIniValue(F, Section, "EquifaxVendorIDCode", .EquifaxVendorIDCode, True)



            Section = STORE_INI_SECTION_PAYPAL_OPTIONS
            WriteIniValue(F, Section, "PayPalUsername", EncodeBase64String(.PayPalUsername), True)
            WriteIniValue(F, Section, "PayPalPassword", EncodeBase64String(.PayPalPassword), True)
            WriteIniValue(F, Section, "PayPalSignature", EncodeBase64String(.PayPalSignature), True)
            WriteIniValue(F, Section, "PayPalAuthKey", EncodeBase64String(.PayPalAuthKey), True)



            Section = STORE_INI_SECTION_ASHLEY_OPTIONS
            WriteIniValue(F, Section, "AshleyID", .AshleyID, True)
            WriteIniValue(F, Section, "AshleyUName", .AshleyUName, True)
            WriteIniValue(F, Section, "AshleyPWord", .AshleyPWord, True)
            WriteIniValue(F, Section, "AshleyPath", .AshleyPath, True)

            WriteIniValue(F, Section, "AshleyATPExternalID", .AshleyATPExternalID, True)
            WriteIniValue(F, Section, "AshleyATPKeyCode", .AshleyATPKeyCode, True)
            WriteIniValue(F, Section, "AshleyATPUserName", .AshleyATPUserName, True)
            WriteIniValue(F, Section, "AshleyATPPassword", .AshleyATPPassword, True)
            WriteIniValue(F, Section, "AshleyATPPICode", .AshleyATPPICode, True)
            WriteIniValue(F, Section, "AshleyATPShipToID", .AshleyATPShipToID)


            Section = STORE_INI_SECTION_AMAZON_OPTIONS
            WriteIniValue(F, Section, "KeyID", .AmazonKeyID)
            WriteIniValue(F, Section, "SecretKey", .AmazonSecretKey)
            WriteIniValue(F, Section, "UserName", .AmazonUserName)
            WriteIniValue(F, Section, "CustomerBucket", .AmazonCustomerBucket)
            WriteIniValue(F, Section, "Password", .AmazonPassword)
            WriteIniValue(F, Section, "AWS Panel Config", .AmazonAWSPanelConfig)
            WriteIniValue(F, Section, "AWS QB Path", .AmazonQBPath)
            WriteIniValue(F, Section, "AWS Misc Backup", .AmazonMisc)

            Section = STORE_INI_SECTION_DISPATCHTRACK_OPTIONS
            WriteIniValue(F, Section, "DispatchTrackLicense", EncodeBase64String(.DispatchTrackLicense))
            WriteIniValue(F, Section, "Code", .DispatchTrackServiceCode)
            WriteIniValue(F, Section, "API-KEY", .DispatchTrackServiceAPI)
            WriteIniValue(F, Section, "ServiceURL", .DispatchTrackServiceURL)

            Section = STORE_INI_SECTION_TRAX_OPTIONS
            WriteIniValue(F, Section, "TRAX License", EncodeBase64String(.TRAXLicense))
            WriteIniValue(F, Section, "TRAX ID", .TRAXID)




            ' +++ Add the option above to the appropriate setting
            ' +++ Take Note of the value of the "Section" variable for where in the INI you want the value to go--it must match the READ section.

        End With

        DoOtherSaves(SI)
    End Function
    Private Function GetStoreInformationINI(Optional ByVal StoreNum As Integer = 0, Optional ByVal Quiet As Boolean = False, Optional ByRef Success As Boolean = False, Optional ByVal AltFileName As String = "") As StoreInfo
        ' Simple function to retrieve basic store info.
        Dim C As Integer, S As String, X As String
        Dim F As String, Section As String
        If StoreNum <= 0 Then StoreNum = StoresSld

        Dim L

        If AltFileName <> "" Then
            F = AltFileName
        Else
            F = StoreINIFile(StoreNum)
        End If

        With GetStoreInformationINI
            On Error Resume Next

            Section = STORE_INI_SECTION_STORE_SETTINGS
            .loadedLicense = ReadIniValue(F, Section, "License")

            .Name = ReadIniValue(F, Section, "Name")
            .Address = ReadIniValue(F, Section, "Address")
            .City = ReadIniValue(F, Section, "City")
            .Phone = ReadIniValue(F, Section, "Phone")

            .Email = ReadIniValue(F, Section, "Email")

            .StoreShipToName = ReadIniValue(F, Section, "ShipToName")
            .StoreShipToAddr = ReadIniValue(F, Section, "ShipToAddr")
            .StoreShipToCity = ReadIniValue(F, Section, "ShipToCity")
            .StoreShipToTele = ReadIniValue(F, Section, "ShipToTele")

            .bStartMaximized = GetSIBool(ReadIniValue(F, Section, "StartMaximized"))
            .bUseBTScanner = GetSIBool(ReadIniValue(F, Section, "UseBTScanner"))

            .bUseQB = GetSIBool(ReadIniValue(F, Section, "UseQB"))
            .bOneCalendar = GetSIBool(ReadIniValue(F, Section, "OneCalendar"))
            .bUseTimeWindows = GetSIBool(ReadIniValue(F, Section, "UseTimeWindows"), True)

            .CCProcessor = ReadIniValue(F, Section, "CCProcessor")
            .CCConfig = ReadIniValue(F, Section, "CCConfig")

            .CashDrawerConfig = ReadIniValue(F, Section, "CashDrawerConfig")
            .ServerLock = GetSIBool(ReadIniValue(F, Section, "ServerLock"))





            Section = STORE_INI_SECTION_SALE_OPTIONS
            .FabSeal = Val(ReadIniValue(F, Section, "FabSeal"))
            .SalesTax = Val(ReadIniValue(F, Section, "SalesTax"))

            .bDeliveryTaxable = GetSIBool(ReadIniValue(F, Section, "DeliveryTaxable"))
            .bLaborTaxable = GetSIBool(ReadIniValue(F, Section, "LaborTaxable"))

            .bNoMerchandisePrice = GetSIBool(ReadIniValue(F, Section, "NoMerchandisePrice"))
            .CalculateList = Val(ReadIniValue(F, Section, "CalculateList"))

            .PrintCopies = Val(ReadIniValue(F, Section, "PrintCopies"))

            ReDim GetStoreInformationINI.SalesCopyID(3)
            .SalesCopyID(0) = ReadIniValue(F, Section, "SalesCopyID1", COPY_FILE)
            .SalesCopyID(1) = ReadIniValue(F, Section, "SalesCopyID2", COPY_CUSTOMER)
            .SalesCopyID(2) = ReadIniValue(F, Section, "SalesCopyID3", COPY_SALESMAN)
            .SalesCopyID(3) = ReadIniValue(F, Section, "SalesCopyID4", COPY_DELIVERY)

            .bManualBillofSaleNo = GetSIBool(ReadIniValue(F, Section, "ManualBillofSaleNo"))
            .bSellFromLoginLocation = GetSIBool(ReadIniValue(F, Section, "SellFromLoginLocation"))

            .bRequireAdvertising = GetSIBool(ReadIniValue(F, Section, "RequireAdvertising"))
            .bRequestEmail = GetSIBool(ReadIniValue(F, Section, "RequestEmail"), True)

            .DelPercent = Val(ReadIniValue(F, Section, "DelPercent"))
            .bUseCashRegisterAddress = GetSIBool(ReadIniValue(F, Section, "UseCashRegisterAddress"))
            .CashRegisterReceiptTailMessage = ReadIniValue(F, Section, "CashRegisterReceiptTailMessage")



            Section = STORE_INI_SECTION_PO_OPTIONS
            .bPOInvoicesLocation = GetSIBool(ReadIniValue(F, Section, "POInvoicesLocation"))
            'L = Array(SS_RecLab_NONE, SS_RecLab_ALL, SS_RecLab_NONE, SS_RecLab_JUSTSTOCK)
            L = New String() {SS_RecLab_NONE, SS_RecLab_ALL, SS_RecLab_NONE, SS_RecLab_JUSTSTOCK}
            .ReceivingLabels = FitList(ReadIniValue(F, Section, "ReceivingLabels"), L, SS_RecLab_NONE)

            .bAPPost = GetSIBool(ReadIniValue(F, Section, "APPost"))
            .bBankManagerPost = GetSIBool(ReadIniValue(F, Section, "BankManagerPost"))

            .bPOSpecialInstr = GetSIBool(ReadIniValue(F, Section, "POSpecialInstr"), True)

            .bPostToLoc1 = GetSIBool(ReadIniValue(F, Section, "PostToLoc1"))
            .bTagIncommingDistinct = GetSIBool(ReadIniValue(F, Section, "TagIncommingDistinct"))

            .PoSpecInstr1 = ReadIniValue(F, Section, "PoSpecInstr1", PO_SPECINSTR_DEFAULT_1)
            .PoSpecInstr2 = ReadIniValue(F, Section, "PoSpecInstr2", PO_SPECINSTR_DEFAULT_2)
            .PoSpecInstr3 = ReadIniValue(F, Section, "PoSpecInstr3", PO_SPECINSTR_DEFAULT_3)
            .PoSpecInstr4 = ReadIniValue(F, Section, "PoSpecInstr4", PO_SPECINSTR_DEFAULT_4)



            Section = STORE_INI_SECTION_COMMISSIONS_OPTIONS
            .CommCode = Val(ReadIniValue(F, Section, "CommCode"))
            .bSeparateCommTables = GetSIBool(ReadIniValue(F, Section, "SeparateCommTables"))

            .bDelIsCommissionable = GetSIBool(ReadIniValue(F, Section, "DelIsCommissionable"))
            .bLabIsCommissionable = GetSIBool(ReadIniValue(F, Section, "LabIsCommissionable"))
            .bNotIsCommissionable = GetSIBool(ReadIniValue(F, Section, "NotIsCommissionable"))




            Section = STORE_INI_SECTION_TAG_OPTIONS
            .TagJustify = FitList(ReadIniValue(F, Section, "TagJustify"), New String() {"Left", "Right", "Center"}, "Center")
            .bPicturesOnTags = GetSIBool(ReadIniValue(F, Section, "PicturesOnTags"))
            .bUseStoreCode = GetSIBool(ReadIniValue(F, Section, "UseStoreCode"))
            .bPrintPoNoCost = GetSIBool(ReadIniValue(F, Section, "PrintPoNoCost"))

            .UseCost = Nothing
            'L = Array(Setup_UseCost_Curr, Setup_UseCost_Aver, Setup_UseCost_FIFO, Setup_UseCost_LIFO, Setup_UseCost_Manu)
            L = New String() {Setup_UseCost_Curr, Setup_UseCost_Aver, Setup_UseCost_FIFO, Setup_UseCost_LIFO, Setup_UseCost_Manu}
            .UseCost = FitList(ReadIniValue(F, Section, "UseCost", .UseCost), L, Setup_UseCost_Curr)

            .bShowRegularPrice = GetSIBool(ReadIniValue(F, Section, "ShowRegularPrice"), True)
            .bShowManufacturer = GetSIBool(ReadIniValue(F, Section, "ShowManufacturer"))
            .bUseStoreCode = GetSIBool(ReadIniValue(F, Section, "UseStoreCode"))
            .bStyleNoInCode = GetSIBool(ReadIniValue(F, Section, "StyleNoInCode"))
            .bCostInCode = GetSIBool(ReadIniValue(F, Section, "CostInCode"))
            .bShowAvailableStock = GetSIBool(ReadIniValue(F, Section, "ShowAvailableStock"))
            .bPrintBarCode = GetSIBool(ReadIniValue(F, Section, "PrintBarCode"), True)

            .bNoListPriceOnTags = GetSIBool(ReadIniValue(F, Section, "NoListPriceOnTags"))
            .bShowPackageItemPrices = GetSIBool(ReadIniValue(F, Section, "ShowPackageItemPrices"))


            Section = STORE_INI_SECTION_INSTALLMENT_OPTIONS
            .loadedInstallmentLicense = ReadIniValue(F, Section, "Installment License")

            .bPaymentBooksMonthly = GetSIBool(ReadIniValue(F, Section, "PaymentBooksMonthly"))
            .GracePeriod = Val(ReadIniValue(F, Section, "GracePeriod", "6"))
            .LateChargePer = Val(ReadIniValue(F, Section, "LateChargePer", .LateChargePer))

            .SimpleInterestRate = Val(ReadIniValue(F, Section, "SimpleInterestRate", .SimpleInterestRate))
            .DocFee = Val(ReadIniValue(F, Section, "DocFee", "50"))

            .MaxLateCharge = GetPrice(ReadIniValue(F, Section, "MaxLateCharge", .MaxLateCharge))
            .bPrintPaymentBooks = GetSIBool(ReadIniValue(F, Section, "PrintPaymentBooks"))

            .MinLateCharge = GetPrice(ReadIniValue(F, Section, "MinLateCharge", .MinLateCharge))

            .bInstallmentInterestIsTaxable = GetSIBool(ReadIniValue(F, Section, "InstallmentInterestIsTaxable"))
            .bAPR = GetSIBool(ReadIniValue(F, Section, "APR"))
            .bJointLife = GetSIBool(ReadIniValue(F, Section, "JointLife"))

            .bEmailLateChargeNotices = GetSIBool(ReadIniValue(F, Section, "EmailLateChargeNotices"))
            .bEmailMonthlyStatements = GetSIBool(ReadIniValue(F, Section, "EmailMonthlyStatements"))

            .ModifiedRevolvingCharge = GetSIBool(ReadIniValue(F, Section, "ModifiedRevolvingCharge"))
            .ModifiedRevolvingRate = Val(ReadIniValue(F, Section, "ModifiedRevolvingRate"))
            .ModifiedRevolvingAPR = GetSIBool(ReadIniValue(F, Section, "ModifiedRevolvingAPR"))
            .ModifiedRevolvingSameAsCash = Val(ReadIniValue(F, Section, "ModifiedRevolvingSameAsCash"))
            .ModifiedRevolvingMinPmt = Val(ReadIniValue(F, Section, "ModifiedRevolvingMinPmt"))

            .bInstallmentRoundUp = GetSIBool(ReadIniValue(F, Section, "RoundUpByDefault"), True)
            .bPostTermInterest = GetSIBool(ReadIniValue(F, Section, "PostTermInterest"), False)
            .PostTermInterestRate = GetPrice(ReadIniValue(F, Section, "PostTermInterestRate"))



            Section = STORE_INI_SECTION_EQUIFAX_OPTIONS
            .CompanyIdent = Left(ReadIniValue(F, Section, "CompanyIdent"), 10)
            .EquifaxAccountNo = Left(ReadIniValue(F, Section, "EquifaxAccountNo"), 10)
            .EquifaxSecurityCode = Left(ReadIniValue(F, Section, "EquifaxSecurityCode"), 3)

            .TransUnionAcctNo = Left(ReadIniValue(F, Section, "TransUnionAcctNo"), 14) 'Altered by Robert 5/24/2017
            .ExperianAcctNo = Left(ReadIniValue(F, Section, "ExperianAcctNo"), 14) 'Along with the designer to reflect the size change

            .EquifaxSiteID = ReadIniValue(F, Section, "EquifaxSiteID")
            .EquifaxPassword = ReadIniValue(F, Section, "EquifaxPassword")

            .EquifaxVendorIDCode = ReadIniValue(F, Section, "EquifaxVendorIDCode")



            Section = STORE_INI_SECTION_PAYPAL_OPTIONS
            .PayPalUsername = DecodeBase64String(ReadIniValue(F, Section, "PayPalUsername"))
            .PayPalPassword = DecodeBase64String(ReadIniValue(F, Section, "PayPalPassword"))
            .PayPalSignature = DecodeBase64String(ReadIniValue(F, Section, "PayPalSignature"))
            .PayPalAuthKey = DecodeBase64String(ReadIniValue(F, Section, "PayPalAuthKey"))



            Section = STORE_INI_SECTION_ASHLEY_OPTIONS
            .AshleyID = ReadIniValue(F, Section, "AshleyID")
            .AshleyUName = ReadIniValue(F, Section, "AshleyUName")
            .AshleyPWord = ReadIniValue(F, Section, "AshleyPWord")
            .AshleyPath = ReadIniValue(F, Section, "AshleyPath")

            .AshleyATPExternalID = ReadIniValue(F, Section, "AshleyATPExternalID")
            .AshleyATPKeyCode = ReadIniValue(F, Section, "AshleyATPKeyCode")
            .AshleyATPUserName = ReadIniValue(F, Section, "AshleyATPUserName")
            .AshleyATPPassword = ReadIniValue(F, Section, "AshleyATPPassword")
            .AshleyATPPICode = ReadIniValue(F, Section, "AshleyATPPICode")
            .AshleyATPShipToID = ReadIniValue(F, Section, "AshleyATPShipToID")

            Section = STORE_INI_SECTION_AMAZON_OPTIONS
            .AmazonKeyID = ReadIniValue(F, Section, "KeyID")
            .AmazonSecretKey = ReadIniValue(F, Section, "SecretKey")
            .AmazonUserName = ReadIniValue(F, Section, "UserName")
            .AmazonCustomerBucket = ReadIniValue(F, Section, "CustomerBucket")
            .AmazonPassword = ReadIniValue(F, Section, "Password")
            .AmazonAWSPanelConfig = ReadIniValue(F, Section, "AWS Panel Config")
            .AmazonQBPath = ReadIniValue(F, Section, "AWS QB Path")
            .AmazonMisc = ReadIniValue(F, Section, "AWS Misc Backup")

            Section = STORE_INI_SECTION_DISPATCHTRACK_OPTIONS
            .DispatchTrackLicense = DecodeBase64String(ReadIniValue(F, Section, "DispatchTrackLicense"))
            .DispatchTrackServiceCode = ReadIniValue(F, Section, "Code")
            .DispatchTrackServiceAPI = ReadIniValue(F, Section, "API-KEY")
            .DispatchTrackServiceURL = ReadIniValue(F, Section, "ServiceURL")

            Section = STORE_INI_SECTION_TRAX_OPTIONS
            .TRAXLicense = DecodeBase64String(ReadIniValue(F, Section, "TRAX License"))
            .TRAXID = ReadIniValue(F, Section, "TRAX ID")
        End With


        ' +++ Add the option above to the appropriate setting
        ' +++ Take Note of the value of the "Section" variable for where in the INI you want the value to go.

        DoOtherLoads(GetStoreInformationINI)


        Success = True
        Exit Function

NoFile:
        If Not Quiet Then MsgBox("Can't load information for Store " & StoreNum & ".", vbExclamation)
        Success = False
    End Function

    Public Property InstallmentLicense() As String
        Get
            InstallmentLicense = GetCDSSetting("InstallmentLicense", "")
        End Get
        Set(value As String)
            SaveCDSSetting("InstallmentLicense", value)
        End Set
    End Property

    Private Sub DoOtherSaves(ByRef SI As StoreInfo)
        With SI
            SaveCDSSetting("CCMachine", IIf(.bUseCCMachine, "1", "0"))
            ServerLock(IIf(.ServerLock, vbTrue, vbFalse))
        End With
    End Sub

    Private Function GetSIBool(ByVal S As String, Optional ByVal Dflt As Boolean = False) As Boolean
        If S = "" Then GetSIBool = Dflt : Exit Function
        On Error Resume Next
        Select Case UCase(Left(S, 1))
            Case "" : GetSIBool = Dflt
            Case "0", "F" : GetSIBool = False
            Case "1", "T" : GetSIBool = True
            Case Else : GetSIBool = False
        End Select
    End Function

    Private Sub DoOtherLoads(ByRef SI As StoreInfo)
        With SI
            .bUseCCMachine = IIf(Val(GetCDSSetting("CCMachine", "0")) = 0, False, True)
            .ServerLock = ServerLock()
        End With
    End Sub

    Public Function ReadStoreSetting(ByVal nStoreNo As Integer, ByVal nSection As IniSections_StoreSettings, ByVal nKey As String, Optional ByVal nDefault As String = "") As String
        Const UnusedValue As String = "#_*SA"
        If nStoreNo = -1 Then
            Dim I As Integer ' Like "broadcast save", but read all files until you find one that has it set.  A value that should never be used is a standin for "default" in this case.
            For I = 1 To ActiveNoOfLocations
                ReadStoreSetting = ReadStoreSetting(I, nSection, nKey, UnusedValue)
                If ReadStoreSetting <> UnusedValue Then Exit Function
            Next
        End If
        ReadStoreSetting = ReadIniValue(StoreINIFile(nStoreNo), StoreSettingSectionKey(nSection), nKey, nDefault)
    End Function

    Public Function InstallINIToStoreSettings(ByVal Source As String, Optional ByVal StoreNo As Integer = 0) As Boolean
        If StoreNo > ActiveNoOfLocations Then Exit Function ' don't install to non-existent INI
        InstallINIToStoreSettings = InstallINIValues(Source, StoreINIFile(StoreNo))
    End Function

    Public Function InstallINIValues(ByVal Source As String, ByVal Destination As String) As Boolean
        Dim X() As String, Y() As String, V As String, L, M, K()
        Dim R As IniSections_StoreSettings, Z As clsHashTable
        On Error GoTo InstallFail
        X = INISections(Source)
        For Each L In X
            R = StoreSettingSectionCode(L)
            Y = INISectionKeys(Source, L)

            If R = IniSections_StoreSettings.iniSection_UNKNOWN Then
                ' Do Nothing on install
            ElseIf R = IniSections_StoreSettings.iniSection_InstallSQL Then
                Z = INISectionAsHashTable(Source, L)
                If Z.Item("RunOnce") = "" Or AllowRunOnce(Z.Item("RunOnce")) Then
                    K = Z.Keys(vbTrue)
                    For Each M In K
                        If Left(M, 3) = "SQL" Then
                            ExecuteRecordsetBySQL(Z.Item(M))
                        End If
                    Next
                End If
            Else  ' For all other sections, install to the store setup file directly
                For Each M In Y
                    V = ReadIniValue(Source, L, M)
                    WriteIniValue(Destination, L, M, V)
                Next
            End If
        Next

        InstallINIValues = True
InstallFail:
    End Function

    Private Function StoreSettingSectionCode(ByVal nSectionNo As String) As IniSections_StoreSettings
        Select Case nSectionNo
            'Case STORE_INI_SECTION_STORE_SETTINGS : StoreSettingSectionCode = IniSections_StoreSettings
            Case STORE_INI_SECTION_SALE_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_SaleOptions
            Case STORE_INI_SECTION_PO_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_POOptions
            Case STORE_INI_SECTION_COMMISSIONS_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_Commissions
            Case STORE_INI_SECTION_TAG_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_TagOptions
            Case STORE_INI_SECTION_INSTALLMENT_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_Installment
            Case STORE_INI_SECTION_EQUIFAX_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_Equifax
            Case STORE_INI_SECTION_PAYPAL_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_PayPal
            Case STORE_INI_SECTION_ASHLEY_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_Ashley
            Case STORE_INI_SECTION_AMAZON_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_Amazon
            Case STORE_INI_SECTION_DISPATCHTRACK_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_DispatchTrack
            Case STORE_INI_SECTION_TRAX_OPTIONS : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_TRAX
            Case STORE_INI_SECTION_PROGRAM_STATE : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_Program
            Case STORE_INI_SECTION_CUSTOM_VALUES : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_Custom

            Case STORE_INI_SECTION_INSTALL_SQL : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_InstallSQL

            Case Else : StoreSettingSectionCode = IniSections_StoreSettings.iniSection_UNKNOWN
        End Select
    End Function

    Public Function LoadFrmSetupFromStoreInformation(ByRef SI As StoreInfo)
        On Error Resume Next
        With SI
            '        frmSetup.txtStoreLocName = .Name
            '        frmSetup.txtStoreAddress = .Address
            '        frmSetup.txtStoreCity = .City
            '        frmSetup.txtStoreTele = .Phone

            '        frmSetup.txtCommissionMethod = .CommCode
            '        frmSetup.hsCommType = Val(.CommCode)
            '        'On Error GoTo 0
            '        frmSetup.txtStainProtection = .FabSeal
            '        frmSetup.txtDefSalesTax = .SalesTax

            '        frmSetup.txtShipLocName = .StoreShipToName
            '        frmSetup.txtShipAddress = .StoreShipToAddr
            '        frmSetup.txtShipCity = .StoreShipToCity
            '        frmSetup.txtShipTele = .StoreShipToTele

            '        frmSetup.chkPOInvoicesLocation = IIf(.bPOInvoicesLocation, 1, 0)
            '        frmSetup.chkDelIsTaxable = IIf(.bDeliveryTaxable, 1, 0)
            '        frmSetup.chkLabIsTaxable = IIf(.bLaborTaxable, 1, 0)
            '        frmSetup.chkPicturesOnTags = IIf(.bPicturesOnTags, 1, 0)

            '        frmSetup.cboTagJustify = .TagJustify
            '        frmSetup.chkNoMerchandisePrice = IIf(.bNoMerchandisePrice, 1, 0)
            '        frmSetup.cboCalculateList.ListIndex = .CalculateList

            '        frmSetup.txtPrintCopies = .PrintCopies
            '        If .bPaymentBooksMonthly Then frmSetup.optPaymentBookMonthly = True Else frmSetup.optPaymentBookWeekly = True

            '        frmSetup.cboSalesCopyID.List(0) = .SalesCopyID(0)
            '        frmSetup.cboSalesCopyID.List(1) = .SalesCopyID(1)
            '        frmSetup.cboSalesCopyID.List(2) = .SalesCopyID(2)
            '        frmSetup.cboSalesCopyID.List(3) = .SalesCopyID(3)

            '        frmSetup.chkDeliveryCommissionable = IIf(.bDelIsCommissionable, 1, 0)
            '        frmSetup.chkLaborCommissionable = IIf(.bLabIsCommissionable, 1, 0)
            '        frmSetup.chkNotesCommissionable = IIf(.bNotIsCommissionable, 1, 0)


            '        frmSetup.chkUseStoreCode = IIf(.bUseStoreCode, 1, 0)

            '        frmSetup.txtGrace = .GracePeriod
            '        frmSetup.cboReceivingLabels.Text = .ReceivingLabels

            '        frmSetup.SetPostAccountsPayableActiveCheck = IIf(.bAPPost, 1, 0)
            '        frmSetup.BankManagerPost = IIf(.bBankManagerPost, 1, 0)
            '        frmSetup.chkPrintPONoCost = IIf(.bPrintPoNoCost, 1, 0)

            '        frmSetup.cboUseCost.Text = .UseCost
            '        frmSetup.chkPOSpecialInst = IIf(.bPOSpecialInstr, 1, 0)

            '        frmSetup.txtSimpleInterestRate = .SimpleInterestRate
            '        frmSetup.txtDocFee = CurrencyFormat(.DocFee)
            '        frmSetup.txtLateChargePerc = .LateChargePer
            '        frmSetup.txtMaxLateCharge = CurrencyFormat(.MaxLateCharge)

            '        frmSetup.chkPrintPaymentBooks = IIf(.bPrintPaymentBooks, 1, 0)
            '        frmSetup.chkManualBillOfSaleNo = IIf(.bManualBillofSaleNo, 1, 0)
            '        frmSetup.chkShowRegularPrice = IIf(.bShowRegularPrice, 1, 0)
            '        frmSetup.chkShowManufacturer = IIf(.bShowManufacturer, 1, 0)
            '        frmSetup.chkUseStoreCode = IIf(.bUseStoreCode, 1, 0)
            '        frmSetup.chkStyleNoInCode = IIf(.bStyleNoInCode, 1, 0)
            '        frmSetup.chkCostInCode = IIf(.bCostInCode, 1, 0)

            '        frmSetup.chkShowAvailStock = IIf(.bShowAvailableStock, 1, 0)
            '        frmSetup.chkPrintBarCode = IIf(.bPrintBarCode, 1, 0)
            '        frmSetup.chkSellFromLoginLoc = IIf(.bSellFromLoginLocation, 1, 0)

            '        frmSetup.txtMinLateCharge = FormatCurrency(.MinLateCharge)

            '        frmSetup.chkPostToLoc1 = IIf(.bPostToLoc1, 1, 0)
            '        frmSetup.chkRequireAdvertising = IIf(.bRequireAdvertising, 1, 0)
            '        frmSetup.chkTagIncommingDistinct = IIf(.bTagIncommingDistinct, 1, 0)
            '        frmSetup.chkUseCCMachine = IIf(.bUseCCMachine, 1, 0)

            '        frmSetup.chkStartMaximized = IIf(.bStartMaximized, 1, 0)
            '        frmSetup.chkUseBTScanner = IIf(.bUseBTScanner, 1, 0)

            '        frmSetup.InitPOSpecInstr
            '        '    frmSetup.SetPoSpecInstr 1, .PoSpecInstr1
            '        '    frmSetup.SetPoSpecInstr 2, .PoSpecInstr2
            '        '    frmSetup.SetPoSpecInstr 3, .PoSpecInstr3
            '        '    frmSetup.SetPoSpecInstr 4, .PoSpecInstr4

            '        frmSetup.txtDelPercent = .DelPercent
            '        frmSetup.chkUseCashRegisterAddress = IIf(.bUseCashRegisterAddress, 1, 0)
            '        frmSetup.txtCashRegisterMessage = .CashRegisterReceiptTailMessage
            '        frmSetup.txtStoreEmail = .Email

            '        frmSetup.chkQB = IIf(.bUseQB, 1, 0)
            '        frmSetup.chkSeparateCommTables = IIf(.bSeparateCommTables, 1, 0) : frmSetup.UpdateSeparateCommSalesmen
            '        frmSetup.chkOneCalendar = IIf(.bOneCalendar, 1, 0)

            '        frmSetup.EquifaxAcctNo = .EquifaxAccountNo
            '        frmSetup.chkInstallmentInterestTaxable = IIf(.bInstallmentInterestIsTaxable, 1, 0)
            '        frmSetup.EquifaxSecCode = .EquifaxSecurityCode

            '        frmSetup.chkAPR = IIf(.bAPR, 1, 0)
            '        frmSetup.EquifaxTransUnion = .TransUnionAcctNo
            '        frmSetup.chkUseTimeWindows = IIf(.bUseTimeWindows, 1, 0)
            '        frmSetup.EquifaxExperian = .ExperianAcctNo
            '        frmSetup.EquifaxCompanyID = .CompanyIdent

            '        frmSetup.chkJointLife = IIf(.bJointLife, 1, 0)

            '        frmSetup.AshleyStoreNo = .AshleyID
            '        frmSetup.AshleyUName = .AshleyUName
            '        frmSetup.AshleyPWord = .AshleyPWord
            '        frmSetup.AshleyPath = .AshleyPath

            '        frmSetup.AshleyATPExternalID = .AshleyATPExternalID
            '        frmSetup.AshleyATPKeyCode = .AshleyATPKeyCode
            '        frmSetup.AshleyATPUserName = .AshleyATPUserName
            '        frmSetup.AshleyATPPassword = .AshleyATPPassword
            '        frmSetup.AshleyATPPICode = .AshleyATPPICode
            '        frmSetup.AshleyATPShipToID = .AshleyATPShipToID

            '        frmSetup.chkEmailLCNotices = IIf(.bEmailLateChargeNotices, 1, 0)
            '        frmSetup.chkEmailMonthlyStatements = IIf(.bEmailMonthlyStatements, 1, 0)

            '        frmSetup.chkEmail = IIf(.bRequestEmail, 1, 0)

            '        frmSetup.EquifaxSiteID = .EquifaxSiteID
            '        frmSetup.EquifaxPassword = .EquifaxPassword
            '        frmSetup.EquifaxVendorIDCode = .EquifaxVendorIDCode

            '        frmSetup.SetupCCConfig.CCConfig, .CCProcessor

            'frmSetup.txtPPUsername = .PayPalUsername
            '        frmSetup.txtPPPassword = .PayPalPassword
            '        frmSetup.txtPPSignature = .PayPalSignature
            '        frmSetup.txtPayPalAuthKey = .PayPalAuthKey
            '        frmSetup.AdjustPayPal

            '        frmSetup.chkNoListPriceOnTags = IIf(.bNoListPriceOnTags, 1, 0)

            '        frmSetup.chkShowPackageItemPrices = IIf(.bShowPackageItemPrices, 1, 0)

            '        frmSetup.CashDrawerConfig.CashDrawerConfig
            '        frmSetup.SetServerLock.ServerLock
            '        frmSetup.RevolvingCharge = IIf(.ModifiedRevolvingCharge, 1, 0)
            '        frmSetup.RevolvingRate = CurrencyFormat(.ModifiedRevolvingRate)
            '        frmSetup.RevolvingAPR = IIf(.ModifiedRevolvingAPR, 1, 0)
            '        frmSetup.RevolvingSAC = .ModifiedRevolvingSameAsCash
            '        frmSetup.RevolvingMinPmt = Round(.ModifiedRevolvingMinPmt, 4)

            '        frmSetup.AmazonKeyID = .AmazonKeyID
            '        frmSetup.AmazonSecretKey = .AmazonSecretKey
            '        frmSetup.AmazonCustomerBucket = .AmazonCustomerBucket
            '        frmSetup.AmazonUserName = .AmazonUserName
            '        frmSetup.AmazonPassword = .AmazonPassword
            '        frmSetup.AmazonQBPath = .AmazonQBPath
            '        frmSetup.AmazonMisc = .AmazonMisc

            '        frmSetup.txtDispatchTrackLicense = .DispatchTrackLicense
            '        frmSetup.DDTServiceCode = .DispatchTrackServiceCode
            '        frmSetup.DDTServiceAPIKey = .DispatchTrackServiceAPI
            '        frmSetup.DDTServiceURL = .DispatchTrackServiceURL

            '        frmSetup.txtTRAXLicense = .TRAXLicense
            '        frmSetup.TRAXID = .TRAXID

            '        frmSetup.chkInstallRoundUp.Value = IIf(.bInstallmentRoundUp, 1, 0)
            '        frmSetup.chkPostTermInterest.Value = IIf(.bPostTermInterest, 1, 0)
            '        frmSetup.txtPostTermInterestRate.Text = FormatQuantity(.PostTermInterestRate)


            ' +++ This is where you add maintenance for a new control on frmSetup.
            ' +++ Take Note of the control type, and default, and adapt from a control above

        End With
    End Function

End Module
