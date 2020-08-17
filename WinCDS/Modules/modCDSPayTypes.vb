Imports Microsoft.VisualBasic.Interaction
Module modCDSPayTypes
    '::::modCDSPayTypes
    ':::SUMMARY
    ': Handling of Payment types throughout the software.
    ':
    ':::DESCRIPTION
    ': All Pay types use the same enum / helper functions.  All predecent uses of hard-coded values should
    ': be updated to use the appropriate enum value.
    ':
    ': The main pay types are stored in the database and should not be changed.
    ':   - Order shouldn't matter...?
    ':   - Skipping sequence shouldn't break (allowing for removal, but not reuse)
    ':   - Should not be REPLACED or REUSED, as there may be existing items with the previous PT
    ':   - Min/Max are defined privately within the module, as is PayTypeList.
    ':
    ': If types are added/removed, they should be added to the appropriate payment forms, either by
    ': referencing the list below, or searching for the configuration function, DoPayType
    ':   - OrdSelect        - Payment on new sales
    ':   - OrdPay           - General payment on sale
    ':   - CustAdjRefund    - Customer adjustments
    ':
    ':::INTERFACE
    '::Data Types
    ':   cdsPayTypeMode     Type enum)
    ':   cdsPayTypes        Pay Type enum
    '::Handlers
    ':   PayTypeName        [enum -> string], several display formats
    ':   PayListItem        Shortcut of special case of PayTypeName
    ':   PayTypeIs          [string -> enum], any supported format
    '::Test Functions
    ':   PayTypeIsIn        Test up to 8 PT strings for a match
    ':   PayTypeIsFinance   Shortcut of PayTypeIsIn to test for financing
    ':   PayTypeIsCC        Shortcut of PayTypeIsIn to test for CC types
    ':   PayTypeIsDebit     Shortcut of PayTypeIsIn to test for debit type
    ':   PayTypeIsStoreCard Shortcut of PayTypeIsIn to test for store card type
    '::Configuration
    ':   DoPayType          whether a store supports a given payment type...  Any type can be enabled/disabled per location & store

    Public Enum cdsPayTypeMode
        cdsPTM_Standard = 1
        cdsPTM_Abbrev = 2
        cdsPTM_PayItem = 3
    End Enum
    Private Const cdsPayTypeMode_Min As Integer = cdsPayTypeMode.cdsPTM_Standard
    Private Const cdsPayTypeMode_Max as integer = cdsPayTypeMode.cdsPTM_PayItem

    Public Enum cdsPayTypes
        cdsPT_NONE = 0
        cdsPT_Cash = 1
        cdsPT_Check = 2
        cdsPT_Visa = 3
        cdsPT_MCard = 4
        cdsPT_Disc = 5
        cdsPT_amex = 6
        cdsPT_BackOrder = 7
        cdsPT_OutsideFinance = 8
        cdsPT_DebitCard = 9
        cdsPT_MiscDiscount = 10
        cdsPT_StoreFinance = 11
        cdsPT_StoreCreditCard = 12
        cdsPT_ECheck = 13
        cdsPT_CompanyCheck = 14
        cdsPT_OutsideFinance2 = 15
        cdsPT_OutsideFinance3 = 16
        cdsPT_OutsideFinance4 = 17
        cdsPT_OutsideFinance5 = 18
    End Enum
    Private Const cdsPayType_Min as integer = cdsPayTypes.cdsPT_Cash
    Private Const cdsPayType_Max as integer = cdsPayTypes.cdsPT_OutsideFinance5
    Public Function PayTypeIs(ByVal PaymentType As String) As cdsPayTypes
        '::::PayTypeIs
        ':::SUMMARY
        ': Returns the [cdsPayTypes] value for the given <PaymentType> string.
        ':::DESCRIPTION
        ':  Transforms string -> enum.
        ':  Performs a table lookup for the appropriate cdsPayTypes enum value for the given Payment Type string.
        ':::PARAMETERS
        ': PaymentMode - The string representing the payment type (e.g., "Cash", "Check", "Store Finance", "AMEX").
        ':::RETURN
        ':  cdsPayTypes - The system-wide enum representing the payment type.
        ':::SEE ALSO
        ':  PayTypeName, PayListItem, PayItemIsIn, Pay TypeIsCC, PayTypeIsDebit, PayTypeIsStoreCard
        ':::NOTES
        ':  The comprehensive function attempts to match ALL possible pay type names used throughout WinCDS, no matter the mode/style used.
        ':  It is imperative that whatever naming convention, capitalization, or abbreviation is employed, it is
        ':  always executed through this function, and likewise, the PayTypeName function, or there will be no guarantee
        ':  that existing and future payment types will always be captured.
        ':  Always be sure to only use these function to reference and supply any payment type throughout WinCDS,
        ':  Including adding a Payment Mode if necessary.
        Dim I as integer
        Dim A(), L

        PayTypeIs = cdsPayTypes.cdsPT_NONE
        If PaymentType = "" Then Exit Function

        If IsNumeric(PaymentType) And PayTypeName(Val(PaymentType), , , True) <> "" Then PayTypeIs = Val(PaymentType)

        A = PayTypeList()

        For Each L In A
            '  For J = cdsPayType_Min To cdsPayType_Max
            For I = cdsPayTypeMode_Min To cdsPayTypeMode_Max
                If UCase(Trim(PaymentType)) = PayTypeName(L, I, True, True) Then PayTypeIs = L : Exit Function
            Next
        Next
    End Function
    Public Function PayTypeIsOutsideFinance(ByVal PayTypeString As String) As Boolean
        PayTypeIsOutsideFinance = PayTypeIsIn(PayTypeString, cdsPayTypes.cdsPT_OutsideFinance, cdsPayTypes.cdsPT_OutsideFinance2, cdsPayTypes.cdsPT_OutsideFinance3, cdsPayTypes.cdsPT_OutsideFinance4, cdsPayTypes.cdsPT_OutsideFinance5)
    End Function
    Public Function PayListItem(ByVal pt As cdsPayTypes) As String
        PayListItem = PayTypeName(pt, cdsPayTypeMode.cdsPTM_PayItem, True)
    End Function
    Public Function PayTypeName(ByVal pt As cdsPayTypes, Optional ByVal Abbr As cdsPayTypeMode = cdsPayTypeMode.cdsPTM_Standard, Optional ByVal Upper As Boolean = True, Optional ByVal doSeek As Boolean = False) As String
        '::::PayTypeName
        ':::SUMMARY
        ':Returns the string representation of the given cdsPayTypes enum value, allowing for different display modes.
        ':::DESCRIPTION
        ':  This is the central look-up function for all WinCDS payment types.  All payment types used in the system,
        ':  whether in the code or in the UI must go through this for guaranteed naming conformity and persistence.
        ':  Even if one of these
        ':::PARAMETERS
        ':  Pt - The Payment Type as a cdsPayTypes enum.
        ':  Abbr - The display mode as a cdsPayTypeMode enum (default=cdsPTM_Standard).
        ':  Upper - Boolean value indicating whether the resulting string should be passed through UCase() before return.
        ':::RETURN
        ':  String - The Payment Type name in the mode and style requested.
        ':::SEE ALSO
        ':  PayTypeName, PayListItem, PayItemIsIn, Pay TypeIsCC, PayTypeIsDebit, PayTypeIsStoreCard
        ':::NOTES
        ':  Persistence is also required here.  If one of these payment types is no longer supported, it is
        ':  important to keep the place in the enum and the string representations because the database
        ':  may or may not still be populated with these values in various places.
        ':  Never delete a Payment Type, only mark it deprecated and possibly flag its use.
        ':  Alternatively, to be thorough, one could make a patch for existing PTs to be removed, if necessary.
        If CLng(pt) > 1000 Then PayTypeName = "" & pt : Exit Function ' FOR G/L ACCOUNT NUMBERS

        Select Case pt
            Case cdsPayTypes.cdsPT_Cash : PayTypeName = "Cash"    ' These ones never change...
            Case cdsPayTypes.cdsPT_Check : PayTypeName = "Check"
            Case cdsPayTypes.cdsPT_Visa : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Visa", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Visa Card", True, "Visa Card")
            Case cdsPayTypes.cdsPT_MCard : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "MCard", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Master Card", True, "MasterCard")
            Case cdsPayTypes.cdsPT_Disc : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Disc", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Discover Card", True, "Discover Card")
            Case cdsPayTypes.cdsPT_amex : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "AMEX", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Amex Card", True, "American Express")
            Case cdsPayTypes.cdsPT_BackOrder : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Back Ord", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Back Order", True, "Back Order")
            Case cdsPayTypes.cdsPT_OutsideFinance : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Outside Fin", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Outside Fin Co", True, "Outside Finance")
            Case cdsPayTypes.cdsPT_DebitCard : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Debit", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Debit Card", True, "Debit Card")
            Case cdsPayTypes.cdsPT_MiscDiscount : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Misc Dscnt", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Misc Discount", True, "Misc Discount")
            Case cdsPayTypes.cdsPT_StoreFinance : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Store Fin", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Store Finance", True, "Store Finance")
            Case cdsPayTypes.cdsPT_StoreCreditCard : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Store Card", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Store Card", True, "Store Credit Card")
            Case cdsPayTypes.cdsPT_ECheck : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "E-Check", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Electronic Check", True, "Electronic Check")
            Case cdsPayTypes.cdsPT_CompanyCheck : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Company Check", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Company Check", True, "Company Check")

            Case cdsPayTypes.cdsPT_OutsideFinance2 : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Outside Fin 2", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Outside Fin 2", True, "Company Finance 2")
            Case cdsPayTypes.cdsPT_OutsideFinance3 : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Outside Fin 3", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Outside Fin 3", True, "Company Finance 3")
            Case cdsPayTypes.cdsPT_OutsideFinance4 : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Outside Fin 4", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Outside Fin 4", True, "Company Finance 4")
            Case cdsPayTypes.cdsPT_OutsideFinance5 : PayTypeName = Switch(Abbr = cdsPayTypeMode.cdsPTM_Abbrev, "Outside Fin 5", Abbr = cdsPayTypeMode.cdsPTM_PayItem, "Outside Fin 5", True, "Company Finance 5")

            Case Else
                If Not doSeek Then DevErr("Unknown Pay Type:  modSetup.PayTypeName - " & CLng(pt))
        End Select
        If Upper Then PayTypeName = UCase(PayTypeName)
    End Function

    Private Function PayTypeList()
        Dim I As Integer, N As Integer
        Static Loaded As Boolean, A()       ' Runtime cache.

        If Not Loaded Then
            For I = cdsPayType_Min To cdsPayType_Max
                If PayTypeName(I, , , True) <> "" Then
                    N = N + 1
                    ReDim Preserve A(0 To N - 1)
                    A(N - 1) = I
                End If
            Next
        End If

        Loaded = True
        PayTypeList = A
    End Function
    Public Function PayTypeIsIn(ByVal PaymentMode As String, Optional ByVal PayT_A As cdsPayTypes = -1, Optional ByVal PayT_B As cdsPayTypes = -1, Optional ByVal PayT_C As cdsPayTypes = -1, Optional ByVal PayT_D As cdsPayTypes = -1, Optional ByVal PayT_E As cdsPayTypes = -1, Optional ByVal PayT_F As cdsPayTypes = -1, Optional ByVal PayT_G As cdsPayTypes = -1, Optional ByVal PayT_H As cdsPayTypes = -1) As Boolean
        PayTypeIsIn = IsIn(PayTypeIs(PaymentMode), PayT_A, PayT_B, PayT_C, PayT_D, PayT_E, PayT_F, PayT_G, PayT_H)
    End Function

    Public Function DoPayType(ByVal PayType As cdsPayTypes) As Boolean
        '::::DoPayType
        ':::SUMMARY
        ': Configuration function to determine whether a this store/location supports a given payment type
        ':::DESCRIPTION
        ':  Transforms string -> enum.
        ':  Performs a table lookup for the appropriate <cdsPayTypes> enum value for the given Payment Type string.
        ':  Returns a T/F value indicating whether the payment type is supported.
        ':::PARAMETERS
        ': PayType - The <cdsPayType> to be tested for support.
        ':::RETURN
        ':  Boolean - True/False whether this store/location supports a given payment type
        ':::SEE ALSO
        ':  PayTypeName, PayListItem, PayItemIsIn, Pay TypeIsCC, PayTypeIsDebit, PayTypeIsStoreCard
        ':  SwipeCards, SwipeCreditCards, SwipeDebitCards, SwipeGiftCards, ProcessCC
        ':::NOTES
        ':  All UI cases that present a list of payment options should use this function for EACH individual payment type.
        ':  No Payment types should be presented to the user without these options
        Select Case PayType
            Case cdsPayTypes.cdsPT_Cash : DoPayType = True
            Case cdsPayTypes.cdsPT_Check : DoPayType = True
            Case cdsPayTypes.cdsPT_Visa : DoPayType = True
            Case cdsPayTypes.cdsPT_MCard : DoPayType = True
            Case cdsPayTypes.cdsPT_Disc
                Select Case StoreSettings.CCProcessor
                    Case CCPROC_TC : DoPayType = InStr(CSVField(StoreSettings.CCConfig, 3), "D") <> 0
                    Case Else : DoPayType = True
                End Select
            Case cdsPayTypes.cdsPT_amex
                Select Case StoreSettings.CCProcessor
                    Case CCPROC_TC : DoPayType = InStr(CSVField(StoreSettings.CCConfig, 3), "A") <> 0
                    Case Else : DoPayType = True
                End Select
            Case cdsPayTypes.cdsPT_BackOrder : DoPayType = True
            Case cdsPayTypes.cdsPT_OutsideFinance : DoPayType = True
            Case cdsPayTypes.cdsPT_DebitCard
                Select Case StoreSettings.CCProcessor
                    Case CCPROC_TC : DoPayType = False
                    Case Else : DoPayType = True
                End Select
            Case cdsPayTypes.cdsPT_MiscDiscount : DoPayType = True
            Case cdsPayTypes.cdsPT_StoreFinance : DoPayType = True
            Case cdsPayTypes.cdsPT_StoreCreditCard : DoPayType = True
            Case cdsPayTypes.cdsPT_ECheck : DoPayType = True
            Case cdsPayTypes.cdsPT_CompanyCheck : DoPayType = True
            Case cdsPayTypes.cdsPT_OutsideFinance2 : DoPayType = DoExtraOutsideFinance(2)
            Case cdsPayTypes.cdsPT_OutsideFinance3 : DoPayType = DoExtraOutsideFinance(3)
            Case cdsPayTypes.cdsPT_OutsideFinance4 : DoPayType = DoExtraOutsideFinance(4)
            Case cdsPayTypes.cdsPT_OutsideFinance5 : DoPayType = DoExtraOutsideFinance(5)

            Case Else
                DevErr("Unknown Payment Type: " & PayType)
                DoPayType = False
        End Select
    End Function

    Public Function DoExtraOutsideFinance(ByVal Idx As Integer) As Boolean
        DoExtraOutsideFinance = IsDevelopment()   ' Default is off except for CDS.
        Select Case Idx
            Case 2
                If IsSidesFurniture Then DoExtraOutsideFinance = True
            Case 3
                If IsSidesFurniture Then DoExtraOutsideFinance = True
            Case 4
                If IsSidesFurniture Then DoExtraOutsideFinance = True
            Case 5
                If IsSidesFurniture Then DoExtraOutsideFinance = True
        End Select
    End Function

    Public Function PayTypeIsFinance(ByVal PayTypeString As String, Optional ByVal IncludeBackorder As Boolean = True) As Boolean
        If IncludeBackorder Then
            PayTypeIsFinance = PayTypeIsIn(PayTypeString, cdsPayTypes.cdsPT_BackOrder, cdsPayTypes.cdsPT_OutsideFinance, cdsPayTypes.cdsPT_StoreFinance, cdsPayTypes.cdsPT_OutsideFinance2, cdsPayTypes.cdsPT_OutsideFinance3, cdsPayTypes.cdsPT_OutsideFinance4, cdsPayTypes.cdsPT_OutsideFinance5)
        Else
            PayTypeIsFinance = PayTypeIsIn(PayTypeString, cdsPayTypes.cdsPT_OutsideFinance, cdsPayTypes.cdsPT_StoreFinance, cdsPayTypes.cdsPT_OutsideFinance2, cdsPayTypes.cdsPT_OutsideFinance3, cdsPayTypes.cdsPT_OutsideFinance4, cdsPayTypes.cdsPT_OutsideFinance5)
        End If
    End Function

    Public Function PayTypeIsCC(ByVal PayTypeString As String) As Boolean
        PayTypeIsCC = PayTypeIsIn(PayTypeString, cdsPayTypes.cdsPT_Visa, cdsPayTypes.cdsPT_MCard, cdsPayTypes.cdsPT_Disc, cdsPayTypes.cdsPT_amex)
        '  PayTypeIsCC = IsIn(PayTypeIs(PayTypeString), cdsPT_Visa, cdsPT_MCard, cdsPT_Disc, cdsPT_amex)
    End Function

    Public Function PayTypeIsDebit(ByVal PayTypeString As String) As Boolean
        PayTypeIsDebit = PayTypeIsIn(PayTypeString, cdsPayTypes.cdsPT_DebitCard)
        '  PayTypeIsDebit = IsIn(PayTypeIs(PayTypeString), cdsPT_DebitCard)
    End Function

    Public Function PayTypeIsStoreCard(ByVal PayTypeString As String) As Boolean
        PayTypeIsStoreCard = PayTypeIsIn(PayTypeString, cdsPayTypes.cdsPT_StoreCreditCard)
        'PayTypeIsStoreCard = IsIn(PayTypeIs(PayTypeString), cdsPT_StoreCreditCard)
    End Function

End Module
