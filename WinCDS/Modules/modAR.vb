Module modAR
    '::::modAR
    ':::SUMMARY
    ': Functions for the general performance of the A/R module.
    ':::DESCRIPTION
    '
    '
    '$1.00 Grace..  Some stores may not want this, in which case 0 zere be used as a property instead
    Public Const LateGraceAmt As Decimal = 1.0#

    ' AR Status Constants
    Public Const arST_Open As String = "O"   ' Open
    Public Const arST_Clos As String = "C"    ' Closed
    Public Const arST_Void As String = "V"    ' Void
    Public Const arST_Writ As String = "W"    ' Write Off
    Public Const arST_Repo As String = "R"    ' Repo
    Public Const arST_Lega As String = "L"    ' Legal
    Public Const arST_Bank As String = "Y"    ' Bankruptcy

    ' AR Status Descriptions
    Public Const arSD_O As String = "Open"
    Public Const arSD_C As String = "Closed"
    Public Const arSD_V As String = "Void"
    Public Const arSD_W As String = "Write Off"
    Public Const arSD_R As String = "Repo"
    Public Const arSD_L As String = "Legal"
    Public Const arSD_Y As String = "Bankruptcy"


    ' [Transactions].[Type], MaxLength=50...
    ' Len50 looks like this:".................................................."
    ' But, we probably want to keep it a little bit shorter...
    Public Const arPT_New As String = "NewSale"
    Public Const arPT_Prv As String = "Previous Payments"
    Public Const arPT_PLC As String = "Previous L/C"
    Public Const arPT_L_C As String = "Late Charge"
    Public Const arPT_Mem As String = "Memo"
    Public Const arPT_PrB As String = "Previous Balance"

    Public Const arPT_stReO As String = "A/R Re-opened"
    Public Const arPT_stClo As String = "A/R Closed"
    Public Const arPT_stVoi As String = "A/R Void"
    Public Const arPT_stWtO As String = "A/R Write Off"
    Public Const arPT_stRep As String = "A/R Repo"
    Public Const arPT_stLeg As String = "A/R Legal"
    Public Const arPT_stBkr As String = "A/R Bankruptcy"

    Public Const arPT_Doc As String = "Doc Fees"
    Public Const arPT_Lif As String = "Life Ins."
    Public Const arPT_Acc As String = "Acc. Ins."
    Public Const arPT_Pro As String = "Prop. Ins."
    Public Const arPT_IUI As String = "IUI Ins."
    Public Const arPT_Int As String = "Interest Chg."
    Public Const arPT_Tax As String = "Int. Sls Tax"
    Public Const arPT_PTI As String = "Post Term Interest"

    Public Const arPT_crDoc As String = "Doc Credit"
    Public Const arPT_crPri As String = "Principal Credit"
    Public Const arPT_crL_C As String = "L/C Credit"
    Public Const arPT_crInt As String = "Interest Credit"
    Public Const arPT_crLif As String = "Life Credit"
    Public Const arPT_crAcc As String = "Acc. Credit"
    Public Const arPT_crPro As String = "Prop. Credit"
    Public Const arPT_crIUI As String = "IUI Credit"
    Public Const arPT_crTax As String = "Sales Tax Credit"

    Public Const arPT_crDoc2 As String = "Credit Doc"
    Public Const arPT_crPri2 As String = "Credit Princ."
    Public Const arPT_crL_C2 As String = "Credit L/C"
    Public Const arPT_crInt2 As String = "Credit Interest"
    Public Const arPT_crLif2 As String = "Credit Life"
    Public Const arPT_crAcc2 As String = "Credit Accident"
    Public Const arPT_crPro2 As String = "Credit Property"
    Public Const arPT_crIUI2 As String = "Credit IUI"
    Public Const arPT_crTax2 As String = "Credit Sales Tax"

    Public Const arPT_dbPri As String = "Principal Debit"
    Public Const arPT_dbL_C As String = "L/C Debit"
    Public Const arPT_dbInt As String = "Interest Debit"

    Public Const arPT_dbPri2 As String = "Debit Prin."
    Public Const arPT_dbL_C2 As String = "Debit L/C"
    Public Const arPT_dbInt2 As String = "Debit Interest"

    Public Const arPT_poDoc As String = "Doc Payoff"
    Public Const arPT_poPri As String = "Principal Payoff"
    Public Const arPT_poL_C As String = "L/C Payoff"
    Public Const arPT_poInt As String = "Interest Payoff"
    Public Const arPT_poLif As String = "Life Payoff"
    Public Const arPT_poAcc As String = "Acc. Payoff"
    Public Const arPT_poPro As String = "Prop. Payoff"
    Public Const arPT_poIUI As String = "IUI Payoff"
    Public Const arPT_poTax As String = "Sls. Tax Payoff"

    Public Const arPT_pyCash As String = "Cash"
    Public Const arPT_pyChck As String = "Check"
    Public Const arPT_pyDebt As String = "Debit Card"
    Public Const arPT_pyVisa As String = "Visa"
    Public Const arPT_pyMCrd As String = "Master CArd"
    Public Const arPT_pyDisc As String = "Discover"
    Public Const arPT_pyAmex As String = "AMEX"

    Public Const arPT_pyL_C As String = "L/C Payment"

    Public Const ArAddOn_Nil As String = ""
    Public Const ArAddOn_New As String = "New"
    Public Const ArAddOn_Add As String = "Add"
    Public Const ArAddOn_AdT As String = "AddToNew"
    Public Const ArAddOn_Rev As String = "Revolving"


    Public Const ArPayoffMethod_Rule_78 As String = "Rule 78"
    Public Const ArPayoffMethod_Rule78b As String = "Rule78b"
    Public Const ArPayoffMethod_ProRata As String = "ProRata"
    Public Const ArPayoffMethod_Anticip As String = "Anticip"

    Public Function AccountHasRecentSaleNotes(ByVal SaleNo As String) As Boolean
        '::::AccountHasRecentSaleNotes
        ':::SUMMARY
        ': Used to check whether the Customer account has recent Sale notes
        ':::DESCRIPTION
        ': This function checks whether the customer account has recent Sale Notes using Sale number
        ':::PARAETERS
        ': - SaleNo
        ':::RETURN
        ': Boolean
        Dim RS As ADODB.Recordset
        On Error Resume Next
        RS = GetRecordsetBySQL("SELECT Count(*) AS Cnt FROM SaleNotes WHERE datediff('d',NoteDate,Date()) Between 0 and 31 AND BillOSale=""" & ProtectSQL(SaleNo) & """")
        AccountHasRecentSaleNotes = (RS("Cnt").Value > 0)
        RS.Close()
        RS = Nothing
    End Function
    Public Function AccountHasRecentARNotes(ByVal MailIndex as integer) As Boolean
        '::::AccountHasRecentARNotes
        ':::SUMMARY
        ': Check if customer has recent AR notes.
        ':::DESCRIPTION
        ': This function checks whether the customer account has recent Ar Notes.
        ':::PARAMETERS
        ': - MailIndex
        ':::RETURN
        ': Boolean
        Dim RS As ADODB.Recordset
        On Error Resume Next
        RS = GetRecordsetBySQL("SELECT Count(*) AS Cnt FROM ARNotes INNER JOIN InstallmentInfo on ARNotes.ArNo=InstallmentInfo.ArNo WHERE datediff('d',NoteDate,Date()) between 0 and 31 AND InstallmentInfo.MailIndex=" & ProtectSQL(MailIndex))
        AccountHasRecentARNotes = (RS("Cnt").Value > 0)
        RS.Close()
        RS = Nothing
    End Function

    Public ReadOnly Property UseAmericanNationalInsurance() As Boolean
        Get
            '  UseAmericanNationalInsurance = UseAmericanNationalInsurance Or IsTreehouse ' Treehouse doesn't use this anymore
            '  UseAmericanNationalInsurance = UseAmericanNationalInsurance Or IsBlueSky
            UseAmericanNationalInsurance = UseAmericanNationalInsurance Or IsMcClure
        End Get
    End Property

    Public ReadOnly Property UseAlabamaSection5_19_3() As Boolean
        Get
            UseAlabamaSection5_19_3 = IsSidesFurniture
        End Get
    End Property

    Public ReadOnly Property AlabamaFinanceCharges(ByVal CA As Decimal) As Decimal
        Get
            AlabamaFinanceCharges = Max(AlabamaFinanceChargesMin(CA), AlabamaFinanceChargesMax(CA))
        End Get
    End Property

    Public ReadOnly Property AlabamaFinanceChargesMin(ByVal CA As Decimal) As Decimal
        Get
            AlabamaFinanceChargesMin = IIf(CA <= 25, 4, 6)
        End Get
    End Property

    Public ReadOnly Property AlabamaFinanceChargesMax(ByVal CA As Decimal) As Decimal
        Get
            AlabamaFinanceChargesMax = IIf(CA < 750, 15 * CA / 100, 15 * 750 / 100) + IIf(CA < 750, 0, 10 * Max(0, (CA - 750)) / 100)
        End Get
    End Property

    Public Function ArAddOnAccount(ByVal vArNo As String) As String
        '::::ArAddOnAccount
        ':::SUMMARY
        ': Generates the appropriate ArNo for an Add on Account
        ':::DESCRIPTION
        ': Adds/modfies ArNo suffix by augmenting the final letter.
        ':
        ': Central way to calculate the appropriate add on account number to use for an add on.  Does not create the add on.
        ':::PARAMETERS
        ': - ArNo
        ':::RETURN
        ': String
        ArAddOnAccount = AugmentByRightLetter(vArNo)
        Do While ArNoExists(ArAddOnAccount)
            ArAddOnAccount = AugmentByRightLetter(ArAddOnAccount)
        Loop
    End Function

End Module
