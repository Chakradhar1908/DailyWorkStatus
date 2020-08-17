Imports cipwin32
Public Class clsChargeItPro
    Private Const CONFIG_FILENAME As String = "ChargeItPro.cnf"

#Const AllowDebit = True
#Const AllowGift = True

    'Private mXCTran As Object
    Private mInt As cipwin32.EasyIntegrator

    Dim a As cipwin32.EasyIntegrator
    Public FormHandle As Integer             ' some form's hWnd
    Public ShowStatus As Boolean
    Public ShowResult As Boolean

    Public Clerk As String
    Public Receipt As String
    Public Amount As Decimal
    Private mRefID As String

    ' Global variables
    Public ErrorMsg As String
    Public ApprovalCode As String
    Public Success As Boolean
    Public mLog As String

    Public AVSResult As String
    Public CVVResult As String

    Public Balance As String
    Public AdditionalFunds As String

    Public Description As String
    Public Swipe As String
    Public Track1 As String
    Public Track2 As String

    'PIN and KEY must be retrieved from a PIN Pad for Debit to work correctly
    Public Pin As String
    Public key As String

    Public CC As String
    Public CCTypeName As String
    Public ExpirationMonth As String
    Public ExpirationYear As String
    Public CardHolderName As String
    Public Zip As String
    Public Address As String
    Public CVV As String


    'Public MerchID As String
    'Public MarketType As String
    Public Recurring As Boolean
    'Public AllowDuplicate As Boolean
    'Public PartialApprovalSupport As Boolean
    Public TransIDResult As String
    Public ApprovedAmountResult As String
    Public BalanceAmountResult As String


    Private IsManuallyEntered As Boolean  ' Is the CCNum privately entered?

    Public Function ExecPurchase(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("ExecPurchase")
        '  IsManuallyEntered = False
        '  If Prompt Then If Not PromptCC Then Exit Function
        LogText("CIP-Tran Object: " & TypeName(CIP))
        '  LogText "TransactionFolder: " & TransactionFolder

        Dim cP As cipwin32.EasyIntegrator
        cP = CIP(True, True)
        cP.TransFields.AmountTotal = Amount
        ExecPurchase = cP.CreditSale()
        Amount = cP.TransFields.AmountTotal  ' For partial approval

        ApprovalCode = cP.ResultsFields.ApprovalNumberResult
        CC = cP.ResultsFields.MaskedAccount
        ExecPurchase = cP.ResultsFields.ResultStatus
        ErrorMsg = cP.ResultsFields.ResultMessage
        TransIDResult = cP.ResultsFields.UniqueTransID
    End Function

#If AllowDebit Then

    Public Function ExecDebitPurchase(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("CIP-DebitPurchase")

        CIP(True, True).TransFields.AmountTotal = Amount
        ExecDebitPurchase = CIP.DebitSale
        Amount = CIP.TransFields.AmountTotal ' For partial approval

        ApprovalCode = CIP.ApprovalNumberResult
        CC = CIP.MaskedAccount
        ExecDebitPurchase = CIP.ResultStatus
        ErrorMsg = CIP.ResultMessage
        TransIDResult = CIP.UniqueTransID
    End Function
    Public Function ExecDebitReturn(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("CIP-DebitReturn")

        CIP(True, True).TransFields.AmountTotal = Amount
        ExecDebitReturn = CIP.DebitReturn

        ApprovalCode = CIP.ResultsFields.ApprovalNumberResult
        CC = CIP.ResultsFields.MaskedAccount
        ExecDebitReturn = CIP.ResultsFields.ResultStatus
        ErrorMsg = CIP.ResultsFields.ResultMessage
        TransIDResult = CIP.ResultsFields.UniqueTransID
    End Function

#End If

    Public Function ExecGiftRedeem(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("CIP-GiftRedeem")
    End Function

    Public ReadOnly Property XCC() As String
        Get
            XCC = CC
        End Get
    End Property

    Private Sub LogStartFunction(ByVal TransactionName As String)
        LogText("------  " & TransactionName & "  ------")
    End Sub

    Private Sub LogText(ByVal Text As String, Optional ByVal Priority As Integer = 4)
        mLog = mLog & IIf(Len(mLog) > 0, vbCrLf, "") & Text
        ActiveLog("clsChargeItPro::" & Text, Priority)
        LogFile("XCharge.txt", Text)
        '  If FileExists(AppFolder & "XCLog.txt") Then
        '    WriteFile AppFolder & "XCLog.txt", Text & vbCrLf
        '  End If
    End Sub


    Public ReadOnly Property CIP(Optional ByVal Reset As Boolean = False, Optional ByVal LoadSetup As Boolean = False) As cipwin32.EasyIntegrator
        Get
            If Reset Then mInt = Nothing
            If IsNothing(mInt) Then mInt = AddControlToForm("cipwin32.EasyIntegrator")
            CIP = mInt
            If LoadSetup Then
                CIP.LoadSetup(ParamString)
                CIP.Clear()
                CIP.ConfigFields.AllowPartialApprovals = True
                CIP.TransFields.Cashier = GetCashierName
                CIP.TransFields.ReportName = RefId
                CIP.TransFields.PartnerSoftwareName = ProgramName
                CIP.TransFields.PartnerSoftwareVersion = SoftwareVersion(False, True)
            End If
        End Get
    End Property

    Public Property ParamString() As String
        Get
            ParamString = ReadEntireFile(ParamStringFile)
        End Get
        Set(value As String)
            WriteFile(ParamStringFile, value, True)
            '  If vData = ParamString Then Exit Property
            '  WriteStoreSetting StoresSld, iniSection_StoreSettings, "CCConfig", vData
            '  ResetStoreSettings
        End Set
    End Property

    Public ReadOnly Property ParamStringFile() As String
        Get
            ParamStringFile = CDSAppDataFolder() & CONFIG_FILENAME
        End Get
    End Property

    Public Property RefId() As String
        Get
            Dim S As String, T As String
            Dim C As Integer

            ' This is only set once (per object)
            If mRefID <> "" Then RefId = mRefID : Exit Property

            T = "" & Random(10000000)
            If IsFormLoaded("BillOSale") Then
                S = BillOSale.BillOfSale.Text
                If S = "" Then S = "New Sale"
                BillOSale.TotalPaymentsByType("", C)
                C = C + 1
                S = S & "." & C
            ElseIf IsFormLoaded("CashRegister") Then
                S = "CashRegister"
            End If

            If S = "" Then S = T
            RefId = S
            mRefID = RefId
        End Get
        Set(value As String)
            mRefID = value
        End Set
    End Property

    Public Function ExecVoid(Optional ByVal Prompt As Boolean = True, Optional ByVal TransID As String = "") As Boolean
        LogStartFunction("CIP-Void")
        CIP(True, True).TransFields.TransactionReference = TransID
        CIP.TransFields.AmountTotal = Amount
        '    ExecVoid = .GenericVoid
        ExecVoid = CIP.CreditReturn

        ApprovalCode = CIP.ResultsFields.ApprovalNumberResult
        CC = CIP.ResultsFields.MaskedAccount
        ExecVoid = CIP.ResultsFields.ResultStatus
        ErrorMsg = CIP.ResultsFields.ResultMessage
        TransIDResult = CIP.ResultsFields.UniqueTransID
    End Function

    Public Function ExecReturn(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("CIP-Return")

        Dim cP As EasyIntegrator
        cP = CIP(True, True)
        cP.TransFields.AmountTotal = Amount
        ExecReturn = cP.CreditReturn()

        ApprovalCode = cP.ResultsFields.ApprovalNumberResult
        CC = cP.ResultsFields.MaskedAccount
        ExecReturn = cP.ResultsFields.ResultStatus
        ErrorMsg = cP.ResultsFields.ResultMessage
        TransIDResult = cP.ResultsFields.UniqueTransID
    End Function

    Public Function ExecGiftReturn(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("CIP-GiftReturn")
    End Function

End Class
