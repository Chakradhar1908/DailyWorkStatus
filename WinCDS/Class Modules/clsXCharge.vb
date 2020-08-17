Imports PINPadDevice
Imports VBA
Public Class clsXCharge
#Const AllowDebit = True
#Const AllowGift = True

    Private mXCTran As Object

    Public FormHandle As Integer             ' some form's hWnd
    Public ShowStatus As Boolean
    Public ShowResult As Boolean

    Public Clerk As String
    Public Receipt As String
    Public Amount As Decimal

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
    Public XCTransIDResult As String
    Public ApprovedAmountResult As String
    Public BalanceAmountResult As String


    Private IsManuallyEntered As Boolean  ' Is the CCNum privately entered?

    Private Enum PinPadResults
        PPR_SUCCESS = 0
        PPR_NOTSUPPORTED = 1
        PPR_COMMERR = 2
        PPR_FCANCELLED = 3
        PPR_FFAILED = 4
        PPR_NODEVICECONFIGURED = 5
        PPR_DEVICENOTCONFIGURED = 6
        PPR_INITFAIL = 7
        PPR_INVALIDCARDTYPE = 8
        PPR_INVALIDCCNUM = 9
        PPR_INVALIDPURCHASEAMOUNT = 10
    End Enum

    Public Function ExecPurchase(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("ExecPurchase")
        IsManuallyEntered = False
        If Prompt Then If Not PromptCC() Then Exit Function
        LogText("XCTran Object: " & TypeName(XCTran))
        LogText("TransactionFolder: " & TransactionFolder)
        If Not MCPAM Then
            If XCTran.XCPurchase(FormHandle, TransactionFolder, "Purchase",
        ShowStatus, ShowResult, Clerk, Receipt, Swipe,
        ExpDate, Swipe, Amount, Zip,
        Address, CVV, ErrorMsg, ApprovalCode, AVSResult, CVVResult) Then
                LogText("ApprovalCode = " & ApprovalCode)
                Success = True
            Else
                LogText("XCPurchase Unsuccessfull - " & ErrorMsg)
            End If
        Else
            If XCTran.XCPurchaseEx2(FormHandle, TransactionFolder, "Purchase",
        ShowStatus, ShowResult, Clerk, Receipt, Swipe,
        ExpDate, Swipe, Amount, Zip,
        Address, CVV,
        MerchID, MarketType, Recurring, AllowDuplicate,
        PartialApprovalSupport, ErrorMsg, ApprovalCode, AVSResult, CVVResult,
        XCTransIDResult, ApprovedAmountResult, BalanceAmountResult) Then
                If PartialApprovalSupport Then
                    If GetPrice(Amount) <> GetPrice(ApprovedAmountResult) Then
                        ' BFH20110612 The amount was partially approved.
                        ' xcharge mandates support..  if we use their result msgbox, it will show the partial approval
                        ' so, we offer them the option of approving or cancelling the partial approval at their request.
                        ' if they do cancel, we must cancel the partial transaction, because it has already gone through.
                        Dim R As VbMsgBoxResult
                        LogText("Partial Approval Amount: " & ApprovedAmountResult & " (Tot=" & GetPrice(Amount) & ")... Cancelling Transaction", 2)
                        R = MsgBox("Amount was partially approved." & vbCrLf & "You can use an additional tender type, or cancel this partial transaction." & vbCrLf2 & "Press OK to continue, or Cancel this transaction.", vbExclamation + vbOKCancel, "Partial Approval on Credit Card")
                        If R = vbCancel Then
                            LogText("Partially Approved Transaction CANCELLED at users request", 1)
                            Success = CancelTransaction()
                            ExecPurchase = Success
                            Exit Function
                        End If
                        Amount = ApprovedAmountResult
                    End If
                End If
                LogText("ApprovalCode = " & ApprovalCode)
                Success = True
            Else
                LogText("XCPurchase Unsuccessfull - " & ErrorMsg)
            End If
        End If
        ExecPurchase = Success
    End Function

    Public Function ExecDebitPurchase(Optional ByVal Prompt As Boolean = True) As Boolean
        Dim Cancelled As Boolean
        LogStartFunction("XCDebitPurchase")
        'ActiveLog "clsXCharge::ExecDebitPurchase(Prompt=" & Prompt & ")", 5
        IsManuallyEntered = False
        If Prompt Then
            If Not PromptDebit() Then Exit Function
        Else
            GetPin(CC, Amount, Pin, key, Cancelled)
        End If
        If Cancelled Then Exit Function

        'ActiveLog "clsXCharge::ExecDebitPurchase:  Swipe=" & Swipe, 8
        'ActiveLog "clsXCharge::ExecDebitPurchase:  Pin=" & Pin, 8
        'ActiveLog "clsXCharge::ExecDebitPurchase:  Key=" & Key, 8
        Debug.Print(Swipe)
        Dim R As Boolean
        If Not MCPAM Then
            R = XCTran.XCDebitPurchase(FormHandle, TransactionFolder, "Debit Card Purchase", ShowStatus, ShowResult, Clerk, Receipt, Amount, "0.00", Amount, Track2, Track2, Pin, key, ErrorMsg, ApprovalCode)
        Else
            'BFH20110612
            ' for whatever reason, though they said Debit must be included,
            ' the XCTransaction2 DLL doesn't provided approved amount in this interface (yet),
            ' so we can't implement this
            R = XCTran.XCDebitPurchaseEx2(FormHandle, TransactionFolder, "Debit Card Purchase", ShowStatus, ShowResult, Clerk, Receipt, Amount, "0.00", Amount, Track2, Track2, Pin, key, MerchID, MarketType, AllowDuplicate, PartialApprovalSupport, ErrorMsg, ApprovalCode, XCTransIDResult, ApprovedAmountResult, BalanceAmountResult)
            If R And PartialApprovalSupport Then
                If GetPrice(Amount) <> GetPrice(ApprovedAmountResult) Then
                    ' BFH20110612 The amount was partially approved.
                    ' xcharge mandates support..  if we use their result msgbox, it will show the partial approval
                    ' so, we offer them the option of approving or cancelling the partial approval at their request.
                    ' if they do cancel, we must cancel the partial transaction, because it has already gone through.
                    Dim RR As VbMsgBoxResult
                    LogText("Partial Approval Debit Amount: " & ApprovedAmountResult & " (Tot=" & GetPrice(Amount) & ")... Cancelling Transaction", 2)
                    RR = MsgBox("Amount was partially approved." & vbCrLf & "You can use an additional tender type, or cancel this partial transaction." & vbCrLf2 & "Press OK to continue, or Cancel this transaction.", vbExclamation + vbOKCancel, "Partial Approval on Credit Card")
                    If RR = vbCancel Then
                        LogText("Partially Approved Transaction CANCELLED at users request", 1)
                        Success = CancelDebitTransaction()
                        ExecDebitPurchase = Success
                        Exit Function
                    End If
                    Amount = ApprovedAmountResult
                End If
            End If
        End If
        If R Then
            LogText("ApprovalCode = " & ApprovalCode)
            Success = True
        Else
            LogText("XCDebitPurchase Unsuccessfull - " & ErrorMsg)
        End If
        ExecDebitPurchase = Success
    End Function

    Public Function ExecGiftRedeem(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("XCGiftRedeem")
        IsManuallyEntered = False
        If Prompt Then If Not PromptGift() Then Exit Function
        If XCTran.XCGiftRedeem(FormHandle, TransactionFolder, "Gift Redeem",
      ShowStatus, ShowResult, CC, "M", CVV, Amount,
      Receipt, Clerk, "", ErrorMsg, Balance, AdditionalFunds) Then
            LogText("Balance = " & Balance)
            LogText("AdditionalFundsRequired = " & AdditionalFunds)
            Success = True
        Else
            LogText("XCGiftRedeem Unsuccessfull - " & ErrorMsg)
        End If
        ExecGiftRedeem = Success
    End Function

    Public ReadOnly Property XCC() As String
        Get
            If Len(CC) <= 10 Then Return String.Empty : Exit Property
            'XCC = String(Len(CC) - 4, "X") & Right(CC, 4)
            XCC = New String("X"c, Len(CC) - 4) & Right(CC, 4)
            XCC = Replace(XCC, "XXXXXXXXX", "X")
        End Get
    End Property

    Private Sub LogStartFunction(ByVal TransactionName As String)
        LogText("------  " & TransactionName & "  ------")
    End Sub

    Public Function PromptCC(Optional ByVal Caption As String = "Enter Card Information or Swipe Card") As Boolean
        On Error Resume Next
Again:
        Success = False
        If Not MCPAM Then
            PromptCC = XCTran.PromptCreditCardEntry(FormHandle, Caption, False,
    False, False, Swipe, Track1, Track2, CC, CCTypeName, ExpirationMonth,
    ExpirationYear, CardHolderName, Zip, Address, CVV)
        Else
            Dim EnableZip As Boolean, EnableCVV As Boolean, EnableAddress As Boolean
            Dim RequireZip As Boolean, RequireCVV As Boolean, RequireAddress As Boolean
            Dim AutoExitAfterSwipe As Boolean
            EnableZip = True : EnableCVV = True : EnableAddress = True
            RequireZip = False : RequireCVV = False : RequireAddress = False
            AutoExitAfterSwipe = True
            PromptCC = XCTran.PromptCreditCardEntryEx(FormHandle, Caption,
    RequireZip, RequireAddress, RequireCVV, EnableZip, EnableAddress, EnableCVV, AutoExitAfterSwipe,
    Swipe, Track1, Track2, CC, CCTypeName, ExpirationMonth,
    ExpirationYear, CardHolderName, Zip, Address, CVV)
        End If
        IsManuallyEntered = CardHolderName = "" ' PromptCC
        If Not PromptCC Then
            '    If MsgBox("Card information not read.", vbRetryCancel, "Card read failed.") = vbRetry Then
            '      GoTo Again
            '    End If
            ErrorMsg = "Credit Card Prompt Failed"
        End If
    End Function

    Private Sub LogText(ByVal Text As String, Optional ByVal Priority As Integer = 4)
        mLog = mLog & IIf(Len(mLog) > 0, vbCrLf, "") & Text
        ActiveLog("clsXCharge::" & Text, Priority)
        LogFile("XCharge.txt", Text)
        '  If FileExists(AppFolder & "XCLog.txt") Then
        '    WriteFile AppFolder & "XCLog.txt", Text & vbCrLf
        '  End If
    End Sub

    Public ReadOnly Property MCPAM() As Boolean
        Get
            MCPAM = Trim(MerchID) <> ""
        End Get
    End Property

    Public ReadOnly Property XCTran() As XCTransaction2.XChargeTransaction
        Get
            If mXCTran Is Nothing Then mXCTran = New XCTransaction2.XChargeTransaction
            XCTran = mXCTran
        End Get
    End Property

    Public ReadOnly Property TransactionFolder(Optional ByVal ForceMoto As Boolean = False, Optional ByVal StoreNo As Integer = 0) As String
        Get
            Dim S As String
            If StoreNo = 0 Then StoreNo = StoresSld
            S = GetXCTransactionFolder(StoreNo, IsManuallyEntered Or ForceMoto)
            If S = "" And StoreNo <> -1 Then S = GetXCTransactionFolder(-1, IsManuallyEntered Or ForceMoto)
            If S = "" Then S = DefaultTransactionFolder(ForceMoto) ' , 1)
            'Debug.Print "Transaction Folder: " & S
            TransactionFolder = S
        End Get
    End Property

    Private ReadOnly Property ExpDate() As String
        Get
            ExpDate = ExpirationYear & ExpirationMonth
        End Get
    End Property

    Public ReadOnly Property MerchID() As String
        Get
            MerchID = IIf(MarketType = "M", MOTOMerchID, RetailMerchID)
        End Get

    End Property

    Public ReadOnly Property MarketType() As String      ' R = Retail, M = MOTO, E = E-Commerce
        Get
            If IsManuallyEntered And XCHasMoto Then
                MarketType = "M"
            Else
                MarketType = "R"
            End If
            If MarketType <> "R" And MarketType <> "M" And MarketType <> "E" Then MarketType = "R"
        End Get

    End Property

    Public ReadOnly Property AllowDuplicate() As Boolean
        Get
            AllowDuplicate = False
        End Get

    End Property

    Public ReadOnly Property PartialApprovalSupport() As Boolean
        Get
            PartialApprovalSupport = XCHasPASupport
        End Get

    End Property

    Private Function CancelTransaction() As Boolean
        Dim VoidIt As Boolean, E As String = "", A As String = "", Tid As String = ""
        VoidIt = XCTran.XCVoidEx(FormHandle, TransactionFolder, "Void Partial Payment",
      False, False, Clerk,
      "", "", "", "",
      XCTransIDResult, ApprovedAmountResult, MerchID, MarketType, E, A, Tid)
        CancelTransaction = False  ' return false, so we can set the actual transaction results.
    End Function

#If AllowDebit Then
    Public Function PromptDebit() As Boolean
        Dim Cancelled As Boolean
        Success = False
        PromptDebit = XCTran.PromptDebitCardEntry(FormHandle, "Debit Card Information", CC, Track2)
        IsManuallyEntered = Track2 = "" 'PromptDebit
        If Not PromptDebit Then Exit Function
        Success = True
        GetPin(CC, 0, Pin, key, Cancelled)
        If Not PromptDebit Then ErrorMsg = "Debit Card Prompt Failed"
        If key = "" Then Cancelled = True
        If Cancelled Then ErrorMsg = "Debit Card Prompt Pin Cancelled" : Success = False : PromptDebit = False
    End Function
#End If

    Private Sub GetPin(ByRef CCNum As String, ByRef Amount As Decimal, ByRef vPIN As String, ByRef vKEY As String, Optional ByRef Cancelled As Boolean = False)
        Dim PP As PINPad, HH As Integer, Res As Integer
        Dim X As Integer
        PP = CreatePinPadObject()
        If PP Is Nothing Then Exit Sub
        On Error Resume Next
        '  PP.Reset 0    ' BFH20081211 - wasn't closing.. trying some thing...
        '  hH = Random(50000)

        Pin = "" : key = ""
        X = PP.Init(HH)
        If X = 5 Then
            MsgBox("No Pin Pad configured." & vbCrLf & "You must set up your Pin Pad Device in the Store Setup section.", vbCritical, "No Pinpad Configured")
            Cancelled = True
            PP = Nothing
            Exit Sub
        End If
        PinPadResult(PP, X)
        On Error GoTo 0
        Res = PP.PromptDebitPIN(HH, CCNum, Amount, vPIN, vKEY)
        Cancelled = (Res = PinPadResults.PPR_FCANCELLED)
        If Not Cancelled Then PinPadResult(PP, Res)
        PinPadResult(PP, PP.Close(HH))

        PP = Nothing
    End Sub

    Private Function CancelDebitTransaction() As Boolean
        Dim VoidIt As Boolean, E As String = "", A As String = "", Tid As String = ""
        VoidIt = XCTran.XCDebitReturnEx(FormHandle, TransactionFolder, "Returning Transaction", False, False, Clerk, "", ApprovedAmountResult, "0.00", ApprovedAmountResult, Track2, Track2, Pin, key, MerchID, MarketType, AllowDuplicate, E, A, Tid)
        CancelDebitTransaction = False  ' return false, so we can set the actual transaction results
    End Function

#If AllowGift Then
    Public Function PromptGift() As Boolean
        Success = False
        PromptGift = XCTran.PromptGiftCardEntry(FormHandle, "Gift Card Information", False, CC, Swipe, CVV)
        IsManuallyEntered = CVV = ""  ' PromptGift
        If Not PromptGift Then ErrorMsg = "Gift Card Prompt Failed" Else Success = True
    End Function
#End If

    Public ReadOnly Property DefaultTransactionFolder(Optional ByVal ForceMoto As Boolean = False, Optional ByVal StoreNo As Integer = -1) As String
        Get
            Const cLT As String = "LocalTran"
            Dim fXC As String, dLT As String
            Dim MotoDir As String
            If StoreNo <= 0 Then StoreNo = StoresSld

            fXC = XChargeFolder()
            dLT = fXC & cLT
            MotoDir = dLT & "-MOTO"

            If DirExists(MotoDir & StoreNo) And (IsManuallyEntered Or ForceMoto) Then
                DefaultTransactionFolder = MotoDir & StoreNo
                Exit Property
            ElseIf DirExists(MotoDir) And (IsManuallyEntered Or ForceMoto) Then
                DefaultTransactionFolder = MotoDir
                Exit Property
            Else
                'bfh20080531 - one checking account means go to store 1 folder
                If StoreNo = 1 Or StoreSettings.bPostToLoc1 Then ' jk maintains original single stores set up
                    DefaultTransactionFolder = dLT
                    If DirExists(DefaultTransactionFolder) Then Exit Property
                Else 'keeps data bases different for multi stores
                    DefaultTransactionFolder = dLT
                    If DirExists(DefaultTransactionFolder) Then Exit Property
                End If
            End If

            MsgBox("Could not locate transaction folder." & vbCrLf & dLT & StoreNo, vbCritical, "X-Charge Not Installed?")
            '  DefaultTransactionFolder = ""
        End Get
    End Property

    Public ReadOnly Property MOTOMerchID() As String
        Get
            MOTOMerchID = CSVField(StoreSettings.CCConfig, 3)
        End Get
    End Property

    Public ReadOnly Property RetailMerchID() As String
        Get
            RetailMerchID = CSVField(StoreSettings.CCConfig, 1)
        End Get
    End Property

    Public ReadOnly Property XCHasMoto() As Boolean
        Get
            XCHasMoto = Trim(MOTOMerchID) <> ""
        End Get

    End Property

    Public ReadOnly Property XCHasPASupport() As Boolean
        Get
            '  XCHasPASupport = InStr(CSVField(StoreSettings.CCConfig, 2), "P") <> 0
            XCHasPASupport = True
        End Get
    End Property

    Private Function CreatePinPadObject() As PINPadDevice.PINPad
        On Error Resume Next
        CreatePinPadObject = New PINPadDevice.PINPad
        If CreatePinPadObject Is Nothing Then
            MsgBox("Could not initiate Pin Pad Device Interface", vbExclamation, "Missing DLL??")
            Exit Function
        End If
    End Function

    Private Sub PinPadResult(ByRef PP As PINPad, ByRef x As Integer)
        If x <> 0 Then MsgBox(PP.GetResultMessage(x), vbExclamation, "Pin Error (" & x & ")")
    End Sub

    Public Function ExecReturn(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("XCReturn")
        IsManuallyEntered = False
        If Prompt Then If Not PromptCC() Then Exit Function
        Dim R As Boolean
        If Not MCPAM Then
            R = XCTran.XCReturn(FormHandle, TransactionFolder, "Return",
      ShowStatus, ShowResult, Clerk, Receipt, Swipe,
      ExpDate, Swipe, Amount,
      ErrorMsg, ApprovalCode)
        Else
            R = XCTran.XCReturnEx(FormHandle, TransactionFolder, "Return",
      ShowStatus, ShowResult, Clerk, Receipt, CC, ExpDate, Swipe,
      Amount, MerchID, MarketType, AllowDuplicate,
      ErrorMsg, ApprovalCode, XCTransIDResult)
        End If
        If R Then
            LogText("ApprovalCode = " & ApprovalCode)
            Success = True
        Else
            LogText("XCReturn Unsuccessfull - " & ErrorMsg)
        End If
        ExecReturn = Success
    End Function

    Public Function ExecDebitReturn(Optional ByVal Prompt As Boolean = True) As Boolean
        Dim Cancelled As Boolean
        LogStartFunction("XCDebitReturn")
        IsManuallyEntered = False
        If Prompt Then
            If Not PromptDebit() Then Exit Function
        Else
            GetPin(CC, Amount, Pin, key, Cancelled)
        End If
        If Cancelled Or (Pin = "" And key = "") Then Exit Function

        Dim R As Boolean
        If Not MCPAM Then
            R = XCTran.XCDebitReturn(FormHandle, TransactionFolder, "Debit Card Return",
      ShowStatus, ShowResult, Clerk, Receipt, Amount,
      "0.00", Amount, Track2, Track2, Pin, key, ErrorMsg, ApprovalCode)
        Else
            R = XCTran.XCDebitReturnEx(FormHandle, TransactionFolder, "Debit Card Return",
      ShowStatus, ShowResult, Clerk, Receipt, Amount,
      "0.00", Amount, Track2, Track2, Pin, key, MerchID, MarketType, AllowDuplicate, ErrorMsg, ApprovalCode, XCTransIDResult)
        End If
        If R Then
            LogText("ApprovalCode = " & ApprovalCode)
            Success = True
        Else
            LogText("XCDebitReturn Unsuccessfull - " & ErrorMsg)
        End If
        ExecDebitReturn = Success
    End Function

    Public Function ExecGiftReturn(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("XCGiftReturn")
        IsManuallyEntered = False
        If Prompt Then If Not PromptGift() Then Exit Function
        If XCTran.XCGiftReturn(FormHandle, TransactionFolder, "Gift Return",
      ShowStatus, ShowResult, CC, "M", CVV, Amount,
      Receipt, Clerk, "", ErrorMsg, Balance) Then
            LogText("Balance = " & Balance)
            Success = True
        Else
            LogText("XCGiftReturn Unsuccessfull - " & ErrorMsg)
        End If
        ExecGiftReturn = Success
    End Function

End Class
