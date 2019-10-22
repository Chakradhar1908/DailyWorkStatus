Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class clsTransactionCentral
#Const aAllowCredit = True
#Const aDoesPinPad = False
#Const aAllowDebit = False
#Const aAllowGift = False

    'Transaction Central Merchant ID (TC ID): 54369
    'Transaction Central RegKey : 8W4LSKJERGZLXTG6
    'Transaction Central Virtual Terminal URL: https://www.oc2net.net/billing/login.asp
    'Transaction Central Virtual Terminal Pwd: XLMH2024

    Public MerchantID As String
    Public RegKey As String

    Public ShowStatus As Boolean
    Public ShowResult As Boolean

    Public Clerk As String
    Public Receipt As String
    Public Amount As Decimal
    Public SettledDate As String
    Public PostedDate As String

    ' Global variables
    Public ErrorMsg As String
    Public TransID As String
    Public CreditID As String
    Public ApprovalCode As String
    Private mRefID As String
    Public Success As Boolean
    Public mLog As String

    Public AVSCode As String
    Public CVVResult As String
    Public Status As String
    Public Message As String
    Public RefCode As String

    Public Balance As String
    Public AdditionalFunds As String

    Public Description As String
    Public Swipe As String
    Public Track1 As String
    Public Track2 As String

    'PIN and KEY must be retrieved from a PIN Pad for Debit to work correctly
    Public EncryptedPin As String
    Public SMID As String
    Public PINFormat As String

    Public CC As String
    Public CCTypeName As String
    Public CVV2 As String
    Public ExpirationMonth As String
    Public ExpirationYear As String
    Public CardHolderName As String
    Public Zip As String
    Public Address As String
    Public CVV As String

    Public CommunicationRetries As Integer

    Private IsManuallyEntered As Boolean  ' Is the CCNum privately entered?

    Public Function ExecVoid(Optional ByVal SaleDate As String = Nothing) As Boolean
        PostedDate = SaleDate
        If Not IsDate(SaleDate) Then SaleDate = DateAdd("d", -2, Today)
        If DateAfter(Today, SaleDate, False) Then
            ExecVoid = ExecCredit()
            Exit Function
        End If

        LogStartFunction("TransCentral-ExecVoid")
#If aAllowCredit Then
        IsManuallyEntered = False

        Dim S As String, A As String, Res As clsHashTable
        CCQSStart(S, A)
        S = S & A & "TransID=" & TransID

        ExecVoid = DoTransaction("VoidCreditCardSale", S, Res)

        If Not ExecVoid Then
            LogText("VoidCreditCardSale unsuccessful")
            Exit Function
        End If


        If Res.Item("TransID") <> "111" Then TransID = Res.Item("TransID")
        ExecVoid = Val(TransID) <> 0
        '  ApprovalCode = Res.Item("AuthCode")

        '  PostedDate = Res.Item("PostedDate")
        '  SettledDate = Res.Item("SettledDate")

        '  Amount = GetPrice(Res.Item("Amount"))
        '  ApprovalCode = Res.Item("AuthCode")
        '  Status = Res.Item("Status")
        '  AVSCode = Res.Item("AVSCode")
#End If
    End Function
    ' after settling only (midnight of sale date)
    Public Function ExecCredit() As Boolean
        LogStartFunction("TransCentral-ExecCredit")
#If aAllowCredit Then
        IsManuallyEntered = False
        '  If Prompt Then If Not PromptCC Then Exit Function
        If RefId = "" Then RefId = "" & Random(10000000)  ' gives a better transaction rate

        Dim S As String, A As String, Res As clsHashTable
        CCQSStart(S, A)
        S = S & A & "Amount=" & Amount
        S = S & A & "TransID=" & TransID
        S = S & A & "RefID=" & RefId

        ExecCredit = DoTransaction("CreditCardCredit", S, Res)
        If Not ExecCredit Then
            LogText("CreditCardCredit unsuccessful")
            Exit Function
        End If


        TransID = Res.Item("TransID")
        ExecCredit = Val(TransID) <> 0
        '  RefID = Res.Item("RefID")

        '  PostedDate = Res.Item("PostedDate")
        '  SettledDate = Res.Item("SettledDate")

        '  Amount = GetPrice(Res.Item("Amount"))
        '  Status = Res.Item("Status")
        '  Message = Res.Item("Message")
#End If
    End Function

    Public Property RefId() As String
        Get
            RefId = mRefID
            If RefId = "" Then RefId = UCase(DetectSaleNo) & "-" & Random(1000)
            If RefId = "" Then RefId = "" & Random(10000000)   ' gives a better transaction rate
        End Get
        Set(value As String)
            mRefID = value
        End Set
    End Property

    Private Sub LogStartFunction(ByVal TransactionName As String)
        LogText("")
        LogText("------  " & TransactionName & "  ------")
    End Sub
    Private Sub CCQSStart(ByRef S As String, ByRef A As String)
        S = ""
        S = S & A & "MerchantID=" & MerchantID
        A = "&"
        S = S & A & "RegKey=" & RegKey
    End Sub
    Private Function DoTransaction(ByVal Action As String, ByVal QS As String, ByRef Res As clsHashTable, Optional ByVal ProgressBar As Boolean = True, Optional ByVal ForceRetries As Integer = -1) As Boolean
        Dim R As String, I As Integer, FL As String
        LogStartFunction("Transaction Central: " & Action)
        If ForceRetries = -1 Then ForceRetries = CommunicationRetries

        'If ProgressBar Then ProgressForm(0, 1, "Processing.  Please wait...")

        If DoExtraLog Then LogText("Request URL: " & SOAPURL(Action) & "?" & QS)
        R = DownloadURLToString(SOAPURL(Action) & "?" & QS)
        If DoExtraLog Then
            LogText("---------------------------  RESPONSE  -------------------------------")
            LogText(R)
            LogText("-------------------------  END RESPONSE  -----------------------------")
        End If

        If R = "" Then
            If ForceRetries > 0 Then        ' useful for some reports, but not transactions
                For I = 1 To ForceRetries
                    R = DownloadURLToString(SOAPURL(Action) & "?" & QS, FL)

                    If R <> "" Then Exit For
                Next
            ElseIf Action = "CreditCardSaleAll" Then
                ' if an exec fails in responding, but goes through, we do a quick background lookup to find the
                ' appropriate information (transid, approval, etc).  We load the info automatically, and return
                ' automatically
                If FindDuplicatedTransction(RefId, Amount) Then
                    DoTransaction = True
                    Exit Function
                End If
            ElseIf Action = "VoidCreditCardSale" Or Action = "CreditCardCredit" Then
                Dim S As String
                S = CheckTransIDStatus(TransID, PostedDate)
                If S = "V" Or S = "C" Then
                    Res = New clsHashTable
                    Res.Add("TransID", "111")
                    Res.Add("Amount", Amount)
                    DoTransaction = True
                    Exit Function
                End If
            End If
        End If     ' we got some response
        'If ProgressBar Then ProgressForm

        DoTransaction = HandleResponse(Action, QS, R, Res, FL)
    End Function

    Private Function HandleResponse(ByVal Action As String, ByVal QS As String, ByVal R As String, ByRef Res As clsHashTable, Optional ByVal FailureMessage As String = "") As Boolean
        Dim TransID As String, Status As String, Msg As String
        If R = "" Then
            MsgBox("Communication Failed." & IIf(FailureMessage <> "", vbCrLf & FailureMessage, ""))
            If DoExtraLog Then LogText("Communication Failed (" & Action & "):  Host Returned an Empty String" & IIf(FailureMessage <> "", vbCrLf & FailureMessage, ""))
            Exit Function
        End If

        If Not ParseResponse(R, Res) Then
            MsgBox("Declined." & IIf(IsDevelopment, vbCrLf & "Parse Failed." & vbCrLf & Action & vbCrLf & QS, ""), vbDefaultButton3)
            LogText("Parse Failed, Msg: " & R)
            Exit Function
        End If

        If InStr(LCase(Action), "report") Then HandleResponse = True : Exit Function

        TransID = Res.Item("TransID")
        CreditID = Res.Item("CreditID")
        If Val(TransID) = 0 Then TransID = CreditID
        Status = Res.Item("Status")
        Msg = Res.Item("Message")

        If Val(TransID) = 0 Then
            If IsDevelopment() Then
                MsgBox("Declined: " & Msg & IIf(IsDevelopment, vbCrLf & Action & vbCrLf & "TransID: 0" & vbCrLf & QS, ""), vbDefaultButton3)
            Else
                MsgBox("Declined.")
            End If
            LogText("Purchase failed: " & Msg & ", TransID: 0" & vbCrLf & QS)
            LogText("QS=" & QS)
            Exit Function
        End If

        LogText("TransID: " & TransID)

        If Status = "Declined" Then
            MsgBox("Declined." & IIf(IsDevelopment, vbCrLf & Action & vbCrLf & "TransID: " & TransID, ""), vbDefaultButton3)
            LogText("Transaction Declined: " & Msg & ", TransID=" & TransID)
            Exit Function
        End If

        HandleResponse = True
    End Function

    Private ReadOnly Property SOAPURL(ByVal Action As String) As String
        Get
            Select Case LCase(Action)
                Case "transactiondetailreport", "transactionsummaryreport"
                    SOAPURL = "https://webservices.primerchants.com/tcreports.asmx/" & Action
                Case Else
                    SOAPURL = "https://webservices.primerchants.com/creditcard.asmx/" & Action
            End Select
        End Get
    End Property

    Public Function CheckTransIDStatus(ByVal nTransID As String, Optional ByVal DoDate As String = NullDateString) As String
        '  LogStartFunction "TransactionDetailreport"
        Dim S As String, A As String, Res As clsHashTable, R As Boolean

        CCQSStart(S, A)
        If DoDate = NullDateString Then DoDate = Today
        If Not IsDate(DoDate) Then DoDate = Today
        S = S & A & "FromDate=" & DateFormat(DoDate)
        S = S & A & "ToDate=" & DateFormat(DateAdd("d", 1, DoDate))
        S = S & A & "TransType=CC"    ' * CC/DC/CD/ACH
        S = S & A & "DisplayType="    '   SC, C, S, R
        S = S & A & "Status="         '   A, C, CO, D, O, OT, P, Q, R, RV, S, V
        S = S & A & "CardType="       '   BC, AM, DI, DN, JC, VI, MC
        S = S & A & "UserID="
        S = S & A & "TransID=" & nTransID
        S = S & A & "RefID="
        S = S & A & "AccountNumber="
        S = S & A & "AccountName="
        S = S & A & "Amount="

        R = DoTransaction("TransactionDetailreport", S, Res, False, 3)

        If Not R Then
            LogText("Transaction Lookup Unsuccessful")
            Exit Function
        End If

        On Error Resume Next
        TransID = 0
        CheckTransIDStatus = Trim(Res.Item(0).Item("Status"))
    End Function

    Public Function FindDuplicatedTransction(ByVal nRefID As String, ByVal nAmount As Decimal, Optional ByVal DoDate As String = NullDateString) As Boolean
        '  LogStartFunction "TransactionDetailreport"
        Dim S As String, A As String, Res As clsHashTable, R As Boolean

        CCQSStart(S, A)
        If DoDate = NullDateString Then DoDate = Today
        If Not IsDate(DoDate) Then DoDate = Today
        S = S & A & "FromDate=" & DateFormat(DoDate)
        S = S & A & "ToDate=" & DateFormat(DateAdd("d", 1, DoDate))
        S = S & A & "TransType=CC"    ' * CC/DC/CD/ACH
        S = S & A & "DisplayType="    '   SC, C, S, R
        S = S & A & "Status="         '   A, C, CO, D, O, OT, P, Q, R, RV, S, V
        S = S & A & "CardType="       '   BC, AM, DI, DN, JC, VI, MC
        S = S & A & "UserID="
        S = S & A & "TransID="
        S = S & A & "RefID=" & nRefID
        S = S & A & "AccountNumber="
        S = S & A & "AccountName="
        S = S & A & "Amount=" & nAmount

        R = DoTransaction("TransactionDetailreport", S, Res, False, 3)

        If Not R Then
            LogText("Transaction Lookup Unsuccessful")
            Exit Function
        End If

        On Error Resume Next
        TransID = 0
        TransID = Res.Item(0).Item("TransID")
        ApprovalCode = Res.Item(0).Item("AuthCode")
        PostedDate = Res.Item(0).Item("Posteddate")
        SettledDate = Res.Item(0).Item("SettledDate")
        Status = Res.Item(0).Item("Status")

        FindDuplicatedTransction = Val(TransID) <> 0
    End Function

    Private Sub LogText(ByVal Text As String)
        mLog = mLog & IIf(Len(mLog) > 0, vbCrLf, "") & Text
        ActiveLog("clsTransactionCentral::" & Text)
        LogFile("TransFirst.txt", Text)
        Exit Sub

        Debug.Print(Text)



        If DoExtraLog Then
            WriteFile(AppFolder() & "TransFirstLog.txt", Text & vbCrLf)
        ElseIf True Then
            If FileExists(AppFolder() & "TCLog.txt") Then
                WriteFile(AppFolder() & "TCLog.txt", Text & vbCrLf)
            End If
        End If
    End Sub
    Private ReadOnly Property DoExtraLog() As Boolean
        Get
            '  DoExtraLog = IsMattressKing
        End Get
    End Property
    Private Function ParseResponse(ByVal Body As String, ByRef Returns As clsHashTable) As Boolean
        Dim A As Integer, B As Integer, Ck As Boolean, X As Integer
        Dim ResultType As String, ResponseType As String
        Dim FieldName As String, FieldValue As String

        If Body = "" Then Exit Function
        On Error GoTo Failure
        Returns = New clsHashTable


        Returns.Add("_Body", Body)


        A = InStr(Body, "?>")
        If A = 0 Then Exit Function
        Body = NLTrim(Mid(Body, A + 2))

        A = InStr(Body, " ")
        B = InStr(Body, ">")
        If A > B Then A = B
        ResponseType = Mid(Body, 2, A - 2)
        Returns.Add("_ResponseType", ResponseType)
        Body = NLTrim(Mid(Body, B + 1))


        A = InStr(Body, "</" & ResponseType & ">")
        If A <> 0 Then Body = NLTrim(Mid(Body, 1, A - 1))

        ' Just the fields should remain...
        Dim Detail As Boolean, dR As Integer

        Do While True
            A = InStr(Body, ">")
            If A = 0 Then Exit Do
            FieldName = Mid(Body, 2, A - 2)

            If FieldName = "TransactionSummaryReport" Then
                Body = Replace(Body, "<TransactionSummaryReport>", "")
                Body = Replace(Body, "</TransactionSummaryReport>", "")
                Body = NLTrim(Body)
                A = InStr(Body, ">")
                If A = 0 Then Exit Do
                FieldName = Mid(Body, 2, A - 2)
            ElseIf FieldName = "TransactionDetailReport" Then
                Detail = True
                Body = NLTrim(Replace(Body, "<TransactionDetailReport>", "", 1, 1))
                GoTo NextOne
            ElseIf FieldName = "/TransactionDetailReport" Then
                dR = dR + 1
                Body = NLTrim(Replace(Body, "</TransactionDetailReport>", "", 1, 1))
                GoTo NextOne
            End If

            B = InStr(A, Body, "</")

            If B = 0 Then Exit Do
            FieldValue = Mid(Body, A + 1, B - A - 1)



            If Not Detail Then
                If Returns.Exists(FieldName) Then Returns.Remove(FieldName)
                Returns.Add(FieldName, FieldValue)
            Else
                If Not Returns.Exists(dR) Then Returns.Add(dR, New clsHashTable)
                If Returns.Item(dR).Exists(FieldName) Then Returns.Item(dR).Remove(FieldName)
                Returns.Item(dR).Add(FieldName, FieldValue)
            End If

            Body = NLTrim(Mid(Body, B + Len(FieldName) + 3))
NextOne:
        Loop

        ParseResponse = True
Failure:
    End Function

    Public Function BlindSwipe(ByVal CCSwipe As String) As Boolean
        On Error Resume Next
        'ActiveLog "clsTransactionCentral::BlindSwipe(CCSwipe=" & CCSwipe & ")", 7
        Swipe = CCSwipe
        BlindSwipe = ParseTrackData(CCSwipe, Track1, Track2, CC, CCTypeName, ExpirationMonth, ExpirationYear, CardHolderName)
    End Function

    Public Function ExecPurchase(Optional ByVal Prompt As Boolean = True) As Boolean
#If aAllowCredit Then

        IsManuallyEntered = False
        If Prompt Then If Not PromptCC() Then Exit Function

        Dim S As String, A As String, Res As clsHashTable
        If RefId = "" Then RefId = "" & Random(10000000) ' gives a better transaction rate

        CCQSStart(S, A)
        S = S & A & "RefID=" & RefId
        S = S & A & "Amount=" & CurrencyFormat(Amount, , , True)
        S = S & A & "SaleTaxAmount=0.00"
        S = S & A & "CardNumber=" & URLEncode(CC)
        S = S & A & "CardHolderName=" & IIf(CardHolderName <> "", CardHolderName, DetectCustomerName)
        S = S & A & "Expiration=" & ExpirationMonth & ExpirationYear
        S = S & A & "ZipCode=" & IIf(Zip <> "", Zip, DetectCustomerZipCode)
        S = S & A & "TrackData=" & URLEncode(Track2) ' Swipe ' Transaction Central, using GET and hence URL encode, can't URL Decode the ^ to its appropriate character.
        S = S & A & "PONumber="
        S = S & A & "CVV2=" & CVV2
        S = S & A & "Address="
        S = S & A & "UserID="
        S = S & A & "UsrDef1="
        S = S & A & "UsrDef2="
        S = S & A & "UsrDef3="
        S = S & A & "UsrDef4="

        ExecPurchase = DoTransaction("CreditCardSaleAll", S, Res)
        If Not ExecPurchase Then
            LogText("CreditCardSaleAll unsuccessful")
            Exit Function
        End If


        '  Dim rAmount as decimal
        '  Dim Status As String
        '  Dim Msg As String

        TransID = Res.Item("TransID")
        ApprovalCode = Res.Item("AuthCode")

        PostedDate = Res.Item("PostedDate")
        SettledDate = Res.Item("SettledDate")

        '  Status = Res.Item("Status")
        '  Msg = Res.Item("Message")
#End If
    End Function

    Public Function ExecDebitPurchase(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("TransCentralDebitPurchase")
#If aAllowDebit Then
  IsManuallyEntered = False
  If Prompt Then If Not PromptCC Then Exit Function
  
  Dim S As String, A As String, Res As clsHashTable
  CCQSStart S, A
  S = S & A & "RefID=" & RefId
  S = S & A & "Amount=" & CurrencyFormat(Amount, , , True)
  S = S & A & "CardNumber=" & CC
  S = S & A & "Expiration=" & ExpirationMonth & ExpirationYear
  S = S & A & "CardHolderName=" & IIf(CardHolderName <> "", CardHolderName, DetectCustomerName)
  S = S & A & "Address=" & Address
  S = S & A & "ZipCode=" & IIf(Zip <> "", Zip, DetectCustomerZipCode)
  S = S & A & "EncryptedPIN=" & EncryptedPin
  S = S & A & "SMID=" & SMID
  S = S & A & "PINFormat=" & PINFormat
  S = S & A & "CashbackAmount="
  S = S & A & "VoucherNumber="
  S = S & A & "UserID="
  S = S & A & "TrackData=" & Swipe
  S = S & A & "UsrDef="

  ExecDebitPurchase = DoTransaction("DebitCardSale", S, Res)
  If Not ExecDebitPurchase Then
    LogText "DebitCardSale unsuccessful"
    Exit Function
  End If
  
  
'  TransID = Res.Item("TransID")
'  ApprovalCode = Res.Item("AuthCode")
#End If
    End Function

    Public Function ExecGiftRedeem(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("TransCentralGiftRedeem")
#If aAllowGift Then
#End If
    End Function

    Public ReadOnly Property XCC() As String
        Get
            XCC = CCXOut(CC)
            XCC = Replace(XCC, "XXXXXXXX", "")
        End Get
    End Property

    Public Function PromptCC(Optional ByVal Caption As String = "Enter Card Information or Swipe Card") As Boolean
#If aAllowCredit Then
        On Error Resume Next
Again:
        Success = False
        Dim x As String
        PromptCC = ManualCCEntry(CC, x, CardHolderName, CVV2, Zip, Swipe)
        ExpirationMonth = Left(x, 2)
        ExpirationYear = Right(x, 2)

        IsManuallyEntered = Swipe = "" ' PromptCC
        If Not IsManuallyEntered Then
            ParseTrackData(Swipe, Track1, Track2)
        End If

        If Not PromptCC Then
            '    If MsgBox("Card information not read.", vbRetryCancel, "Card read failed.") = vbRetry Then
            '      GoTo Again
            '    End If
            ErrorMsg = "Credit Card Prompt Failed"
        End If
#End If
    End Function

    Public Function ExecBlindCredit(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("TransCentral-BlindCredit")
#If aAllowCredit Then
        IsManuallyEntered = False
        If Prompt Then If Not PromptCC() Then Exit Function

        '  If Prompt Then If Not PromptCC Then Exit Function
        If RefId = "" Then RefId = "" & Random(10000000)  ' gives a better transaction rate

        Dim S As String, A As String, Res As clsHashTable
        CCQSStart(S, A)
        S = S & A & "TransID=" & TransID
        S = S & A & "Amount=" & CurrencyFormat(Amount, , , True)
        S = S & A & "CardNumber=" & URLEncode(CC)
        S = S & A & "CardHolderName=" & IIf(CardHolderName <> "", CardHolderName, DetectCustomerName)
        S = S & A & "Expiration=" & ExpirationMonth & ExpirationYear
        S = S & A & "CVV2=" & CVV2
        S = S & A & "RefID=" & RefId
        S = S & A & "Address="
        S = S & A & "ZipCode=" & IIf(Zip <> "", Zip, DetectCustomerZipCode)
        S = S & A & "UserID="

        ExecBlindCredit = DoTransaction("BlindCredit", S, Res)
        If Not ExecBlindCredit Then
            LogText("Blind Credit unsuccessful")
            Exit Function
        End If

        TransID = Res.Item("TransID")
        If Val(TransID) = 0 Then TransID = Res.Item("CreditID")
        ExecBlindCredit = Val(TransID) <> 0
#End If
    End Function

End Class
