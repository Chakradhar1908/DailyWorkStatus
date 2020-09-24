Imports VBA
Public Class clsXChargeXpressLink
#Const aAllowCredit = True
#Const aDoesPinPad = False
#Const aAllowDebit = True
#Const aAllowGift = True

    'Public Clerk As String
    Public Receipt As String
    Public Amount As Decimal

    ' Global variables
    Public ErrorMsg As String
    Public TransID As String
    Public CreditID As String
    Public ApprovalCode As String
    Private mRefID As String
    Public mLog As String

    Public Description As String

    Public Zip As String, Address As String

    Public Success As Boolean
    Public XCTransIDResult As String
    Public ApprovedAmountResult As Decimal
    Public BalanceAmountResult As Decimal
    'Public Balance As String
    Public AdditionalFundsResult As Decimal
    Public CardHolderName As String
    Public CardExpiration As String
    Public XCC As String
    Public CCTypeName As String

    Private mResultFile As String

    Public Function ExecPurchase() As Boolean ' Optional ByVal Prompt As Boolean = True
        LogStartFunction("XCXL-ExecPurchase-" & DescribeState())
#If aAllowCredit Then
        Dim S As String, Res As clsHashTable
        S = ""
        S = S & ArgDefault()
        S = S & ArgTransType("PURCHASE")
        S = S & ArgAmount(Amount)

        Res = eXpressLink(S)
        If Res Is Nothing Then Exit Function
        If Not Res.Item("SUCCESS") Then
            Dim Msg As String
            Msg = "Failure: "
            Msg = Msg & Res.Item("REASON")  '  can fail...
            Msg = Msg & Res.Item("DESCRIPTION")
            'MsgBox(">> " & Msg)
            MessageBox.Show(">> " & Msg)
            Exit Function
        End If
        ExecPurchase = True

        If GetPrice(ApprovedAmountResult) And GetPrice(Amount) <> GetPrice(ApprovedAmountResult) Then
            ' BFH20110612 The amount was partially approved.
            ' xcharge mandates support..  if we use their result msgbox, it will show the partial approval
            ' so, we offer them the option of approving or cancelling the partial approval at their request.
            ' if they do cancel, we must cancel the partial transaction, because it has already gone through.
            Dim R As VbMsgBoxResult
            LogText("Partial Approval Amount: " & ApprovedAmountResult & " (Tot=" & GetPrice(Amount) & ")... Cancelling Transaction")
            R = MessageBox.Show("Amount was partially approved." & vbCrLf & "You can use an additional tender type, or cancel this partial transaction." & vbCrLf2 & "Press OK to continue, or Cancel this transaction.", "Partial Approval on Credit Card", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation)
            If R = vbCancel Then
                LogText("Partially Approved Transaction CANCELLED at users request")
                ExecPurchase = False
                Exit Function
            End If
            Amount = ApprovedAmountResult
        End If

        '  TransID = Res.item("TransID")
        '  PostedDate = Res.item("PostedDate")
        '  SettledDate = Res.item("SettledDate")
#End If
    End Function

    Public Function ExecDebitPurchase() As Boolean ' Optional ByVal Prompt As Boolean = True
        LogStartFunction("XCXL-DebitPurchase-" & DescribeState())
#If aAllowDebit Then
        Dim S As String, Res As clsHashTable
        S = ""
        S = S & ArgDefault()
        S = S & ArgTransType("DEBITPURCHASE")
        S = S & ArgAmount(Amount)

        Res = eXpressLink(S)
        If Res Is Nothing Then Exit Function
        If Not Res.Item("SUCCESS") Then Exit Function
        ExecDebitPurchase = True
#End If
    End Function

    Public Function ExecGiftRedeem() As Boolean ' Optional ByVal Prompt As Boolean = True
        LogStartFunction("XCXL-GiftRedeem-" & DescribeState())
#If aAllowGift Then
        Dim S As String, Res As clsHashTable
        S = ""
        S = S & ArgDefault()
        S = S & ArgTransType("GIFTREDEEM")
        S = S & ArgAmount(Amount)

        Res = eXpressLink(S)
        If Res Is Nothing Then Exit Function
        If Not Res.Item("SUCCESS") Then Exit Function
        ExecGiftRedeem = True
#End If
    End Function

    Private Sub LogStartFunction(ByVal TransactionName As String)
        LogText("")
        LogText("------  " & TransactionName & "  ------")
    End Sub

    Private Function DescribeState() As String
        Dim Q As String

        If Clerk <> "" Then Q = Q & ", Clerk=" & Clerk
        If Receipt <> "" Then Q = Q & ", Receipt=" & Receipt
        If Amount <> 0 Then Q = Q & ", Amount=" & Amount

        If Zip <> "" Then Q = Q & ", Zip=" & Zip
        If Address <> "" Then Q = Q & ", Address=" & Address

        If Len(Q) > 2 Then Q = Mid(Q, 3)

        DescribeState = Q
    End Function

    Private Function ArgDefault() As String
        Dim Q As String
        Q = ""
        ' Store Setup
        Q = Q & SafeArg("TITLE", "WinCDS POS Software")
        '  Q = Q & SafeArg("MID", MerchantID, True)
        Q = Q & SafeArg("USERID", XCUsername) & SafeArg("PASSWORD", XCPassword)
        ' Optional Configuration
        If Zip <> "" Or Address <> "" Then Q = Q & SafeArg("ZIP", Zip, True)
        If Address <> "" Then Q = Q & SafeArg("ADDRESS", Address, True)
        If Clerk <> "" Then Q = Q & SafeArg("CLERK", Clerk, True)
        If Receipt <> "" Then Q = Q & SafeArg("RECEIPT", Receipt, True)
        ' Appearance
        Q = Q & SafeArg("SMALLWINDOW")
        Q = Q & SafeArg("STAYONTOP") & SafeArg("SMARTAUTOPROCESS")
        Q = Q & SafeArg("HIDEMAINWINDOW")
        Q = Q & SafeArg("HIDEMAINMENU")
        Q = Q & SafeArg("HIDETOOLBAR")
        Q = Q & SafeArg("TOOLBAREXITBUTTON")
        Q = Q & SafeArg("HIDEWINDOWBORDER")
        ' Behavior
        Q = Q & SafeArg("LARGEPROCESSBUTTON")
        Q = Q & SafeArg("AUTOCLOSE")
        Q = Q & SafeArg("EXITWITHESCAPEKEY")
        Q = Q & SafeArg("EXITAFTERFAILEDLOGIN")
        ' Processing/Result Info
        Q = Q & SafeArg("RESULTFILE", ResultFile)
        Q = Q & SafeArg("XMLRESULTFILE")
        Q = Q & SafeArg("RECEIPTINRESULT")

        ArgDefault = Q
        '/TRANSACTIONTYPE:Purchase /LOCKTRANTYPE /AMOUNT:10.00 /LOCKAMOUNT /ZIP:89015 "/ADDRESS:1234 Easy Street" /RECEIPT:RC001 /LOCKRECEIPT /CLERK:Clerk /LOCKCLERK /MID:ABCDEFG /LOCKMID /STAYONTOP /SMARTAUTOPROCESS /AUTOCLOSE /HIDEMAINMENU /SMALLWINDOW "/TITLE:WinCDS POS Software" /TOOLBAREXITBUTTON /EXITWITHESCAPEKEY /USERID:system /PASSWORD:system /RESULTFILE:C:\Users\Owner\AppData\Local\Temp\ResultFile.txt
    End Function

    Private Function ArgTransType(ByVal tType As String) As String
        If Not IsIn(tType, "PURCHASE", "RETURN", "DEBITPURCHASE", "DEBITORCREDITPURCHASE", "FORCE", "PREAUTH", "ADJUSTMENT", "VOID", "CHECKVERIFY", "GIFTISSUE", "GIFTREDEEM", "GIFTRETURN", "GIFTBALANCEQUERY", "GIFTVOID", "ARCHIVEVAULTADD", "ARCHIVEVAULTUPDATE", "ARCHIVEVAULTDELETE", "ARCHIVEVALUEQUERY", "EBTSALE", "EBTCASHBENEFIT", "EBTRETURN", "EBTFORCE", "CHECKSALE", "CHECKCREDIT", "CHECKVERIFICATION", "ARCHIVECHECKVAULTADD", "ARCHIVECHECKVAULTDELETE", "ARCHIVECHECKVAULTQUERY", "DISPLAYREPORTSCREEN", "REPORTDATA") Then
            DevErr("Invalid Transaction Type: " & tType)
            Exit Function
        End If
        ArgTransType = SafeArg("TRANSACTIONTYPE", tType) & SafeArg("LOCKTRANTYPE")
    End Function

    Private Function ArgAmount(ByVal Amt As Decimal) As String
        ArgAmount = SafeArg("AMOUNT", SQLCurrency(Amt), True)
    End Function

    Private Function eXpressLink(ByVal Args As String) As clsHashTable
        Dim T As String
        Dim C As Object, R As Integer
        On Error Resume Next

        LogText("XCXL - eXpressLink: " & Args)

        C = CreateObject("eXpressLink.XCeXpressLink")
        If C.IsXCClientRunning Then
            MessageBox.Show("XCharge Client is already running." & vbCrLf & "Close the current window before processing Credit Card.")
            LogText("XCXL - eXpressLink: Already Running")
            Exit Function
        End If

        DeleteFileIfExists(ResultFile)

        R = C.ExecuteXCClient(Args, 1)
        If R = 0 Then
            MessageBox.Show("Failed to start XCharge", "WinCDS")
            LogText("XCXL - eXpressLink: Failed to start")
            Exit Function
        End If
        T = ReadEntireFileAndDelete(ResultFile)

        eXpressLink = GetXMLFields(T)
        LogText("XCXL - eXpressLink: " & Replace(eXpressLink.ContentString, vbCrLf, " / "))
    End Function

    Private Sub LogText(ByVal Text As String)
        mLog = mLog & IIf(Len(mLog) > 0, vbCrLf, "") & Text
        ActiveLog("clsXChargeXpressLink::" & Text)
        LogFile("XCXL.txt", Text)
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

    Private ReadOnly Property Clerk() As String
        Get
            Clerk = GetCashierName
            ' X-Charge eXpressLink doesn't like these.
            Clerk = Replace(Clerk, "[", "")
            Clerk = Replace(Clerk, "]", "")
            Clerk = Replace(Clerk, "(", "")
            Clerk = Replace(Clerk, ")", "")
            Clerk = Replace(Clerk, "{", "")
            Clerk = Replace(Clerk, "}", "")
        End Get
    End Property

    Private Function SafeArg(ByVal Field As String, Optional ByVal Val As String = "", Optional ByVal vLOCK As Boolean = False) As String
        If Left(Field, 1) = "/" Then Field = Mid(Field, 2)
        Field = Replace(UCase(Trim(Field)), " ", "")
        SafeArg = "/" & Field & IIf(Val <> "", ":" & Val, "")
        If InStr(SafeArg, " ") <> 0 Then SafeArg = """" & SafeArg & """"
        SafeArg = SafeArg & " "
        If vLOCK Then SafeArg = SafeArg & "/LOCK" & Field & " "
    End Function

    Private ReadOnly Property XCUsername() As String
        Get
            XCUsername = CSVField(StoreSettings.CCConfig, 4)
        End Get
    End Property

    Private ReadOnly Property XCPassword() As String
        Get
            XCPassword = CSVField(StoreSettings.CCConfig, 5)
        End Get
    End Property

    Public ReadOnly Property ResultFile() As String
        Get
            ResultFile = mResultFile
        End Get
    End Property

    Private Function GetXMLFields(ByVal tXML As String) As clsHashTable
        Dim T As String, A() As Object, L As Object
        GetXMLFields = New clsHashTable

        'A = Array("ACCOUNT", "ACCOUNTTYPE", "ADDITIONALFUNDSREQUIRED", "AMOUNT", "APPROVALCODE", "APPROVEDAMOUNT", "ARCHIVEADDOPTIONS1", "ARCHIVEADDOPTION2", "ARCHIVEADDOPTIONS3", "ARCHIVEPURGEDATE", "AVSRESULT", "BALANCE", "CASHBACKAMOUNT", "CHECKACCOUNTNO", "CHECKNO", "CHECKROUTINGNO", "CLERK", "CVRESULT", "DESCRIPTION", "EXPIRATION", "IIASTRANSACTION", "NAME", "RECEIPT", "RESULT", "SIGNATURE", "SWIPED", "TIPAMOUNT", "TYPE", "XCACCOUNTID", "XCTRANSACTIONID")
        A = New String() {"ACCOUNT", "ACCOUNTTYPE", "ADDITIONALFUNDSREQUIRED", "AMOUNT", "APPROVALCODE", "APPROVEDAMOUNT", "ARCHIVEADDOPTIONS1", "ARCHIVEADDOPTION2", "ARCHIVEADDOPTIONS3", "ARCHIVEPURGEDATE", "AVSRESULT", "BALANCE", "CASHBACKAMOUNT", "CHECKACCOUNTNO", "CHECKNO", "CHECKROUTINGNO", "CLERK", "CVRESULT", "DESCRIPTION", "EXPIRATION", "IIASTRANSACTION", "NAME", "RECEIPT", "RESULT", "SIGNATURE", "SWIPED", "TIPAMOUNT", "TYPE", "XCACCOUNTID", "XCTRANSACTIONID"}
        For Each L In A
            T = PullXMLTag(tXML, "<" & L & ">", "</" & L & ">")
            If T <> "" Then GetXMLFields.Add(L, T)
        Next
        ' The above are all the default fields supplied by X-Charge
        ' We also want a Boolean indicator that can be checked without
        ' other knowledge.
        If GetXMLFields.Exists("RESULT") Then
            If GetXMLFields.Item("RESULT") = "SUCCESS" Then
                GetXMLFields.Add("SUCCESS", True)
            End If
        End If
        If Not GetXMLFields.Exists("SUCCESS") Then GetXMLFields.Add("SUCCESS", False)

        On Error Resume Next
        Description = GetXMLFields.Item("DESCRIPTION")
        Success = GetXMLFields.Item("SUCCESS")
        If Not Success Then ErrorMsg = Description Else ErrorMsg = ""
        ApprovalCode = GetXMLFields.Item("APPROVALCODE")
        TransID = GetXMLFields.Item("XCTRANSACTIONID")
        BalanceAmountResult = GetPrice(GetXMLFields.Item("BALANCE"))
        AdditionalFundsResult = GetPrice(GetXMLFields.Item("ADDITIONALFUNDSREQUIRED"))
        ApprovedAmountResult = GetPrice(GetXMLFields.Item("APPROVEDAMOUNT"))

        CardHolderName = GetXMLFields.Item("NAME")
        XCC = GetXMLFields.Item("ACCOUNT")
        CCTypeName = GetXMLFields.Item("ACCOUNTTYPE")

        Debug.Print(GetXMLFields.ContentString)
    End Function

    Private ReadOnly Property DoExtraLog() As Boolean
        Get
            '  DoExtraLog = IsMattressKing
            Return False
        End Get
    End Property

    Public Function ExecReturn(Optional ByVal Prompt As Boolean = True) As Boolean
        LogStartFunction("XCXL-Return-" & DescribeState())
#If aAllowCredit Then
        Dim S As String, Res As clsHashTable
        S = ""
        S = S & ArgDefault()
        S = S & ArgTransType("RETURN")
        S = S & ArgAmount(Amount)

        Res = eXpressLink(S)
        If Res Is Nothing Then Exit Function
        If Not Res.Item("SUCCESS") Then Exit Function
        ExecReturn = True
#End If
    End Function

    Public Function ExecDebitReturn() As Boolean ' Optional ByVal Prompt As Boolean = True
        LogStartFunction("XCXL-DebitReturn-" & DescribeState())
#If aAllowDebit Then
        Dim S As String, Res As clsHashTable
        S = ""
        S = S & ArgDefault()
        S = S & ArgTransType("DEBITRETURN")
        S = S & ArgAmount(Amount)

        Res = eXpressLink(S)
        If Res Is Nothing Then Exit Function
        If Not Res.Item("SUCCESS") Then Exit Function
        ExecDebitReturn = True
#End If
    End Function

    Public Function ExecGiftReturn() As Boolean 'Optional ByVal Prompt As Boolean = True
        LogStartFunction("XCXL-GiftReturn-" & DescribeState())
#If aAllowGift Then
        Dim S As String, Res As clsHashTable
        S = ""
        S = S & ArgDefault()
        S = S & ArgTransType("GIFTRETURN")
        S = S & ArgAmount(Amount)

        Res = eXpressLink(S)
        If Res Is Nothing Then Exit Function
        If Not Res.Item("SUCCESS") Then Exit Function
        ExecGiftReturn = True
#End If
    End Function
End Class
