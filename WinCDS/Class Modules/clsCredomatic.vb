Public Class clsCredomatic
#Const aAllowCredit = True
#Const aDoesPinPad = False
#Const aAllowDebit = False
#Const aAllowGift = False

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
    Public ApprovalCode As String
    Public RefId As String
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

    Public Function ExecVoid(Optional ByVal SaleDate As String = "") As Boolean
        If IsDate(SaleDate) Then
            If DateAfter(Today, SaleDate, False) Then
                ExecVoid = ExecCredit()
                Exit Function
            End If
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


        TransID = Res.Item("TransID")
        ApprovalCode = Res.Item("AuthCode")

        PostedDate = Res.Item("PostedDate")
        SettledDate = Res.Item("SettledDate")

        Amount = GetPrice(Res.Item("Amount"))
        ApprovalCode = Res.Item("AuthCode")
        Status = Res.Item("Status")
        AVSCode = Res.Item("AVSCode")
#End If
    End Function

    Public ReadOnly Property XCC() As String
        Get
            XCC = CCXOut(CC)
        End Get
    End Property

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
        If Not ExecVoid() Then
            LogText("CreditCardCredit unsuccessful")
            Exit Function
        End If


        TransID = Res.Item("TransID")
        RefId = Res.Item("RefID")

        PostedDate = Res.Item("PostedDate")
        SettledDate = Res.Item("SettledDate")

        Amount = GetPrice(Res.Item("Amount"))
        Status = Res.Item("Status")
        Message = Res.Item("Message")
#End If
    End Function

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

    Private Function DoTransaction(ByVal Action As String, ByVal QS As String, ByRef Res As clsHashTable) As Boolean
        Dim R As String, I As Integer
        LogStartFunction("Transaction Central: " & Action)

        ProgressForm(0, 1, "Processing.  Please wait...")
        R = INETGET(SOAPURL(Action) & "?" & QS)

        If R = "" And CommunicationRetries > 0 Then
            For I = 1 To CommunicationRetries
                R = INETGET(SOAPURL(Action) & "?" & QS)
                If R <> "" Then Exit For
            Next
        End If
        ProgressForm()

        DoTransaction = HandleResponse(Action, QS, R, Res)
    End Function

    Private Sub LogText(ByVal Text As String)
        Debug.Print(Text)
        LogFile("Credomatic.txt", Text & vbCrLf)
        mLog = mLog & IIf(Len(mLog) > 0, vbCrLf, "") & Text
    End Sub

    Private ReadOnly Property SOAPURL(ByVal Action As String) As String
        Get
            Select Case LCase(Action)
                Case Else
                    SOAPURL = "https://secure.redfinnet.com/SmartPayments/transact.asmx?ProcessCreditCard"
            End Select
        End Get
    End Property

    Private Function HandleResponse(ByVal Action As String, ByVal QS As String, ByVal R As String, ByRef Res As clsHashTable) As Boolean
        Dim TransID As String, Status As String, Msg As String
        If R = "" Then
            MessageBox.Show("Communication Failed.")
            Exit Function
        End If

        If Not ParseResponse(R, Res) Then
            MessageBox.Show("Declined." & IIf(IsDevelopment, vbCrLf & "Parse Failed." & vbCrLf & QS, ""))
            LogText("Parse Failed, Msg: " & R)
            Exit Function
        End If

        If InStr(LCase(Action), "report") Then HandleResponse = True : Exit Function

        TransID = Res.Item("TransID")
        Status = Res.Item("Status")
        Msg = Res.Item("Message")

        If Val(TransID) = 0 Then
            MessageBox.Show("Declined." & IIf(IsDevelopment, vbCrLf & "TransID: 0" & vbCrLf & QS, ""))
            LogText("Purchase failed: " & Msg)
            LogText("QS=" & QS)
            Exit Function
        End If

        LogText("TransID: " & TransID)

        If Status = "Declined" Then
            'MsgBox "Declined." & IIf(IsDevelopment, vbCrLf & "TransID: " & TransID, ""), vbDefaultButton4
            MessageBox.Show("Declined." & IIf(IsDevelopment, vbCrLf & "TransID: " & TransID, ""))
            LogText("Transaction Declined: " & Msg)
            Exit Function
        End If

        HandleResponse = True
    End Function

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

            'B = InStr(Body, "</", A)
            B = InStr(A, Body, "</",)

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
End Class
