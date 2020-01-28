Module modPayPal
    Private Const PayPal_Version As String = "51.0"   ' 20110501
    Private Const PayPal_UseSandbox As Boolean = False

    Public Function PayPalCheckSales(Optional ByVal D As Date = NullDate) As Integer
        Dim A As Object, L As Object

        If Not IsServer() Then Exit Function                                            ' only the server makes requests
        If PayPal_APIUserName = "" Or IsFormLoaded("BilloSale") Then Exit Function    ' only if one of the stores is setup
        If D = NullDate Then D = Today
        A = modPayPal.PPTransactionSearch(D, D)
        On Error GoTo NoneItems
        For Each L In A
            If Not OnlineSaleIDExists("PP:" & L) Then modPayPal.PPGetTransactionDetails(L, True)
            PayPalCheckSales = PayPalCheckSales + 1
        Next
        Exit Function
NoneItems:
        PayPalCheckSales = -1
    End Function

    Public ReadOnly Property PayPal_APIUserName() As String
        Get
            Dim I As Integer
            For I = 1 To ActiveNoOfLocations
                PayPal_APIUserName = StoreSettings(I).PayPalUsername
            Next
        End Get
    End Property

    Public Function PPTransactionSearch(ByVal StartDate As String, ByVal EndDate As String) As Object
        Dim TransID As String, R As clsHashTable, N As Integer, S As String, A() As Object

        StartDate = PayPalDateFormat(StartDate)
        EndDate = PayPalDateFormat(EndDate, True)

        R = PPHttpPost("TransactionSearch", "&TRANSACTIONID=" & TransID & "&STARTDATE=" & StartDate & "&ENDDATE=" & EndDate & "&TRANSACTIONCLASS=All&STATUS=Success")
        PPTransactionSearch = PPSuccess(R)
        If Not PPTransactionSearch Then Exit Function

        Do While True
            S = R.Item("L_TRANSACTIONID" & N)
            If S = "" Then Exit Do
            ReDim Preserve A(0 To N)
            A(N) = S
            N = N + 1
        Loop
        PPTransactionSearch = A
    End Function

    Private Function PayPalDateFormat(Optional ByVal DD As String = "", Optional ByVal EndOfDay As Boolean = False) As String
        If Not IsDate(DD) Then DD = "" & Now
        '  PayPalDateFormat = Format(DD, "YYYY-MM-DD hh:mm:ss")
        PayPalDateFormat = Format(DD, "YYYY-MM-DD") & IIf(EndOfDay, " 23:59:59", " 00:00:00")
    End Function

    Public Function PPGetTransactionDetails(ByVal TransID As String, Optional ByVal CreateSale As Boolean = False) As Boolean
        Dim R As clsHashTable

        R = PPHttpPost("GetTransactionDetails", "&TRANSACTIONID=" & TransID)
        PPGetTransactionDetails = PPSuccess(R)

        ' BFH20110802
        ' transaction search isn't returning the same transid as the transaction lookup...
        ' this is different than the transid we looked up, for some reason... so we check both
        ' transids, but this one is the one that we should keep from getting duplicated.
        If OnlineSaleIDExists("PP:" & R.Item("TRANSACTIONID")) Then Exit Function

        If PPGetTransactionDetails And CreateSale Then
            CreatePayPalSale(R)
        End If
    End Function

    Private Function PPSuccess(ByVal R As clsHashTable, Optional ByVal AllowWarn As Boolean = True, Optional ByVal VerboseErr As Boolean = True) As Boolean
        Dim X As String, L As Object, EC As String

        X = UCase(R.Item("ACK"))
        PPSuccess = X = "SUCCESS" Or (AllowWarn And X = "SUCCESSWITHWARNING")
        If Not PPSuccess Then
            EC = R.Item("L_ERRORCODE0")
            If EC = "10004" Then GoTo NoMsg
            MessageBox.Show("PayPal lookup Failed:" & vbCrLf & R.Item("L_SEVERITYCODE0") & " (" & EC & ")" & vbCrLf & R.Item("L_SHORTMESSAGE0") & vbCrLf & R.Item("L_LONGMESSAGE0"))
NoMsg:
        End If

        If Not PPSuccess And VerboseErr Then
            X = ""
            'MsgBox R.ContentString
        End If
    End Function

    Private Function PPHttpPost(ByVal methodName As String, Optional ByVal nvpStr As String = "", Optional ByRef ErrStr As String = "") As clsHashTable
        Dim S As String, L As Object, R As Object

        PPHttpPost = New clsHashTable
        S = DownloadURLToString(PayPal_EndPoint & "?" & PayPal_NVP(methodName, nvpStr))

        If S = "" Then
            ErrStr = "Communication Failed."
            Exit Function
        End If

        '  PPHttpPost.Add "_", S

        For Each L In Split(S, "&")
            R = Split(L, "=")
            If UBound(R) >= 1 Then
                PPHttpPost.Add(R(0), UrlDecode(R(1)))
            End If
        Next
    End Function

    Private ReadOnly Property PayPal_NVP(ByVal Method As String, ByVal X As String)
        Get
            Dim S As String
            If X <> "" And Left(X, 1) <> "&" Then X = "&" & X
            S = ""
            S = S & "METHOD=" & Method
            S = S & "&VERSION=" & PayPal_Version
            S = S & "&PWD=" & PayPal_APIPassword
            S = S & "&USER=" & PayPal_APIUserName
            S = S & "&SIGNATURE=" & PayPal_APISignature
            S = S & X
            PayPal_NVP = S
        End Get
    End Property

    Private ReadOnly Property PayPal_APISignature() As String
        Get
            Dim I As Integer
            For I = 1 To ActiveNoOfLocations
                PayPal_APISignature = StoreSettings(I).PayPalSignature
            Next
        End Get
    End Property

    Private ReadOnly Property PayPal_APIPassword() As String
        Get
            Dim I As Integer
            For I = 1 To ActiveNoOfLocations
                PayPal_APIPassword = StoreSettings(I).PayPalPassword
            Next
        End Get
    End Property

    Private ReadOnly Property PayPal_EndPoint() As String
        Get
            'PayPal_EndPoint = IIf(PayPal_SandBox = "", "https://api-3t.paypal.com/nvp", "https://api-3t." & PayPal_SandBox & ".paypal.com/nvp")
            PayPal_EndPoint = "https://api-3t." & PayPal_SandBox(".") & "paypal.com/nvp"
        End Get
    End Property

    Private ReadOnly Property PayPal_SandBox(Optional ByVal Dot As String = "", Optional ByVal Force As Boolean = False) As String
        Get
            If PayPal_UseSandbox Or Force Then
                PayPal_SandBox = "sandbox" & Dot
            Else
                PayPal_SandBox = ""
            End If
        End Get
    End Property

    Private Function CreatePayPalSale(ByVal R As clsHashTable) As String
        Dim OrderItem As OnlineSaleOrderItem  ' Item Fields
        Dim Order As OnlineSaleOrder          ' Order Fields
        Dim N As Integer

        Order.XML = R.ContentString
        Order.ServiceProvider = "PayPal"

        Order.Time = Now

        Order.LoginID = R.Item("RECEIVERBUSINESS")
        Order.ID = "PP:" & R.Item("TRANSACTIONID")
        Order.BillTo.Last = R.Item("LASTNAME")
        Order.BillTo.First = R.Item("LASTNAME")
        Order.BillTo.Zip = R.Item("SHIPTOZIP")
        Order.BillTo.City = R.Item("SHIPTOCITY")
        Order.BillTo.Address1 = R.Item("SHIPTOSTREET")
        Order.BillTo.State = R.Item("SHIPTOSTATE")
        Order.BillTo.Email = R.Item("EMAIL")


        Order.Total = GetPrice(R.Item("AMT"))
        Order.Tax = GetPrice(R.Item("TAXAMT"))
        Order.SubTotal = Order.Total - Order.Tax

        Order.ServiceFee = GetPrice(R.Item("FEEAMT"))
        Order.Total = Order.Total '- .ServiceFee

        Do While True
            OrderItem.Quantity = Val(R.Item("L_QTY" & N))
            If OrderItem.Quantity = 0 Then Exit Do

            OrderItem.Style = R.Item("L_NUMBER" & N)
            OrderItem.Description = R.Item("L_NAME" & N)
            OrderItem.UnitPrice = GetPrice(R.Item("L_AMT" & N)) / OrderItem.Quantity
            OrderItem.Taxable = True
            AddOrderItemToOnlineSale(Order, OrderItem)
            N = N + 1
        Loop

        TotalizeOnlineSaleOrder(Order, PayPal_StoreNo)
        CreatePayPalSale = Order.SaleNo
        WI_SendSale(CreatePayPalSale)

        'TIMESTAMP=2011-05-02T00:59:40Z
        'VERSION = 51.0
        'BUILD = 1824201
        'ACK = Success
        'TRANSACTIONID=5NT57886WT257853E
        'CORRELATIONID = f3bb511122f8
        'TRANSACTIONTYPE = cart
        'RECEIVERBUSINESS=bhoogt_1303614066_biz@yahoo.com
        'RECEIVEREMAIL=bhoogt_1303614066_biz@yahoo.com
        'RECEIVERID=7KZZAMD9RJXWS
        'PAYMENTSTATUS = Completed
        'PENDINGREASON = NONE
        'REASONCODE = NONE
        'ORDERTIME=2011-04-28T12:32:06Z
        'PAYMENTTYPE = instant
        'PAYERSTATUS = verified
        'PAYERID = A5X9ZYV9VYW4Q
        'CURRENCYCODE = USD
        'SALESTAX = 0.00
        'FEEAMT = 1.11
        'TAXAMT = 0#
        'AMT = 27.95
        'L_NUMBER0 = B-001
        'L_NAME0 = Book
        'L_CURRENCYCODE0 = USD
        'L_QTY0 = 1
        'L_AMT0 = 14.95
        'L_NUMBER1 = P-001
        'L_NAME1 = Perfume
        'L_QTY1 = 1
        'L_AMT1 = 13.00
        'L_CURRENCYCODE1 = USD
        'COUNTRYCODE = US
        'ADDRESSOWNER = PayPal
        'ADDRESSSTATUS = Confirmed
        'LASTNAME = User
        'FIRSTNAME = Test
        'EMAIL=bhoogt_1303614514_per@yahoo.com
        'SHIPTOZIP = 95131
        'SHIPTONAME=Test User
        'SHIPTOCOUNTRYCODE = US
        'SHIPTOCITY=San Jose
        'SHIPTOSTREET=1 Main St
        'SHIPTOCOUNTRYNAME=United States
        'SHIPTOSTATE = CA
    End Function

    Public ReadOnly Property PayPal_StoreNo() As String
        Get
            Dim I As Integer
            PayPal_StoreNo = 1
            For I = 1 To ActiveNoOfLocations
                If StoreSettings(I).PayPalUsername <> "" Then PayPal_StoreNo = I : Exit Property
            Next
        End Get
    End Property
End Module
