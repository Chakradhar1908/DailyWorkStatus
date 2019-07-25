Module modProgramState
    Private mStoresSld As Integer
    Private mOrder As String
    Private mMail As String
    Private mArSelect As String
    Private mInven As String
    Private mTerminal As String
    Private mPO As String
    Private mReports As String
    '    Public Property Get StoresSld() as integer
    '    If mStoresSld <= 0 Then mStoresSld = 1
    '  StoresSld = mStoresSld
    'End Property

    '    Public Property Let StoresSld(ByVal nStore as integer)
    '  mStoresSld = FitRange(1, nStore, Setup_MaxStores)

    '  If ssMaxStore < 1 Then
    '    License = LICENSE_STORES_1
    ''    MsgBox "Store setup was not completed properly.  The program will not behave properly until this is fixed.", vbCritical, "Error!"
    '    Exit Property
    '    End If
    '    End Property
    Public Property StoresSld() As Integer
        Get
            If mStoresSld <= 0 Then mStoresSld = 1
            StoresSld = mStoresSld
        End Get
        Set(value As Integer)
            mStoresSld = FitRange(1, value, Setup_MaxStores)

            If ssMaxStore < 1 Then
                License = LICENSE_STORES_1
                '    MsgBox "Store setup was not completed properly.  The program will not behave properly until this is fixed.", vbCritical, "Error!"
                Exit Property
            End If
        End Set
    End Property
    Public Function OrderMode(ParamArray List()) As Boolean
        Dim A()
        A = List
        OrderMode = IsInArray(Order, A)
    End Function
    'Public Property Get Order() As String :   Order = mOrder: End Property
    'Public Property Let Order(ByVal vData As String) :  mOrder = vData: UpdatePermStatus: DescribeOrderMode: End Property
    Public Property Order() As String
        Get
            Order = mOrder
        End Get
        Set(value As String)
            mOrder = value
            UpdatePermStatus
            DescribeOrderMode
        End Set
    End Property
    Public Function DescribeOrderMode(Optional ByVal nOrderMode As String = "") As String
        If nOrderMode = "" Then nOrderMode = OrderModeIs
        Select Case nOrderMode
            Case "" : DescribeOrderMode = ""
            Case "A" : DescribeOrderMode = "New Sale"
            Case "AshleyASN" : DescribeOrderMode = "Ashley ASN"
            Case "AshleyEDI" : DescribeOrderMode = "Ashley EDI"
            Case "Audit" : DescribeOrderMode = "Daily Audit Report"
            Case "B" : DescribeOrderMode = "Deliver"
            Case "C" : DescribeOrderMode = "Void Sale"
            Case "CashRegister" : DescribeOrderMode = "Cash Register Interface"
            Case "Credit" : DescribeOrderMode = "Adjustments"
            Case "D" : DescribeOrderMode = "Payment On Account"
            Case "E" : DescribeOrderMode = "View Sale"
            Case "F" : DescribeOrderMode = "Item Preview"
            Case "R" : DescribeOrderMode = "Undelivered Sales Report"
            Case "L" : DescribeOrderMode = "Lay-A-Way"
            Case "C" : DescribeOrderMode = "Credit Sales"
            Case "H" : DescribeOrderMode = "Customer History Report"
            Case "ST" : DescribeOrderMode = "Sales Tax Reports"
            Case "ATR" : DescribeOrderMode = "Advertizing Report"
            Case "S" : DescribeOrderMode = "Service (Service)"
            Case "SDam" : DescribeOrderMode = "Damaged Stock (Service)"
            Case "SParts" : DescribeOrderMode = "Parts Order (Service)"
            Case "SPR" : DescribeOrderMode = "Open Parts order (Service)"
            Case "SBR" : DescribeOrderMode = "Parts order Billing (Service)"
            Case "SBU" : DescribeOrderMode = "Unpaid Billing (Service)"
            Case "STTOT" : DescribeOrderMode = "Sales Tax Report (Totals)"

            Case Else : DescribeOrderMode = "" : DevNotifyUnknownMode("Order", nOrderMode)
        End Select
    End Function
    Public ReadOnly Property OrderModeIs() As String
        Get
            OrderModeIs = Order
        End Get
    End Property
    Public Function DevNotifyUnknownMode(ByVal Ty As String, ByVal X As String)
        Dim S As String
        '  If True Then Exit Function                        ' Enable to disable
        DevNotifyUnknownMode = Nothing

        If Not IsDevelopment Then Exit Function
        If Not IsIDE Then Exit Function
        If Not IsCDSComputer("laptop") Then Exit Function

        S = "modProgramState.DevNotifyUnknownMode: " & Ty & ": " & X '& vbCrLf & "modProgramState.DevNotifyUnknownMode"
        Debug.Print(S)
        '  DevErr S
    End Function
    Public Function MailMode(ParamArray List()) As Boolean
        Dim A()
        A = List
        MailMode = IsInArray(Mail, A)
    End Function
    Public Function ArMode(ParamArray List()) As Boolean
        Dim A()
        A = List
        ArMode = IsInArray(ArSelect, A)
    End Function
    Public Property Mail() As String
        Get
            Mail = mMail
        End Get
        Set(value As String)
            mMail = value
            UpdatePermStatus()
            DescribeMailMode()
        End Set
    End Property
    Public Function DescribeMailMode(Optional ByVal nMailMode As String = "") As String
        If nMailMode = "" Then nMailMode = MailModeIs
        Select Case nMailMode
            Case "" : DescribeMailMode = ""
            Case "ADD/Edit" : DescribeMailMode = "Add/Edit Names"
            Case "Book" : DescribeMailMode = "Mail Book"
            Case "Fix" : DescribeMailMode = "Fix Mailing Addresses"
            Case Else : DescribeMailMode = "" : DevNotifyUnknownMode("Mail", nMailMode)
        End Select
    End Function
    Public ReadOnly Property MailModeIs() As String
        Get
            MailModeIs = Mail
        End Get
    End Property

    Public Property ArSelect() As String
        Get
            ArSelect = mArSelect
        End Get
        Set(value As String)
            mArSelect = value
            UpdatePermStatus()
            DescribeArMode()
        End Set
    End Property
    Public Function DescribeArMode(Optional ByVal nARMode As String = "") As String
        'If nARMode = "" Then nARMode = ArModeIs
        Select Case nARMode
            Case "" : DescribeArMode = ""
            Case "A" : DescribeArMode = "AR Ageing Report"
            Case "B" : DescribeArMode = "Print Monthly Statements"
            Case "BK" : DescribeArMode = "Bankruptcy (unused?)"
            Case "D" : DescribeArMode = "AR Delinquent Report"
            Case "E" : DescribeArMode = "AR Payment Estimator"
            Case "EA" : DescribeArMode = "Edit AR Account"
            Case "Edit" : DescribeArMode = "Edit Ar Account (uses ArCard)"
            Case "L" : DescribeArMode = "Late Charges and Notices"
            Case "LG" : DescribeArMode = "Legal Report (unused?)"
            Case "N" : DescribeArMode = "New Account Report (Installment)"
            Case "NP" : DescribeArMode = "Non Payment Report"
            Case "O" : DescribeArMode = "Closed Account Report"
            Case "P" : DescribeArMode = "Payment on Account"
            Case "R" : DescribeArMode = "Repo Report (unused?)"
            Case "REPRINT" : DescribeArMode = "Reprint Contract"
            Case "RestoreAR" : DescribeArMode = "Restore Voided Account"
            Case "Revolving" : DescribeArMode = "Revolving Accounts"
            Case "S" : DescribeArMode = "Old Account Setup"
            Case "T" : DescribeArMode = "AR Trial Balance Report"
            Case "W" : DescribeArMode = "Status Report"
            Case "WHOLATE" : DescribeArMode = "Whos Late Report"
            Case "V" : DescribeArMode = "Void AR Account No"
            Case "X" : DescribeArMode = "Metro426 (Credit Bureau) Export"
            Case Else : DescribeArMode = "" : DevNotifyUnknownMode("AR", nARMode)
        End Select
    End Function

    Public Property Inven() As String
        Get
            Inven = mInven
        End Get
        Set(value As String)
            mInven = value
            UpdatePermStatus()
            DescribeInvenMode
        End Set
    End Property

    Public Function OrderMode(ByVal ParamArray List() As Array) As Boolean
        Dim A()
        A = List
        OrderMode = IsInArray(Order, A)
    End Function
    Public Function DescribeInvenMode(Optional ByVal nInvenMode As String = "") As String
        If nInvenMode = "" Then nInvenMode = InvenModeIs
        Select Case nInvenMode
            Case "" : DescribeInvenMode = ""
            Case "A" : DescribeInvenMode = "New Items"
            Case "AK" : DescribeInvenMode = "POReport - Not Acknowledge"
            Case "AK-E" : DescribeInvenMode = "not acknoweldge report, email"
            Case "B" : DescribeInvenMode = "Price Changes"
            Case "BS" : DescribeInvenMode = "Best Sellers Report"
            Case "D" : DescribeInvenMode = "Factory Shipments"
            Case "E" : DescribeInvenMode = "View Stock"
            Case "FPO" : DescribeInvenMode = "Fax PO"
            Case "H" : DescribeInvenMode = "Change Contents"
            Case "EPO" : DescribeInvenMode = "Email PO"
            Case "IRep" : DescribeInvenMode = "Inventory Reports"
            Case "L" : DescribeInvenMode = "Load original inventory"
            Case "MRep" : DescribeInvenMode = "Margin Reports"
            Case "ML" : DescribeInvenMode = "Inventory Manuf Report"
            Case "OverdueOrders-E" : DescribeInvenMode = "overdue orders report"
            Case "P" : DescribeInvenMode = "POs"
            Case "PhysicalInventory" : DescribeInvenMode = "Take Physical Inventory"
            Case "R" : DescribeInvenMode = "Restore Delete Items"
            Case "T" : DescribeInvenMode = "Store Transfers, Schedule"
            Case "PRec" : DescribeInvenMode = "PO Receive Shipment"
            Case "TAGS" : DescribeInvenMode = "Print tags"
            Case "View Transfer" : DescribeInvenMode = "View Store Transfer"
            Case Else : DescribeInvenMode = "" : DevNotifyUnknownMode("Inven", nInvenMode)
        End Select
    End Function
    Public ReadOnly Property InvenModeIs() As String
        Get
            InvenModeIs = Inven
        End Get
    End Property
    Public Property Terminal() As String
        Get
            If mTerminal = "" Then mTerminal = GetCDSSetting("Terminal", GetLocalComputerName)
            Terminal = mTerminal
        End Get
        Set(value As String)
            mTerminal = value
            SaveCDSSetting("Terminal", value)
        End Set
    End Property
    Public Function ClearProgramState() As Boolean
        Inven = ""
        Order = ""
        ArSelect = ""
        Mail = ""
        PurchaseOrder = ""
        Reports = ""
    End Function
    Public Property PurchaseOrder() As String
        Get
            PurchaseOrder = mPO
        End Get
        Set(value As String)
            mPO = value
            UpdatePermStatus()
            DescribePOMode
        End Set
    End Property

    Public Function DescribePOMode(Optional ByVal nPOMode As String = "") As String
        If nPOMode = "" Then nPOMode = POModeIs
        Select Case nPOMode
            Case "" : DescribePOMode = ""
            Case "OverdueOrders-E" : DescribePOMode = "Overdue PO Orders Report"
            Case "ReOpen" : DescribePOMode = "PO Re-order"
            Case "OrderMinimum" : DescribePOMode = "PO Order Minimum"
            Case "poorderdemand" : DescribePOMode = "Overdue PO Orders Report"
            Case "OrderDemand" : DescribePOMode = "PO On Demand Order"
            Case "CombinePO" : DescribePOMode = "Combine POs"
            Case "REC" : DescribePOMode = "Receive POs"
            Case "Void" : DescribePOMode = "Void POs"
            Case "EDIT" : DescribePOMode = "Edit POs"
            Case Else : DescribePOMode = "" : DevNotifyUnknownMode("PO", nPOMode)
        End Select
    End Function
    Public ReadOnly Property POModeIs() As String
        Get
            POModeIs = PurchaseOrder
        End Get
    End Property
    Public Property Reports() As String
        Get
            Reports = mReports
        End Get
        Set(value As String)
            mReports = value
            UpdatePermStatus()
            DescribeReportsMode
        End Set
    End Property
    Public Function DescribeReportsMode(Optional ByVal nReportsMode As String = "") As String
        If nReportsMode = "" Then nReportsMode = ReportsModeIs
        Select Case nReportsMode
            Case "MT" : DescribeReportsMode = "Make KITs"
            Case "ET" : DescribeReportsMode = "Edit KITs"
            Case "RT" : DescribeReportsMode = "KIT List"
            Case "CS" : DescribeReportsMode = "Kit Lookup"

            Case "LoadOriginalInvByBarcodes" : DescribeReportsMode = "Load Orig Inv (Barcodes)"

            Case "O" : DescribeReportsMode = "[Order Entry Report]"
            Case "I" : DescribeReportsMode = "[Inventory Report]"
            Case "H" : DescribeReportsMode = "Customer History Report"

            Case "VSS" : DescribeReportsMode = "Special Special Report"

            Case "Ashley" : DescribeReportsMode = "Ashley Open PO Report"

            Case "DESIGNTAG" : DescribeReportsMode = "Custom Tag Designer"
            Case "Store Catalog" : DescribeReportsMode = "Store Catalog"
            Case "Mini-Barcode Scanner" : DescribeReportsMode = "Mini-Barcode Scanner"

            Case "" : DescribeReportsMode = ""
            Case Else : DescribeReportsMode = "" : DevNotifyUnknownMode("Reports", nReportsMode)
        End Select
    End Function
    Public ReadOnly Property ReportsModeIs() As String
        Get
            ReportsModeIs = Reports
        End Get
    End Property
    Public Function ReportsMode(ParamArray List() As Object) As Boolean
        Dim A() As Object
        A = List
        ReportsMode = IsInArray(Reports, A)
    End Function

    Public Function InvenMode(ParamArray List() As Object) As Boolean
        Dim A() As Object
        A = List
        InvenMode = IsInArray(Inven, A)
    End Function

End Module
