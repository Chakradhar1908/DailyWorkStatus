Module modHoldings
    Public Const SlSt_Open As String = "O"
    Public Const SlSt_Dlvd As String = "D"
    Public Const SlSt_LyWy As String = "L"
    Public Const SlSt_Ly30 As String = "1"
    Public Const SlSt_Ly60 As String = "2"
    Public Const SlSt_Ly90 As String = "3"
    Public Const SlSt_Ly12 As String = "4"
    Public Const SlSt_OpCr As String = "E"
    Public Const SlSt_ClCr As String = "C"
    Public Const SlSt_OpFi As String = "S"
    Public Const SlSt_ClFi As String = "F"
    Public Const SlSt_BkOr As String = "B"
    Public Const SlSt_Void As String = "V"

    Public g_Holding As New cHolding
    Public Const SALE_STATUS_OPEN As String = "OPEN"
    Public Const SALE_STATUS_DELIVERED As String = "DELIVERED"
    Public Const SALE_STATUS_LAYAWAY As String = "Lay-A-Way"
    Public Const SALE_STATUS_LAW30 As String = "30 Day Lay-A-Way"
    Public Const SALE_STATUS_LAW60 As String = "60 Day Lay-A-Way"
    Public Const SALE_STATUS_LAW90 As String = "90 Day Lay-A-Way"
    Public Const SALE_STATUS_LAW120 As String = "120 Day Lay-A-Way"
    Public Const SALE_STATUS_OPEN_CREDIT As String = "Open Credit"
    Public Const SALE_STATUS_CREDIT As String = "CREDIT"
    Public Const SALE_STATUS_BORD As String = "BACK ORD."
    Public Const SALE_STATUS_VOID As String = "Void"

    Public Const HoldNew_FILE As String = "Holding" & ".exe"
    Public Const HoldNew_TABLE As String = "Holding"
    Public Const HoldNew_INDEX As String = "LeaseNo"
    Public Const SALE_STATUS_FINANCE As String = "Store Finance"
    Public Const SALE_STATUS_OPENFINANCE As String = "Open Store Finance"

    Public Function DescribeHoldingStatus(ByVal hStat As String) As String
        '#**Returns the string description of the one-letter Holding Status field.
        '#Parameters:
        '#  hStat - The Holding Status (a one-character string)
        '#Returns:
        '#  String - The User-friendly description of the one-character Holding status field
        '#Description:
        '#  Provides a status character lookup table.  Used for both lookup and user-display.
        '#  Current Codes are:
        '#    * O - Open
        '#    * D - Delivered
        '#    * L - Lay-A-Way
        '#    * 1 - Lay-A-Way 30
        '#    * 2 - Lay-A-Way 60
        '#    * 3 - Lay-A-Way 90
        '#    * 4 - Lay-A-Way 120
        '#    * E - Open Finance
        '#    * C - Credit
        '#    * F - Store Finance (Delivered)
        '#    * S - Open Store Finance
        '#    * B - Back Ordered
        '#    * V - Void
        '#
        '#  Description values are also available in the constants SALE_STATUS_XXX.
        '#See Also:
        '#  HoldingStatusRepresents
        '#DEV NOTE:
        '#  Persistence is also required here.

        Select Case UCase(hStat)
            Case SlSt_Open : DescribeHoldingStatus = SALE_STATUS_OPEN
            Case SlSt_Dlvd : DescribeHoldingStatus = SALE_STATUS_DELIVERED
            Case SlSt_LyWy : DescribeHoldingStatus = SALE_STATUS_LAYAWAY
            Case SlSt_Ly30 : DescribeHoldingStatus = SALE_STATUS_LAW30
            Case SlSt_Ly60 : DescribeHoldingStatus = SALE_STATUS_LAW60
            Case SlSt_Ly90 : DescribeHoldingStatus = SALE_STATUS_LAW90
            Case SlSt_Ly12 : DescribeHoldingStatus = SALE_STATUS_LAW120
            Case SlSt_OpCr : DescribeHoldingStatus = SALE_STATUS_OPEN_CREDIT
            Case SlSt_ClCr : DescribeHoldingStatus = SALE_STATUS_CREDIT
            Case SlSt_ClFi : DescribeHoldingStatus = SALE_STATUS_FINANCE
            Case SlSt_OpFi : DescribeHoldingStatus = SALE_STATUS_OPENFINANCE ' BFH20060803
            Case SlSt_BkOr : DescribeHoldingStatus = SALE_STATUS_BORD
            Case SlSt_Void : DescribeHoldingStatus = SALE_STATUS_VOID
            Case Else : DescribeHoldingStatus = hStat
        End Select
    End Function

    Public Function LeaseNoExists(ByVal LeaseNo As String) As Boolean
        Dim objHolding As cHolding
        objHolding = New cHolding
        LeaseNoExists = objHolding.Load(LeaseNo)
        DisposeDA(objHolding)
    End Function

    Public Function GetLeaseNumber(Optional ByVal ForceAutomatic As Boolean = False, Optional ByVal Specified As String = "") As String
        ' This Function gets a Sale Number.
        ' It still uses a file, but now checks for duplicates and
        ' uses a return value instead of a global variable.

        If Not ForceAutomatic And StoreSettings.bManualBillofSaleNo Then  'Manual BillofSale
            If LeaseNoExists(Specified) Then  ' First, check for duplicates.
                '###  Add something here to the calling function know it's failed?
                MessageBox.Show("This lease number already exists.  Please try again.", "Duplicate Sale Number")
            Else
                GetLeaseNumber = Specified
            End If
            Exit Function
        End If

        ' ### Replace this with a database access..
        ' ### Except Jerry really likes the file-based approach because it lets him edit it at will.
        GetLeaseNumber = GetFileAutonumber(BOSFile, Val(Trim(StoresSld) + "0000"))
        If LeaseNoExists(GetLeaseNumber) Then
            ' BillSale.Dat is out of synch with the database,
            ' or there was a specified autonumber we happened to hit?
            ' We need to generate a different sale number.
            ' We don't care if it takes an extra minute, no more save-overs are allowed..
            GetLeaseNumber = GetNextLeaseNumber(GetLeaseNumber)
        End If
    End Function

    Public Function HoldingStatusRepresents(ByVal HS As String) As String
        Dim K As Object
        For Each K In HoldingStatusList()
            If UCase(DescribeHoldingStatus(K)) = UCase(HS) Then HoldingStatusRepresents = K : Exit Function
        Next
    End Function

    Public Function HoldingStatusList() As Object
        'HoldingStatusList = Array(SlSt_Open, SlSt_Dlvd, SlSt_LyWy, SlSt_Ly30, SlSt_Ly60, SlSt_Ly90, SlSt_Ly12, SlSt_OpCr, SlSt_ClCr, SlSt_OpFi, SlSt_ClFi, SlSt_BkOr, SlSt_Void)
        HoldingStatusList = New String() {SlSt_Open, SlSt_Dlvd, SlSt_LyWy, SlSt_Ly30, SlSt_Ly60, SlSt_Ly90, SlSt_Ly12, SlSt_OpCr, SlSt_ClCr, SlSt_OpFi, SlSt_ClFi, SlSt_BkOr, SlSt_Void}
    End Function

    Private Function GetNextLeaseNumber(ByVal LeaseNo As String) As String
        ' This function is called when we have a LeaseNo conflict.
        ' The file-based autonumber has run into a preexisting sale.
        ' First step is to lock the sale number generator file.
        ' Next, determine the correct next sale number.
        ' Then write it to the file, and unlock it.

        Dim FNum As Integer, fName As String
        On Error GoTo BadFile
        FNum = FreeFile()
        fName = BOSFile()

        'Open fName For Output As #FNum
        FileOpen(FNum, fName, OpenMode.Output)
        Dim objHolding As cHolding
        objHolding = New cHolding
        objHolding.DataAccess.Records_OpenSQL("SELECT * FROM Holding WHERE Val(LeaseNo)>Val('" & ProtectSQL(LeaseNo) & "') ORDER BY Val(LeaseNo)")
        GetNextLeaseNumber = LeaseNo + 1
        Do While objHolding.DataAccess.Records_Available
            If IsNumeric(LeaseNo) Then
                LeaseNo = CStr(Val(LeaseNo) + 1)
                If objHolding.LeaseNo = LeaseNo Then
                    ' Keep searching..
                    GetNextLeaseNumber = Val(LeaseNo) + 1 ' just in case we run out..
                Else
                    GetNextLeaseNumber = LeaseNo
                    ' Write this value to the file..
                    ' And stop searching.
                    Exit Do
                End If
            Else
                ' This had better not happen with file-based autonumbers.
                Err.Raise(-11332, , "Invalid sale number: " & LeaseNo)
            End If
        Loop
        DisposeDA(objHolding)

        If GetNextLeaseNumber = "" Then
            Err.Raise(-11333, , "Unable to generate sale number.")
        End If

        ' Write the new autonumber to the file.
        Print(FNum, GetNextLeaseNumber)
        'Close #FNum
        FileClose(FNum)

        Exit Function
BadFile:
        Select Case Err.Number
            Case 52, 53, 62 ' File not found, bad file name or number, input past end of file
                ' These errors mean the file didn't exist, but can be created.
                Resume Next
            Case 70 'Permission denied.
                'If another computer is reading or writing the file, you may get this error.
                Application.DoEvents()
                Resume
            Case 75 'Path/File access error
                Application.DoEvents()
                Resume
            Case 76 ' Path not found
                If MessageBox.Show("Can't access " & fName & ", try again?", "File Error", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                    Resume
                Else
                    End
                End If
            Case Else  ' An error we didn't foresee.
                Debug.Assert(False)
                'Debug.Print(Err, error)
                '      Resume
        End Select
    End Function

    Public Function HoldNew_GetBalance(ByVal LeaseNo As String, Optional ByVal StoreNo As Integer = 0) As String
        Dim objHolding As cHolding
        objHolding = New cHolding
        If StoreNo < 0 Or StoreNo <> StoresSld Then
            objHolding.DataAccess.DataBase = GetDatabaseAtLocation(StoreNo)
        End If
        If objHolding.Load(Trim(LeaseNo), "LeaseNo") Then
            HoldNew_GetBalance = Format(objHolding.Sale - objHolding.Deposit, "###,##0.00")
        Else  ' Can't find sale!
            MessageBox.Show("Could not find holding record for Sale No: " & LeaseNo & vbCrLf & "Balance will be $0.00", "Sale Number Not Located in Holding", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            HoldNew_GetBalance = "0"
        End If
        DisposeDA(objHolding)
    End Function

    Public Function GetLeaseNoStatus(ByVal LeaseNo As String, Optional ByVal Desc As Boolean = True) As String
        Dim objHolding As cHolding
        objHolding = New cHolding
        If Not objHolding.Load(LeaseNo) Then
            GetLeaseNoStatus = ""
        Else
            GetLeaseNoStatus = IIf(Desc, DescribeHoldingStatus(objHolding.Status), objHolding.Status)
        End If
        DisposeDA(objHolding)
    End Function

    Public Function GetLeaseNoMailIndex(ByVal LeaseNo As String) As Integer
        Dim C As CGrossMargin
        C = New CGrossMargin
        If C.Load(LeaseNo, "SaleNo") Then GetLeaseNoMailIndex = C.Index
        DisposeDA(C)
    End Function

End Module
