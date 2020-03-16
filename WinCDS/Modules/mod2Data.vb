Module mod2Data
    '::::mod2Data
    ':::SUMMARY
    ': Contains different functions , required to perform different functions in WinCDS Software.
    ':::DESCRIPTION
    ': This module contains multiple functions which helps to get information related to Vendors, Manufacturers, SalesTax2, PO's, Discount, DataBase etc.
    ': This module contains functions mostly used to perform different types of operations like Inserting, Updating, Deleting, Accessing data in 2Data table.
    ':::SEE ALSO
    ': - mod2DataTransfers, mod2DataPictures, mod2DataVendors
    ': - modDBUtilities

    Public Const KIT_LOC As Integer = 1              ' Which Store DB is the [InvKit] table located in.
    Public Const KIT_PFX As String = "KIT-"
    Public Const KIT_PFX_LEN As Integer = 4          ' Len(KIT_PFX)
    'Public Const KIT_STYLE_MAXLEN As integer = Setup_2Data_StyleMaxLen - KIT_PFX_LEN
    Public Const KIT_OFS As Integer = 5              ' KIT OFFSET:  For Mid(xxx, KIT_OFS)
    Public Const cdspicType_PartsOrder As Integer = 2

    Private m2Data As Collection
    Public TicketCodeFor As Integer, TicketCode As String

    Public Function TableExists(ByVal Location As Integer, ByVal TableName As String) As Boolean
        '::::TableExists
        ':::SUMMARY
        ': Checks whether required Table exists or not.
        ':::DESCRIPTION
        ': This function is used to check whether required table exist or not from Database.
        ':::PARAMETERS
        ': - TableName - Indicates  the name of table.
        ':::RETURN
        ': Boolean - Returns whether it is true or not.
        TableExists = DatabaseTableExists(IIf(Location <= 0, GetDatabaseInventory, GetDatabaseAtLocation(Location)), TableName)
    End Function
    Public Function DatabaseTableExists(ByVal DatabaseName As String, ByVal TableName As String) As Boolean
        '::::DatabaseTableExists
        ':::SUMMARY
        ': Used to check whether the Database Table is exists or not.
        ':::DESCRIPTION
        ': This function is used to check whether the specified Database Table from specified Database through sql statement.
        ':::PARAMETERS
        ': - DatabaseName - Indicates the Datebase name.
        ': - TableName - Indicates the Table name.
        ':::RETURN
        ': Boolean - Returns whether it is true or false.
        Dim RS As ADODB.Recordset
        On Error Resume Next
        If Left(TableName, 1) = "[" Then TableName = Mid(TableName, 2)
        If Right(TableName, 1) = "]" Then TableName = Left(TableName, Len(TableName) - 1)
        If Dir(DatabaseName) = "" Then Exit Function
        RS = GetRecordsetBySQL("SELECT * FROM [" & TableName & "] WHERE FALSE=TRUE", , DatabaseName, True)
        DatabaseTableExists = (Not RS Is Nothing)
        RS.Close()
        RS = Nothing
    End Function
    Public Function QuerySalesTax2(ByVal ind As Integer) As String
        QuerySalesTax2 = GetTax2String(ind + 1, True)
    End Function
    Public Sub LoadAdvTypesIntoComboBox(ByRef Cbo As ComboBox, Optional ByVal StoreNum As Integer = 0, Optional ByVal NoExtras As Boolean = False)
        '::::LoadAdvTypesIntoComboBox
        ':::SUMMARY
        ': Loads Advertising types into Combo Box.
        ':::DESCRIPTION
        ': This function is used to Load Advertising Types into ComboBox
        ':::PARAMETERS
        ': - cbo - Indicates the combo Box.
        ': - StoreNum - Indicates the Store Number.
        ': - NoExtras - Indicates whether it is true or false
        Dim cAd As clsAdvertisingType
        cAd = New clsAdvertisingType
        Dim FF As Integer
        If StoreNum = 0 Then StoreNum = StoresSld

        cAd.DataAccess.Records_Open("AdvertisingTypeID", "Error opening Advertising Types table.")
        'Cbo.Clear
        Cbo.Items.Clear()
        'FromAdvertisingType = True
        Do While cAd.DataAccess.Records_Available
            cAd.cDataAccess_GetRecordSet(cAd.DataAccess.RS)
            AddItemToComboBox(Cbo, cAd.AdType, cAd.ID)
            'AddItemToComboBox(Cbo, AdType, ID)
            'Cbo.itemData(Cbo.NewIndex) = cAd.ID

            '--> Note: replaced above two lines with the below one. created custom class ItemDataclass to implement
            '--> itemData property of vb6 in vb.net
            'Cbo.Items.Add(New ItemDataClass(AddItemToComboBox(Cbo, cAd.AdType, cAd.ID), cAd.ID))
        Loop
        'FromAdvertisingType = False
        DisposeDA(cAd)

        If Not NoExtras Then
            Dim idc As ItemDataClass
            FF = 1 ' bfh20051212
            'Do While Cbo.ListCount < 10
            Do While Cbo.Items.Count < 10
                'Cbo.AddItem "New Item " & FF
                'Cbo.itemData(Cbo.NewIndex) = -1

                '--> Note: replaced above two lines with the below one. created custom class ItemDataclass to implement
                '--> itemData property of vb6 in vb.net
                idc = New ItemDataClass("New Item" & FF, -1)
                'Cbo.Items.Add(New ItemDataClass("New Item " & FF, -1))
                Cbo.Items.Add(idc.Itemname)
                FF = FF + 1
            Loop
        End If
    End Sub
    Public Sub LoadSalesTax2IntoComboBox(ByRef Cbo As ComboBox, Optional ByVal StoreNum As Integer = 0, Optional ByVal Separated As Boolean = False, Optional ByVal OnlyName As Boolean = False)
        '::::LoadSalesTax2IntoComboBox
        ':::SUMMARY
        ': Used to Load SalesTax2 into ComboBox.
        ':::DESCRIPTION
        ': This function is used to load SalesTax2 into combobox based on store number.
        ':::PARAMETERS
        ': - cbo - Indicates the combo Box.
        ': - StoreNum - Indicates the Store Number.
        ': - Separated - Indicates whether it is true or false.
        ': - OnlyName - Indicates whether it is true or false.
        ':::RETURN
        Dim Mfl() As String, I As Integer
        Dim idc As ItemDataClass

        If StoreNum <= 0 Then StoreNum = StoresSld
        Mfl = GetSalesTax2(StoreNum)
        Cbo.Items.Clear()

        On Error GoTo SalesTaxError
        If UBound(Mfl) < 0 Then Exit Sub

        For I = LBound(Mfl) To UBound(Mfl)
            idc = New ItemDataClass(Trim(Mfl(I)), I)
            If Separated Then
                'Cbo.AddItem Trim(Mfl(I))
                'Cbo.itemData(Cbo.NewIndex) = I

                '--> Note: replaced above two lines with the below one. created custom class ItemDataclass to implement
                '--> itemData property of vb6 in vb.net
                'Cbo.Items.Add(New ItemDataClass(Trim(Mfl(I)), I))
                Cbo.Items.Add(idc.Itemname)

            Else
                ' Don't worry about alignment or spacing while there are only 10 departments.
                'Cbo.AddItem I & "   " & Trim(Mfl(I))
                'Cbo.itemData(Cbo.NewIndex) = I

                '--> Note: replaced above two lines with the below one. created custom class ItemDataclass to implement
                '--> itemData property of vb6 in vb.net
                'Cbo.Items.Add(New ItemDataClass(I & "   " & Trim(Mfl(I)), I))
                Cbo.Items.Add(idc.Itemname)
            End If
        Next
SalesTaxError:
    End Sub
    Public Function GetRNByStyle(ByVal Style As String) As Integer
        '::::GetRNByStyle
        ':::SUMMARY
        ': Used to get RN by using Style.
        ':::DESCRIPTION
        ': This function is used to get RN with Style.
        ':::PARAMETERS
        ': - Style -Indicates the Style.
        ':::RETURN
        ': integer - Returns the RN as a integer.
        Dim X As CInvRec
        X = New CInvRec
        If X.Load(Style, "Style") Then GetRNByStyle = X.RN
        DisposeDA(X)
    End Function
    Public Function QuickShowPOForStyle(ByVal Style As String, Optional ByVal OpenClosed As TriState = vbTrue) As Boolean
        '::::QuickShowPOForStyle
        ':::SUMMARY
        ': Used to display quick show of PO's.
        ':::DESCRIPTION
        ': This function is used to get Quick Show of Po's for required Style.
        ':::PARAMETERS
        ': Style - Indicates the Style.
        ':::RETURN
        ': Boolean - Returns the result whether it is true or false.
        Dim R As ADODB.Recordset, S As String, A() As Object, L As String, N As Integer, ST As String
        S = ""
        S = S & "SELECT * FROM [PO] WHERE 1=1"
        S = S & " AND Style='" & Style & "'"
        S = S & " AND [PrintPo] <> 'v'"
        If OpenClosed <> vbUseDefault Then
            S = S & " AND Posted" & IIf(Not OpenClosed, "=", "<>") & "'X'"
        End If
        R = GetRecordsetBySQL(S, , GetDatabaseInventory)

        ReDim A(R.RecordCount - 1)
        Do While Not R.EOF
            ST = "Open"
            If R("posted").Value = "X" Then ST = "Closed"
            If R("printpo").Value = "v" Then ST = "Void"
            L = AlignString(R("pono").Value, 7) & " " & AlignString(R("podate").Value, 10) & " " & AlignString(ST, 8) & " " & AlignString(IfNullThenNilString(R("DueDate")), 10)
            A(N) = L
            N = N + 1
            R.MoveNext()
        Loop

        S = SelectOptionArray("View PO", frmSelectOption.ESelOpts.SelOpt_ToItem Or frmSelectOption.ESelOpts.SelOpt_List, A)
        If S <> "" Then
            S = Trim(Left(S, 7))
            'EditPO.QuickViewPO(S)
        End If
    End Function
    Public Function GetSalesTax2(Optional ByVal StoreNum As Integer = 0) As String()
        ':::: GetSalesTax2
        ':::SUMMARY
        ': Used to get SalesTax2 file.
        ':::DESCRIPTION
        ': This function is used to get SalesTax2 file based on Store Number and SalesTax2 is different for different countries, like for HU it is 0.60 percent etc.
        ': Must be in each store directory for sales tax rates
        ':::PARAMETERS
        ': StoreNum - Indicates the Store Number.
        ':::RETURN
        ': String - Returns SalesTax2 as a string.

        Dim D() As String, Line As String, I As Integer, FNum As Integer
        On Error GoTo SalesTaxError

        If StoreNum <= 0 Then StoreNum = StoresSld
        FNum = FreeFile()
        'Open(SalesTax2File For Input As #FNum)
        FileOpen(FNum, SalesTax2File, OpenMode.Input)
        I = 0
        Do Until EOF(FNum)
            'Input( #FNum, Line)
            Input(FNum, Line)
            Line = Trim(Line)
            If Line <> "" Then
                ReDim Preserve D(I)
                D(I) = Line
                I = I + 1
            End If
        Loop
        'Close(#FNum)
        FileClose(FNum)
        GetSalesTax2 = D
SalesTaxError:
    End Function
    Public Function GetTax2String(ByVal tL As Integer, Optional ByVal QuietError As Boolean = False) As String
        '::::GetTax2String
        ':::SUMMARY
        ': Used to get Tax2String.
        ':::DESCRIPTION
        ': This function is used to get Tax2 string.
        ':::PARAMETERS
        ': - tL - Indicates the location.
        ': - QuietError - Indicates whether it is true or false.
        ':::RETURN
        ': String - Returns the Tax2 String.
        Dim Taxes As Object
        On Error GoTo GetTax2StringFailure
        Taxes = GetSalesTax2()
        tL = tL - 1
        If tL < LBound(Taxes) Or tL > UBound(Taxes) Then
            If Not QuietError Then MsgBox("Invalid tax rate: " & tL & ".", vbCritical) : Exit Function ' Error!
        End If
        GetTax2String = Taxes(tL)
GetTax2StringFailure:
    End Function
    Public Function SalesTax2File(Optional ByVal StoreNum As Integer = 0) As String
        '::::SalesTax2File
        ':::SUMMARY
        ': Used to display SalesTax2 File.
        ':::DESCRIPTION
        ': This function is used to get SalesTax2 file based on Store Number.
        ':::PARAMETERS
        ': - StoreNum - Indicates the Store number.
        ':::RETURN
        If StoreNum = 0 Then StoreNum = StoresSld
        SalesTax2File = StoreFolder(StoreNum) & "SALESTAX2.DAT"
    End Function
    Public Function GetVendorNoFromName(ByVal VendorName As String) As String
        '::::GetVendorNoFromName
        ':::SUMMARY
        ': Gets Vendor numbe based on Vendor Name.
        ':::DESCRIPTION
        ': This function is used to get Vendor Number from 2Data table based on Vendor Name through Sql statement.
        ':::PARAMETERS
        ': - VendorName - Indicates the Vendor name.
        ':::RETURN
        ': String - Returns the vendor number as a string.
        Dim RS As ADODB.Recordset
        RS = GetRecordsetBySQL("SELECT DISTINCT VendorNo FROM [2Data] WHERE Vendor='" & ProtectSQL(UCase(Trim(VendorName)), False) & "'", , GetDatabaseInventory())
        If Not RS.EOF Then
            GetVendorNoFromName = RS("VendorNo").Value
        End If
        RS.Close()
        RS = Nothing
    End Function
    Public Function GetDeptNoFromStyle(ByVal Style As String) As String
        '::::GetDeptNoFromStyle
        ':::SUMMARY
        ': Used to get Department Number based on Style.
        ':::DESCRIPTION
        ': This function is used to display Department Number based on Style.
        ':::PARAMETERS
        ': - Style - Indicates the Style.
        ':::RETURN
        ': String - Returns the Department Number as a String.
        Dim T As Integer
        'T = GetRNFromStyleNo(Style)
        'If T <> 0 Then GetDeptNoFromStyle = GetDeptNoFromRn(T)
    End Function
    Public Function GetDeptFromStyleNo(ByVal Style As String) As String
        '::::GetDeptFromStyleNo
        ':::SUMMARY
        ': Gets Department based on Style Number.
        ':::DESCRIPTION
        ': This function is used to get Departments based on Style Number from 2Data table through sql statement.
        ':::PARAMETERS
        ': - Style - Indicates the Style number.
        ':::RETURN
        ': String - Returns Department as a String.
        Dim RS As ADODB.Recordset
        On Error Resume Next
        RS = GetRecordsetBySQL("SELECT [Dept] FROM [2Data] WHERE Style='" & Style & "'", , GetDatabaseInventory)
        If Not RS Is Nothing Then
            If Not RS.EOF Then 'Return an empty if we can --- Robert
                GetDeptFromStyleNo = RS("Dept").Value
            End If
        End If
        RS = Nothing
    End Function
    Public Function GetVendorByStyle(ByVal Style As String, Optional ByRef VendorNo As String = "", Optional ByRef DeptNo As String = "") As String
        '::::GetVendorByStyle
        ':::SUMMARY
        ': Used to get Vendor number by using Style.
        ':::DESCRIPTION
        ': This function is used to get Vendor number,Department number using Style.
        ':::PARAMETERS
        ': - Style - Indicates the Style.
        ': - VendorNo - Indicates the Vendor number.
        ': - DeptNo - Indicates the Department number.
        ':::RETURN
        ': String - Returns the Vendor number as a string.
        Dim X As CInvRec
        X = New CInvRec
        If X.Load(Style, "Style") Then
            GetVendorByStyle = X.Vendor
            VendorNo = X.VendorNo
            DeptNo = X.DeptNo
        Else
            VendorNo = ""
            DeptNo = ""
            GetVendorByStyle = ""
        End If
        DisposeDA(X)
    End Function

    Public Function QuerySalesTax2List() As String()
        QuerySalesTax2List = GetSalesTax2()
    End Function

    Public Sub LoadMfgNamesIntoListBox(ByRef Cbo As ListBox, Optional ByVal Vendor As String = "", Optional ByVal Separated As Boolean = False, Optional ByVal OnlyName As Boolean = False)
        '::::LoadMfgNamesIntoListBox
        ':::SUMMARY
        ': Load Manufacturer Names into ListBox.
        ':::DESCRIPTION
        ': This function is used to represent Manufacturer Names in ListBox.
        ':::PARAMETERS
        ': - cbo - Indicates the List Box.
        ': - Vendor - Indicates the Vendor name.
        ': - Separated - Indicates whether it is true or false.
        ': - OnlyName - Indicates whether it is true or false.
        ':::RETURN
        Dim Mfl As Object, I As Integer
        Mfl = GetManufacturerList(Vendor)
        Cbo.Items.Clear()
        If UBound(Mfl, 1) <= 0 Then Exit Sub
        For I = LBound(Mfl, 2) To UBound(Mfl, 2)
            If IsNothing(Mfl(0, I)) Then Mfl(0, I) = 0
            If Separated Then
                'Cbo.AddItem Trim("" & Mfl(1, I))
                'Cbo.itemData(Cbo.NewIndex) = Val(Mfl(0, I))
                Cbo.Items.Add(New ItemDataClass(Trim("" & Mfl(1, I)), Val(Mfl(0, I))))
            Else
                'Cbo.AddItem Mfl(0, I) & "   " & Trim(Mfl(1, I))
                'Cbo.itemData(Cbo.NewIndex) = Val(Mfl(0, I))
                Cbo.Items.Add(New ItemDataClass(Mfl(0, I) & "   " & Trim(Mfl(1, I)), Val(Mfl(0, I))))
            End If
        Next
    End Sub

    Public Function GetManufacturerList(Optional ByVal Vendor As String = "") As Object
        '::::GetManufacturerList
        ':::SUMMARY
        ': Gets Manufacturer list.
        ':::DESCRIPTION
        ': This function is used to get Manufacturer list through sql statements.
        ':::PARAMETERS
        ': - Vendor - Indicates the available vendors in database.
        ':::RETURN
        Dim RS As ADODB.Recordset
        Dim SQL As String

        SQL = "SELECT [2Data].VendorNo, First([2Data].Vendor) as FV FROM [2Data] INNER JOIN Search " &
    "ON [2Data].Rn=Search.Rn "
        If Trim(Vendor) <> "" Then SQL = SQL & "WHERE left([2Data].Vendor, " & Len(Vendor) & ")=""" & ProtectSQL(Vendor) & """ "
        SQL = SQL & "GROUP BY [2Data].VendorNo ORDER BY First([2Data].Vendor), [2Data].VendorNo"
        RS = GetRecordsetBySQL(SQL, , GetDatabaseInventory)
        If RS.EOF Then
            RS = Nothing
            Dim GMLtmp(0, 0) '@NO-LINT-NTYP
            GetManufacturerList = GMLtmp
            Exit Function
        End If
        GetManufacturerList = RS.GetRows
        RS.Close()
        RS = Nothing
    End Function

    Public Function SelectKitStatusAndQuantity(ByRef Status As String, ByRef nQuantity As Double, ByVal KitStyle As String, Optional ByVal LeaveOpen As Boolean = True) As Boolean
        '::::SelectKitStatusAndQuantity
        ':::SUMMARY
        ': Used to Select Status and Quantity of Kit.
        ':::DESCRIPTION
        ': This function is used to select Kit Status and Quantity based on parameters listed above.
        ':::PARAMETERS
        ': - Status - Indicates the Status.
        ': - nQuantity - Indicates the Quantity number.
        ': - KitStyle - Indicates the Kit Style.
        ': - LeaveOpen - Indicates whether it is true or false.
        ':::RETURN
        ': Boolean - Returns whether the result is true or false.
        'Load frmKitLevels
        If Not IsIn(Status, "ST", "SO", "LAW", "SS", "DELTW") Then Status = "ST"
        frmKitLevels.AllowStatusChange = True
        frmKitLevels.AllowItemStatusChange = LeaveOpen
        frmKitLevels.LoadKit(StoresSld, Status, KitStyle)
        On Error Resume Next

        'frmKitLevels.Show vbModal
        frmKitLevels.ShowDialog()

        If Not frmKitLevels.Cancelled Then
            Status = frmKitLevels.Status
            nQuantity = frmKitLevels.Quantity

            SelectKitStatusAndQuantity = True
        End If

        If Not LeaveOpen Then
            'Unload frmKitLevels
            frmKitLevels.Close()
        End If
    End Function

    ' these are for searching, in case you forget the name of this function
    ' :loadmanufnamesintocombobox, loadvendornamesintocombobox, loadvendorsintocombobox
    Public Sub LoadMfgNamesIntoComboBox(ByRef Cbo As ComboBox, Optional ByVal Vendor As String = "", Optional ByVal Separated As Boolean = False, Optional ByVal OnlyName As Boolean = False)
        '::::LoadMfgNamesIntoComboBox
        ':::SUMMARY
        ': Loads Manufacturer Names into Combobox.
        ':::DESCRIPTION
        ': This function is used to load Manufacturer Names in to Combo Box.
        ':::PARAMETERS
        ': - cbo - Indicates the combo Box.
        ': - Vendor - Indicates the Vendor name.
        ': - Separated - Indicates whether it is true or false.
        ': - OnlyName - Indicates whether it is true or false.
        ':::RETURN
        Dim Mfl As Object, I As Integer
        Mfl = GetManufacturerList(Vendor)
        If UBound(Mfl, 1) <= 0 Then Exit Sub
        Cbo.Items.Clear()

        For I = LBound(Mfl, 2) To UBound(Mfl, 2)
            If IsNothing(Mfl(0, I)) Then Mfl(0, I) = 0
            If Separated Then
                'Cbo.AddItem Trim("" & IfNullThenNilString(Mfl(1, I)))
                'Cbo.itemData(Cbo.NewIndex) = Val(IfNullThenNilString(Mfl(0, I)))
                Cbo.Items.Add(New ItemDataClass(Trim("" & IfNullThenNilString(Mfl(1, I))), Val(IfNullThenNilString(Mfl(0, I)))))
            Else
                If OnlyName Then
                    'Cbo.AddItem Trim(IfNullThenNilString(Mfl(1, I)))
                    Cbo.Items.Add(New ItemDataClass(Trim(IfNullThenNilString(Mfl(1, I))), Val(IfNullThenNilString(Mfl(0, I)))))
                Else
                    'Cbo.AddItem Mfl(0, I) & "   " & Trim(IfNullThenNilString(Mfl(1, I)))
                    Cbo.Items.Add(New ItemDataClass(Mfl(0, I) & "   " & Trim(IfNullThenNilString(Mfl(1, I))), Val(IfNullThenNilString(Mfl(0, I)))))
                End If
                'Cbo.itemData(Cbo.NewIndex) = Val(IfNullThenNilString(Mfl(0, I)))
            End If
        Next
    End Sub

    Public Function DepartmentFile(Optional ByVal StoreNum As Integer = 0) As String
        '::::DepartmentFile
        ':::SUMMARY
        ': This function is used to display Departments file.
        ':::DESCRIPTION
        ': This function is used to get Departments from DEPT.DAT file based on Store Number.
        ': We can update Departments in DEPT.DAT file according to any store requirement.
        ':::PARAMETERS
        ': - StoreNum - Indicates the Store number.
        ':::RETURN
        ': String - Returns Department File as a string.
        If StoreNum = 0 Then StoreNum = StoresSld
        If StoreNum > Setup_MaxStores Or StoreNum < 1 Then StoreNum = 1
        DepartmentFile = StoreFolder(StoreNum) & "DEPT.DAT"
    End Function

    Public Sub LoadDiscountTypesIntoComboBox(ByRef Cbo As ComboBox, Optional ByVal StoreNum As Integer = 0, Optional ByVal NoneSelected As String = "-", Optional ByVal DisableWhenEmpty As Boolean = True)
        '::::LoadDiscountTypesIntoComboBox
        ':::SUMMARY
        ': Loads types of Discounts into ComboBox.
        ':::DESCRIPTION
        ': This function is used to display Types of Discounts into Combo Box.
        ':::PARAMETERS
        ': - cbo - Indiactes the Combo Box.
        ': - StoreNum - Indicates the Store Number.
        ': - NoneSelected -  Indicates that nothing is selected.
        ': - DisableWhenEmpty - Indicates whether it is true or false.
        ':::RETURN
        Dim N As Integer, I As Integer
        N = DiscountTypeCount()

        Cbo.Items.Clear()

        If NoneSelected <> "-" Then
            'Cbo.AddItem NoneSelected, 0
            'Cbo.itemData(Cbo.NewIndex) = 0
            Cbo.Items.Insert(0, New ItemDataClass(NoneSelected, 0))
        End If
        If N = 0 Then
            Cbo.Enabled = IIf(DisableWhenEmpty, False, True)
        Else
            Cbo.Enabled = True
            For I = 1 To N
                'Cbo.AddItem DiscountType(I)
                'Cbo.itemData(Cbo.NewIndex) = I
                Cbo.Items.Add(New ItemDataClass(DiscountType(I), I))
            Next
            Cbo.SelectedIndex = 0
        End If
    End Sub

    Public Function DiscountType(ByVal Index As Integer, Optional ByRef vName As String = "", Optional ByRef vType As String = "", Optional ByRef vPercent As String = "", Optional ByRef vExtra As String = "") As String
        '::::DiscountType
        ':::SUMMARY
        ': Used to display the type of Discount.
        ':::DESCRIPTION
        ': This function is used to display every term related to DiscountType.
        ':::PARAMETERS
        ':::RETURN
        ': String - Returns the result as a String.
        Dim T As Object, S As String, X() As String, N As Integer
        On Error Resume Next
        T = GetDiscounts()
        S = T(Index - 1)
        If S <> "" Then
            X = Split(S, ":")
            N = UBound(X) - LBound(X) + 1
            If N >= 1 Then vName = X(0)
            If N >= 2 Then vType = X(1)
            If N >= 3 Then vPercent = X(2)
            If N >= 4 Then vExtra = X(3)
            DiscountType = X(0)
            If DiscountType = "" Then DiscountType = S
        Else
            DiscountType = ""
        End If
    End Function

    Public Function DiscountTypeCount() As Integer
        '::::DiscountTypeCount
        ':::SUMMARY
        ': Used to count of types of Discount.
        ':::DESCRIPTION
        ': This function is used to count types of Discount.
        ':::PARAMETERS
        ':::RETURN
        ': Long - Returns the result as a Long.
        On Error Resume Next
        Dim T As Object
        T = GetDiscounts()
        DiscountTypeCount = UBound(T) - LBound(T) + 1
    End Function

    Public Function GetDiscounts() As Object
        '::::GetDiscounts
        ':::SUMMARY
        ': Used to get Discounts.
        ':::DESCRIPTION
        ': This function is used to display Discounts.
        ':::PARAMETERS
        ':::RETURN
        Dim X As String
        X = ReadFile(BOSDiscountFile)
        X = Replace(X, vbLf, "")
        GetDiscounts = Split(X, vbCr)
    End Function

    Public Function GetCubesOnSale(ByVal SaleNo As String, Optional ByVal OnDate As String = "", Optional ByVal Loc As Integer = 0) As Double
        '::::GetCubesOnSale
        ':::SUMMARY
        ':::DESCRIPTION
        ':::PARAMETERS
        ':::RETURN
        Dim C As CGrossMargin, D As CInvRec
        C = New CGrossMargin

        If SaleNo = "" Then Exit Function
        If Loc = 0 Then Loc = StoresSld
        C.DataAccess.DataBase = GetDatabaseAtLocation(Loc)
        C.DataAccess.Records_OpenSQL("SELECT * FROM GrossMargin WHERE SaleNo='" & SaleNo & "' AND NOT Style IN ('STAIN', 'DEL', 'LAB', 'TAX1', 'TAX2', 'NOTES', 'SUB', 'PAYMENT', 'VOID', 'Style', '--- Adj ---', '') AND NOT Status IN ('VOID','VD')")

        Do
            D = New CInvRec
            If Not IsItem(C.Style) Or IsVoid(C.Status) Or IsReturned(C.Status) Then GoTo Skip
            If Not D.Load(C.Style, "Style") Then GoTo Skip
            If Not IsDate(C.DDelDat) Then GoTo Skip
            If IsDate(OnDate) Then
                If DateDiff("d", DateValue(OnDate), C.DDelDat) <> 0 Then GoTo Skip
            End If

            GetCubesOnSale = GetCubesOnSale + C.Quantity * D.Cubes
Skip:
            DisposeDA(D)
        Loop While C.DataAccess.Records_Available

        DisposeDA(C)
    End Function

    Public Function GetCubesByStyle(ByVal Style As String, Optional ByVal Qty As Double = 1) As Double
        '::::GetCubesByStyle
        ':::SUMMARY
        ':::DESCRIPTION
        ':::PARAMETERS
        ':::RETURN
        Dim C As CInvRec
        C = New CInvRec
        If C.Load(Style, "Style") Then GetCubesByStyle = C.Cubes * Qty
        DisposeDA(C)
    End Function

    Public Function GetDescByStyle(ByVal Style As String) As String
        '::::GetDescByStyle
        ':::SUMMARY
        ': Used to gets description of each item.
        ':::DESCRIPTION
        ': This function is used to get description
        ':::PARAMETERS
        ': - Style - Indicates the Style.
        ':::RETURN
        ': String - Returns description as a string.
        Dim X As CInvRec
        X = New CInvRec
        If X.Load(Style, "Style") Then GetDescByStyle = X.Desc
        DisposeDA(X)
    End Function

    Public Function GetTax2Rate(ByVal tL As Integer) As String
        '::::GetTax2Rate
        ':::SUMMARY
        ': Gets Tax2 Rate.
        ':::DESCRIPTION
        ': This function is used to calculate Tax2 Rate using formula given below.
        ':::PARAMETERS
        ': - tL - Indicates the location.
        ':::RETURN
        ': String - Returns Tax2 Rate as a String.
        Dim T As String, N As Integer
        On Error Resume Next
        T = GetTax2String(tL)
        N = InStr(T, " ")
        If N = 0 Then GetTax2Rate = T Else GetTax2Rate = Mid(T, 1, N - 1)
    End Function

    Public Function SalesTax2Count() As Integer
        '::::SalesTax2Count
        ':::SUMMARY
        ': Used to count SalesTax2.
        ':::DESCRIPTION
        ': This function is used to count Sales Tax2 even when we get errors.
        ':::PARAMETERS
        ':::RETURN
        On Error Resume Next
        Dim X() As String
        SalesTax2Count = 0
        X = GetSalesTax2()
        SalesTax2Count = UBound(X) - LBound(X) + 1
    End Function

    Public Function QuerySalesTax2Rate(ByVal ind As Integer) As Double
        QuerySalesTax2Rate = GetTax2Rate(ind + 1)
    End Function
End Module
