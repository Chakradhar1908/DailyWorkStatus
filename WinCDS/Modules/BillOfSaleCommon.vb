Imports Microsoft.VisualBasic.Compatibility.VB6
Module BillOfSaleCommon
    Public Enum BillColumns
        eStyle = 0
        eManufacturer
        eLoc
        eStatus
        eQuant
        eDescription
        ePrice
        eManufacturerNo
        eTransID
    End Enum

    Public Enum BillProcessValidation
        eOK = 0
        eNoTax = 1
        eNoItems = 2
        eSSNoVendor = 3
    End Enum
    Public Function PrintSale(ByVal SaleNo As String, Optional ByVal Store As Integer = 0, Optional ByVal CopyID As String = "", Optional ByVal Copies As Integer = 1) As Boolean
        '::::PrintSale
        ':::SUMMARY
        ': Print a given sale at a given location.
        ':::DESCRIPTION
        ':By comparing all conditions, this function is used to print the sale.
        ':::PARAMETERS
        ':-SaleNo-Represents the Sale number.
        ':-Store-Represents the Store
        ':-CopyID-Represents Unique id of copy.
        ':-Copies-Represents the number of copies.
        ':::RETURN
        ':Boolean-Denotes whether it is true or false.

        Dim S As sSale, X As Integer
        S = New sSale
        If Copies < 0 Then Exit Function
        If Store <= 0 Then Store = StoresSld
        X = StoresSld
        If StoresSld <> Store Then StoresSld = Store
        If Not LeaseNoExists(SaleNo) Then
            If StoresSld <> X Then StoresSld = X
            Exit Function
        End If
        If Val(CopyID) > 0 Then CopyID = StoreSettings.SalesCopyID(FitRange(0, Val(CopyID) - 1, 3))
        S.PrintInvoice(CopyID, Copies, SaleNo)
        If StoresSld <> X Then StoresSld = X
        DisposeDA(S)
        PrintSale = True
    End Function
    Public Function CalculateGM(ByVal Sale As Decimal, ByVal Landed As Decimal, Optional ByVal DefaultValue As Double = 100, Optional ByVal RoundTo As Integer = -1) As Double
        '::::CalculateGM
        ':::SUMMARY
        ': Calculate Gross Margin
        ':::DESCRIPTION
        ': Calculates Gross Margin based on Sale and Landed cost.
        ':::PARAMETERS
        ':-Sale-Denotes the current sale.
        ':-Landed-Denotes the landed cost of item.
        ':-DefaultValue-Denotes the default value and it is equal to 100.
        ':-RoundTo-Used to roundup the value to nearest round value.
        ':::RETURN
        ':Double-Returns the Gross margin in a double.

        If Sale = 0 Then CalculateGM = DefaultValue : Exit Function
        CalculateGM = (1 - Landed / Sale) * 100
        If RoundTo >= 0 Then CalculateGM = Math.Round(CalculateGM, RoundTo)
    End Function
    Public Function AddNewMarginRecord(ByVal SaleNo As String, ByVal Style As String, ByVal Desc As String,
  Optional ByVal Quantity As Double = 0, Optional ByVal SellPrice As Decimal = 0,
  Optional ByVal Vendor As String = "", Optional ByVal DeptNo As String = "", Optional ByVal VendorNo As String = "", Optional ByVal Cost As Decimal = 0,
  Optional ByVal ItemFreight As Decimal = 0, Optional ByVal RN As Integer = 0,
  Optional ByVal PorD As String = "", Optional ByVal Commission As String = "", Optional ByVal Status As String = "", Optional ByVal Salesman As String = "",
  Optional ByVal Location As String = "", Optional ByVal SellDte As String = "", Optional ByVal DDelDat As String = "", Optional ByVal Store As String = "",
  Optional ByVal Name As String = "", Optional ByVal ShipDte As String = "", Optional ByVal Phone As String = "", Optional ByVal Index As String = "",
  Optional ByVal GM As String = "", Optional ByVal Detail As Integer = 0, Optional ByVal SS As String = "", Optional ByVal DelPrint As String = "", Optional ByVal PullPrint As String = "",
  Optional ByVal CommPd As Date = Nothing, Optional ByVal TransID As String = "", Optional ByVal SalesSplit As String = "100.0 0.0 0.0"
  ) As Boolean
        '::::AddNewMarginRecord
        ':::SUMMARY
        ':It is used to add new margin record.
        ':::DESCRIPTION
        ':Margin is a record from parent sale.It is used to add new margin record.
        ':::RETURN
        ':Boolean-Represents whether it is true or false.
        ':::SEE ALSO
        ':SaveNewMarginRecord

        Dim X As CGrossMargin
        X = SaveNewMarginRecord(SaleNo, Style, UCase(Desc), Quantity, SellPrice, Vendor, DeptNo, VendorNo, Cost, ItemFreight, RN, PorD, Commission, Status, Salesman, Location, SellDte, DDelDat, Store, Name, ShipDte, Phone, Index, GM, Detail, SS, DelPrint, PullPrint, CommPd, TransID, SalesSplit)
        DisposeDA(X)
        AddNewMarginRecord = True

    End Function

    Public Function SaveNewMarginRecord(ByVal SaleNo As String, ByVal Style As String, ByVal Desc As String,
  Optional ByVal Quantity As Double = 0, Optional ByVal SellPrice As Decimal = 0,
  Optional ByVal Vendor As String = "", Optional ByVal DeptNo As String = "", Optional ByVal VendorNo As String = "", Optional ByVal Cost As Decimal = 0,
  Optional ByVal ItemFreight As Decimal = 0, Optional ByVal RN As Integer = 0,
  Optional ByVal PorD As String = "", Optional ByVal Commission As String = "", Optional ByVal Status As String = "", Optional ByVal Salesman As String = "",
  Optional ByVal Location As String = "", Optional ByVal SellDte As String = "", Optional ByVal DDelDat As String = "", Optional ByVal Store As String = "",
  Optional ByVal Name As String = "", Optional ByVal ShipDte As String = "", Optional ByVal Phone As String = "", Optional ByVal Index As String = "",
  Optional ByVal GM As String = "", Optional ByVal Detail As Integer = 0, Optional ByVal SS As String = "", Optional ByVal DelPrint As String = "", Optional ByVal PullPrint As String = "",
  Optional ByVal CommPd As Date = Nothing, Optional ByVal TransID As String = "", Optional ByVal SalesSplit As String = "100.0 0.0 0.0"
  ) As CGrossMargin
        Dim C As CGrossMargin
        C = New CGrossMargin

        If Salesman = "" Or SalesSplit = "" Or Index = "" Or Name = "" Or Phone = "" Then
            If SaleNo <> "" Then
                C.DataAccess.DataBase = GetDatabaseAtLocation(StoresSld)
                C.Load(SaleNo, "SaleNo")
                If C.DataAccess.Record_Count > 0 Then
                    C.DataAccess.Records_MoveAbsolute(1)
                    C.cDataAccess_GetRecordSet(C.DataAccess.RS)
                End If
            End If
        End If

        SaveNewMarginRecord = New CGrossMargin
        SaveNewMarginRecord.SaleNo = SaleNo
        SaveNewMarginRecord.Style = Style
        SaveNewMarginRecord.Desc = UCase(Desc)
        SaveNewMarginRecord.Quantity = Val(Quantity)
        SaveNewMarginRecord.SellPrice = GetPrice(SellPrice)
        SaveNewMarginRecord.Vendor = Vendor
        SaveNewMarginRecord.DeptNo = DeptNo
        SaveNewMarginRecord.VendorNo = VendorNo
        SaveNewMarginRecord.Cost = GetPrice(Cost)
        SaveNewMarginRecord.ItemFreight = GetPrice(ItemFreight)
        SaveNewMarginRecord.RN = Val(RN)
        SaveNewMarginRecord.PorD = PorD
        SaveNewMarginRecord.Commission = Commission
        SaveNewMarginRecord.Status = Status
        SaveNewMarginRecord.Salesman = IIf(Salesman <> "", Salesman, C.Salesman)
        SaveNewMarginRecord.Location = IIf(Val(Location) = 0, StoresSld, Val(Location))
        SaveNewMarginRecord.SellDte = SellDte
        ' BFH20120327 - we don't verify the setting of this one, simply to make sure it's "safe"... could be an item we want them to set up the ddeldat for
        SaveNewMarginRecord.DDelDat = DDelDat ' IIf(DDelDat <> "", DDelDat, C.DDelDat)
        SaveNewMarginRecord.Store = Val(Store)
        SaveNewMarginRecord.Name = IIf(Name <> "", Name, C.Name)
        SaveNewMarginRecord.ShipDte = ShipDte
        SaveNewMarginRecord.Phone = IIf(Phone <> "", Phone, C.Phone)
        SaveNewMarginRecord.Index = IIf(Index <> "", Index, C.Index)
        SaveNewMarginRecord.GM = GM
        SaveNewMarginRecord.Detail = Detail
        SaveNewMarginRecord.SS = SS
        SaveNewMarginRecord.DelPrint = DelPrint
        SaveNewMarginRecord.PullPrint = PullPrint
        SaveNewMarginRecord.CommPd = CommPd
        SaveNewMarginRecord.SalesSplit = IIf(SalesSplit <> "", SalesSplit, C.SalesSplit)
        SaveNewMarginRecord.TransID = TransID
        SaveNewMarginRecord.Save()

        DisposeDA(C)
    End Function

    Public Function CreateDetailRecord(ByRef InvData As CInvRec, ByVal SaleNo As String,
      ByVal Name As String, ByVal Qty As Double, ByVal Status As String, ByVal Loc As Integer,
      ByVal DelDate As String, ByVal PoNo As Integer, Optional ByVal MarginLine As Integer = 0) As CInventoryDetail
        '::::CreateDetailRecord
        ':::SUMMARY
        ':This function is used  to create a record in detail table of Inventory.
        ':::DESCRIPTION
        ':This function is used to create and update records in Detail table and also used to handle errors.
        ':::RETURN
        ':Boolean-Denotes whether it is true or false.
        ':::SEE ALSO
        ':
        Dim InvDetail As CInventoryDetail

        On Error GoTo GeneralErr
        InvDetail = New CInventoryDetail
        InvDetail.Style = IfNullThenNilString(InvData.Style)
        InvDetail.Lease1 = Trim(SaleNo)
        InvDetail.Name = Trim(Name)
        InvDetail.InvRn = Val(InvData.RN)
        InvDetail.Store = StoresSld ' Loc 'added 04-05-2002 ' Changed 20040226
        InvDetail.MarginRn = MarginLine

        If Trim(Status) = "PO" Then
            InvDetail.Trans = "PO"
        Else
            InvDetail.Trans = "NS"
        End If
        UpdateQuarterlySales(InvData, Qty, Today)     ' put into proper month

        If IsDelivered(Status) Then
            InvDetail.DDate1 = DelDate                ' Don't populate Date on new sales.
            InvDetail.SO1 = 0 '""
            InvDetail.Trans = "DS"
            InvData.OnHand = InvData.OnHand - Qty
        End If

        If Status = "SO" Or Status = "PO" Then      'Or Style = "NOTES" Then  ' Impossible
            If Status = "SO" Then
                InvDetail.Name = Name & " " & PoNo
            End If
            If Status = "PO" Then
                InvDetail.Name = Name
                InvData.PoSold = InvData.PoSold + Qty
            End If

            InvDetail.AmtS1 = 0 '""
            InvDetail.Ns1 = 0 ' ""
            InvDetail.SO1 = Qty
            InvDetail.LAW = 0 '""
            UpdateDetailLocation(InvDetail, Loc, Qty)
        End If

        If Status = "LAW" Then
            InvDetail.AmtS1 = 0 '""
            InvDetail.Ns1 = 0 '""
            InvDetail.SO1 = 0 '""
            InvDetail.LAW = Qty
            UpdateDetailLocation(InvDetail, Loc, Qty)
        End If

        If Status = "ST" Or IsDelivered(Status) Then
            InvDetail.AmtS1 = Qty
            InvDetail.Ns1 = 0
            InvDetail.SO1 = 0
            InvDetail.LAW = 0
            UpdateDetailLocation(InvDetail, Loc, Qty)

            InvData.AddLocationQuantity(Loc, -Qty)
            InvData.Available = IfNullThenZeroDouble(InvData.Available) - Qty
        End If
        InvData.Save()
        '  Set InvData = Nothing

        ' Details Record
        On Error GoTo HandleErr
        '  InvDetail.DataAccess.Records_Add
        InvDetail.Save()
        CreateDetailRecord = InvDetail
        InvDetail = Nothing ' do not dispose, is return value in previous line

        Exit Function

GeneralErr:
        MessageBox.Show("Error in CreateDetailRecord: " & Err.Description)
        Err.Clear()
        Resume Next

HandleErr:
        'MsgBox("ERROR in Detail BOS2.PrintRec: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        MessageBox.Show("ERROR in Detail BOS2.PrintRec: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        Err.Clear()
        Resume Next
    End Function

    Public Sub UpdateQuarterlySales(ByRef InvData As CInvRec, ByVal Qty As Double, ByVal SaleDate As Date)
        '::::UpdateQuarterlySales
        ':::SUMMARY
        ':This function is used to put sales in proper quarter.
        ':::DESCRIPTION
        ':This function is mostly used to update sales quarterly.
        ':::PARAMETERS
        ':-InvData--Represents the stock of items in Inventory.
        ':-Qty-Represents the Quantity of items.
        ':-SaleDate-Represents the SaleDate of Item.
        ':::RETURN
        'Select Case Format(SaleDate, "Q")
        'MessageBox.Show(DatePart(DateInterval.Quarter, SaleDate))

        Select Case DatePart(DateInterval.Quarter, SaleDate)
            Case 1 : InvData.Sales1 = InvData.Sales1 + Qty
            Case 2 : InvData.Sales2 = InvData.Sales2 + Qty
            Case 3 : InvData.Sales3 = InvData.Sales3 + Qty
            Case 4 : InvData.Sales4 = InvData.Sales4 + Qty
        End Select
    End Sub
    Private Sub UpdateDetailLocation(ByRef InvDetail As CInventoryDetail, ByVal Location As Integer, ByVal Qty As Double)
        InvDetail.SetLocationQuantity(Location, Qty)
    End Sub

    Public Function QuerySaleLocation() As String
        '::::QuerySaleLocation
        ':::SUMMARY
        ': Determine where an item will be sold from
        ':::DESCRIPTION
        ': This function is used to get active store location,if the sale from login location.
        ':::PARAMETERS
        ':::RETURN
        ':STRING-Returns the Active store location as a string.
        Dim tmpLoc As Integer
        ' Sell from Logged In Location.
        If StoreSettings.bSellFromLoginLocation Then tmpLoc = StoresSld
        If tmpLoc < 1 Or tmpLoc > Setup_MaxStores Then QuerySaleLocation = "1" Else QuerySaleLocation = CStr(tmpLoc)
        'setup_maxstore=active store locations
    End Function

    Public Function CalculateGMROI(ByVal Style As String, Optional ByVal DStart As String = "", Optional ByVal DEnd As String = "", Optional ByVal UseEndingInventory As Boolean = False, Optional ByRef AverageInventory As Double = 0, Optional ByRef GMDollars As Decimal = 0, Optional ByRef InvDollars As Decimal = 0, Optional ByRef Sales As Decimal = 0, Optional ByRef COGS As Decimal = 0, Optional ByRef EndingInventory As Double = 0) As Decimal
        '::::CalculateGMROI
        ':::SUMMARY
        ':This function is used to calculate Gross Margin Return on Inventory.
        ':::DESCRIPTION
        ':This function is used to calculate the Gross Margin Return on Inventory (GMROI) for the given items over the specified date range and used to access through sql._
        ':GMROI is one of Inventory Reports.
        ':::PARAMETERS
        ':-Style-Represents the Style number.
        ':-DStart-Represents the starting date.
        ':-DEnd-Represents the Ending date.
        ':-UseEndingInventory
        ':-AverageInventory
        ':-GMDollars
        ':-InvDollars
        ':-Sales
        ':-COGS
        ':-EndingInventory
        ':::RETURN
        ':Currency-Returns GMROI in Currency.
        Dim M As CGrossMargin
        Dim StartingInventory As Double
        Dim pAS As Double, pNI As Double
        Dim AmtSold As Double, LndSold As Decimal, SlsSold As Decimal, NewItems As Double
        Dim LastD As String, Per As Integer
        Dim RS As ADODB.Recordset, SQL As String, X As String, R As Double
        Dim T As Double
        Dim DD As Integer, First As Boolean

        If Not IsDate(DEnd) Then DEnd = Today
        If Not IsDate(DStart) Then DStart = YearStart(DEnd)

        StartingInventory = 0
        If IsDate(DStart) Then
            SQL = ""
            SQL = SQL & "SELECT * FROM [Detail]"
            SQL = SQL & " WHERE [Style]='" & Style & "'"
            SQL = SQL & " AND [Ddate1] < #" & DateFormat(DStart) & "#"
            RS = GetRecordsetBySQL(SQL, , GetDatabaseInventory)

            Do While Not RS.EOF
                X = UCase(Trim(IfNullThenNilString(RS("Trans"))))
                StartingInventory = StartingInventory + IfNullThenZeroDouble(RS("NewStock")) - IfNullThenZeroDouble(RS("AmtSold"))
                '      If X = "IN" Then
                '        StartingInventory = StartingInventory + IfNullThenZeroDouble(RS("NewStock"))
                '      ElseIf IsIn(X, "DS", "NS") Then
                '        StartingInventory = StartingInventory - IfNullThenZeroDouble(RS("AmtSold"))
                '      End If
                RS.MoveNext()
            Loop
        End If
        If StartingInventory < 0 Then StartingInventory = 0


        SQL = ""
        SQL = SQL & "SELECT * FROM [Detail]"
        SQL = SQL & " WHERE [Style]='" & Style & "'"
        SQL = SQL & " AND [Trans] IN ('IN','NS','DS')"
        SQL = SQL & " AND ([Ddate1] BETWEEN #" & DateFormat(DStart) & "# AND #" & DateFormat(DEnd) & "#)"
        SQL = SQL & " ORDER BY Ddate1"
        RS = GetRecordsetBySQL(SQL, , GetDatabaseInventory)
        If RS.EOF Then Exit Function

        AverageInventory = 0
        EndingInventory = StartingInventory
        NewItems = StartingInventory
        AmtSold = 0
        LndSold = 0
        SlsSold = 0
        pNI = StartingInventory
        pAS = 0
        First = True

        LastD = DStart
        Do While Not RS.EOF
            If Not IsNothing(RS("DDate1").Value) Then
                DD = DateDiff("d", LastD, RS("DDate1").Value)
                If DD >= 1 Then
                    AverageInventory = AverageInventory + ((NewItems - AmtSold) * DD)
                    pNI = 0
                    pAS = 0
                End If

                If IsDate(IfNullThenNilString(RS("DDate1").Value)) Then
                    LastD = IfNullThenNilString(RS("DDate1").Value)
                End If

            End If

            First = False

            X = UCase(Trim(IfNullThenNilString(RS("Trans"))))
            R = IfNullThenZeroDouble(RS("NewStock")) '+ IfNullThenZeroDouble(RS("SpecOrd"))
            T = IfNullThenZeroDouble(RS("AmtSold")) '+ IfNullThenZeroDouble(RS("SpecOrd"))
            If X = "IN" And R >= 0 Then
                NewItems = NewItems + R
                pNI = pNI + R
            ElseIf X = "IN" And R < 0 Then
                AmtSold = AmtSold - R
                pAS = pAS - R
            ElseIf IsIn(X, "DS", "NS") Then
                AmtSold = AmtSold + T
                pAS = pAS + T

                M = New CGrossMargin
                M.DataAccess.DataBase = GetDatabaseAtLocation(RS("Store").Value)
                If M.Load(RS("MarginRn").Value, "#MarginLine") Then
                    LndSold = LndSold + M.Cost
                    If IsPatchApplied("Calculate Packages") Then
                        SlsSold = SlsSold + M.ItemFreight + M.PackSell
                    Else
                        SlsSold = SlsSold + M.ItemFreight + CalculateKitItemSellPrice(M.SaleNo, M.MarginLine)
                    End If
                End If
                DisposeDA(M)
            End If

            EndingInventory = StartingInventory + NewItems - AmtSold

            RS.MoveNext()
        Loop

        DD = DateDiff("d", LastD, DEnd) + 1
        AverageInventory = AverageInventory + (EndingInventory * DD)

        ' for ByRef params
        COGS = LndSold
        Sales = SlsSold

        Per = DateDiff("d", DStart, DEnd) + 1
        If Per <> 0 Then AverageInventory = AverageInventory / Per Else AverageInventory = StartingInventory '1

        InvDollars = GetItemCost(Style, , False, IIf(UseEndingInventory, EndingInventory, AverageInventory))
        GMDollars = SlsSold - LndSold
        If InvDollars = 0 Then CalculateGMROI = 0 Else CalculateGMROI = GMDollars / InvDollars
    End Function

    Public Function CalculateKitItemSellPrice(ByVal SaleNo As String, ByVal MarginLine As Integer) As Decimal
        '::::CalculateKitItemSellPrice
        ':::SUMMARY
        ': Calculate Kit Item selling price.
        ':::DESCRIPTION
        ':This function is  used to  get selling price of items in kit, based on kit total cost and number of total kits sold._
        'The formulas are given below.
        ':::PARAMETERS
        ':-SaleNo-Represents the sale number.
        ':-MarginLine
        '::::RETURN
        ':Currency-Represents the selling price of items in kit in a currency.
        Dim M As CGrossMargin, S As String
        Dim KitCostTotal As Decimal, KitSoldTotal As Decimal, IsKit As Boolean
        Dim FoundML As Boolean, ItemCost As Decimal

        M = New CGrossMargin
        S = "SELECT * FROM [GrossMargin] WHERE SaleNo='" & SaleNo & "' ORDER BY MarginLine"

        If M.DataAccess.Records_OpenSQL(S) Then
            Do While M.DataAccess.Records_Available
                If Val(M.MarginLine) = MarginLine Then
                    If IsKit Or GetPrice(M.SellPrice) = 0 Then  ' only let it work for real kits
                        FoundML = True
                        ItemCost = GetPrice(M.Cost)
                    Else
                        'if it's not a kit, just return sell price
                        CalculateKitItemSellPrice = GetPrice(M.SellPrice)
                        DisposeDA(M)
                        Exit Function
                    End If
                End If
                If GetPrice(M.SellPrice) = 0 Then     ' somewhere in a kit
                    IsKit = True
                    KitCostTotal = KitCostTotal + GetPrice(M.Cost)
                    ' sell price is zero, so this doesn't matter
                ElseIf IsKit Then                     ' end of kit
                    KitSoldTotal = KitSoldTotal + GetPrice(M.SellPrice)
                    KitCostTotal = KitCostTotal + GetPrice(M.Cost)

                    If FoundML Then
                        If KitCostTotal = 0 Then
                            CalculateKitItemSellPrice = 0
                        Else
                            ' formula is ratio of target item cost to total kit cost (should be <1.0), times kit sold total..
                            CalculateKitItemSellPrice = ItemCost / KitCostTotal * KitSoldTotal
                        End If
                        DisposeDA(M)
                        Exit Function
                    End If

                    IsKit = False
                    KitSoldTotal = 0
                    KitCostTotal = 0
                End If
            Loop
        End If

        DisposeDA(M)
    End Function

    ' directly reversible from Calculate GM
    Public Function CalculateSalePrice(ByVal Landed As Decimal, ByVal GM As Double) As Decimal
        '::::CalculateSalePrice
        ':::SUMMARY
        ': Reverse calculates Sale price.
        ':::DESCRIPTION
        ': Calculate sale price based on Landed and a given GM.
        ':::PARAMETERS
        ':-Landed-Denotes the landed cost of item.
        ':-GM-Denotes the gross margin.
        ':::RETURN
        ':Currency-Returns the sales price in a currency.
        ':::SEE ALSO
        ':CalculateOnSale

        If GM >= 100 Or GM <= 0 Then CalculateSalePrice = Landed : Exit Function
        If GM <= 10 Then
            CalculateSalePrice = Landed * GM   ' BFH20060420 this clause added to reflect InvenA
        Else
            CalculateSalePrice = Landed / (1 - GM / 100)
        End If
    End Function

End Module
