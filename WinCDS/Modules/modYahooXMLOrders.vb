Module modYahooXMLOrders
    Private Const StoreNumForOnlineSale As Integer = 1
    Public Structure OnlineSaleOrderItem
        Dim Num As String
        Dim ID As String
        Dim Style As String
        Dim Quantity As Integer
        Dim UnitPrice As Decimal
        Dim Description As String
        Dim URL As String
        Dim Taxable As Boolean
    End Structure
    Public Structure OnlineSaleAddress
        Dim Type As String
        Dim MailIndex As Integer

        Dim First As String
        Dim Last As String
        Dim Address1 As String
        Dim Address2 As String
        Dim City As String
        Dim State As String
        Dim Country As String
        Dim Zip As String
        Dim Phone As String
        Dim Email As String
    End Structure

    Public Structure OnlineSaleOrder
        Dim XML As String
        Dim SaleNo As String

        Dim ID As String
        Dim Currency As String

        Dim Time As String
        Dim NumericTime As String

        Dim Referer As String
        Dim EntryPoint As String

        Dim CouponID As String
        Dim CouponDesc As String
        Dim CouponAmount As String

        Dim ServiceProvider As String
        Dim LoginID As String

        Dim ShippingType As String

        Dim CCExp As String
        Dim CCType As String


        Dim BillTo As OnlineSaleAddress
        Dim ShipTo As OnlineSaleAddress

        Dim OrderItems() As OnlineSaleOrderItem
        Dim OrderItemCount As Integer

        Dim SpecInstr As String

        '            giftwrap, giftcertificate, coupon, taxableamt, nontaxableamt
        Dim Discount As String
        Dim MiscAdjustment As String
        Dim ServiceFee As Decimal     ' e.g. PayPal's fee
        Dim SubTotal As String
        Dim ShippingAmount As String
        Dim Tax As String
        Dim Credit As String
        Dim Total As String

        Dim Warning As String
        Dim Suspect As String
        Dim Bogus As Boolean
    End Structure

    Public Function OnlineSaleIDExists(ByVal ID As String) As Boolean
        Dim R As ADODB.Recordset
        R = GetRecordsetBySQL("SELECT [OrderID] FROM [OnlineOrderRecord] WHERE [OrderID]='" & ID & "'", , GetDatabaseInventory)
        OnlineSaleIDExists = R.RecordCount > 0
        DisposeDA(R)
    End Function

    Public Function AddOrderItemToOnlineSale(ByRef Order As OnlineSaleOrder, ByRef Item As OnlineSaleOrderItem) As Integer
        Dim X As Integer
        On Error Resume Next
        X = -1
        X = UBound(Order.OrderItems)
        X = X + 1
        ReDim Preserve Order.OrderItems(X)
        Order.OrderItems(X) = Item
        Order.OrderItemCount = X + 1
        AddOrderItemToOnlineSale = Order.OrderItemCount
    End Function

    Public Function TotalizeOnlineSaleOrder(ByRef Order As OnlineSaleOrder, Optional ByVal StoreNo As Integer = 0) As Boolean
        ' BFH20050425
        ' To save an order from the web, we're going through BOS and BOS2..
        ' There are just too many things to keep track of and to keep one copy of each
        ' it just seems like it makes the most sense..
        ' we load (but hide) all the forms...
        '   Step 1:  Load BOS and commit
        '   Step 2:  Load BOS2 and process the sale
        '   Step 3:  Cleanup and log the internet sale in a new Web Tracking table


        Dim L As Integer, oI As OnlineSaleOrderItem, I As Integer
        Dim RS As ADODB.Recordset, Mfg As String, Dsc As String, Sty As String
        Dim cM As New clsMailRec
        Dim oSN As Integer
        Dim ServNm As String
        Dim cInv As CInvRec

        ServNm = UCase(Order.ServiceProvider)
        'Exit Function
#Const HideBOS = True

        If StoreNo = 0 Then StoreNo = StoresSld
        oSN = StoresSld
        StoresSld = StoreNo

#If HideBOS Then
        BillOSale.Hide()
        BillOSale.Visible = False
        BillOSale.BOS2IsHidden = True
        BillOSale.Hide()
        BillOSale.Visible = False
#Else
  BillOSale.Show
#End If

        BillOSale.ClearBillOfSale()
        BillOSale.IsInternetSale = True  ' these will be cleared b/c we unload BOS and BOS2 at the end of this!

        modProgramState.Order = "A"

        If cM.Load(Order.BillTo.Phone, "Tele") Then
            BillOSale.Index = Trim(cM.Index)
        Else
            BillOSale.Index = 0
        End If
        DisposeDA(cM)

        If StoreSettings.bManualBillofSaleNo Then
            BillOSale.txtSaleNo.Text = GetLeaseNumber(True)  'if normally requires manual BOS, force the next one on it..
        End If

        BillOSale.CustomerFirst.Text = UCase(Order.BillTo.First)
        BillOSale.CustomerLast.Text = UCase(Order.BillTo.Last)
        BillOSale.Email.Text = Order.BillTo.Email
        BillOSale.CustomerAddress.Text = UCase(Order.BillTo.Address1)
        BillOSale.AddAddress.Text = UCase(Order.BillTo.Address2)
        BillOSale.CustomerCity.Text = UCase(Order.BillTo.City)
        BillOSale.CustomerZip.Text = UCase(Order.BillTo.Zip)
        BillOSale.CustomerPhone1.Text = UCase(Order.BillTo.Phone)

        BillOSale.txtSpecInst.Text = Order.SpecInstr

        If Order.BillTo.First <> Order.ShipTo.First Or
     Order.BillTo.Last <> Order.ShipTo.Last Or
     Order.BillTo.Address1 <> Order.ShipTo.Address1 Or
     Order.BillTo.Address2 <> Order.ShipTo.Address2 Or
     Order.BillTo.City <> Order.ShipTo.City Or
     Order.BillTo.State <> Order.ShipTo.State Or
     Order.BillTo.Zip <> Order.ShipTo.Zip Or
     Order.BillTo.Phone <> Order.ShipTo.Phone _
    Then
            BillOSale.ShipToFirst.Text = UCase(Order.ShipTo.First)
            BillOSale.ShipToLast.Text = UCase(Order.ShipTo.Last)
            BillOSale.CustomerAddress2.Text = UCase(Order.ShipTo.Address1)
            BillOSale.CustomerCity2.Text = UCase(Order.ShipTo.City)
            BillOSale.CustomerZip2.Text = UCase(Order.ShipTo.Zip)
            BillOSale.CustomerPhone3.Text = UCase(Order.ShipTo.Phone)
        End If

        'BillOSale.cmdApplyBillOSale.Value = True               '  Step one complete!
        BillOSale.cmdApplyBillOSale.PerformClick()
#If HideBOS Then
        BillOSale.Hide()             ' would clicking OK on BOS make BOS2 visible?
        BillOSale.Visible = False
        OrdSelect.Hide()
        OrdSelect.Visible = False
#End If
        For L = 0 To Order.OrderItemCount - 1
            oI = Order.OrderItems(L)

            Dim iK As cInvKit, iKitItem As CInvRec, LineStatus As String
            DisposeDA(cInv)
            cInv = New CInvRec
            If cInv.Load(oI.Style, "Style") Then
                LineStatus = IIf(cInv.QueryStock(StoreNo) < oI.Quantity, "SO", "ST")
                Mfg = cInv.Vendor
                Dsc = cInv.Desc
            Else
                LineStatus = "SS"
                Mfg = ""
                Dsc = oI.Description

                iK = New cInvKit
                iK.DataAccess.DataBase = GetDatabaseAtLocation(1)
                iK.DataAccess.Records_OpenSQL("SELECT * FROM InvKit WHERE KitStyleNo='" & oI.Style & "' OR KitSKU='" & oI.Style & "'")

                If iK.DataAccess.Record_Count >= 1 Then
                    LineStatus = "ST"
                    iK.DataAccess.Records_MoveAbsolute(1)
                    For I = 1 To 10
                        If iK.Item(I) = "" Then Exit For
                        iKitItem = New CInvRec
                        iKitItem.Load(iK.Item(I), "Style")
                        AddBOS2Line(iK.Item(I), iKitItem.Vendor, StoreNo, LineStatus, iK.Quantity(I) * oI.Quantity, iKitItem.Desc, 0)
                        Mfg = iKitItem.Vendor
                        DisposeDA(iKitItem)
                    Next
                    oI.Style = iK.KitStyleNo   ' after all the items are added, the KIT- line will be added with the price
                End If
            End If
            DisposeDA(cInv, iK)

            AddBOS2Line(oI.Style, Mfg, StoreNumForOnlineSale, LineStatus, oI.Quantity, Dsc, oI.Quantity * oI.UnitPrice)
        Next
        DisposeDA(cInv)

        AddBOS2Note("WEBSITE SALE" & IIf(Order.LoginID <> "", " (" & ServNm & " Login=" & Order.LoginID, ""))

        If Order.CouponID <> "" Then AddBOS2Note(ServNm & " COUPON (" & Order.CouponID & "): " & Order.CouponDesc, Order.CouponAmount)
        If Order.Warning <> "" Then AddBOS2Note(ServNm & " WARNING: " & Order.Warning)
        If Order.Suspect <> "" Then AddBOS2Note(ServNm & " SUSPECT: " & Order.Suspect)
        If Order.Bogus Then AddBOS2Note(ServNm & " MARKED AS BOGUS")

        If Order.Discount <> "" Then AddBOS2Note(ServNm & " DISCOUNT", Order.Discount)
        If Order.MiscAdjustment <> "" Then AddBOS2Note(ServNm & " MISC ADJ", Order.MiscAdjustment)

        If Order.Credit <> "" Then AddBOS2Note(ServNm & " CREDIT", Order.Credit)
        If Order.ShippingAmount <> "" Then AddBOS2Line("DEL", , , , , "DELIVERY CHARGE (" & Order.ShippingType & ")", Order.ShippingAmount)


        AddBOS2Line("SUB", , , , , "               Sub Total =", Order.SubTotal)

        If GetPrice(Order.Tax) <> 0 Then
            AddBOS2Line("TAX1", , , , 1, "SALES TAX ", Order.Tax)
            AddBOS2Line("SUB", , , , , "               Sub Total =", Order.Total)
        End If

        ' bfh20050426
        ' BOS2 (OrdSelect really) puts the payment type into the quantity field..
        ' because we don't know for sure what Yahoo will send us other than 'Visa', we check
        ' the first letter of the cctype because that should determine it...
        ' this could be expanded, but it should be OK for now too
        Dim C As String, Q As String
        C = LCase(Left(Order.CCType, 1))  ' we don't know what they'll look like yet!!!
        Select Case C
            Case "v" : Q = "3" ' visa
            Case "m" : Q = "4" ' mcard
            Case "d" : Q = IIf(LCase(Mid(Order.CCType, 2, 1)) = "i", "5", "9") ' discover / debit
            Case "a" : Q = "6" ' amex
            Case Else : Q = "1"
        End Select
        AddBOS2Line("PAYMENT", , , , Q, "PAYPAL SALE", Order.Total)

        'Unload OrdSelect
        OrdSelect.Close()
        'BillOSale.cmdProcessSale.Value = True                 '  This should be all of step 2!!
        BillOSale.cmdProcessSale.PerformClick()
        Order.SaleNo = BillOSale.BillOfSale.Text                            '  store for tracking

        If Order.ServiceFee <> 0 Then
            ' add cash table entry for "Bank Fees", 40250
            '      AddBOS2Note ServNm & " FEE " & FormatCurrency(Order.ServiceFee)
        End If


        BillOSale.BOS2IsHidden = False            ' unloading the forms doesn't clear these!!!
        BillOSale.IsInternetSale = False
        'Unload BillOSale
        BillOSale.Close()
        StoresSld = oSN

        TotalizeOnlineSaleOrder = True
        ' Record The Sale
        Dim SQL As String
        SQL = "INSERT INTO [OnlineOrderRecord] (OrderID, ServiceProvider, SaleNo, Location, OrderTime, OrderItems, Total, XML) VALUES " &
                        "(""" & ProtectSQL(Order.ID) & """," &
                        """" & ProtectSQL(Order.ServiceProvider) & """," &
                        """" & ProtectSQL(Order.SaleNo) & """," &
                        """" & ProtectSQL(StoreNumForOnlineSale) & """," &
                        """" & ProtectSQL(Order.Time) & """," &
                        Order.OrderItemCount & "," &
                        IIf(Order.Total = "", "0", Order.Total) & "," &
                        """" & ProtectSQL(Order.XML) & """)"  ' ")" '
        ExecuteRecordsetBySQL(SQL, , GetDatabaseInventory, True)
    End Function

    Private Function AddBOS2Line(Optional ByVal Style As String = "", Optional ByVal Mfg As String = "", Optional ByVal Loc As String = "", Optional ByVal Status As String = "", Optional ByVal Quan As String = "", Optional ByVal Desc As String = "", Optional ByVal Price As String = "") As Integer
        AddBOS2Line = BillOSale.NewStyleLine
        BillOSale.RowClear(AddBOS2Line)

        BillOSale.SetStyle(AddBOS2Line, UCase(Style))
        BillOSale.SetMfg(AddBOS2Line, UCase(Mfg))
        BillOSale.SetLoc(AddBOS2Line, Loc)
        BillOSale.SetStatus(AddBOS2Line, Status)
        BillOSale.SetQuan(AddBOS2Line, Quan)
        BillOSale.SetDesc(AddBOS2Line, UCase(Desc))
        If GetPrice(Price) <> 0# Then BillOSale.SetPrice(AddBOS2Line, Price)

        BillOSale.StyleAddEnd()
    End Function

    Private Function AddBOS2Note(Optional ByVal Desc As String = "", Optional ByVal Amt As String = "") As Integer
        Dim L As String, Q As String
        L = IIf(Amt = "", "", StoreNumForOnlineSale)
        Q = IIf(Amt = "", "", 1)
        AddBOS2Note = AddBOS2Line("NOTES", , L, , Q, Desc, Amt)
    End Function

End Module
