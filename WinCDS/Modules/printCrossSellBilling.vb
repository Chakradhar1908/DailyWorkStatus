Module printCrossSellBilling
    Public Sub printCrossSellBilling_PrintRecords(ByVal MinStore As Integer, ByVal nMaxStore As Integer, ByVal DelDate As Date, ByVal EndDate As Date, ByVal Juice As Double, ByVal ToPrinter As Boolean, ByVal DriverPickup As Boolean)
        On Error GoTo HandleErr
        Dim RS As ADODB.Recordset, SQL As String
        Dim Store As Integer, LastStore As Integer, LastSale As String
        Dim SaleTotal As Decimal, StoreTotal As Decimal
        Dim OO As Object
        Dim JuiceAmt As Decimal, X As Integer

        If ToPrinter Then
            OO = Printer
        Else
            OutputObject = frmPrintPreviewDocument.picPicture
            OO = OutputObject
            frmPrintPreviewDocument.CallingForm = InvPull
        End If

        If MinStore = 0 Then MinStore = 1
        If nMaxStore = 0 Then nMaxStore = ssMaxStore

        ' For each (selected) store, get a list of items sent to other stores.
        For Store = MinStore To nMaxStore
            SQL = ""
            SQL = SQL & " SELECT Location, DelDate, Style, Status, Desc, Vendor, SaleNo, MarginLine"
            SQL = SQL & "  , Name, Mail.First, Cost, ItemFreight, Quantity"
            SQL = SQL & " From GrossMargin LEFT JOIN Mail ON GrossMargin.MailIndex = Mail.Index"
            SQL = SQL & " WHERE "
            SQL = SQL & "  GrossMargin.Location<>" & Store & " AND GrossMargin.Location<>0 "
            SQL = SQL & "  AND InStr(1,',ST,TW,DELST,SO,SOREC,SS,SSREC,SSLAW,PO,POREC,LAW,FND,', ',' & trim(GrossMargin.Status) & ',')>0"
            'BFH20070522... was out, put back in.. wasn't sure why it was out.
            '    SQL = SQL & "  AND DelDate BETWEEN #" & DelDate & "# AND #" & EndDate & "#"
            'BFH20080213 Made it check SellDate if no delivery date is set up...
            SQL = SQL & "  AND iif(isnull(DelDate),SellDAte,DelDate) BETWEEN #" & DelDate & "# AND #" & EndDate & "#"
            SQL = SQL & " ORDER BY GrossMargin.Location, GrossMargin.SaleNo" ' GrossMargin.DelDate, GrossMargin.Name"

            RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(Store))

            Do Until RS.EOF
                ' Go to a new page if the store changes, or if we would pass the margin.
                'If LastSale = "25605B" Then Stop
                If (LastStore <> 0 And LastStore <> RS("Location").Value) Or (OO.CurrentY + 2 * OO.TextHeight("X") > Printer.ScaleHeight) Then
                    If ToPrinter Then
                        OO.NewPage
                    Else
                        frmPrintPreviewDocument.NewPage()
                    End If
                End If

                If OO.CurrentY = 0 Then    ' New page, print headers.
                    If Not DriverPickup Then
                        PrintReportHeader("Multi-Store Cross Selling Billing", DelDate, EndDate, OO)
                        '          PrintCompanyInformation Store, OO        ' Source store goes at the top.
                        '          PrintCompanyBilledInformation RS("Location"), OO ' Destination store below it.
                        PrintCompanyInformation(RS("Location").Value, OO)        ' Source store goes at the top.
                        PrintCompanyBilledInformation(Store, OO) ' Destination store below it.
                        OO.Print("")

                        OO.FontSize = 10
                        OO.FontBold = True

                        PrintToPosition(OO, "Sale No: " & RS("SaleNo").Value, 0, VBRUN.AlignConstants.vbAlignLeft, False)
                        PrintToPosition(OO, "Delivery Date: " & RS("DelDate").Value, 2500, VBRUN.AlignConstants.vbAlignLeft, False)
                        PrintToPosition(OO, "Cust: " & Trim(RS("First").Value) & " " & Trim(RS("Name").Value), 5000, VBRUN.AlignConstants.vbAlignLeft, False)
                        PrintToPosition(OO, "Landed:", 11500, VBRUN.AlignConstants.vbAlignRight, True)
                        X = 0
                    Else
                        PrintReportHeader("Multi-Store Cross Driver Pickup", DelDate, EndDate, OO)
                        PrintCompanyInformation(Store, OO)        ' Source store goes at the top.
                        PrintToPosition(OO, "Pickup From Loc" & RS("Location").Value & ", " & StoreSettings(Val(RS("Location"))).Name, 0, VBRUN.AlignConstants.vbAlignLeft, True)
                        PrintToPosition(OO, Space(35) & StoreSettings(Val(RS("Location"))).Address, 0, VBRUN.AlignConstants.vbAlignLeft, True)
                        PrintToPosition(OO, "Deliver to Loc" & Store & ", " & StoreSettings(Store).Name, 0, VBRUN.AlignConstants.vbAlignLeft, True)
                        PrintToPosition(OO, Space(35) & StoreSettings(Store).Address, 0, VBRUN.AlignConstants.vbAlignLeft, True)

                        OO.Print("")

                        OO.FontSize = 10
                        OO.FontBold = True
                        PrintToPosition(OO, "SaleNo", 200, VBRUN.AlignConstants.vbAlignLeft, False)
                        PrintToPosition(OO, "DelDate", 1100, VBRUN.AlignConstants.vbAlignLeft, False)
                        X = 2000
                        PrintToPosition(OO, "#", X + 200, VBRUN.AlignConstants.vbAlignRight, False)
                        PrintToPosition(OO, "Style", X + 350, VBRUN.AlignConstants.vbAlignLeft, False)
                        PrintToPosition(OO, "Vendor", X + 2500, VBRUN.AlignConstants.vbAlignLeft, False)
                        PrintToPosition(OO, "St", X + 4300, VBRUN.AlignConstants.vbAlignLeft, False)
                        PrintToPosition(OO, "Loc", X + 4800, VBRUN.AlignConstants.vbAlignLeft, False)
                        PrintToPosition(OO, "Desc", X + 5600, VBRUN.AlignConstants.vbAlignLeft, True)
                    End If
                End If

                OO.FontBold = False
                OO.FontSize = 8
                If DriverPickup Then
                    PrintToPosition(OO, RS("SaleNo").Value, 200, VBRUN.AlignConstants.vbAlignLeft, False)
                    PrintToPosition(OO, RS("DelDate").Value, 1100, VBRUN.AlignConstants.vbAlignLeft, False)
                End If
                PrintToPosition(OO, RS("Quantity").Value, X + 200, VBRUN.AlignConstants.vbAlignRight, False)
                PrintToPosition(OO, RS("Style").Value, X + 350, VBRUN.AlignConstants.vbAlignLeft, False)
                PrintToPosition(OO, RS("Vendor").Value, X + 2500, VBRUN.AlignConstants.vbAlignLeft, False)
                PrintToPosition(OO, RS("Status").Value, X + 4300, VBRUN.AlignConstants.vbAlignLeft, False)
                PrintToPosition(OO, "Loc: " & RS("Location").Value, X + 4800, VBRUN.AlignConstants.vbAlignLeft, False)
                PrintToPosition(OO, RS("Desc").Value, X + 5600, VBRUN.AlignConstants.vbAlignLeft, DriverPickup)
                If Not DriverPickup Then PrintToPosition(OO, CurrencyFormat(RS("ItemFreight").Value + RS("Cost").Value), X + 11500, VBRUN.AlignConstants.vbAlignRight, True)

                LastSale = RS("SaleNo").Value
                LastStore = RS("Location").Value
                SaleTotal = SaleTotal + RS("Cost").Value + RS("ItemFreight").Value
                StoreTotal = StoreTotal + RS("Cost").Value + RS("ItemFreight").Value

                RS.MoveNext()

                If RS.EOF Then
                    ' At the last record, close the sale and store data displays.
                    If Not DriverPickup Then
                        JuiceAmt = SaleTotal * Juice
                        If JuiceAmt <> 0 Then PrintToPosition(OutputObject, "Warehse Chg:              " & Format(JuiceAmt, "###,###.00"), 11500, VBRUN.AlignConstants.vbAlignRight, True)
                        PrintToPosition(OO, "Sale Total: " & FormatCurrency(SaleTotal + JuiceAmt), 11500, VBRUN.AlignConstants.vbAlignRight, True)
                        OO.Print
                        OO.FontBold = True
                        OO.FontSize = 14
                        PrintToPosition(OO, "Store Total: " & FormatCurrency(StoreTotal + JuiceAmt), 11500, VBRUN.AlignConstants.vbAlignRight, True)
                        OO.FontBold = False
                        SaleTotal = 0
                        StoreTotal = 0
                    End If
                ElseIf (LastSale <> RS("SaleNo").Value And Not DriverPickup) Or LastStore <> RS("Location").Value Then

                    If Not DriverPickup Then
                        JuiceAmt = SaleTotal * Juice
                        If JuiceAmt <> 0 Then PrintToPosition(OutputObject, "Warehse Chg:              " & Format(JuiceAmt, "###,###.00"), 11500, VBRUN.AlignConstants.vbAlignRight, True)

                        PrintToPosition(OO, "Sale Total: $" & Format(SaleTotal, "0.00"), 11500, VBRUN.AlignConstants.vbAlignRight, True)
                        SaleTotal = 0

                        If LastStore <> RS("Location").Value Then
                            OO.Print
                            OO.FontBold = True
                            OO.FontSize = 14
                            PrintToPosition(OO, "Store Total: $" & Format(StoreTotal, "0.00"), 11500, VBRUN.AlignConstants.vbAlignRight, True)
                            OO.FontBold = False
                            StoreTotal = 0
                        ElseIf OO.CurrentY + 5 * OO.TextHeight("X") < OO.ScaleHeight Then
                            OO.Print
                            OO.FontSize = 10
                            OO.FontBold = True
                            PrintToPosition(OO, "Sale No: " & RS("SaleNo").Value, 0, VBRUN.AlignConstants.vbAlignLeft, False)
                            PrintToPosition(OO, "Delivery Date: " & RS("DelDate").Value, 2500, VBRUN.AlignConstants.vbAlignLeft, False)
                            PrintToPosition(OO, "Cust: " & Trim(RS("First").Value) & " " & Trim(RS("Name").Value), 5000, VBRUN.AlignConstants.vbAlignLeft, False)
                            PrintToPosition(OO, "Landed", 11500, VBRUN.AlignConstants.vbAlignRight, True)
                            SaleTotal = 0
                        Else
                            ' Next customer won't fit nicely on this page, start a new one.
                            ' Headers will be added by the next loop.
                            OO.NewPage
                        End If
                    End If
                Else
                    ' Sale continues on the next line..
                End If
            Loop
            RS.Close()
            RS = Nothing
            LastStore = -1  ' Always start the next store on a new page.
        Next

        If ToPrinter Then
            Printer.EndDoc()
        Else
            InvPull.Hide()
            frmPrintPreviewDocument.DataEnd()
        End If
        Exit Sub

HandleErr:
        Resume Next
    End Sub

    Private Sub PrintReportHeader(ByVal RptTitle As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal OutObj As Object)
        Dim MinY As Integer
        OutObj.FontSize = 20
        OutObj.FontBold = True
        OutObj.CurrentX = (Printer.ScaleWidth - OutObj.TextWidth(RptTitle)) / 2
        OutObj.Print(RptTitle)
        MinY = OutObj.CurrentY

        OutObj.FontSize = 8
        OutObj.FontBold = False
        OutObj.CurrentX = 9500
        OutObj.Print("From:  " & DateFormat(StartDate))
        OutObj.CurrentX = 9500
        OutObj.Print("To:      " & DateFormat(EndDate))
        OutObj.CurrentY = MinY
    End Sub

    Private Sub PrintCompanyInformation(ByVal Store As Integer, ByVal OutObj As Object)
        On Error Resume Next
        OutObj.FontName = "Arial"
        OutObj.FontSize = 14
        OutObj.FontBold = False
        OutObj.DrawWidth = 20
        OutObj.CurrentX = (Printer.ScaleWidth - OutObj.TextWidth(StoreSettings(Store).Name)) / 2
        OutObj.Print(StoreSettings(Store).Name)
        OutObj.CurrentX = (Printer.ScaleWidth - OutObj.TextWidth(StoreSettings(Store).Address)) / 2
        OutObj.Print(StoreSettings(Store).Address)
        OutObj.CurrentX = (Printer.ScaleWidth - OutObj.TextWidth(StoreSettings(Store).City)) / 2
        OutObj.Print(StoreSettings(Store).City)
        OutObj.CurrentX = (Printer.ScaleWidth - OutObj.TextWidth(StoreSettings(Store).Phone)) / 2
        OutObj.Print(StoreSettings(Store).Phone)
        OutObj.Print
    End Sub

    Private Sub PrintCompanyBilledInformation(ByVal Store As Integer, ByVal OutObj As Object)
        On Error Resume Next
        OutObj.FontSize = 14
        OutObj.CurrentX = 150
        OutObj.Print("Bill To:")

        OutObj.FontSize = 12
        OutObj.CurrentX = 300
        OutObj.Print(StoreSettings(Store).Name)
        OutObj.CurrentX = 300
        OutObj.Print(StoreSettings(Store).Address)
        OutObj.CurrentX = 300
        OutObj.Print(StoreSettings(Store).City)
        OutObj.CurrentX = 300
        OutObj.Print(StoreSettings(Store).Phone)
        OutObj.Print
    End Sub
End Module
