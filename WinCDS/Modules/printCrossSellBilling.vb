Module printCrossSellBilling
    Public Sub printCrossSellBilling_PrintRecords(ByVal MinStore As Long, ByVal nMaxStore As Long, ByVal DelDate As Date, ByVal EndDate As Date, ByVal Juice As Double, ByVal ToPrinter As Boolean, ByVal DriverPickup As Boolean)
        On Error GoTo HandleErr
        Dim RS As ADODB.Recordset, SQL As String
        Dim Store As Long, LastStore As Long, LastSale As String
        Dim SaleTotal As Currency, StoreTotal As Currency
        Dim OO As Object
        Dim JuiceAmt As Currency, X As Long

        If ToPrinter Then
    Set OO = Printer
  Else
    Set OutputObject = frmPrintPreviewDocument.picPicture
    Set OO = OutputObject
    Set frmPrintPreviewDocument.CallingForm = InvPull
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
    

    Set RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(Store))
    
    Do Until RS.EOF
                ' Go to a new page if the store changes, or if we would pass the margin.
                'If LastSale = "25605B" Then Stop
                If (LastStore <> 0 And LastStore <> RS("Location")) Or (OO.CurrentY + 2 * OO.TextHeight("X") > Printer.ScaleHeight) Then
                    If ToPrinter Then
                        OO.NewPage
                    Else
                        frmPrintPreviewDocument.NewPage()
                    End If
                End If

                If OO.CurrentY = 0 Then    ' New page, print headers.
                    If Not DriverPickup Then
                        PrintReportHeader "Multi-Store Cross Selling Billing", DelDate, EndDate, OO
'          PrintCompanyInformation Store, OO        ' Source store goes at the top.
                        '          PrintCompanyBilledInformation RS("Location"), OO ' Destination store below it.
                        PrintCompanyInformation RS("Location"), OO        ' Source store goes at the top.
                        PrintCompanyBilledInformation Store, OO ' Destination store below it.
                        OO.Print ""

          OO.FontSize = 10
                        OO.FontBold = True

                        PrintToPosition OO, "Sale No: " & RS("SaleNo"), 0, vbAlignLeft, False
          PrintToPosition OO, "Delivery Date: " & RS("DelDate"), 2500, vbAlignLeft, False
          PrintToPosition OO, "Cust: " & Trim(RS("First")) & " " & Trim(RS("Name")), 5000, vbAlignLeft, False
          PrintToPosition OO, "Landed:", 11500, vbAlignRight, True
          X = 0
                    Else
                        PrintReportHeader "Multi-Store Cross Driver Pickup", DelDate, EndDate, OO
          PrintCompanyInformation Store, OO        ' Source store goes at the top.
                        PrintToPosition OO, "Pickup From Loc" & RS("Location") & ", " & StoreSettings(Val(RS("Location"))).Name, 0, vbAlignLeft, True
          PrintToPosition OO, Space(35) & StoreSettings(Val(RS("Location"))).Address, 0, vbAlignLeft, True
          PrintToPosition OO, "Deliver to Loc" & Store & ", " & StoreSettings(Store).Name, 0, vbAlignLeft, True
          PrintToPosition OO, Space(35) & StoreSettings(Store).Address, 0, vbAlignLeft, True

          OO.Print ""

          OO.FontSize = 10
                        OO.FontBold = True
                        PrintToPosition OO, "SaleNo", 200, vbAlignLeft, False
          PrintToPosition OO, "DelDate", 1100, vbAlignLeft, False
          X = 2000
                        PrintToPosition OO, "#", X + 200, vbAlignRight, False
          PrintToPosition OO, "Style", X + 350, vbAlignLeft, False
          PrintToPosition OO, "Vendor", X + 2500, vbAlignLeft, False
          PrintToPosition OO, "St", X + 4300, vbAlignLeft, False
          PrintToPosition OO, "Loc", X + 4800, vbAlignLeft, False
          PrintToPosition OO, "Desc", X + 5600, vbAlignLeft, True
        End If
                End If

                OO.FontBold = False
                OO.FontSize = 8
                If DriverPickup Then
                    PrintToPosition OO, RS("SaleNo"), 200, vbAlignLeft, False
        PrintToPosition OO, RS("DelDate"), 1100, vbAlignLeft, False
      End If
                PrintToPosition OO, RS("Quantity"), X + 200, vbAlignRight, False
      PrintToPosition OO, RS("Style"), X + 350, vbAlignLeft, False
      PrintToPosition OO, RS("Vendor"), X + 2500, vbAlignLeft, False
      PrintToPosition OO, RS("Status"), X + 4300, vbAlignLeft, False
      PrintToPosition OO, "Loc: " & RS("Location"), X + 4800, vbAlignLeft, False
      PrintToPosition OO, RS("Desc"), X + 5600, vbAlignLeft, DriverPickup
      If Not DriverPickup Then PrintToPosition OO, CurrencyFormat(RS("ItemFreight") + RS("Cost")), X + 11500, vbAlignRight, True

      LastSale = RS("SaleNo")
                LastStore = RS("Location")
                SaleTotal = SaleTotal + RS("Cost") + RS("ItemFreight")
                StoreTotal = StoreTotal + RS("Cost") + RS("ItemFreight")

                RS.MoveNext()

                If RS.EOF Then
                    ' At the last record, close the sale and store data displays.
                    If Not DriverPickup Then
                        JuiceAmt = SaleTotal * Juice
                        If JuiceAmt <> 0 Then PrintToPosition OutputObject, "Warehse Chg:              " & Format(JuiceAmt, "###,###.00"), 11500, vbAlignRight, True
          PrintToPosition OO, "Sale Total: " & FormatCurrency(SaleTotal + JuiceAmt), 11500, vbAlignRight, True
          OO.Print
                        OO.FontBold = True
                        OO.FontSize = 14
                        PrintToPosition OO, "Store Total: " & FormatCurrency(StoreTotal + JuiceAmt), 11500, vbAlignRight, True
          OO.FontBold = False
                        SaleTotal = 0
                        StoreTotal = 0
                    End If
                ElseIf (LastSale <> RS("SaleNo") And Not DriverPickup) Or LastStore <> RS("Location") Then

                    If Not DriverPickup Then
                        JuiceAmt = SaleTotal * Juice
                        If JuiceAmt <> 0 Then PrintToPosition OutputObject, "Warehse Chg:              " & Format(JuiceAmt, "###,###.00"), 11500, vbAlignRight, True

          PrintToPosition OO, "Sale Total: $" & Format(SaleTotal, "0.00"), 11500, vbAlignRight, True
          SaleTotal = 0

                        If LastStore <> RS("Location") Then
                            OO.Print
                            OO.FontBold = True
                            OO.FontSize = 14
                            PrintToPosition OO, "Store Total: $" & Format(StoreTotal, "0.00"), 11500, vbAlignRight, True
            OO.FontBold = False
                            StoreTotal = 0
                        ElseIf OO.CurrentY + 5 * OO.TextHeight("X") < OO.ScaleHeight Then
                            OO.Print
                            OO.FontSize = 10
                            OO.FontBold = True
                            PrintToPosition OO, "Sale No: " & RS("SaleNo"), 0, vbAlignLeft, False
            PrintToPosition OO, "Delivery Date: " & RS("DelDate"), 2500, vbAlignLeft, False
            PrintToPosition OO, "Cust: " & Trim(RS("First")) & " " & Trim(RS("Name")), 5000, vbAlignLeft, False
            PrintToPosition OO, "Landed", 11500, vbAlignRight, True
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
                Set RS = Nothing
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

End Module
