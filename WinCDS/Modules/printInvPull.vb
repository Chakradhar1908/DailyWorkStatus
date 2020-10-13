Module printInvPull
    Public Sub printInvPull_PrintRecords(ByVal StoreCount As Integer, ByVal DeliveryDate As String, Optional ByVal PrintAll As Boolean = False) ' Pull List
        '::::printInvPull_PrintRecords
        ':::SUMMARY
        ': Print Pull Load
        ':::DESCRIPTION
        ': Print Inventory Pull Loads
        ':::PARAMETERS
        ': - StoreCount - Indicates the number of Stores.
        ': - DeliveryDate - Indicates the Delivery Date String.
        ': - PrintAll - Used to print all records.

        Dim InvData As CInvRec
        Dim Mail As MailNew, Mail2 As MailNew2
        Dim TFS As Integer, TFN As String, Ty As Integer, TTY As Integer

        Dim Margin As CGrossMargin ': Set Margin = cTable  ' OPTIONAL
        Dim cTa As CDataAccess ': Set cta = cTable.DataAccess()
        Dim SQL As String

        Dim Store As Integer, StartStore As Integer, EndStore As Integer
        Dim previousSaleNo As String
        Dim employeeId As String
        Dim Printing As Boolean

        OutputToPrinter = True

        On Error GoTo HandleErr
        SQL = ""
        SQL = SQL & "SELECT   GrossMargin.* "
        SQL = SQL & "From GrossMargin "
        SQL = SQL & " WHERE   "
        SQL = SQL & " (  (GrossMargin.DelDate=#" & DeliveryDate & "#)  "
        SQL = SQL & "    AND (trim(Status) IN ('ST','SOREC','SSREC','SSLAW','POREC','FND'"

        If ShowPOSO() Then SQL = SQL & ",'PO','SO','SS'"

        'BFH20060731 LAW taken out
        'BFH20060808 LAW added back for F1 (BFH20061012 united too)
        If IsFurnOne() Or IsUFO() Then SQL = SQL & ",'LAW'"

        SQL = SQL & ")"

        SQL = SQL & "    or (trim(Style) IN ('NOTES','STAIN','DEL','LAB') AND trim(Status)=''))"
        SQL = SQL & "    " & IIf(PrintAll, "", " AND (PullPrint IS Null OR PullPrint<>'X')")
        SQL = SQL & " ) "
        SQL = SQL & " ORDER BY GrossMargin.SaleNo, GrossMargin.MarginLine "
        SQL = SQL & ";"
        If StoreCount <= 0 Then
            StartStore = 1
            ' for some reason, endstore was being set to mainmenu.storesld.
            ' should have been the licensed number of stores...
            ' bfh20050628
            EndStore = LicensedNoOfStores()
        Else
            StartStore = StoreCount
            EndStore = StoreCount
        End If

        If IsPuritan() Then
            InvData = New CInvRec
        End If

        For Store = StartStore To EndStore
            Margin = New CGrossMargin
            cTa = Margin.DataAccess

            cTa.DataBase = GetDatabaseAtLocation(Store)
            If cTa.Records_OpenSQL(SQL) Then
                If (cTa.Record_Count <> 0) Then
                    previousSaleNo = "-1"
                    Do While cTa.Records_Available()
                        If previousSaleNo = "-1" Then
                            If Printing Then Printer.NewPage()
                            Print_Header(Store, DeliveryDate, DescribeTimeWindow(Margin.StopStart, Margin.StopEnd))
                            Printing = True
                            PrintOut(FontBold:=False, X:=0, Y:=500)
                        End If

                        If (previousSaleNo <> Trim(Margin.SaleNo)) Then
                            previousSaleNo = Trim(Margin.SaleNo)
                            PrintOut(FontSize:=10, FontBold:=False, BlankLines:=1)
                            Mail_GetAtIndex(Margin.Index, Mail, Store)
                            Mail2_GetAtIndex(Margin.Index, Mail2, Store)

                            employeeId = GetFirstItem(Margin.Salesman)  ' =  .rs("Salesman")
                            Ty = Printer.CurrentY + 300

                            Printer.Print(" Sales Name:", getSalesName(employeeId, Store))
                            Printer.Print(Mail.First, " ", Mail.Last, TAB(50), Mail2.ShipToFirst & " " & Mail2.ShipToLast)
                            Printer.Print(Mail.Address, TAB(50), Mail2.Address2)
                            Printer.Print(Mail.AddAddress)
                            Printer.Print(Mail.City, "    ", Mail.Zip, TAB(50), Mail2.City2)
                            Printer.Print(Mail.Tele, "    ", Mail.Tele2, TAB(50), Mail2.Tele3, TAB(65)) '; "Sales: "; SalesStaff
                            Dim SpLoop As Object
                            For Each SpLoop In Split(WrapLongTextByPrintWidth(Printer, Mail.Special, Printer.ScaleWidth - 100), vbCrLf)
                                Printer.CurrentX = 100
                                Printer.Print(CStr(SpLoop))
                            Next
                            Printer.CurrentY = Printer.CurrentY + 175

                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            '             BFH20050509
                            '             Print Larger SaleNo and Barcode!
                            TTY = Printer.CurrentY
                            Printer.CurrentY = Ty
                            Printer.CurrentX = 9000
                            If SelectBarcodeFont(, , TFN, TFS) Then
                                Printer.Print(PrepareBarcode(EncodeSaleNoBarcode(Margin.SaleNo)))
                            End If
                            Printer.CurrentX = 9000
                            Printer.FontName = TFN
                            Printer.FontSize = 16
                            Printer.Print(Margin.SaleNo)
                            Printer.FontName = TFN
                            Printer.FontSize = TFS
                            Printer.CurrentY = TTY
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        End If

                        If Printer.CurrentY >= 13500 Then ' 14300
                            Printer.NewPage()
                            PrintOut(FontSize:=14, FontBold:=True, DrawWidth:=20, X:=8200, Y:=100, Text:="Del: " & DeliveryDate & "  Page: " & Printer.Page + 1)
                            Printer.FontBold = False
                        End If

                        PrintOut(FontSize:=10, FontBold:=False)

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '           Prints the stylenumber as a barcode
                        Printer.CurrentX = 300
                        SelectBarcodeFont(1, 16, TFN, TFS)
                        Printer.Print(PrepareBarcode(Margin.Style))
                        Printer.FontSize = TFS
                        Printer.FontName = TFN
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        Printer.Print(Margin.PorD, TAB(5), Margin.Quantity, TAB(12), Margin.Style, TAB(34), Margin.Vendor, TAB(54), "Sts: " & Margin.Status, TAB(68), "Loc: " & Margin.Location, TAB(77), "Pulled? ________ By:________ Drop#:______")
                        Printer.Print(TAB(28), Margin.Desc)
                        If IsPuritan() Then
                            ' Puritan wants to print the item's comments..
                            If InvData.Load(Margin.Style, "Style") Then
                                If Trim(InvData.Comments) <> "" Then
                                    Printer.Print(TAB(28), Trim(InvData.Comments))
                                End If
                            End If
                        End If

                        PrintItemLocations(Margin.Style)

                        Margin.PullPrint = "X"
                        Margin.Save()
                    Loop
                End If
                cTa.Records_Close()
            End If

            Margin = Nothing
            Printer.EndDoc()
        Next
        Exit Sub

HandleErr:
        Resume Next
    End Sub

    Private Sub PrintItemLocations(ByVal StyleNo As String)
        Dim S As String, il As clsItemLocation
        Dim TFN As String, TFS As Integer
        Dim C(), N As Integer '@NO-LINT-NTYP
        il = New clsItemLocation

        If Not HasItemLocationTable() Then Exit Sub

        S = "SELECT * FROM [ItemLocation] WHERE StyleNo='" & StyleNo & "' AND Status=" & ItemLocationStatus.ItemLocationStatus_Stocked & " ORDER BY Status ASC, [StockDate] ASC"
        'C = Array(300, 850, 1400, 1950, 2700, 4000, 7000)
        C = New Integer() {300, 850, 1400, 1950, 2700, 4000, 7000}

        il.DataAccess.Records_OpenSQL(S)

        If il.DataAccess.Record_Count > 0 Then
            Printer.FontBold = True
            Printer.FontUnderline = True
            Printer.CurrentX = C(0) : Printer.Print("BLD")
            Printer.CurrentX = C(1) : Printer.Print("ROW")
            Printer.CurrentX = C(2) : Printer.Print("LVL")
            Printer.CurrentX = C(3) : Printer.Print("BAY")
            Printer.CurrentX = C(4) : Printer.Print("Stock Date")
            Printer.CurrentX = C(5) : Printer.Print("Serial No.")
            Printer.CurrentX = C(6) : Printer.Print("Barcoded Serial No.")
            Printer.Print("")
            Printer.FontUnderline = False
            Printer.FontBold = False
        End If

        N = 1
        Do While il.DataAccess.Records_Available
            If N <> 1 Then Printer.Print("") ' extra line between them
            Printer.CurrentX = C(0) : Printer.Print(il.Bld)
            Printer.CurrentX = C(1) : Printer.Print(il.Row)
            Printer.CurrentX = C(2) : Printer.Print(il.Lvl)
            Printer.CurrentX = C(3) : Printer.Print(il.Bay)
            Printer.CurrentX = C(4) : Printer.Print(il.StockDate)
            Printer.CurrentX = C(5) : Printer.Print(il.SerialNo)

            If SelectBarcodeFont(1, 14, TFN, TFS) And Len(Trim(IfNullThenNilString(il.SerialNo))) > 0 Then
                Printer.CurrentX = C(6) : Printer.Print(PrepareBarcode(il.SerialNo))
            Else
                Printer.Print("")
            End If

            Printer.FontSize = TFS
            Printer.FontName = TFN
            N = N + 1
        Loop

        il = Nothing
    End Sub

    Private Function HasItemLocationTable() As Boolean
        HasItemLocationTable = TableExists(0, "ItemLocation")
    End Function

    '  ShowPOSO()
    '  Normally, status of PO and SO are not shown
    '  A customer (FURNITURE ONE) desired them to be shown, so we make
    '  an exception for them.  It is functionalized to allow others to be added easily.
    '  used in printInvPull_PrintRecords (just below)
    '  It is also called from printDeliveryTickets
    Public Function ShowPOSO() As Boolean
        '::::ShowPOSO
        ':::SUMMARY
        ': Configure Pull Load and Delivery Tickets (PO/SO)
        ':::DESCRIPTION
        ': Normally, status of PO and SO are not shown, but when any customer desired them to be shown, so we make an exception for them.
        ':::PARAMETERS
        ':::RETURN
        ': Boolean - Returns True.

        ShowPOSO = False
        If IsFurnOne() Then ShowPOSO = True
    End Function

    Private Sub Print_Header(ByVal Store As Integer, ByVal DeliveryDate As String, ByVal Window As String)
        Dim DelDate As Date
        PrintOut(X:=0, Y:=100, FontName:="Arial", FontSize:=18, FontBold:=True, DrawWidth:=20)
        PrintOut(XCenter:=True, FontBold:=True, Text:="Delivery Pull List")
        PrintOut(FontSize:=8, FontBold:=False, X:=10, Y:=100, Text:="Date: " & DateFormat(Now))
        PrintOut(X:=10, Text:="Time: " & Format(Now, "h:mm:ss am/pm"))
        PrintOut(BlankLines:=2)
        DelDate = DateFormat(DeliveryDate)
        PrintOut(FontSize:=14, FontBold:=True, X:=8500, Y:=100, Text:="Del: " & DelDate & "; " & WeekdayName(DeliveryDate))
        PrintOut(FontSize:=12, FontBold:=False, X:=9250, Y:=400, Text:=Window)
        PrintOut(X:=9250, Y:=700, FontBold:=True, Text:="Store #" & Store)
    End Sub
End Module
