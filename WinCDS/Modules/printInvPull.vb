Module printInvPull
    Public Sub printInvPull_PrintRecords(ByVal StoreCount As Long, ByVal DeliveryDate As String, Optional ByVal PrintAll As Boolean = False) ' Pull List
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
        Dim TFS As Long, TFN As String, Ty As Long, TTY As Long

        Dim Margin As CGrossMargin ': Set Margin = cTable  ' OPTIONAL
        Dim cTa As CDataAccess ': Set cta = cTable.DataAccess()
        Dim SQL As String

        Dim Store As Long, StartStore As Long, EndStore As Long
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

        If ShowPOSO Then SQL = SQL & ",'PO','SO','SS'"

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
    Set InvData = New CInvRec
  End If

        For Store = StartStore To EndStore
    Set Margin = New CGrossMargin
    Set cTa = Margin.DataAccess
      
    cTa.DataBase = GetDatabaseAtLocation(Store)
            If cTa.Records_OpenSQL(SQL) Then
                If (cTa.Record_Count <> 0) Then
                    previousSaleNo = "-1"
                    Do While cTa.Records_Available()
                        If previousSaleNo = "-1" Then
                            If Printing Then Printer.NewPage()
                            Print_Header Store, DeliveryDate, DescribeTimeWindow(Margin.StopStart, Margin.StopEnd)
            Printing = True
                            PrintOut FontBold:=False, X:=0, Y:=500
          End If

                        If (previousSaleNo <> Trim(Margin.SaleNo)) Then
                            previousSaleNo = Trim(Margin.SaleNo)
                            PrintOut FontSize:=10, FontBold:=False, BlankLines:=1
            Mail_GetAtIndex Margin.Index, Mail, Store
            Mail2_GetAtIndex Margin.Index, Mail2, Store

            employeeId = GetFirstItem(Margin.Salesman)  ' =  .rs("Salesman")

                            Ty = Printer.CurrentY + 300

                            Printer.Print " Sales Name:"; getSalesName(employeeId, Store)
            Printer.Print Mail.First; " "; Mail.Last; Tab(50); Mail2.ShipToFirst & " " & Mail2.ShipToLast
            Printer.Print Mail.Address; Tab(50); Mail2.Address2
            Printer.Print Mail.AddAddress
            Printer.Print Mail.City; "    "; Mail.Zip; Tab(50); Mail2.City2
            Printer.Print Mail.Tele; "    "; Mail.Tele2; Tab(50); Mail2.Tele3; Tab(65) '; "Sales: "; SalesStaff
            Dim SpLoop As Variant
                            For Each SpLoop In Split(WrapLongTextByPrintWidth(Printer, Mail.Special, Printer.ScaleWidth - 100), vbCrLf)
                                Printer.CurrentX = 100
                                Printer.Print CStr(SpLoop)
            Next
                            Printer.CurrentY = Printer.CurrentY + 175

                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            '             BFH20050509
                            '             Print Larger SaleNo and Barcode!
                            TTY = Printer.CurrentY
                            Printer.CurrentY = Ty
                            Printer.CurrentX = 9000
                            If SelectBarcodeFont(, , TFN, TFS) Then
                                Printer.Print PrepareBarcode(EncodeSaleNoBarcode(Margin.SaleNo))
            End If
                            Printer.CurrentX = 9000
                            Printer.FontName = TFN
                            Printer.FontSize = 16
                            Printer.Print Margin.SaleNo
            Printer.FontName = TFN
                            Printer.FontSize = TFS
                            Printer.CurrentY = TTY
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        End If

                        If Printer.CurrentY >= 13500 Then ' 14300
                            Printer.NewPage()
                            PrintOut _
               FontSize:=14, FontBold:=True _
              , DrawWidth:=20 _
              , X:=8200, Y:=100 _
              , Text:="Del: " & DeliveryDate & "  Page: " & Printer.Page + 1
            Printer.FontBold = False
                        End If

                        PrintOut FontSize:=10, FontBold:=False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '           Prints the stylenumber as a barcode
                        Printer.CurrentX = 300
                        SelectBarcodeFont 1, 16, TFN, TFS
          Printer.Print PrepareBarcode(Margin.Style)
          Printer.FontSize = TFS
                        Printer.FontName = TFN
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        Printer.Print _
            Margin.PorD _
            ; Tab(5); Margin.Quantity _
            ; Tab(12); Margin.Style _
            ; Tab(34); Margin.Vendor _
            ; Tab(54); "Sts: " & Margin.Status _
            ; Tab(68); "Loc: " & Margin.Location _
            ; Tab(77); "Pulled? ________ By:________ Drop#:______"
          Printer.Print TAB(28); Margin.Desc
          If IsPuritan() Then
                            ' Puritan wants to print the item's comments..
                            If InvData.Load(Margin.Style, "Style") Then
                                If Trim(InvData.Comments) <> "" Then
                                    Printer.Print TAB(28); Trim(InvData.Comments)
              End If
                            End If
                        End If

                        PrintItemLocations Margin.Style

          Margin.PullPrint = "X"
                        Margin.Save()
                    Loop
                End If
                cTa.Records_Close()
            End If
    
    Set Margin = Nothing
    Printer.EndDoc()
        Next
        Exit Sub

HandleErr:
        Resume Next
    End Sub

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

End Module
