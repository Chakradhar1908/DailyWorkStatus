Module printDeliveryTickets
    Public Sub printDeliveryTickets_PrintRecords(ByVal StoreCount As Integer, ByVal DeliveryDate As String, Optional ByVal PrintAll As Boolean = False, Optional ByRef imgLogo As PictureBox = Nothing, Optional ByVal PrintSingle As String = "") ' Pull List
        '::::printDeliveryTickets_PrintRecords
        ':::SUMMARY
        ': Print Delivery Tickets
        ':::DESCRIPTION
        ': The function will print out a delivery ticket for each sale. It shows most of the same information as the original Bill of Sale, except it shows only the pieces that have not been delivered.
        ': After accessing all the sales records from database through sql statements,sales from different stores locations which are to be delivered will be printed with help of this function.
        ': This function is also used to handle errors.
        ':::PARAMETERS
        ': - StoreCount - Indicates the Number of Stores.
        ': - DeliveryDate - Indicates the Delivery Date String.
        ': - PrintAll - Prints all the sales records which are to be delivered in given date.
        ': - imgLogo - Indicates the Logo Image to display in Delivery Ticket, given nothing.
        ': - PrintSingle - Prints only one individual sale.
        Dim Mail As MailNew, Mail2 As MailNew2
        Dim BalanceDue As String, DelPrint As String, previousSaleNo As String
        Dim cTable As CGrossMargin, cTa As CDataAccess
        Dim SQL As String
        Dim Store As Integer, StartStore As Integer, EndStore As Integer
        Dim NeedHeader As Boolean, LineCount As Integer, PageNum As String, PD As String

        '<CT>
        Dim CY As Integer = 5500
        '</CT>

        cTable = New CGrossMargin

        On Error GoTo HandleErr

        cTa = cTable.DataAccess()
        '        & "    AND (InStr(1,',ST,SOREC,SSREC,SSLAW,POREC,LAW,FND,',',' & trim([Status]) & ',')>0) "

        SQL = ""
        SQL = SQL & "SELECT   GrossMargin.* "
        SQL = SQL & " From GrossMargin "
        SQL = SQL & " WHERE   "
        If PrintSingle <> "" Then  ' BFH20060824
            SQL = SQL & " ( (GrossMargin.SaleNo='" & PrintSingle & "') and iif(isnull(porD),'',pord)<>'' "
        Else
            SQL = SQL & " (  (GrossMargin.DelDate=#" & DeliveryDate & "#)  "
        End If
        SQL = SQL & "    AND (trim(Status) IN ('ST','SOREC','SSREC','SSLAW','POREC','FND'"

        ' bfh20051223 - added 'SS' to the list included by ShowPOSO (in printInvPull.printInvPull_PrintRecords as well)
        '    ,SO,SS,PO  ' These have to be received to be on a Delivery Ticket.
        If printInvPull.ShowPOSO Then SQL = SQL & ",'PO','SO','SS'"

        'BFH20060731 LAW taken out
        'BFH20060808 LAW put back in for F1 (BFH20061012 for united as well)
        If IsFurnOne() Or IsUFO() Then SQL = SQL & ",'LAW'"

        SQL = SQL & ")"

        SQL = SQL & "    or (trim(Style) IN ('NOTES','STAIN','DEL','LAB') AND trim(Status)=''))"
        If PrintSingle = "" Then SQL = SQL & "    " & IIf(PrintAll, "", " AND (DelPrint IS Null OR DelPrint<>'X')")
        SQL = SQL & " ) "
        SQL = SQL & " ORDER BY GrossMargin.SaleNo, GrossMargin.MarginLine "

        If StoreCount <= 0 Then
            StartStore = 1
            EndStore = LicensedNoOfStores() 'StoresSld
        Else
            StartStore = StoreCount
            EndStore = StoreCount
        End If

        For Store = StartStore To EndStore
            With cTa
                .DataBase = GetDatabaseAtLocation(Store)
                If .Records_OpenSQL(SQL) Then
                    If (.Record_Count = 0) Then
                        MessageBox.Show("There are no Delivery Tickets for this period for store #" & Store & ".", "Nothing to print")
                    Else
                        If PrintSingle Then
                            On Error Resume Next
                            DeliveryDate = .RS.Fields("deldate").Value
                        End If
                        previousSaleNo = "-1"
                        Do While .Records_Available()
                            cTable.cDataAccess_GetRecordSet(cTa.RS)
                            If NeedHeader Or (previousSaleNo <> Trim(cTable.SaleNo)) Then
                                If NeedHeader Then PageNum = PageNum + 1 Else PageNum = 1
                                LineCount = 0
                                NeedHeader = False
                                previousSaleNo = Trim(cTable.SaleNo)
                                BalanceDue = HoldNew_GetBalance(cTable.SaleNo, Store)
                                If (.CurrentIndex <> 0) Then Printer.EndDoc()

                                'print the header, with custom logo if available
                                imgLogo.Tag = IIf(modStores.LoadStoreLogo(imgLogo, Store, False), "LOADED", "")
                                PD = "Page: " & PageNum '& " of " & ???
                                If IsIn(GetLeaseNoStatus(cTable.SaleNo, False), "S", "F") Then BalanceDue = "#"
                                Print_Header(Store, DeliveryDate, DescribeTimeWindow(cTable.StopStart, cTable.StopEnd), BalanceDue, imgLogo, PD)
                                imgLogo.Tag = ""

                                PrintOut(FontSize:=14, FontBold:=True, BlankLines:=1, X:=0, Y:=2000)
                                Mail_GetAtIndex(cTable.Index, Mail, Store)
                                Mail2_GetAtIndex(cTable.Index, Mail2, Store)

                                Printer.Print(Mail.First & " " & Mail.Last, TAB(60), "SHIP TO:", TAB(70), Mail2.ShipToFirst & " " & Mail2.ShipToLast)
                                Printer.Print(Mail.Address, TAB(50), Trim(Mail2.Address2))
                                Printer.Print(Mail.AddAddress)
                                Printer.Print(Mail.City & " " & Mail.Zip, TAB(50), Trim(Mail2.City2), " ", Mail2.Zip2)
                                Printer.Print(Mail.Tele, "    ", DressAni(CleanAni(Mail.Tele2)), TAB(50), DressAni(CleanAni(Mail2.Tele3)))
                                Printer.Print()

                                ''------------------------------------------------------
                                ' BFH20050203
                                ' all this stuff to limit mail.special to only 3 lines because we don't have the space!
                                ' bfh20050819 - changed font to not-bold, Arial, 10 and added WrapLongText(...)
                                Printer.Print("Special Instructions:")
                                Dim Lines, Sp, lineCnt As Integer

                                Printer.FontSize = 10
                                Printer.FontBold = True

                                lineCnt = 0
                                If Mail.Special <> "" Then
                                    For Each Sp In Split(WrapLongTextByPrintWidth(Printer, Mail.Special, Printer.ScaleWidth), vbCrLf)
                                        Printer.Print(Sp)
                                        lineCnt = lineCnt + 1
                                        If lineCnt >= 3 Then Exit For
                                    Next
                                End If

                                Printer.FontBold = True
                                Printer.FontSize = 14
                                ''------------------------------------------------------

                                Printer.Print(TAB(5), "Sales: " & " " & TranslateSalesmen(cTable.Salesman, Store))
                                Printer.Print()

                                If cTable.PorD = "D" Then
                                    PrintOut(FontSize:=10, FontBold:=False, X:=885, Y:=200, Text:="X")
                                End If
                                If cTable.PorD = "P" Then
                                    PrintOut(FontSize:=10, FontBold:=False, X:=885, Y:=500, Text:="X")
                                End If

                                PrintOut(FontSize:=18, FontBold:=True, X:=10150, Y:=100, Text:=cTable.SaleNo)
                                PrintOut(FontSize:=10, FontBold:=True, X:=0, Y:=5500)  ' Was 5000

                                '<CT>
                                'PrintToPosition(Printer, "Quan:", 750, VBRUN.AlignConstants.vbAlignRight, False)
                                'PrintToPosition(Printer, "Style:", 900, VBRUN.AlignConstants.vbAlignLeft, False)
                                'PrintToPosition(Printer, "Mfg:", 3000, VBRUN.AlignConstants.vbAlignLeft, False)
                                'PrintToPosition(Printer, "Status:", 5000, VBRUN.AlignConstants.vbAlignLeft, False)
                                'PrintToPosition(Printer, "Loc:", 6500, VBRUN.AlignConstants.vbAlignRight, False)
                                'PrintToPosition(Printer, "Description:", 6750, VBRUN.AlignConstants.vbAlignLeft, True)

                                PrintToPosition2(Printer, "Quan:", 750, VBRUN.AlignConstants.vbAlignRight, False, CY)
                                PrintToPosition2(Printer, "Style:", 900, VBRUN.AlignConstants.vbAlignLeft, False, CY)
                                PrintToPosition2(Printer, "Mfg:", 3000, VBRUN.AlignConstants.vbAlignLeft, False, CY)
                                PrintToPosition2(Printer, "Status:", 5000, VBRUN.AlignConstants.vbAlignLeft, False, CY)
                                PrintToPosition2(Printer, "Loc:", 6500, VBRUN.AlignConstants.vbAlignRight, False, CY)
                                PrintToPosition2(Printer, "Description:", 6750, VBRUN.AlignConstants.vbAlignLeft, True, CY)
                                '</CT>
                                '                Printer.Print "Quan:" _
                                '                  ; Tab(9); "Style:" _
                                '                  ; Tab(29); "Mfg:" _
                                '                  ; Tab(48); "Status:" _
                                '                  ; Tab(60); "Loc:" _
                                '                  ; Tab(67); "Description:"
                            End If

                            PrintOut(FontBold:=False)

                            'bfh20051031
                            If cTable.Style = "NOTES" And IsIn(Left(cTable.Desc, 21), "PRICE WITH TAX BACKED", "ADDITIONAL ADJUSTMENT") Then GoTo SkipLine

                            If cTable.Style = "NOTES" Then Printer.FontItalic = True : Printer.FontBold = True 'BFH20140826 Italics added

                            '<CT>
                            'PrintToPosition(Printer, cTable.Quantity, 750, VBRUN.AlignConstants.vbAlignRight, False)
                            'PrintToPosition(Printer, cTable.Style, 900, VBRUN.AlignConstants.vbAlignLeft, False)
                            'PrintToPosition(Printer, cTable.Vendor, 3000, VBRUN.AlignConstants.vbAlignLeft, False)
                            'PrintToPosition(Printer, cTable.Status, 5000, VBRUN.AlignConstants.vbAlignLeft, False)
                            'PrintToPosition(Printer, cTable.Location, 6500, VBRUN.AlignConstants.vbAlignRight, False)
                            CY = CY + 230
                            PrintToPosition2(Printer, cTable.Quantity, 750, VBRUN.AlignConstants.vbAlignRight, False, CY)
                            PrintToPosition2(Printer, cTable.Style, 900, VBRUN.AlignConstants.vbAlignLeft, False, CY)
                            PrintToPosition2(Printer, cTable.Vendor, 3000, VBRUN.AlignConstants.vbAlignLeft, False, CY)
                            PrintToPosition2(Printer, cTable.Status, 5000, VBRUN.AlignConstants.vbAlignLeft, False, CY)
                            PrintToPosition2(Printer, cTable.Location, 6500, VBRUN.AlignConstants.vbAlignRight, False, CY)
                            '</CT>
                            Printer.FontSize = 8

                            '<CT>
                            'PrintToPosition(Printer, Mid(cTable.Desc, 1, 46), 6750, VBRUN.AlignConstants.vbAlignLeft, False)
                            PrintToPosition2(Printer, Mid(cTable.Desc, 1, 46), 6750, VBRUN.AlignConstants.vbAlignLeft, False, CY)
                            '</CT>
                            Printer.FontItalic = False : Printer.FontBold = False
                            LineCount = LineCount + 1
                            Printer.FontSize = 10
                            '<CT>
                            'Printer.Print()
                            'CY = CY + 220
                            '</CT>
                            If Len(cTable.Desc) > 48 Then
                                Printer.FontSize = 8
                                '<CT>
                                'PrintToPosition(Printer, Mid(cTable.Desc, 47, 46), 6750, VBRUN.AlignConstants.vbAlignLeft, False)
                                CY = CY + 220
                                PrintToPosition2(Printer, Mid(cTable.Desc, 47, 46), 6750, VBRUN.AlignConstants.vbAlignLeft, False, CY)
                                '</CT>
                                Printer.FontSize = 10
                                '<CT>
                                'Printer.Print()
                                CY = CY + 120
                                '</CT>
                                LineCount = LineCount + 1
                            End If
                            'CY = CY + 120
                            If Len(cTable.Desc) > 96 Then
                                Printer.FontSize = 8
                                '<CT>
                                'PrintToPosition(Printer, Mid(cTable.Desc, 93, 46), 6750, VBRUN.AlignConstants.vbAlignLeft, False)
                                CY = CY + 220
                                PrintToPosition2(Printer, Mid(cTable.Desc, 93, 46), 6750, VBRUN.AlignConstants.vbAlignLeft, False, CY)
                                '</CT>
                                Printer.FontSize = 10
                                '<CT>
                                'Printer.Print()
                                CY = CY + 120
                                '</CT>
                                LineCount = LineCount + 1
                            End If

                            If LineCount >= 18 Then LineCount = 0 : NeedHeader = True

                            '              Printer.Print _
                            '                Tab(2); cTable.Quantity; _
                            '                Tab(7); cTable.Style; _
                            '                Tab(31); cTable.Vendor; _
                            '                Tab(54); cTable.Status; _
                            '                Tab(64); cTable.Location; _
                            '                Tab(68); cTable.Desc
                            Printer.Print()
SkipLine:
                            cTable.DelPrint = "X"
                            cTable.Save()
                        Loop
                        Printer.EndDoc()
                    End If
                    .Records_Close()
                End If
            End With
        Next
        Printer.EndDoc()

        DisposeDA(cTable)
        Exit Sub

HandleErr:
        Resume Next
    End Sub

    Public Function printDeliveryTicketHTML(ByRef SaleNo As String, Optional ByRef StoreNo As Integer = 0, Optional ByRef imgLogo As PictureBox = Nothing) As String
        '::::printDeliveryTicketHTML
        ':::SUMMARY
        ':HTML Version of delivery ticket
        ':::DESCRIPTION
        ':Creates an HTML file for Delivery Ticket.
        ': - Often used for email
        ':
        ':::PARAMETERS
        ':- SaleNo - Sale number to generate html file for Delivery ticket.
        ':- StoreNo - The Store number of the Delivery ticket.
        ':- imgLogo - Logo to display in the page.
        ':::RETURN
        ':  String

        Dim Mail As MailNew, Mail2 As MailNew2
        Dim cTable As CGrossMargin, cTa As CDataAccess
        Dim SQL As String, S As String, BalanceDue As String
        Dim ImW As Integer, ImH As Integer
        Dim DeliveryDate As Object
        Dim RS As ADODB.Recordset

        cTable = New CGrossMargin
        cTa = cTable.DataAccess()

        If StoreNo = 0 Then StoreNo = StoresSld

        SQL = "SELECT * FROM GrossMargin WHERE SaleNo ='" & ProtectSQL(SaleNo) & "' AND Store =" & StoreNo
        If getRecordsetCountBySQL(SQL) = 0 Then
            MessageBox.Show("SaleNo and StoreNo combination does not exist")
            Exit Function
        End If

        SQL = "SELECT * From GrossMargin WHERE (SaleNo ='" & ProtectSQL(SaleNo) & "')"
        SQL = SQL & " AND (trim(Status) IN ('ST','SOREC','SSREC','SSLAW','POREC','FND'"
        If printInvPull.ShowPOSO Then SQL = SQL & ",'PO','SO','SS'"
        If IsFurnOne() Or IsUFO() Then SQL = SQL & ",'LAW'"
        SQL = SQL & ")"
        SQL = SQL & " or (trim(Style) IN ('NOTES','STAIN','DEL','LAB') AND trim(Status)='' ))"
        SQL = SQL & " ORDER BY SaleNo, MarginLine"

        RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(StoreNo), True)
        DeliveryDate = IfNullThenNullDate(RS("deldate"))
        DeliveryDate = DateFormat(DeliveryDate)
        BalanceDue = HoldNew_GetBalance(RS("SaleNo").Value, StoreNo)

        If Not imgLogo Is Nothing Then
            imgLogo.Tag = IIf(modStores.LoadStoreLogo(imgLogo, StoreNo, False), "LOADED", "")
        End If

        If IsIn(GetLeaseNoStatus(SaleNo, False), "S", "F") Then BalanceDue = "#"

        S = S & "" & vbCrLf
        S = S & "<html>" & vbCrLf
        S = S & "<head>" & vbCrLf
        S = S & "<title>Order #" & SaleNo & " - " & StoreSettings(StoreNum).Name & "</title>" & vbCrLf
        S = S & "</head>" & vbCrLf
        S = S & "<body>" & vbCrLf

        S = S & "<table align=center>" & vbCrLf
        S = S & "<tr valign=top>"
        If Not imgLogo Is Nothing Then
            If imgLogo.Tag = "LOADED" Then
                ImW = 500
                ImH = 10
                'S = S & "<td width=500 height=10><img src=" & imgLogo.Picture & " width=" & ImW & " height=" & ImH & "</img></td>"
                S = S & "<td width=500 height=10><img src=" & imgLogo.Image.ToString & " width=" & ImW & " height=" & ImH & "</img></td>"
            Else
                S = S & "<td width=500 height=10></td>"
            End If
        Else
            S = S & "<td width=500 height=10></td>"
        End If
        S = S & "</tr></table>" & vbCrLf

        S = S & "<table width=42% align=center>" & vbCrLf
        S = S & "<tr valign=top>"
        S = S & "<td valign=top><span style='padding-RIGHT: 10px;'>Delivery:</span><font size=+1><b>" & DeliveryDate & "</b></font><br><span style='padding-RIGHT: 12px;'>Pick Up:</span><font size=+1><b>" & WeekdayName(DeliveryDate) & "</b></font></td>"

        S = S & "<td align=center><font size=+3><b>" & StoreSettings(StoreNum).Name & "</b></font><br>" & StoreSettings(StoreNum).Address & "<br>" & StoreSettings(StoreNum).City & "<br>" & StoreSettings(StoreNum).Phone & "</td>"
        S = S & "<td><font size=+1><b><span style='padding-LEFT: 20px;'>" & SaleNo & "</span></b></font><br><span style='padding-LEFT: 20px;'>SaleNo:</span></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "</tr>" & vbCrLf

        Mail_GetAtIndex(RS("MailIndex").Value, Mail, StoreNo)
        Mail2_GetAtIndex(RS("MailIndex").Value, Mail2, StoreNo)

        S = S & "<tr valign=top>" & vbCrLf
        S = S & "<td><font size=+1><b>" & Mail.First & Space(1) & Mail.Last & "<br><font size=+1><b>" & Mail.Address & "</b></font></td>"
        S = S & "<td align=center><font size=+1><b><span style='padding-LEFT: 30px;'>SHIP TO:</span></b></font></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td><font size=+1><b>" & Mail.City & Space(1) & Mail.Zip & Space(1) & Trim(Mail2.City2) & Space(1) & Mail2.Zip2 & "</b></font><br><font size=+1><b>" & Mail.Tele & Space(10) & DressAni(CleanAni(Mail.Tele2)) & Space(10) & DressAni(CleanAni(Mail2.Tele3)) & "</b></font></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "<tr>" & vbCrLf
        S = S & "<td colspan=2><font size=+1><b>Special Instructions: " & "</b>" & Mail.Special & "</font><br><font size=+1><b><span style='padding-LEFT: 30px;'>Sales:</span>" & TranslateSalesmen(RS("Salesman"), StoreNo) & "</b></font></td>"
        S = S & "</tr>" & vbCrLf

        S = S & "</table>" & vbCrLf
        S = S & "<br><br>" & vbCrLf

        S = S & "<table width=41% align=center>" & vbCrLf
        S = S & "<tr>" & vbCrLf

        S = S & "<td width=40><font size=-1><b>Quan:</b></font></td>"
        S = S & "<td width=88><font size=-1><b><span style='padding-LEFT: 10px;'>Style:</span></b></font></td>"
        S = S & "<td width=148><font size=-1><b><span style='padding-LEFT: 8px;'>Mfg:</span></b></font></td>"
        S = S & "<td width=48><font size=-1><b>Status:</b></font></td>"
        S = S & "<td width=30><font size=-1><b>Loc:</b></font></td>"
        S = S & "<td><font size=-1><b>Description:</b></font></td>"
        S = S & "</tr>" & vbCrLf

        Do While Not RS.EOF
            S = S & "<tr>" & vbCrLf
            S = S & "<td><font size=-1><span style='padding-LEFT: 24px;'>" & IfNullThenZero(RS("Quantity")) & "</span></font></td>"
            S = S & "<td><font size=-1><span style='padding-LEFT: 10px;'>" & IfNullThenNilString(RS("Style")) & "</font></td>"
            S = S & "<td><font size=-1><span style='padding-LEFT: 10px;'>" & IfNullThenNilString(RS("Vendor")) & "</span></font></td>"
            S = S & "<td><font size=-1>" & IfNullThenNilString(RS("Status")) & "</font></td>"
            S = S & "<td><font size=-1><span style='padding-LEFT: 12px;'>" & IfNullThenZero(RS("Location")) & "</span></font></td>"

            If RS("Style").Value = "NOTES" Then
                S = S & "<td nowrap><font size=-1><b><i>" & IfNullThenNilString(RS("Desc")) & "</i></b></font></td></tr>"
            Else
                S = S & "<td nowrap><font size=-1>" & IfNullThenNilString(RS("Desc")) & "</font></td></tr>"
            End If
            RS.MoveNext()
        Loop
        S = S & "</table>" & vbCrLf

        S = S & "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>" & vbCrLf
        S = S & "<table align=center width=45%>" & vbCrLf
        S = S & "<tr valign=top>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td></td>"
        S = S & "<td width=480>"
        S = S & "<div style='border: 3px solid black;'>"
        S = S & "<font size=-2><b><span style='padding-LEFT: 10px;'>All Items Received in good condition!</span></b></font>"
        S = S & "<font size=-1><span style='padding-LEFT: 140px;'>Rec. Date:</span></font><br><br><br><br>"
        S = S & "</td>"
        S = S & "<td></td><td></td><td></td><td></td><td></td><td></td>"
        S = S & "<td width=182 align=center><div style='border: 3px solid black;'><font size=-1>Balance Due:</font><br><font size=+1><b>" & BalanceDue & "</b></font><br><br><br></td>"
        S = S & "</tr>"
        S = S & "</table>" & vbCrLf

        S = S & "</body>" & vbCrLf
        S = S & "</html>" & vbCrLf

        WriteFile(LocalDesktopFolder() & "Delticket.html", S, True)
        printDeliveryTicketHTML = S
    End Function

    Private Sub Print_Header(ByVal Store As Integer, ByVal DeliveryDate As String, ByVal Window As String, ByVal BalanceDue As String, Optional ByVal imgLogo As PictureBox = Nothing, Optional ByVal PageDescriptor As String = "")
        Dim PrintedLogo As Boolean, ImW As Integer, ImH As Integer
        On Error Resume Next ' BFH20051223 - FurnOne's-Store 3 seemed to be failing in this Sub
        Err.Clear()
        PrintedLogo = False

        PrintOut(FontName:="Arial", FontSize:=18, FontBold:=True, DrawWidth:=20, X:=0, Y:=100)

        If Not imgLogo Is Nothing Then
            If imgLogo.Tag = "LOADED" Then
                ImW = 6000
                ImH = 2000
                'Printer.PaintPicture(imgLogo.Image, Printer.Width / 2 - ImW / 2, 75, ImW, ImH) : PrintedLogo = True
                Printer.PaintPicture(imgLogo.Image, 4000, 200, 5000, 5000, 1200, 1000, 35000, 35000) : PrintedLogo = True
            End If
        End If

        'If Not PrintedLogo Or Err.Number <> 0 Then PrintCompanyInformation(Store)
        If Not PrintedLogo Or Err.Number <> 0 Then PrintCompanyInformation(Store)

        '  Printer.Print
        PrintOut(FontSize:=10, FontBold:=False, X:=0, Y:=200, Text:="Delivery:")
        PrintOut(FontSize:=10, FontBold:=False, X:=0, Y:=500, Text:=" Pick Up:")

        PrintOut(FontSize:=14, FontBold:=True, X:=1200, Y:=200, Text:=DeliveryDate)
        PrintOut(FontSize:=12, FontBold:=True, X:=600, Y:=800, Text:=Window)

        PrintOut(FontSize:=14, FontBold:=True, X:=1200, Y:=500, Text:=WeekdayName(DeliveryDate))
        PrintOut(FontSize:=10, FontBold:=False, X:=10000, Y:=500, Text:="    Sale No:")

        If IsUFO() Then
            PrintOut(FontSize:=10, X:=200, Y:=13800)
            Printer.Print(" I accept UFO Furniture Warehouse Polices")
        End If

        'Printer.Line(1000, 13500)-Step(7500, 1100), QBColor(0), B
        Printer.Line(1000, 13500, 8500, 14700, QBColor(0), True)
        If BalanceDue <> "#" Then
            'Printer.Line(9000, 13500)-Step(2400, 1100), QBColor(0), B
            Printer.Line(9000, 13500, 11500, 14700, QBColor(0), True)
        End If
        'PrintOut(FontSize:=8, X:=1250, Y:=13550)
        PrintOut(FontSize:=8, X:=1250, Y:=13400)
        '  If  IsDevelopment Then
        If False Then
            MainMenu.rtbn.FileRead(False, DeliveryTicketMessageFile)
            Dim SS As String
            SS = MainMenu.rtbn.RichTextBox.Text
            Printer.Print(SS)
        Else
            MainMenu.rtbn.DoPrintFile(DeliveryTicketMessageFile, -1, -1, 7100, 1100, True, False)
            '<CT>
            Printer.FontBold = True
            Printer.FontSize = 10
            Printer.CurrentX = 1250
            Printer.CurrentY = 13580
            Printer.Print(DeliveryticketMessageFileText)
            '</CT>
        End If
        '  Else   ' old way, dev on
        '    Printer.Print " Received in good condition! ";
        '  End If
        If BalanceDue <> "#" Then
            '<CT>
            Printer.FontBold = False
            Printer.FontSize = 8
            Printer.CurrentX = 7500
            Printer.CurrentY = 13580
            'Printer.Print(TAB(170), "Rec. Date:", TAB(208), " Balance Due: ")
            Printer.Print("Rec. Date:", TAB(40), " Balance Due: ")
            '</CT>
        End If

        If PageDescriptor <> "" Then
            PrintOut(FontSize:=8, X:=200, Y:=14900)
            Printer.Print(PageDescriptor)
        End If

        PrintOut(FontSize:=20, FontBold:=True, X:=9000, Y:=13800)
        ' bfh20050919 - attempted to fix a new alignment problem iun delivery tickets..
        '  PrintToPosition Printer, BalanceDue, 10200, vbAlignTop, False
        If BalanceDue <> "#" Then
            PrintOut(9300, 13800, BalanceDue, , , True, 20)
        End If
        'Printer.Print BalanceDue
    End Sub

    Private Sub PrintCompanyInformation(ByVal Store As Integer)
        On Error Resume Next
        With StoreSettings(Store)
            PrintOut(FontName:="Arial", FontSize:=18, FontBold:=True, DrawWidth:=20, X:=0)
            PrintOut(XCenter:=True, Text:= .Name)
            PrintOut(XCenter:=True, Text:= .Address)
            PrintOut(XCenter:=True, Text:= .City)
            PrintOut(XCenter:=True, Text:= .Phone)
        End With
    End Sub
End Module
