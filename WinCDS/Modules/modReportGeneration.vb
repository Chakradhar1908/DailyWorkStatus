Imports VBRUN
Module modReportGeneration
    ' Report printing module, MJK 20030610
    Public Sub AdvertisingReport(ByRef theDate As String, ByRef ToTheDate As String, ByRef PrintIt As Integer, ByRef StoreNum As Integer, Optional ByRef GroupByZip As Boolean = False, Optional ByRef SortByZip As Boolean = False)
        '::::AdvertisingReport
        ':::SUMMARY
        ': Used to print the Advertising Report.
        ':::DESCRIPTION
        ': This function is used to print the Advertising Report after validating Inputs, preparing output object, gathering the filtered data from database through Sql statement.
        ': This function includes properities defining how to print Advertising Report.
        ': This function is also used to handle errors.
        ':::PARAMETERS
        ': - theDate - Indicates the String, used to validate the Inputs.
        ': - ToTheDate - Indicates the String, used to validate the Inputs.
        ': - PrintIt - Used to prepare Output Object.
        ': - StoreNum - Indicates the Store Number.
        ': - GroupByZip -
        ': - SortByZip - Used to Sort the Date based on Zip.
        ':::RETURN

        Dim RecsPerPage As Integer
        Dim Cy As Integer

        RecsPerPage = 70  ' Save at least 2 lines for totals, just in case.

        ' Validate inputs
        theDate = DateFormat(theDate)
        ToTheDate = DateFormat(ToTheDate)

        If DateAfter(theDate, ToTheDate) Then MessageBox.Show("End date is earlier than start date.  Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Exit Sub

        If StoreNum < 1 Or StoreNum > ssMaxStore Then
            MessageBox.Show("Please select a store.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        ' Prepare output object.
        Dim OutputObject As Object, PageController As Object, Page As Integer, Line As Integer, LastType As String, PrintTotals As Boolean
        If PrintIt <> 0 Then
            OutputObject = Printer
            PageController = Printer
        Else
            OutputObject = frmPrintPreviewDocument.picPicture
            PageController = frmPrintPreviewDocument
            frmPrintPreviewDocument.ReportName = "Advertising Report"
        End If

        ' Output report page header.
        Page = 1
        Line = 1
        AdvertisingReportPageHeader(OutputObject, Page, theDate, ToTheDate, StoreNum)

        ' Gather the data.
        Dim RS As ADODB.Recordset, SQL As String
        Dim TotSold As Decimal, TotDel As Decimal, TotCust As Integer
        Dim PrintLine As Boolean
        Dim LastAdv As Integer, LastCity As String, LastZip As String, LastLoc As Integer, LastCust As Integer
        Dim SubSold As Decimal, SubDel As Decimal
        Dim GrandSold As Decimal, GrandDel As Decimal
        Dim SIP As Boolean, DIP As Boolean, VIP As Boolean

        Dim B As String
        'B = " between #" & Format(theDate, "mm/dd/yyyy") & "# and #" & Format(ToTheDate, "mm/dd/yyyy") & "# "
        B = " between #" & Format(Convert.ToDateTime(theDate), "MM/dd/yyyy") & "# and #" & Format(Convert.ToDateTime(ToTheDate), "MM/dd/yyyy") & "# "

        SQL = ""
        SQL = SQL & "SELECT GrossMargin.Style, GrossMargin.Rn, GrossMargin.SellPrice, "
        SQL = SQL & "Mail.Zip, Mail.City, Mail.CustType, "
        SQL = SQL & "GrossMargin.Location, GrossMargin.Status, GrossMargin.SellDate, "
        SQL = SQL & "GrossMargin.DelDate, GrossMargin.ShipDate, MailIndex "
        SQL = SQL & "FROM GrossMargin LEFT JOIN Mail ON GrossMargin.MailIndex = Mail.[Index] "
        SQL = SQL & "WHERE "
        SQL = SQL & "NOT Trim(Style) IN ('SUB','TAX1','TAX2','PAYMENT','--- Adj ---','STAIN','NOTES','DEL','LAB') "
        SQL = SQL & " AND (CustType <> 1) "   'BFH20080108 No "Get Mail"

        SQL = SQL & "AND "

        SQL = SQL & "("
        SQL = SQL & "SellDate" & B
        SQL = SQL & "OR "
        SQL = SQL & "(Not IsNull(DelDate) AND DelDate" & B & "AND (Status='VDDEL' or left(STATUS,3)='DEL')) "
        SQL = SQL & "OR "
        SQL = SQL & "(Not IsNull(ShipDate) AND ShipDate" & B & ") "
        SQL = SQL & ")"

        '  SQL = SQL & "("
        '  SQL = SQL & "GrossMargin.SellDate Between #" & theDate & "# And #" & ToTheDate & "# "
        '  SQL = SQL & "OR GrossMargin.DelDate Between #" & theDate & "# And #" & ToTheDate & "# "
        '  SQL = SQL & "OR GrossMargin.ShipDate Between #" & theDate & "# And #" & ToTheDate & "# "
        '  SQL = SQL & ") "

        SQL = SQL & "ORDER BY Mail.CustType, " & IIf(SortByZip, "", "Mail.City, ")
        SQL = SQL & "Mail.Zip, GrossMargin.Location, MailIndex"
        ' Loop, checking if CustType, Zip, and maybe City changed.
        ' Advertising type has to be checked differently...
        ' Count changes in MailIndex to get an accurate customer count..
        '    (Location<>0 or Style IN ('STAIN','DEL','LAB','TAX1','TAX2'))  -- Matches Audit report.  But we don't necessarily want to do that.

        'On Error GoTo QueryError
        RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation(StoreNum))
        If Not RS.EOF Then
            LastAdv = IfNullThenZero(RS("CustType").Value)
            LastCity = IfNullThenNilString(RS("City").Value)
            LastZip = IfNullThenNilString(RS("Zip").Value)
            LastLoc = IfNullThenZero(RS("Location").Value) : If LastLoc = 0 Then LastLoc = 1
        End If

        Cy = 800
        Do Until RS.EOF
            ' Check for page wrap
            If Line > RecsPerPage Then
                Page = Page + 1
                Line = 1
                PageController.NewPage
                AdvertisingReportPageHeader(OutputObject, Page, theDate, ToTheDate, StoreNum)
            End If

            ' Add to the line totals..
            SIP = RS("SellDate").Value >= CDate(theDate) And RS("SellDate").Value <= CDate(ToTheDate)
            'If IsNull(RS("DelDate").Value) Then ' Or (RS("Rn") = 0 And Not RS("Style") Like "KIT-*") Then
            If RS("DelDate").Value.ToString = "" Then ' Or (RS("Rn") = 0 And Not RS("Style") Like "KIT-*") Then
                DIP = False
            Else
                DIP = (RS("Status").Value = "VDDEL" Or IsDelivered(RS("Status").Value)) And RS("DelDate").Value >= CDate(theDate) And RS("DelDate").Value <= CDate(ToTheDate)
            End If
            If RS("ShipDate").Value.ToString = "" Then
                VIP = False
            Else
                VIP = RS("ShipDate").Value >= CDate(theDate) And RS("ShipDate").Value <= CDate(ToTheDate)
            End If

            If SIP Then SubSold = SubSold + RS("SellPrice").Value
            If DIP Then SubDel = SubDel + RS("SellPrice").Value
            If VIP Then
                If IsVoid(RS("Status").Value) Or Left(RS("Status").Value, 1) = "x" Then SubSold = SubSold - RS("SellPrice").Value
                If RS("Status").Value = "VDDEL" Then SubDel = SubDel - RS("SellPrice").Value
            End If

            If LastCust <> RS("MailIndex").Value Then
                TotCust = TotCust + 1
                LastCust = RS("MailIndex").Value
            End If

            RS.MoveNext()        ' *** We have to switch records to tell if this line gets printed! ***

            PrintLine = False
            If RS.EOF Then
                PrintTotals = True
                PrintLine = True
            Else
                If IfNullThenZero(RS("CustType").Value) <> LastAdv Then
                    PrintLine = True ' Print if AdvType changes..
                    PrintTotals = True
                End If
                If IfNullThenNilString(RS("Zip").Value) <> LastZip Then PrintLine = True ' Print if Zip changes..
                If IfNullThenZero(RS("Location").Value) <> LastLoc And LastLoc <> 1 And IfNullThenZero(RS("Location").Value) <> 0 Then PrintLine = True ' Print if Location changes..
                If Not GroupByZip And IfNullThenNilString(RS("City").Value) <> LastCity Then PrintLine = True ' Print if City changes, and we're not ignoring City.
            End If

            If PrintLine Then
                ' Print out line data.
                'PrintTo(OutputObject, QueryAdvertisingType(LastAdv, StoreNum), 0, AlignConstants.vbAlignLeft, False) ' Type
                PrintTo(OutputObject, QueryAdvertisingType(LastAdv, StoreNum), 0, AlignConstants.vbAlignLeft, False, Cy) ' Type
                'PrintTo(OutputObject, LastCity, 30, AlignConstants.vbAlignLeft, False) ' City
                PrintTo(OutputObject, LastCity, 30, AlignConstants.vbAlignLeft, False, Cy) ' City
                'PrintTo(OutputObject, LastZip, 60, AlignConstants.vbAlignLeft, False) ' Zip
                PrintTo(OutputObject, LastZip, 60, AlignConstants.vbAlignLeft, False, Cy) ' Zip
                'PrintTo(OutputObject, LastLoc, 75, AlignConstants.vbAlignLeft, False) ' Loc
                PrintTo(OutputObject, LastLoc, 75, AlignConstants.vbAlignLeft, False, Cy) ' Loc
                'PrintTo(OutputObject, Format(SubSold, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False) ' Sold
                PrintTo(OutputObject, Format(SubSold, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False, Cy) ' Sold
                'PrintTo(OutputObject, Format(SubDel, "$###,##0.00"), 105, AlignConstants.vbAlignRight, False) ' Delivered
                PrintTo(OutputObject, Format(SubDel, "$###,##0.00"), 105, AlignConstants.vbAlignRight, False, Cy) ' Delivered
                'PrintTo(OutputObject, TotCust, 107, AlignConstants.vbAlignLeft, True)
                PrintTo(OutputObject, TotCust, 107, AlignConstants.vbAlignLeft, True, Cy)

                TotSold = TotSold + SubSold
                TotDel = TotDel + SubDel
                SubSold = 0
                SubDel = 0
                Line = Line + 1
                If Not RS.EOF Then
                    ' There's more data, which means we switched because it changed.
                    LastAdv = IfNullThenZero(RS("CustType").Value)
                    LastCity = IfNullThenNilString(RS("City").Value)
                    LastZip = IfNullThenNilString(RS("Zip").Value)
                    LastLoc = IfNullThenZero(RS("Location").Value) : If LastLoc = 0 Then LastLoc = 1
                    TotCust = 0
                    LastCust = 0
                End If
                Cy = Cy + 200
            End If

            If PrintTotals Then
                OutputObject.FontBold = True
                'PrintTo(OutputObject, "Totals:", 60, AlignConstants.vbAlignLeft, False)
                PrintTo(OutputObject, "Totals:", 60, AlignConstants.vbAlignLeft, False, Cy)
                'PrintTo(OutputObject, Format(TotSold, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False)
                PrintTo(OutputObject, Format(TotSold, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False, Cy)
                'PrintTo(OutputObject, Format(TotDel, "$###,##0.00"), 105, AlignConstants.vbAlignRight, True)
                PrintTo(OutputObject, Format(TotDel, "$###,##0.00"), 105, AlignConstants.vbAlignRight, True, Cy)
                OutputObject.FontBold = False
                GrandSold = GrandSold + TotSold
                GrandDel = GrandDel + TotDel
                TotSold = 0
                TotDel = 0
                PrintTotals = False
                Line = RecsPerPage + 1  ' Trigger a new page next time.
            End If
        Loop

        RS.Close()
        RS = Nothing

        OutputObject.Print
        OutputObject.FontBold = True
        Cy = OutputObject.CurrentY
        'PrintTo(OutputObject, "Grand Totals:", 60, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "Grand Totals:", 60, AlignConstants.vbAlignLeft, False, Cy)
        'PrintTo(OutputObject, Format(GrandSold, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, Format(GrandSold, "$###,##0.00"), 90, AlignConstants.vbAlignRight, False, Cy)
        'PrintTo(OutputObject, Format(GrandDel, "$###,##0.00"), 105, AlignConstants.vbAlignRight, True)
        PrintTo(OutputObject, Format(GrandDel, "$###,##0.00"), 105, AlignConstants.vbAlignRight, True, Cy)
        OutputObject.FontBold = False

        ' This is handled by DateForm, of all things.
        '  If PrintIt <> 0 Then
        '    OutputObject.EndDoc
        '  Else
        '    frmPrintPreviewDocument.DataEnd
        '  End If
        Exit Sub

QueryError:
        MessageBox.Show("Error generating Advertising Report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Err.Clear()
        If PrintIt <> 0 Then OutputObject.KillDoc
    End Sub

    Private Sub AdvertisingReportPageHeader(ByRef OutputObject As Object, ByRef PageNum As Integer, ByRef StartDate As String, ByRef EndDate As String, ByRef StoreNum As Integer)
        Dim Cy As Integer

        ' Print store header with report title.
        PrintStoreHeader(OutputObject, "Advertising Report")
        PrintTo(OutputObject, "Date Range: " & StartDate, Printer.Width - OutputObject.TextWidth("Date Range: " & StartDate) - 1000, AlignConstants.vbAlignLeft, True)
        PrintTo(OutputObject, "thru " & EndDate, Printer.Width - OutputObject.TextWidth("thru " & EndDate) - 1000, AlignConstants.vbAlignLeft, True)

        ' Print column titles.
        OutputObject.FontBold = True
        OutputObject.FontSize = 8
        OutputObject.CurrentY = 600
        Cy = OutputObject.CurrentY
        'PrintTo(OutputObject, "Type", 0, AlignConstants.vbAlignLeft, False) ' Type
        PrintTo(OutputObject, "Type", 0, AlignConstants.vbAlignLeft, False, Cy) ' Type
        'PrintTo(OutputObject, "City", 30, AlignConstants.vbAlignLeft, False) ' City
        PrintTo(OutputObject, "City", 30, AlignConstants.vbAlignLeft, False, Cy) ' City
        '  PrintTo OutputObject, "State", 50, alignconstants.vbalignleft, False ' State
        'PrintTo(OutputObject, "Zip", 60, AlignConstants.vbAlignLeft, False) ' Zip
        PrintTo(OutputObject, "Zip", 60, AlignConstants.vbAlignLeft, False, Cy)
        'PrintTo(OutputObject, "Loc", 75, AlignConstants.vbAlignLeft, False) ' Loc
        PrintTo(OutputObject, "Loc", 75, AlignConstants.vbAlignLeft, False, Cy) ' Loc
        'PrintTo(OutputObject, "Written", 90, AlignConstants.vbAlignRight, False) ' Sold
        PrintTo(OutputObject, "Written", 90, AlignConstants.vbAlignRight, False, Cy) ' Sold
        'PrintTo(OutputObject, "Delivered", 105, AlignConstants.vbAlignRight, False) ' Delivered
        PrintTo(OutputObject, "Delivered", 105, AlignConstants.vbAlignRight, False, Cy) ' Delivered
        'PrintTo(OutputObject, "Customers", 107, AlignConstants.vbAlignLeft, True)  ' Removed because I can't get the right number.
        PrintTo(OutputObject, "Customers", 107, AlignConstants.vbAlignLeft, True, Cy)  ' Removed because I can't get the right number.
        OutputObject.FontBold = False
    End Sub

    Private Sub PrintStoreHeader(ByRef OutputObject As Object, ByRef Title As String)
        OutputObject.FontName = "Arial"
        OutputObject.DrawWidth = 2
        OutputObject.FontSize = 18
        OutputObject.CurrentY = 100
        OutputObject.FontBold = True
        'PrintTo(OutputObject, Title, 80, AlignConstants.vbAlignTop, True) ' Centered
        PrintTo(OutputObject, Title, 80, AlignConstants.vbAlignTop, True, OutputObject.CurrentY) ' Centered

        OutputObject.FontBold = False
        OutputObject.FontSize = 8
        OutputObject.CurrentY = 100
        'PrintTo(OutputObject, "Date: " & DateFormat(Now), 0, AlignConstants.vbAlignLeft, True)
        PrintTo(OutputObject, "Date: " & DateFormat(Now), 0, AlignConstants.vbAlignLeft, True, OutputObject.CurrentY)
        'PrintTo(OutputObject, "Time: " & Format(Now, "h:mm:ss am/pm"), 0, AlignConstants.vbAlignLeft, True)
        PrintTo(OutputObject, "Time: " & Format(Now, "h:mm:ss tt"), 0, AlignConstants.vbAlignLeft, True, OutputObject.CurrentY)
    End Sub

End Module
