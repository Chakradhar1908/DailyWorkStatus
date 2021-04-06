Module modReportGeneration
    ' Report printing module, MJK 20030610

    Public Sub AdvertisingReport(ByRef theDate As String, ByRef ToTheDate As String, ByRef PrintIt As Long, ByRef StoreNum As Long, Optional ByRef GroupByZip As Boolean = False, Optional ByRef SortByZip As Boolean = False)
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

        Dim RecsPerPage As Long
        RecsPerPage = 70  ' Save at least 2 lines for totals, just in case.

        ' Validate inputs
        theDate = DateFormat(theDate)
        ToTheDate = DateFormat(ToTheDate)

        If DateAfter(theDate, ToTheDate) Then MsgBox "End date is earlier than start date.  Please try again.", vbCritical, "Error": Exit Sub

        If StoreNum < 1 Or StoreNum > ssMaxStore Then
            MsgBox "Please select a store.", vbExclamation
    Exit Sub
        End If

        ' Prepare output object.
        Dim OutputObject As Object, PageController As Object, Page As Long, Line As Long, LastType As String, PrintTotals As Boolean
        If PrintIt <> 0 Then
    Set OutputObject = Printer
    Set PageController = Printer
  Else
    Set OutputObject = frmPrintPreviewDocument.picPicture
    Set PageController = frmPrintPreviewDocument
    frmPrintPreviewDocument.ReportName = "Advertising Report"
        End If

        ' Output report page header.
        Page = 1
        Line = 1
        AdvertisingReportPageHeader OutputObject, Page, theDate, ToTheDate, StoreNum

  ' Gather the data.
        Dim RS As ADODB.Recordset, SQL As String
        Dim TotSold As Currency, TotDel As Currency, TotCust As Long
        Dim PrintLine As Boolean
        Dim LastAdv As Long, LastCity As String, LastZip As String, LastLoc As Long, LastCust As Long
        Dim SubSold As Currency, SubDel As Currency
        Dim GrandSold As Currency, GrandDel As Currency
        Dim SIP As Boolean, DIP As Boolean, VIP As Boolean

        Dim B As String
        B = " between #" & Format(theDate, "mm/dd/yyyy") & "# and #" & Format(ToTheDate, "mm/dd/yyyy") & "# "

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
  Set RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation(StoreNum))
  If Not RS.EOF Then
            LastAdv = IfNullThenZero(RS("CustType"))
            LastCity = IfNullThenNilString(RS("City"))
            LastZip = IfNullThenNilString(RS("Zip"))
            LastLoc = IfNullThenZero(RS("Location")) : If LastLoc = 0 Then LastLoc = 1
        End If
        Do Until RS.EOF
            ' Check for page wrap
            If Line > RecsPerPage Then
                Page = Page + 1
                Line = 1
                PageController.NewPage
                AdvertisingReportPageHeader OutputObject, Page, theDate, ToTheDate, StoreNum
    End If

            ' Add to the line totals..
            SIP = RS("SellDate") >= CDate(theDate) And RS("SellDate") <= CDate(ToTheDate)
            If IsNull(RS("DelDate")) Then ' Or (RS("Rn") = 0 And Not RS("Style") Like "KIT-*") Then
                DIP = False
            Else
                DIP = (RS("Status") = "VDDEL" Or IsDelivered(RS("Status"))) And RS("DelDate") >= CDate(theDate) And RS("DelDate") <= CDate(ToTheDate)
            End If
            If IsNull(RS("ShipDate")) Then
                VIP = False
            Else
                VIP = RS("ShipDate") >= CDate(theDate) And RS("ShipDate") <= CDate(ToTheDate)
            End If

            If SIP Then SubSold = SubSold + RS("SellPrice")
            If DIP Then SubDel = SubDel + RS("SellPrice")
            If VIP Then
                If IsVoid(RS("Status")) Or Left(RS("Status"), 1) = "x" Then SubSold = SubSold - RS("SellPrice")
                If RS("Status") = "VDDEL" Then SubDel = SubDel - RS("SellPrice")
            End If

            If LastCust <> RS("MailIndex") Then
                TotCust = TotCust + 1
                LastCust = RS("MailIndex")
            End If

            RS.MoveNext()        ' *** We have to switch records to tell if this line gets printed! ***

            PrintLine = False
            If RS.EOF Then
                PrintTotals = True
                PrintLine = True
            Else
                If IfNullThenZero(RS("CustType")) <> LastAdv Then
                    PrintLine = True ' Print if AdvType changes..
                    PrintTotals = True
                End If
                If IfNullThenNilString(RS("Zip")) <> LastZip Then PrintLine = True ' Print if Zip changes..
                If IfNullThenZero(RS("Location")) <> LastLoc And LastLoc <> 1 And IfNullThenZero(RS("Location")) <> 0 Then PrintLine = True ' Print if Location changes..
                If Not GroupByZip And IfNullThenNilString(RS("City")) <> LastCity Then PrintLine = True ' Print if City changes, and we're not ignoring City.
            End If

            If PrintLine Then
                ' Print out line data.
                PrintTo OutputObject, QueryAdvertisingType(LastAdv, StoreNum), 0, vbAlignLeft, False ' Type
                PrintTo OutputObject, LastCity, 30, vbAlignLeft, False ' City
                PrintTo OutputObject, LastZip, 60, vbAlignLeft, False ' Zip
                PrintTo OutputObject, LastLoc, 75, vbAlignLeft, False ' Loc
                PrintTo OutputObject, Format(SubSold, "$###,##0.00"), 90, vbAlignRight, False ' Sold
                PrintTo OutputObject, Format(SubDel, "$###,##0.00"), 105, vbAlignRight, False ' Delivered
                PrintTo OutputObject, TotCust, 107, vbAlignLeft, True

      TotSold = TotSold + SubSold
                TotDel = TotDel + SubDel
                SubSold = 0
                SubDel = 0
                Line = Line + 1
                If Not RS.EOF Then
                    ' There's more data, which means we switched because it changed.
                    LastAdv = IfNullThenZero(RS("CustType"))
                    LastCity = IfNullThenNilString(RS("City"))
                    LastZip = IfNullThenNilString(RS("Zip"))
                    LastLoc = IfNullThenZero(RS("Location")) : If LastLoc = 0 Then LastLoc = 1
                    TotCust = 0
                    LastCust = 0
                End If
            End If

            If PrintTotals Then
                OutputObject.FontBold = True
                PrintTo OutputObject, "Totals:", 60, vbAlignLeft, False
      PrintTo OutputObject, Format(TotSold, "$###,##0.00"), 90, vbAlignRight, False
      PrintTo OutputObject, Format(TotDel, "$###,##0.00"), 105, vbAlignRight, True
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
          Set RS = Nothing
  
  OutputObject.Print
        OutputObject.FontBold = True
        PrintTo OutputObject, "Grand Totals:", 60, vbAlignLeft, False
  PrintTo OutputObject, Format(GrandSold, "$###,##0.00"), 90, vbAlignRight, False
  PrintTo OutputObject, Format(GrandDel, "$###,##0.00"), 105, vbAlignRight, True
  OutputObject.FontBold = False

        ' This is handled by DateForm, of all things.
        '  If PrintIt <> 0 Then
        '    OutputObject.EndDoc
        '  Else
        '    frmPrintPreviewDocument.DataEnd
        '  End If
        Exit Sub

QueryError:
        MsgBox "Error generating Advertising Report.", vbCritical, "Error"
  Err.Clear()
        If PrintIt <> 0 Then OutputObject.KillDoc
    End Sub

End Module
