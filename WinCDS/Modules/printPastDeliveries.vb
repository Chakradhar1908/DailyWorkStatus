Module printPastDeliveries
    Public Sub printPastDeliveries_PrintRecords(ByVal StoreNo As Integer, ByVal FromDate As Date, ByVal toDate As Date, ByVal ToPrinter As Boolean)
        Dim OO As Object
        Dim SQL As String
        Dim RS As ADODB.Recordset
        Dim PageNo As Integer

        If ToPrinter Then
            OO = Printer
        Else
            OutputObject = frmPrintPreviewDocument.picPicture
            OO = OutputObject
            frmPrintPreviewDocument.CallingForm = InvPull
        End If

        SQL = "SELECT gm.DelDate, gm.SaleNo, m.Last as Name, m.City, gm.Quantity, gm.Style, gm.Desc, gm.SellPrice "
        SQL = SQL & "FROM GrossMargin gm INNER JOIN Mail m "
        SQL = SQL & "on gm.MailIndex = m.Index "
        SQL = SQL & "WHERE gm.Store =" & StoreNo
        SQL = SQL & "AND gm.PorD = 'D'"
        SQL = SQL & "AND gm.DelDate >=#" & FromDate & "# AND gm.DelDate <=#" & toDate & "#"

        RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(StoreNo))

        PageNo = 1
        Do Until RS.EOF
            If (OO.CurrentY + 2 * OO.TextHeight("X") > Printer.ScaleHeight) Then
                If ToPrinter Then
                    OO.NewPage
                Else
                    frmPrintPreviewDocument.NewPage()
                    PageNo = PageNo + 1
                End If
            End If

            If OO.CurrentY = 0 Then    ' New page, print headers.
                PrintReportHeader("Past Deliveries", StoreSettings(StoreNo).Name, StoreSettings(StoreNo).Address & " " & StoreSettings(StoreNo).Phone, FromDate, toDate, OO, PageNo)
                OO.Print("")
                OO.FontSize = 10
                OO.FontBold = True
                ReportColumnHeading(OO)
            End If

            ShowReportData(RS, OO)
            RS.MoveNext()
        Loop

        If ToPrinter Then
            Printer.EndDoc()
        Else
            InvPull.Hide()
            frmPrintPreviewDocument.DataEnd()
        End If
    End Sub

    Private Sub PrintReportHeader(ByVal RptTitle As String, ByVal StoreName As String, ByVal StoreDetails As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal OutObj As Object, ByVal PageNo As Integer)
        PrintOut(OutObj:=OutObj, FontName:="Arial", FontSize:=10, FontBold:=True, DrawWidth:=20, X:=200, Y:=100)
        OutObj.Print("Date From: ", StartDate, TAB(95), "Date: ", Now)
        PrintOut(OutObj:=OutObj, FontSize:=20, FontBold:=True, X:=200, Y:=100)
        OutObj.CurrentX = (Printer.ScaleWidth - OutObj.TextWidth(RptTitle)) / 2
        OutObj.Print(RptTitle)

        PrintOut(OutObj:=OutObj, FontName:="Arial", FontSize:=10, FontBold:=True, DrawWidth:=20, X:=200, Y:=300)
        OutObj.Print("Date To: ", EndDate, TAB(95), "Page No.: ", PageNo)

        OutObj.FontSize = 15
        OutObj.CurrentX = (Printer.ScaleWidth - OutObj.TextWidth(StoreName)) / 2
        OutObj.Print(StoreName)

        OutObj.FontSize = 10
        OutObj.CurrentX = (Printer.ScaleWidth - OutObj.TextWidth(StoreDetails)) / 2
        OutObj.Print(StoreDetails)
    End Sub

    Private Sub ReportColumnHeading(ByVal OO As Object)
        PrintToPosition(OO, "DelDate", 200, VBRUN.AlignConstants.vbAlignLeft, False)
        PrintToPosition(OO, "SaleNo", 1400, VBRUN.AlignConstants.vbAlignLeft, False)
        PrintToPosition(OO, "Name", 2300, VBRUN.AlignConstants.vbAlignLeft, False)
        PrintToPosition(OO, "Qty", 3600, VBRUN.AlignConstants.vbAlignLeft, False)
        PrintToPosition(OO, "Style", 4000, VBRUN.AlignConstants.vbAlignLeft, False)
        PrintToPosition(OO, "Description", 5800, VBRUN.AlignConstants.vbAlignLeft, False)
        PrintToPosition(OO, "Sale Price", 10000, VBRUN.AlignConstants.vbAlignLeft, True)
    End Sub

    Private Sub ShowReportData(ByVal RS As ADODB.Recordset, ByVal OO As Object)
        If IsItem(RS("Style").Value) = True Then
            PrintToPosition(OO, RS("DelDate").Value, 200, VBRUN.AlignConstants.vbAlignLeft, False)
            PrintToPosition(OO, RS("SaleNo").Value, 1400, VBRUN.AlignConstants.vbAlignLeft, False)
            PrintToPosition(OO, RS("Name").Value, 2300, VBRUN.AlignConstants.vbAlignLeft, False)
            PrintToPosition(OO, RS("Quantity").Value, 3800, VBRUN.AlignConstants.vbAlignRight, False)
            PrintToPosition(OO, RS("Style").Value, 4000, VBRUN.AlignConstants.vbAlignLeft, False)

            If Len(RS("Desc").Value) > 35 Then
                PrintToPosition(OO, Left(RS("Desc").Value, 35), 5800, VBRUN.AlignConstants.vbAlignLeft, False)
            Else
                PrintToPosition(OO, RS("Desc").Value, 5800, VBRUN.AlignConstants.vbAlignLeft, False)
            End If
            PrintToPosition(OO, CurrencyFormat(RS("SellPrice").Value), 11000, VBRUN.AlignConstants.vbAlignRight, True)
        End If
    End Sub
End Module
