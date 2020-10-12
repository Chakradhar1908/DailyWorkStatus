Module printPastDeliveries
    Public Sub printPastDeliveries_PrintRecords(ByVal StoreNo As Long, ByVal FromDate As Date, ByVal toDate As Date, ByVal ToPrinter As Boolean)
        Dim OO As Object
        Dim SQL As String
        Dim RS As ADODB.Recordset
        Dim PageNo As Long

        If ToPrinter Then
    Set OO = Printer
  Else
    Set OutputObject = frmPrintPreviewDocument.picPicture
    Set OO = OutputObject
    Set frmPrintPreviewDocument.CallingForm = InvPull
  End If

        SQL = "SELECT gm.DelDate, gm.SaleNo, m.Last as Name, m.City, gm.Quantity, gm.Style, gm.Desc, gm.SellPrice "
        SQL = SQL & "FROM GrossMargin gm INNER JOIN Mail m "
        SQL = SQL & "on gm.MailIndex = m.Index "
        SQL = SQL & "WHERE gm.Store =" & StoreNo
        SQL = SQL & "AND gm.PorD = 'D'"
        SQL = SQL & "AND gm.DelDate >=#" & FromDate & "# AND gm.DelDate <=#" & toDate & "#"

  Set RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(StoreNo))

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
                PrintReportHeader "Past Deliveries", StoreSettings(StoreNo).Name, StoreSettings(StoreNo).Address & " " & StoreSettings(StoreNo).Phone, FromDate, toDate, OO, PageNo
      OO.Print ""

      OO.FontSize = 10
                OO.FontBold = True

                ReportColumnHeading OO
    End If

            ShowReportData RS, OO
    RS.MoveNext()
        Loop

        If ToPrinter Then
            Printer.EndDoc()
        Else
            InvPull.Hide()
            frmPrintPreviewDocument.DataEnd()
        End If
    End Sub

End Module
