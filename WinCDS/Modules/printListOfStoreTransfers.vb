Module printListOfStoreTransfers
    Private Const ITEMS_PER_PAGE As Integer = 35
    Public Sub printListOfStoreTransfers_PrintRecords(ByVal StoreCount As Integer, ByVal DelDate As Date, ByVal EndDate As Date, ByVal ToPrinter As Boolean, Optional ByVal IncludeSoldTransfers As Boolean = False)
        Dim cTable As New CSQLAccess
        Dim cTa As CDataAccess
        Dim SQL As String
        Dim OO As Object
        Dim I As Integer

        If ToPrinter Then
            OO = Printer
        Else
            OO = frmPrintPreviewDocument.picPicture
            OutputObject = OO
            frmPrintPreviewDocument.CallingForm = InvPull
        End If

        On Error Resume Next
        cTa = cTable.DataAccess()

        SQL = ""
        SQL = SQL & " SELECT"
        SQL = SQL & "      Trans , Style ,Misc, Ddate1"
        For I = 1 To Setup_MaxStores : SQL = SQL & ", Loc" & I : Next
        SQL = SQL & " FROM Detail"
        SQL = SQL & " WHERE "
        SQL = SQL & "  ("
        SQL = SQL & "  (Trans IN ('TP')) AND [DDate1] BETWEEN #" & DelDate & "# AND #" & EndDate & "#"
        SQL = SQL & "  )"
        If IncludeSoldTransfers Then
            SQL = SQL & "  OR "
            SQL = SQL & "   ("
            SQL = SQL & "   Trans IN ('DS','NS','TR') AND ([DDate1] BETWEEN #" & DelDate & "# AND #" & EndDate & "#)"
            SQL = SQL & "   AND (FALSE " & vbCrLf
            For I = 1 To Setup_MaxStores
                SQL = SQL & "   OR (Store=" & I & " AND Loc" & I & "<>AmtSold)" & vbCrLf
            Next
            SQL = SQL & "   )"
            SQL = SQL & "  )"
        End If

        cTa.DataBase = GetDatabaseInventory()
        cTa.Records_OpenSQL(SQL)
        If cTa.Record_Count <> 0 Then
            'If IsUFO() Then Printer.Copies = 2
            TransferHeading(DelDate, EndDate, OO)
            Do While cTa.Records_Available
                PrintTransferList(cTa.RS, OO)
                If (cTa.Record_Index + 1) Mod ITEMS_PER_PAGE = 0 Then
                    If ToPrinter Then
                        Printer.NewPage()
                    Else
                        frmPrintPreviewDocument.NewPage()
                    End If

                    TransferHeading(DelDate, EndDate, OO)
                End If
            Loop
        End If
        If ToPrinter Then
            Printer.EndDoc()
        Else
            InvPull.Hide()
            frmPrintPreviewDocument.DataEnd()
        End If
        Printer.Orientation = 1
        cTa.Records_Close()
        cTa = Nothing
    End Sub

    Private Sub TransferHeading(ByVal FromDate As Date, ByVal toDate As Date, ByRef OutObj As Object)
        On Error Resume Next
        Dim I As Integer, N As Integer, PN As Integer

        OutObj.Orientation = 2
        PrintOut(OutObj:=OutObj, FontName:="Arial", FontSize:=10, FontBold:=True, DrawWidth:=20, X:=200, Y:=100)
        OutObj.Print("From: ", WeekdayName(FromDate) & " " & FromDate, TAB(110), "  To: ", WeekdayName(toDate) & " " & toDate)
        OutObj.CurrentX = 200
        PN = PageNumber
        PN = OutObj.Page
        OutObj.Print("Time: ", Now, TAB(110), "Page: ", PN)

        PrintOut(OutObj:=OutObj, FontSize:=18, FontBold:=True, X:=200, Y:=100)
        PrintOut(OutObj:=OutObj, XCenter:=True, Text:="Store Transfer Summary")
        PrintOut(OutObj:=OutObj, XCenter:=True, FontSize:=10, X:=200, Y:=500, Text:=StoreSettings.Name & "    " & StoreSettings.Address & "    " & StoreSettings.City)

        PrintOut(OutObj:=OutObj, X:=200, Y:=800)

        OutObj.Print("Trans No:")

        PrintToTab(OutObj, "Trans Date:", 14)
        PrintToTab(OutObj, "Style:", 30)

        N = ActiveNoOfLocations
        OutObj.FontBold = True
        For I = 1 To N
            PrintToTab(OutObj, "L" & I, 50 + ((I - 1) * 6), , I = N)
        Next
        OutObj.FontBold = False
    End Sub

    Private Sub PrintTransferList(ByRef RS As ADODB.Recordset, ByRef OutObj As Object)
        Dim I As Integer, N As Integer
        PrintOut(OutObj:=OutObj, FontSize:=10, X:=200)

        If Not IsNothing(RS("MISC").Value) Then
            OutObj.Print(RS("Misc").Value)
        End If
        If Not IsNothing(RS("Ddate1").Value) Then
            PrintToTab(OutObj, IfNullThenZeroDate(RS("Ddate1").Value), 14)
        End If
        PrintToTab(OutObj, RS("Style").Value, 30)
        For I = 1 To ActiveNoOfLocations
            PrintToTab(OutObj, RS("Loc" & I).Value, 50 + ((I - 1) * 6), , I = N)
        Next
        OutObj.Print
    End Sub
End Module
