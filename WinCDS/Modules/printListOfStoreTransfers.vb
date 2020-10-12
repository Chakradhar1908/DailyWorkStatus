Module printListOfStoreTransfers
    Public Sub printListOfStoreTransfers_PrintRecords(ByVal StoreCount As Long, ByVal DelDate As Date, ByVal EndDate As Date, ByVal ToPrinter As Boolean, Optional ByVal IncludeSoldTransfers As Boolean)
        Dim cTable As New CSQLAccess
        Dim cTa As CDataAccess
        Dim SQL As String
        Dim OO As Object
        Dim I As Long

        If ToPrinter Then
    Set OO = Printer
  Else
    Set OO = frmPrintPreviewDocument.picPicture
    Set OutputObject = OO
    Set frmPrintPreviewDocument.CallingForm = InvPull
  End If

        On Error Resume Next
  Set cTa = cTable.DataAccess()
  
    
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
        cTa.Records_OpenSQL SQL
  If cTa.Record_Count <> 0 Then
            'If IsUFO() Then Printer.Copies = 2
            TransferHeading DelDate, EndDate, OO
    Do While cTa.Records_Available
                PrintTransferList cTa.RS, OO
      If (cTa.Record_Index + 1) Mod ITEMS_PER_PAGE = 0 Then
                    If ToPrinter Then
                        Printer.NewPage()
                    Else
                        frmPrintPreviewDocument.NewPage()
                    End If

                    TransferHeading DelDate, EndDate, OO
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
          Set cTa = Nothing
End Sub

End Module
