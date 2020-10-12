Module modStoreTransfers
    Public Function TransferViewRowColor(ByVal Status As String) As Color
        '::::TransferViewRowColor
        ':::SUMMARY
        ': Used to describe color to Transfer Row.
        ':::DESCRIPTION
        ': This function is used to describe color for Transfer Row based on status of Transfer.
        ':::PARAMETERS
        ': - Status - Indicates the Status of Transfer.
        ':::RETURN
        ': Long : Returns transfer row color as a long.
        Select Case LCase(Status)
            Case "open", "tp" : TransferViewRowColor = Color.Magenta
            Case "closed", "tr" : TransferViewRowColor = Color.Black
            Case Else : TransferViewRowColor = Color.Red  'void, etc
        End Select
    End Function

    Public Function DescribeTransferStatus(ByVal Trans As String) As String
        '::::DescribeTransferStatus
        ':::SUMMARY
        ': Used to Describe Transfer Status.
        ':::DESCRIPTION
        ': This function is used to describe Transfer Status.
        ':::PARAMETERS
        ': - Trans - Indicates the Transfer Status.
        ':::RETURN
        ': String : Returns Status as a String.
        Select Case Trans
            Case "TV" : DescribeTransferStatus = "Void"
            Case "TP" : DescribeTransferStatus = "Open"
            Case Else : DescribeTransferStatus = "Closed"
        End Select
    End Function

    Public Function GetTransferRSByTID(ByVal Tid As Integer, Optional ByVal ExtraSQL As String = "", Optional ByVal OrderBy As String = "ORDER BY DetailID") As ADODB.Recordset
        '::::GetTransferRSByTID
        ':::SUMMARY
        ': Used to get Transfer RecordSet using Transfer ID.
        ':::DESCRIPTION
        ': This function is used to get Transfer Recordset using Transfer ID through SQL Statement.
        ':::PARAMETERS
        ': - Tid - Indicates the Transfer Id.
        ': - ExtraSQL - Indicates the extra SQL statement to get Transfer Recordset.
        ': - OrderBy - Indicates the Order By condition to filter the records in RecordSet.
        ':::RETURN
        ': Recordset : Returns Transfer Recordset.

        '  Set GetTransferRSByTID = GetRecordsetBySQL("SELECT * FROM StoreTransfers WHERE TransferID=" & TID & " " & ExtraSQL & " " & OrderBy, , GetDatabaseInventory)
        GetTransferRSByTID = GetRecordsetBySQL("SELECT * FROM Detail WHERE Trans IN ('TP','TR','TV') AND DetailID=" & Tid & " " & ExtraSQL & " " & OrderBy, , GetDatabaseInventory)
    End Function

    Public Function PrintTransfer(ByVal TN As String, Optional ByVal wCost As Boolean = False, Optional ByVal Juice As String) As Boolean
        '::::PrintTransfer
        ':::SUMMARY
        ': Used to Print Transfers.
        ':::DESCRIPTION
        ': This function is used to Print Transfer by defining properities to it.
        ':::PARAMETERS
        ': - TN - Indicates the Transfer Number.
        ': - wCost
        ': - Juice - Indicates the amount of interest to be pay, if customer has not pay the loan with in time.
        ':::RETURN
        ': Boolean
        Dim X As Recordset, P As Object, C As Long, R As Long, I As Long, cInv As CInvRec, NA As Long, F As Long
        Dim TLanded As Currency, TJuice As Currency, TTotal As Currency

        If Not TransferNoExists(TN) Then Exit Function
  Set X = GetTransferRSByTN(TN)
  NA = NoOfActiveLocations
  
  Set P = OutputObject
  R = Printer.ScaleWidth
        C = R / 2
        P.FontSize = 20
        P.FontBold = True
        PrintToPosition P, "Store Transfer #" & TN, C, 5, True
  P.FontSize = 16
        PrintToPosition P, StoreSettings(1).Name, C, 5, True
  P.FontSize = 14
        P.FontBold = False
        PrintToPosition P, StoreSettings(1).Address, C, 5, False
'  P.FontSize = 10
        '  PrintToPosition P, "Created: " & X("Ddate1"), R, vbAlignRight, True

        PrintToPosition P, "", , , True
  PrintToPosition P, "", , , True

  P.FontSize = 10
        P.FontBold = True
        PrintToPosition P, "Style", 100, vbAlignLeft
  PrintToPosition P, "Scheduled", 1500, vbAlignLeft
  PrintToPosition P, "Status", 4000, vbAlignLeft
  For I = 1 To NA
            PrintToPosition P, "L" & I, 5500 + (I - 1) * 400, vbAlignRight
  Next

        F = 5500 + NA * 400 + 400
        If wCost Then
            PrintToPosition P, "Cost", F, vbAlignRight
    PrintToPosition P, "Juice", F + 900, vbAlignRight
    PrintToPosition P, "Total", F + 1800, vbAlignRight
  End If


        Dim RRR As Long
        PrintToPosition P, "", , , True
  RRR = OutputObject.CurrentY
        P.FontBold = False
        On Error Resume Next
        P.Line(50, Printer.CurrentY)-(Printer.ScaleWidth - 50, Printer.CurrentY)
On Error GoTo 0
        OutputObject.CurrentY = RRR

        Do While Not X.EOF
            PrintToPosition P, IfNullThenNilString(X("Style")), 100, vbAlignLeft
    PrintToPosition P, IfNullThenNilString(X("Ddate1")), 1500, vbAlignLeft
    PrintToPosition P, DescribeTransferStatus(IfNullThenNilString(X("Trans"))), 4000, vbAlignLeft

    For I = 1 To NA
                PrintToPosition P, "" & IfNullThenZero(X("Loc" & I)), 5500 + (I - 1) * 400, vbAlignRight
    Next

            If wCost Then
                Dim CC As CInvRec
      Set CC = New CInvRec
      If CC.Load(X("Style"), "Style") Then
                    PrintToPosition P, CurrencyFormat(CC.Landed), F, vbAlignRight
        PrintToPosition P, CurrencyFormat(CC.Landed * GetPercent(Juice) * 0.01), F + 900, vbAlignRight
        PrintToPosition P, CurrencyFormat(CC.Landed + GetPercent(Juice) * 0.01), F + 1800, vbAlignRight
        TLanded = TLanded + CC.Landed
                    TJuice = TJuice + CC.Landed * GetPercent(Juice) * 0.01
                    TTotal = TTotal + CC.Landed + CC.Landed * GetPercent(Juice) * 0.01
                End If
                DisposeDA CC
    End If

            PrintToPosition P, "", , , True
    Set cInv = New CInvRec
    If cInv.Load(X("Style"), "Style") Then
                PrintToPosition P, cInv.Desc, 200, , True
    End If
            DisposeDA cInv

    X.MoveNext
        Loop

        P.FontBold = True
        PrintToPosition P, CurrencyFormat(TLanded), F, vbAlignRight
  PrintToPosition P, CurrencyFormat(TJuice), F + 900, vbAlignRight
  PrintToPosition P, CurrencyFormat(TTotal), F + 1800, vbAlignRight
  P.FontBold = False

        If OutputToPrinter Then
            P.EndDoc
        Else
            frmPrintPreviewDocument.MousePointer = 0
            frmPrintPreviewDocument.DataEnd()
        End If
        PrintTransfer = True
    End Function

End Module
