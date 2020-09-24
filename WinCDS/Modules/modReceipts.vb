Module modReceipts
    Public Enum eReceiptTypes
        ert_SaleNo = 0
        ert_ArNo = 1
    End Enum

    Public Sub MakeReceipt(
      ByVal TransDate As String, ByVal ReceiptType As eReceiptTypes,
      ByVal ItemNo As String, ByVal First As String, ByVal Last As String,
      ByVal Add1 As String, ByVal Add2 As String, ByVal City As String, ByVal Zip As String,
      ByVal PreviousBalance As Decimal,
      ByVal PayType As String, ByVal Amount As Decimal, ByVal Balance As Decimal,
      ByVal Note As String,
      Optional ByVal Approval As String = "", Optional ByVal NextPaymentDate As String = "", Optional ByVal Arrears As Decimal = 0,
      Optional ByVal CheckRevolving As Boolean = False,
      Optional ByVal REPRINT As Boolean = False)
        Dim CustRec As Integer
        CustRec = GetReceipt()
        MakeReceiptPart(True, CustRec, TransDate, ReceiptType, ItemNo, First, Last, Add1, Add2, City, Zip, PreviousBalance, PayType, Amount, Balance, Note, Approval, NextPaymentDate, Arrears, CheckRevolving, REPRINT)
        Printer.CurrentY = 7733
        Printer.Print("----                                                                                                                                                                  ----")
        MakeReceiptPart(False, CustRec, TransDate, ReceiptType, ItemNo, First, Last, Add1, Add2, City, Zip, PreviousBalance, PayType, Amount, Balance, Note, Approval, NextPaymentDate, Arrears, CheckRevolving, REPRINT)
        Printer.EndDoc()
    End Sub

    Public Function GetReceipt() As String
        GetReceipt = GetFileAutonumber(CustRecFile, 999)
    End Function

    Private Sub MakeReceiptPart(ByVal FirstCopy As Boolean, ByVal RcptNo As Integer,
      ByVal TransDate As String, ReceiptType As eReceiptTypes,
      ByVal ItemNo As String, ByVal First As String, ByVal Last As String,
      ByVal Add1 As String, ByVal Add2 As String, ByVal City As String, ByVal Zip As String,
      ByVal PreviousBalance As Decimal,
      ByVal PayType As String, ByVal Amount As Decimal, ByVal Balance As Decimal,
      ByVal Note As String,
      Optional ByVal Approval As String = "", Optional ByVal NextPaymentDate As String = "", Optional ByVal Arrears As Decimal = 0,
      Optional ByVal CheckRevolving As Boolean = False,
      Optional ByVal REPRINT As Boolean = False)

        Dim T As String, Y As Integer, cy As Integer
        Y = IIf(FirstCopy, 400, 8200)
        Printer.FontName = "Arial"
        Printer.FontSize = 12
        Printer.FontBold = False

        ' frame around receipt
        Printer.DrawWidth = 12
        'Printer.Line(300, Y)-Step(11000, 7000), vbBlack, B
        Printer.Line(300, Y, 11000, 7000, , True)
        Printer.CurrentX = 0

        Printer.CurrentY = Y + 400
        Printer.FontBold = True
        Printer.Print(TAB(42), IIf(FirstCopy, "Customer Copy", "Store Copy"))
        If REPRINT Then
            Printer.Print(TAB(44), "REPRINT")
        End If
        Printer.FontBold = False

        Printer.FontSize = 14
        Printer.CurrentY = Y + 800
        Printer.Print(TAB(6), Trim(StoreSettings.Name), TAB(70))
        Printer.FontSize = 12
        Printer.Print("Date: ", TransDate)

        Printer.Print(TAB(7), Trim(StoreSettings.Address))
        Printer.Print(TAB(7), StoreSettings.City, TAB(76), "Receipt No: ", RcptNo)
        Printer.Print(TAB(7), StoreSettings.Phone)
        Select Case ReceiptType
            Case eReceiptTypes.ert_SaleNo
                Printer.Print(TAB(79), "Sale No: ")
            Case eReceiptTypes.ert_ArNo
                Printer.Print(TAB(71), "A/R Account No: ")
        End Select
        Printer.Print(ItemNo)

        Dim BlankRows As Integer, I As Integer
        BlankRows = 3
        I = 1
        Do Until I > BlankRows
            Printer.Print()
            I = I + 1
        Loop

        Printer.Print(TAB(10), Trim(First), " ", Trim(Last))
        cy = Printer.CurrentY

        '05/05/2004   took out boyd didn't want it and neither did Wilkenfeld
        '05/11/2006   Rogers does weekly, so they're removed from this as well
        'BFH20061130  Chicago (new age) didn't want it either
        On Error Resume Next
        Printer.FontBold = True
        Printer.FontSize = 14

        If ReceiptType = eReceiptTypes.ert_ArNo Then
            If Not (IsBoyd() Or IsWilkenfeld() Or IsRogers() Or IsChicago()) Then
                If NextPaymentDate <> "" Then
                    Printer.Print(TAB(40), "Next Regular Payment Due: ", NextPaymentDate)
                Else
                    Printer.Print()
                End If
                If Arrears > 0 Then
                    Printer.Print(TAB(54), "  Arrearages: " & FormatCurrency(Arrears))
                End If
            End If
        End If
        Printer.FontBold = False
        Printer.FontSize = 12

        Printer.CurrentY = cy
        Printer.Print()
        On Error GoTo 0


        Printer.Print(TAB(10), Add1)
        If Len(Add2) > 0 Then
            Printer.Print(TAB(10), Add2)
        End If
        Printer.Print(TAB(10), City, " ", Zip)
        If Len(Add2) = 0 Then
            Printer.Print()
        End If

        BlankRows = 3
        If CheckRevolving And IsFormLoaded("ArCard") Then
            BlankRows = BlankRows - ArCard.PayCount
        End If
        I = 1
        Do Until I > BlankRows
            Printer.Print()
            I = I + 1
        Loop
        Printer.Print(TAB(74), "Previous:                ", TAB(95), FormatCurrency(PreviousBalance))

        If CheckRevolving Then
            If IsFormLoaded("ArCard") Then
                I = 1
                Do While I <= ArCard.PayCount
                    ' This could cause strangeness if more than 3 sales are paid.
                    Printer.Print(TAB(74), IIf(ArCard.QueryPayLogSale(I) = "Interest" Or ArCard.QueryPayLogSale(I) = "Account", "", "Sale "), ArCard.QueryPayLogSale(I), ":", TAB(95), FormatCurrency(ArCard.QueryPayLogAmount(I)))
                    I = I + 1
                    BlankRows = BlankRows - 1
                Loop
            End If
        End If

        Printer.FontSize = 16
        Printer.Print(TAB(10), PayType & "  ")
        Printer.FontSize = 10
        Printer.Print(Approval)
        Printer.FontSize = 16
        Printer.Print(TAB(62), FormatCurrency(Amount))

        Printer.FontSize = 12
        Printer.DrawWidth = 4
        'Printer.Line(1000, Y + 5600)-(8000, Y + 5600)
        Printer.Line(1000, Y + 5600, 8000, Y + 5600)
        'Printer.Line(8800, Y + 5600)-(11000, Y + 5600)
        Printer.Line(8800, Y + 5600, 11000, Y + 5600)
        Printer.CurrentY = Y + 5450

        Printer.Print(TAB(75), "Balance:                 ", TAB(95), FormatCurrency(Balance))


        Printer.Print()

        If FirstCopy Or Approval = "" Then    ' request signature
            Printer.Print(TAB(10), "Rec By: ___________________________________  Note: ", Note)
        Else
            Printer.Print(TAB(16), "X ___________________________________  Note: ", Note)
            Printer.Print(TAB(20), "I authorize the above transaction")
        End If

        Printer.Print(TAB(75), "Thank You")
    End Sub
End Module
