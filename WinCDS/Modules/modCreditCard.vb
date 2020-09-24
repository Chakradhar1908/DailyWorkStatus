Public Module modCreditCard

    Public Function SwipeCards() As Boolean
        SwipeCards = False
        If Not StoreSettings.bUseCCMachine Then Exit Function
        If StoreSettings.CCProcessor = CCPROC_NONE Or StoreSettings.CCProcessor = CCPROC_NA Then Exit Function
        SwipeCards = True
    End Function

    Public Function SwipeCreditCards() As Boolean
        SwipeCreditCards = False
        If Not SwipeCards() Then Exit Function
        SwipeCreditCards = True
    End Function

    Public Function SwipeDebitCards() As Boolean
        If Not SwipeCards() Then
            SwipeDebitCards = False
            Exit Function
        End If
        SwipeDebitCards = True
        ' BFH20110309 - transaction central does not support Debit cards at this time
        If StoreSettings.CCProcessor = CCPROC_TC Then SwipeDebitCards = False
    End Function

    Public Function SwipeGiftCards() As Boolean
        If Not SwipeCards() Then
            SwipeGiftCards = False
            Exit Function
        End If
        SwipeGiftCards = False '### not suppored yet
    End Function

    Public Function ProcessCC(ByRef Amount As Decimal, ByRef pType As String, ByRef Approval As String, Optional ByRef CardTypeCode As Integer = -1, Optional ByRef TransID As String = "", Optional ByVal RefId As String = "", Optional ByRef Balance As Decimal = 0) As Boolean
        Dim Pmt As clsSaleItem
        Dim DBG As String
        On Error GoTo Handler
        DBG = "A"
        If IsFormLoaded("BillOSale") Then BillOSale.cmdProcessSale.Enabled = False : BillOSale.cmdMainMenu.Enabled = False : BillOSale.cmdNextSale.Enabled = False
        DBG = "Ba"
        'Load frmCashRegisterQuantity
        DBG = "Bb"
        frmCashRegisterQuantity.RefId = RefId
        DBG = "Bc"
        Pmt = frmCashRegisterQuantity.GetQuantityAndPrice("PAYMENT", "3", Amount, True)
        DBG = "C"
        If IsFormLoaded("BilloSale") Then BillOSale.cmdProcessSale.Enabled = True : BillOSale.cmdMainMenu.Enabled = True : BillOSale.cmdNextSale.Enabled = True
        DBG = "D"
        Approval = ""
        If Pmt Is Nothing Then Exit Function

        DBG = "E"
        Approval = Pmt.Extra1
        DBG = "F"
        pType = Pmt.Desc
        DBG = "G: PTYPE=" & pType
        Select Case UCase(Left(pType, 1))
            Case "V" : CardTypeCode = cdsPayTypes.cdsPT_Visa
            Case "M" : CardTypeCode = cdsPayTypes.cdsPT_MCard
            Case "D" : CardTypeCode = cdsPayTypes.cdsPT_Disc
            Case "A" : CardTypeCode = cdsPayTypes.cdsPT_amex
            Case Else : CardTypeCode = cdsPayTypes.cdsPT_amex
        End Select

        DBG = "H"
        Amount = Pmt.Price
        DBG = "I"
        TransID = Pmt.TransID
        Balance = Pmt.Balance
        DBG = "J"
        ProcessCC = True
        Exit Function
Handler:
        Dim X As String
        X = frmCashRegisterQuantity.DBG
        MessageBox.Show("Error in ProcessCC (" & Err.Number & ") [" & DBG & "-" & X & "]:" & Err.Description)
    End Function

    Public Function ProcessDebit(ByRef Amount As Decimal, ByRef pType As String, ByRef Approval As String, Optional ByRef CardTypeCode As Integer = 0, Optional ByRef TransID As String = "", Optional ByVal RefId As String = "") As Boolean
        Dim Pmt As clsSaleItem
        'Load frmCashRegisterQuantity
        frmCashRegisterQuantity.RefId = RefId
        Pmt = frmCashRegisterQuantity.GetQuantityAndPrice("PAYMENT", "9", Amount, True)
        Approval = ""
        If Pmt Is Nothing Then Exit Function

        Approval = Pmt.Extra1
        pType = Pmt.Desc
        CardTypeCode = 9

        Amount = Pmt.Price
        TransID = Pmt.TransID
        ProcessDebit = True
    End Function

    Public Function ProcessGiftCard(ByRef Amount As Decimal, ByRef pType As String, ByRef Approval As String, Optional ByRef CardTypeCode As Integer = 0, Optional ByVal RefId As String = "") As Boolean
        Dim Pmt As clsSaleItem
        'Load frmCashRegisterQuantity
        frmCashRegisterQuantity.RefId = RefId
        Pmt = frmCashRegisterQuantity.GetQuantityAndPrice("PAYMENT", "12", Amount, True)
        Approval = ""
        If Pmt Is Nothing Then Exit Function

        Approval = Pmt.Extra1
        pType = Pmt.Desc
        CardTypeCode = 9

        Amount = Pmt.Price
        ProcessGiftCard = True
    End Function

    Public Function CreditCardSwipeValid(ByVal vSwipe As String, Optional ByVal RequireTrack2 As Boolean = False) As Boolean
        Dim T1 As String, T2 As String
        CreditCardSwipeValid = ParseTrackData(vSwipe, T1, T2)
        If RequireTrack2 And T2 = "" Then CreditCardSwipeValid = False
    End Function

    Public Function CCXOut(ByVal CC As String) As String
        If Len(CC) <= 10 Then Exit Function
        CCXOut = New String("X"c, Len(CC) - 4) & Right(CC, 4)
    End Function

    Public Function ParseTrackData(Optional ByRef CCSwipe As String = "", Optional ByRef vTrack1 As String = "", Optional ByRef vTrack2 As String = "", Optional ByRef vCCNumber As String = "", Optional ByRef vCCTypeName As String = "", Optional ByRef vExpMonth As String = "", Optional ByRef vExpYear As String = "", Optional ByRef vCardHolderName As String = "") As Boolean
        Dim Sent1 As String, FormatCode As String, PAN1 As String
        Dim C As String, A As Integer, B As Integer

        If CCSwipe = "" Then Exit Function

        A = InStr(CCSwipe, "?")
        If A = 0 Then
            vTrack1 = CCSwipe
            vTrack2 = ""
        Else
            vTrack1 = Mid(CCSwipe, 1, A)
            vTrack2 = Mid(CCSwipe, A + 1)
        End If

        Sent1 = Left(vTrack1, 1)          ' %
        FormatCode = Mid(vTrack1, 2, 1)   ' B

        C = Mid(vTrack1, 3)

        A = InStr(C, "^")
        If A = 0 Then Exit Function
        vCCNumber = Left(C, A - 1)
        If IsDevelopment() And Left(vCCNumber, 1) = "%" Then Stop

        vCCTypeName = GetCCTypeName(vCCNumber)
        C = Mid(C, A + 1)

        A = InStr(C, "^")
        If A = 0 Then Exit Function
        vCardHolderName = Left(C, A - 1)
        C = Mid(C, A + 1)

        vExpYear = Mid(C, 1, 2)
        vExpMonth = Mid(C, 3, 2)
        C = Mid(C, 5)


        ParseTrackData = True
    End Function

    Public Function GetCCTypeName(ByVal CC As String) As String
        Select Case Val(Left(CC, 2))
            Case 34, 37 : GetCCTypeName = "AMEX"
'    Case 30, 36: GetCCTypeName = "Diners"
            Case 40 To 48 : GetCCTypeName = "Visa"
            Case 51 To 55 : GetCCTypeName = "MasterCard"
            Case 60, 62, 64, 65 : GetCCTypeName = "Discover"
            Case Else : GetCCTypeName = ""
        End Select
    End Function

    Public Function ManualCCEntry(Optional ByRef CCNumber As String = "", Optional ByRef ExpDate As String = "", Optional ByRef CardHolderName As String = "", Optional ByRef CVV2 As String = "", Optional ByRef ZipCode As String = "", Optional ByRef Swipe As String = "") As Boolean
        ManualCCEntry = frmCCManualEntry.GetManualCCEntry(CCNumber, ExpDate, CardHolderName, CVV2, ZipCode, Swipe)
        'Unload frmCCManualEntry
        frmCCManualEntry.Close()
    End Function

    Public Function ProcessCCReturn(ByRef Amount As Decimal, ByRef Approval As String, Optional ByRef TransID As String = "", Optional ByVal SaleDate As String = "", Optional ByVal RefId As String = "") As Boolean
        Dim Pmt As clsSaleItem, Tid As String
        If Amount < 0 Then Amount = -Amount

        If StoreSettings.CCProcessor = CCPROC_XC Then
            'Load frmCashRegisterQuantity
            frmCashRegisterQuantity.RefId = RefId
            Pmt = frmCashRegisterQuantity.DoReturn(Amount, "3", TransID, SaleDate)
        ElseIf StoreSettings.CCProcessor = CCPROC_XL Then
            'Load frmCashRegisterQuantity
            frmCashRegisterQuantity.RefId = RefId
            Pmt = frmCashRegisterQuantity.DoReturn(Amount, "3", TransID, SaleDate)
        ElseIf StoreSettings.CCProcessor = CCPROC_TC Then
            '    If TransID = "" Then Exit Function ' can't return a payment w/o a transid for TC

            Pmt = New clsSaleItem

            Dim TC As clsTransactionCentral
            TC = New clsTransactionCentral
            TC.Amount = Amount
            TC.RefId = RefId
            If Left(TransID, 1) = "#" Then                      ' Refund an amount on a sale number...  good for blind credit on adjustments
                Pmt = frmCashRegisterQuantity.DoReturn(Amount, "3", Tid, Today)
                If Pmt Is Nothing Then Exit Function
                Pmt.Desc = Trim("Payment Voided. " & TC.ApprovalCode)
            Else                                                ' Void an individual transaction.. good for void interface
                TC.TransID = TransID
                If TC.ExecVoid(SaleDate) Then
                    Pmt.TransID = TC.TransID
                    Pmt.Extra1 = TC.ApprovalCode
                    Pmt.Extra2 = TC.RefId
                    Pmt.Price = -Amount
                    Pmt.DisplayPrice = -Amount
                    Pmt.Desc = Trim("Payment Voided. " & TC.ApprovalCode)
                End If
            End If
        ElseIf StoreSettings.CCProcessor = CCPROC_CI Then
            Pmt = New clsSaleItem

            Dim CI As clsChargeItPro
            CI = New clsChargeItPro
            CI.Amount = Amount
            If Left(TransID, 1) = "#" Then                      ' Refund an amount on a sale number...  good for blind credit on adjustments
                Pmt = frmCashRegisterQuantity.DoReturn(Amount, "3", Tid, Today)
                If Pmt Is Nothing Then Exit Function
                Pmt.Desc = Trim("Payment Voided. " & CI.ApprovalCode)
            Else                                                ' Void an individual transaction.. good for void interface
                CI.Amount = Amount
                CI.RefId = TransID
                Pmt.TransID = TransID
                If CI.ExecVoid(, TransID) Then
                    Pmt.TransID = CI.RefId
                    Pmt.Extra1 = CI.ApprovalCode
                    Pmt.Price = -Amount
                    Pmt.DisplayPrice = -Amount
                    Pmt.Desc = Trim("Payment Returned. " & CI.ApprovalCode)
                End If
            End If
            DisposeDA(CI)
        ElseIf StoreSettings.CCProcessor = CCPROC_CM Then
            Dim cM As clsCredomatic
            cM = New clsCredomatic
            DisposeDA(cM)
        End If
        Approval = ""
        If Pmt Is Nothing Then Exit Function

        Approval = Pmt.Extra1
        TransID = Pmt.TransID
        ProcessCCReturn = True

        DisposeDA(Pmt)
    End Function

    Public Function ProcessDebitReturn(ByRef Amount As Decimal, ByRef Approval As String, Optional ByRef TransID As String = "", Optional ByVal RefId As String = "") As Boolean
        Dim Pmt As clsSaleItem
        If Amount < 0 Then Amount = -Amount
        'Load frmCashRegisterQuantity
        frmCashRegisterQuantity.RefId = RefId
        Pmt = frmCashRegisterQuantity.DoReturn(Amount, "9", TransID, "")
        Approval = ""
        If Pmt Is Nothing Then Exit Function

        Approval = Pmt.Extra1
        TransID = Pmt.TransID
        ProcessDebitReturn = True
    End Function

    Public Function ProcessGiftCardReturn(ByRef Amount As Decimal, ByRef Approval As String, Optional ByVal RefId As String = "") As Boolean
        Dim Pmt As clsSaleItem
        If Amount < 0 Then Amount = -Amount
        'Load frmCashRegisterQuantity
        frmCashRegisterQuantity.RefId = RefId
        Pmt = frmCashRegisterQuantity.DoReturn(Amount, "12", "", "")
        Approval = ""
        If Pmt Is Nothing Then Exit Function

        Approval = Pmt.Extra1
        ProcessGiftCardReturn = True
    End Function

End Module
