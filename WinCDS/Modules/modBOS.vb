Module modBOS
    Public Function IsItemNontaxable(ByVal Style As String, Optional ByVal TaxOnSale As Boolean = False) As Boolean
        '::::IsItemNontaxable
        ':::SUMMARY
        ': Return whether the style is non-taxable
        ':::DESCRIPTION
        ': Based on store settings, returns whether the given item is non-taxable.
        ':::PARAMETERS
        ': - Style - The style to verify.
        ': - TaxOnSale - Indicxates whether there is tax on the sale.  Unused.
        ':::RETURN
        ': Boolean - Returns whether the style is non-taxable

        If Style = "DEL" And Not StoreSettings.bDeliveryTaxable Then IsItemNontaxable = True
        If Style = "LAB" And Not StoreSettings.bLaborTaxable Then IsItemNontaxable = True
        If Style = "TAX1" Or Style = "TAX2" Then IsItemNontaxable = True
    End Function
    Public Function DescribeTimeWindow(ByVal twA, ByVal twB) As String
        '::::DescribeTimeWindow
        ':::SUMMARY
        ': Describe a time window
        ':::DESCRIPTION
        ': Returns a text description of the time window
        ':::PARAMETERS
        ': - twA - First Window
        ': - twB - Second Window
        ':::RETURN
        ': String - Returns the result as a String.
        If Not IsDate(twA) Then twA = ""
        If Not IsDate(twB) Then twB = ""
        If twA = "" And twB = "" Then Exit Function

        If twA = "" Then
            DescribeTimeWindow = "Before " & Format(TimeValue(twB), "h:mm ampm")
        ElseIf twB = "" Then
            DescribeTimeWindow = "After " & Format(TimeValue(twA), "h:mm ampm")
        Else
            DescribeTimeWindow = Format(twA, "h:mm ampm") & " to " & Format(twB, "h:mm ampm")
        End If
    End Function
    Public Function DetectSaleNo() As String
        '::::DetectSaleNo
        ':::SUMMARY
        ': Attempt to detect customer information, if possible
        ':::DESCRIPTION
        ': Checks for open forms which may contain the current customer information.
        ':
        ': Safely checks the state of various forms to see if they are loaded and pulls the relevant information, if possible.
        ':::RETURN
        ': String - Sale Number, if available
        If IsFormLoaded("BillOSale") Then DetectSaleNo = Trim(BillOSale.BillOfSale.Text) : Exit Function
        If IsFormLoaded("Service") Then DetectSaleNo = Trim(Service.lblSaleNo.Text) : Exit Function
        If IsFormLoaded("MailCheck") Then DetectSaleNo = Trim(MailCheck.SaleNo) : Exit Function
        '...
    End Function

    Public Function QueryPaymentDescription(ByVal PmtType As Integer) As String
        '::::QueryPaymentDescription
        ':::SUMMARY
        ': Returns a text description of the payment type
        ':::DESCRIPTION
        ': Returns a string describing the payment type
        ':::PARAMETERS
        ': - PmtType - Indicates the payment type.
        ':::RETURN
        ': String - Returns the result as a String.
        Select Case PmtType
            Case 1 : QueryPaymentDescription = "CASH"
            Case 2 : QueryPaymentDescription = "CHECK"
            Case 3 : QueryPaymentDescription = "VISA CARD"
            Case 4 : QueryPaymentDescription = "MASTER CARD"
            Case 5 : QueryPaymentDescription = "DISCOVER CARD"
            Case 6 : QueryPaymentDescription = "AMEX CARD"
            Case 9 : QueryPaymentDescription = "DEBIT CARD"
            Case 11 : QueryPaymentDescription = "STORE FINANCE"
            Case 12 : QueryPaymentDescription = "STORE CARD"
            Case Else : QueryPaymentDescription = ""
                ' Returns "" for invalid payment type.

        End Select
    End Function

    Public Function DetectCustomerName() As String
        '::::DetectCustomerName
        ':::SUMMARY
        ': Attempt to detect customer information, if possible
        ':::DESCRIPTION
        ': Checks for open forms which may contain the current customer information.
        ':
        ': Safely checks the state of various forms to see if they are loaded and pulls the relevant information, if possible.
        ':::RETURN
        ': String - Customer name, if available
        If IsFormLoaded("BillOSale") Then DetectCustomerName = Trim(BillOSale.CustomerFirst.Text & " " & BillOSale.CustomerLast.Text) : Exit Function
    End Function

    Public Function DetectCustomerZipCode() As String
        '::::DetectCustomerZipCode
        ':::SUMMARY
        ': Attempt to detect customer information, if possible
        ':::DESCRIPTION
        ': Checks for open forms which may contain the current customer information.
        ':
        ': Safely checks the state of various forms to see if they are loaded and pulls the relevant information, if possible.
        ':::RETURN
        ': String - Customer Zip Code, if available
        If IsFormLoaded("BillOSale") Then DetectCustomerZipCode = BillOSale.CustomerZip.Text : Exit Function
        If IsFormLoaded("frmCashRegister") Then DetectCustomerZipCode = frmCashRegister.MailZip : Exit Function
        '...
    End Function

    Public Function MerchandisePrice(ByVal Price As Decimal) As Decimal
        '::::MerchandisePrice
        ':::SUMMARY
        ': Return a 'merchandised' price
        ':::DESCRIPTION
        ': Given a price of 100.00, will create a 'merchandised' price of '99.95'.  Can round up or
        ': down depending on how close the price is to the threshhold.
        ':
        ': Used only when the corresponding setting is enabled in the store setup.
        ':
        ': Does nothing if price is below 75$ (customized for some stores)
        ':
        ':::EXAMPLES
        ': - MerchandisePrice(100)    ==  99.95
        ': - MerchandisePrice(101)    == 105.50
        ': - MerchandisePrice(105)    == 105.50
        ': - MerchandisePrice(105.50) == 105.50
        ': - MerchandisePrice(106)    == 109.95
        ':::PARAMETERS
        ': - Price - The price to be merchandised.
        ':::RETURN
        ': Currency - Returns the Merchandised price.
        Dim Dollars As Integer
        If StoreSettings.CalculateList = 2 Then
            MerchandisePrice = RoundUp(Price)
            Exit Function
        End If
        If StoreSettings.bNoMerchandisePrice Then
            MerchandisePrice = Price
            Exit Function
        End If
        If Price > 0 And Price < 70 And IsGrizzlys() Then
            MerchandisePrice = Int(Price) + 0.95
            Exit Function
        End If
        If Price < 75 Then
            MerchandisePrice = Price
            Exit Function
        End If

        Dollars = Int(Price)
        Select Case Right(Dollars, 1)
            Case 0
                ' Go down to the last 9.95.
                MerchandisePrice = Int(Dollars / 10) * 10 - 0.05
            Case 1, 2, 3, 4, 5
                ' Go up to 5.5
                MerchandisePrice = Int(Dollars / 10) * 10 + 5.5
            Case Else
                ' Go up to 9.95
                MerchandisePrice = Int(Dollars / 10) * 10 + 9.95
        End Select
    End Function

End Module
