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

    Public Function NonItemStyleString(Optional ByVal AllowSLD As Boolean = False, Optional ByVal AllowNOTES As Boolean = False) As String
        '::::NonItemStyleString
        ':::SUMMARY
        ': Returns a SQL IN() ready string of Non-item styles
        ':::DESCRIPTION
        ': A helper function for writing SQL statements, this will return a configurable formatted clause
        ': to place within a SQL `...` IN (`...`) block.
        ':::PARAMETERS
        ': - AllowSLD - STAIN/LAB/DEL included
        ': - AllowNOTES - NOTES included
        ':::RETURN
        ': String - The IN (`...`) clause ready for inclusion in a SQL.
        NonItemStyleString = "'" & Join(NonItemStyles(AllowSLD, AllowNOTES), "', '") & "'"
    End Function

    Public Function NonItemStyles(Optional ByVal AllowSLD As Boolean = False, Optional ByVal AllowNOTES As Boolean = False) As Object
        '::::NonItemStyles
        ':::SUMMARY
        ': Returns a list of 'non-item' styles.
        ':::DESCRIPTION
        ': Returns an array of 'non-item' styles, partially configurable.
        ':
        ': Some Examples of non-item styles returned in the array are:
        ':   - TAX1, TAX2, PAYMENT, SUB, VOID, ...
        ':
        ': Can Optionally include/exclude the SLD styles or NOTES.
        ':::PARAMETERS
        ': - AllowSLD - Stain/Labor/Delivery
        ': - AllowNOTES - NOTES
        ':::RETURN
        ': Variant - Returns a String() of non-item styles.

        'If AllowSLD And AllowNOTES Then NonItemStyles = Array("TAX1", "TAX2", "SUB", "PAYMENT", "VOID", "Style", "--- Adj ---")
        If AllowSLD And AllowNOTES Then NonItemStyles = New String() {"TAX1", "TAX2", "SUB", "PAYMENT", "VOID", "Style", "--- Adj ---"}
        'If AllowSLD And Not AllowNOTES Then NonItemStyles = Array("NOTES", "TAX1", "TAX2", "SUB", "PAYMENT", "VOID", "Style", "--- Adj ---")
        If AllowSLD And Not AllowNOTES Then NonItemStyles = New String() {"NOTES", "TAX1", "TAX2", "SUB", "PAYMENT", "VOID", "Style", "--- Adj ---"}
        'If Not AllowSLD And AllowNOTES Then NonItemStyles = Array("STAIN", "DEL", "LAB", "TAX1", "TAX2", "SUB", "PAYMENT", "VOID", "Style", "--- Adj ---")
        If Not AllowSLD And AllowNOTES Then NonItemStyles = New String() {"STAIN", "DEL", "LAB", "TAX1", "TAX2", "SUB", "PAYMENT", "VOID", "Style", "--- Adj ---"}
        'If Not AllowSLD And Not AllowNOTES Then NonItemStyles = Array("STAIN", "DEL", "LAB", "NOTES", "TAX1", "TAX2", "SUB", "PAYMENT", "VOID", "Style", "--- Adj ---")
        If Not AllowSLD And Not AllowNOTES Then NonItemStyles = New String() {"STAIN", "DEL", "LAB", "NOTES", "TAX1", "TAX2", "SUB", "PAYMENT", "VOID", "Style", "--- Adj ---"}
    End Function

    Public Function GetWinCDSCity(ByVal CityState As String) As String
        '::::GetWinCDSCity
        ':::SUMMARY
        ': Get City from a CitySTZip field
        ':::DESCRIPTION
        ': Given a CityStZip field, returns the City
        ':::PARAMETERS
        ': - CityState - The field to be parsed.
        ':::RETURN
        ': String - Returns the city
        CitySTZip(CityState)
        GetWinCDSCity = CityState
    End Function

    Public Function GetWinCDSState(ByVal CityState As String) As String
        '::::GetWinCDSState
        ':::SUMMARY
        ': Get State from a CitySTZip field
        ':::DESCRIPTION
        ': Given a CityStZip field, returns the State
        ':::PARAMETERS
        ': - CityState - The field to be parsed
        ':::RETURN
        ': String - Returns the state
        Dim S As String
        CitySTZip(CityState, S)
        GetWinCDSState = S
    End Function

    Public Function GetWinCDSZip(ByVal CityState As String) As String
        '::::GetWinCDSZip
        ':::SUMMARY
        ': Get Zip from a CitySTZip field
        ':::DESCRIPTION
        ': Given a CityStZip field, returns the Zip
        ':::PARAMETERS
        ': - CityState - The field to be parsed
        ':::RETURN
        ': String - Returns the zip
        Dim Z As String
        CitySTZip(CityState, , Z)
        GetWinCDSZip = Z
    End Function

    'to split "Toledo, OH" to "Toledo" and "OH"
    Public Sub CitySTZip(ByRef City As String, Optional ByRef ST As String = "", Optional ByRef Zip As String = "")
        '::::CitySTZip
        ':::SUMMARY
        ': Parse a CitySTZip field into constituent parts
        ':::DESCRIPTION
        ': Given a text field containing City, ST, and Zip, split the field into individual components
        ':
        ':Returns results ByRef
        ':::PARAMETERS
        ': - City - Indicates the City Name given by User. ByRef.
        ': - ST - Indicates the State Name given by User. ByRef.
        ': - Zip - Indicates the Zip Code given by User. ByRef.
        Dim X As Integer
        ST = ""
        Zip = ""
        City = Replace(City, ",", " ")
        City = CleanAddress(City, False, True)
        On Error Resume Next
        If Len(City) < 2 Then Exit Sub
        If InStr(City, " ") <= 0 Then Exit Sub
        X = InStrRev(City, " ")
        Zip = Trim(Mid(City, X + 1))
        If Not IsNumeric(Zip) Then
            Zip = ""
        Else
            City = Trim(Left(City, X - 1))
        End If
        X = InStrRev(City, " ")
        ST = Trim(Mid(City, X + 1))
        ST = CleanState(ST)
        City = Trim(Left(City, X - 1))
    End Sub

    Public Function IsNonItemStyle(ByVal Style As String) As Boolean
        '::::IsNonItemStyle
        ':::SUMMARY
        ': Return whether the given style is a non-item
        ':::DESCRIPTION
        ': Return true if the style specified is a non-item (system reserved) style
        ':
        ': Does NOT include S/L/D or NOTES.
        ':::PARAMETERS
        ': - Style - Indicates the Style.
        ':::RETURN
        ': Boolean - Returns True if the Style is a non-item.  False otherwise.

        Dim El As Object
        For Each El In NonItemStyles()
            If El = Style Then
                IsNonItemStyle = True
                Exit Function
            End If
        Next
    End Function

End Module
