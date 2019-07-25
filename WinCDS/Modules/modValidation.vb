Module modValidation
    Public Const styNOT As String = "NOTES"
    Public Const styDEL As String = "DEL"
    Public Const styLAB As String = "LAB"
    Public Const stySTA As String = "STAIN"
    Public Const styDIS As String = "DISCOUNT"
    Public Const stySUB As String = "SUB"
    Public Const styPAY As String = "PAYMENT"
    Public Const styTX1 As String = "TAX1"
    Public Const styTX2 As String = "TAX2"
    Public Const styADJ As String = "--- Adj ---"
    Public Const styVOI As String = "VOID"
    Public Const staPFDEL As String = "DEL"
    Public Const staPFVoi As String = "VD"
    Public Const staPFRet As String = "x"
    Public Const staVoi As String = "VOID"
    Public Const staDel As String = "DEL"

    Public Function IsItem(ByVal Style As String) As Boolean
        '::::IsItem
        ':::SUMMARY
        ': Is Item (Style)
        ':::DESCRIPTION
        ': Returns whether style is an item (not a control type)
        ':::PARAMETERS
        ': - Style
        ':::RETURN
        ': Boolean
        Select Case Trim(Style)
            Case ""
                IsItem = False
            Case styDIS, stySTA, styDEL, styLAB, styTX1, styTX2, styNOT, stySUB, styPAY, styADJ, styVOI
                IsItem = False
            Case Else
                IsItem = True
        End Select
    End Function
    Public Function IsNote(ByVal Style As String) As Boolean
        '::::IsNote
        ':::SUMMARY
        ': Is Note (Style)
        ':::DESCRIPTION
        ': Returns whether status is NOTE type
        ':::PARAMETERS
        ': - Style - Indicates the Input value Style String.
        ':::RETURN
        ': Boolean - Returns True.
        IsNote = (Trim(Style) = styNOT)
    End Function
    Public Function IsDLS(ByVal Style As String) As Boolean
        ':::SUMMARY
        ': Is Delivery / Labor / Stain Protection Charge (Style)
        ':::DESCRIPTION
        ': Returns whether given style is a DEL / LAB / STAIN
        ':::PARAMETERS
        ': - Style
        ':::RETURN
        ': Boolean
        IsDLS = Trim(Style) = styDEL Or Trim(Style) = styLAB Or Trim(Style) = stySTA
    End Function
    Public Function IsDelivered(ByVal Status As String) As Boolean
        '::::IsDelivered
        ':::SUMMARY
        ': Is Delivered (Status)
        ':::DESCRIPTION
        ': Whether Status is of delivered type
        ':::PARAMETERS
        ': - Status
        ':::RETURN
        ': Boolean

        IsDelivered = False
        If Left(Trim(Status), 3) = staPFDEL Then IsDelivered = True
        '  If Trim(Status) = "VDDEL" Then IsDelivered = True   ' I'm not sure this is right.  No, it's very bad.  Make a special case when VDDEL is treated as delivered.
    End Function
    Public Function IsVoid(ByVal Status As String) As Boolean
        '::::IsVoid
        ':::SUMMARY
        ': Is Void (Status)
        ':::DESCRIPTION
        ': Returns whether Status is Void
        ':::PARAMETERS
        ': - Status
        ':::RETURN
        ': Boolean
        IsVoid = (Trim(Status) = staVoi Or Left(Trim(Status), 2) = staPFVoi)
    End Function
    Public Function IsPayment(ByVal Style As String) As Boolean
        '::::IsPayment
        ':::SUMMARY
        ': Is Payment (style)
        ':::DESCRIPTION
        ': Returns whether given style is a payment
        ':::PARAMETERS
        ': - Style
        ':::RETURN
        ': Boolean
        IsPayment = (Trim(Style) = styPAY)
    End Function
    Public Function IsDiscount(ByVal Style As String) As Boolean
        '::::IsDiscount
        ':::SUMMARY
        ': Is Discount (Style)
        ':::DESCRIPTION
        ': Returns whether style is a discount.
        ':::PARAMETERS
        ': - Style
        ':::RETURN
        ': Boolean
        IsDiscount = Trim(Style) = styDIS
    End Function
    Public Function IsADJ(ByVal Style As String) As Boolean
        '::::IsADJ
        ':::SUMMARY
        ': Is Adjustment (Style)
        ':::DESCRIPTION
        ': Returns whether the style is an adjustment
        ':::PARAMETERS
        ': - Style
        ':::RETURN
        ': Boolean
        IsADJ = (Trim(Style) = styADJ)
    End Function

End Module
