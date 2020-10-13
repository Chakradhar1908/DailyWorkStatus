Public Module modItemLocation
    Public Const LocationBarcodeSep As String = "$"
    Public Enum ItemLocationStatus
        ItemLocationStatus_Stocked = 0
        ItemLocationStatus_Pulled = 1
        ItemLocationStatus_Deleted = 2
    End Enum

    Public Function EncodeSaleNoBarcode(ByVal SaleNo As String) As String
        '  EncodeSaleNoBarcode = PrepareBarcode("$$" & SaleNo)
        EncodeSaleNoBarcode = SaleNo
    End Function

    Public Sub DecodeLocation(ByVal Loc As String, ByRef Bld As String, ByRef Row As String, ByRef Lvl As String, ByRef Bay As String)
        Dim Div As String, T() As String
        Div = LocationBarcodeSep
        If Left(Loc, 1) <> Div Then Exit Sub
        T = Split(Loc, Div)
        If UBound(T) - LBound(T) + 1 <> 5 Then Exit Sub
        Bld = T(1)
        Row = T(2)
        Lvl = T(3)
        Bay = T(4)
    End Sub

    Public ReadOnly Property ItemLocBldMaxLen() As Integer
        Get
            ItemLocBldMaxLen = 3
        End Get
    End Property

    Public Function EncodeLocation(ByVal Bld As String, ByVal Row As String, ByVal level As String, ByVal Bay As String) As String
        Dim Div As String
        Div = LocationBarcodeSep
        ' bfh20050509
        ' i find out we can't have >16 chars in a barcode...
        ' each element is limited to 3 characters, divider before each one
        '  EncodeLocation = Div & Div & "CDSLOC1" & Div & Bld & Div & Row & Div & Level & Div & Bay & Div
        EncodeLocation = Div & Bld & Div & Row & Div & level & Div & Bay
    End Function

    Public ReadOnly Property ItemLocRowMaxLen() As Integer
        Get
            ItemLocRowMaxLen = 3
        End Get
    End Property

    Public ReadOnly Property ItemLocLvlMaxLen() As Integer
        Get
            ItemLocLvlMaxLen = 3
        End Get
    End Property

    Public ReadOnly Property ItemLocBayMaxLen() As Integer
        Get
            ItemLocBayMaxLen = 3
        End Get
    End Property
End Module
