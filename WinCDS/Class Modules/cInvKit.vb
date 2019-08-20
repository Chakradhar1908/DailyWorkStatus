Public Class cInvKit
    'Public Item() As String
    'Public Quantity() As Double
    Public KitStyleNo As String
    Public KitSKU As String
    Public Heading As String
    Public MemoArea As String
    Public Item1 As String
    Public Item1Rec As Integer
    Public Quan1 As Double
    Public Item2 As String
    Public Item2Rec As Integer
    Public Quan2 As Double
    Public Item3 As String
    Public Item3Rec As Integer
    Public Quan3 As Double
    Public Item4 As String
    Public Item4Rec As Integer
    Public Quan4 As Double
    Public Item5 As String
    Public Item5Rec As Integer
    Public Quan5 As Double
    Public Item6 As String
    Public Item6Rec As Integer
    Public Quan6 As Double
    Public Item7 As String
    Public Item7Rec As Integer
    Public Quan7 As Double
    Public Item8 As String
    Public Item8Rec As Integer
    Public Quan8 As Double
    Public Item9 As String
    Public Item9Rec As Integer
    Public Quan9 As Double
    Public Item10 As String
    Public Item10Rec As Integer
    Public Quan10 As Double
    Public OrgGM As Double
    Public Landed As Decimal
    Public OnSale As Decimal
    Public List As Decimal
    Public SaleGM As Double
    Public PackPrice As Decimal
    Public OptionValue As Integer

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME As String = "InvKit"
    Private Const TABLE_INDEX As String = "KitStyleNo"

    Public Function Load(ByVal KeyVal As String, Optional ByVal KeyName As String = "") As Boolean
        ' Checks the database for a matching LeaseNo.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        ' Search for the Style
        If KeyName = "" Then
            DataAccess.Records_OpenIndexAt(KeyVal)
        ElseIf Left(KeyName, 1) = "#" Then
            ' This allows searching by AutoNumber - specialized to query by number
            ' since Access is exceptionally picky about quotation marks.
            DataAccess.Records_OpenFieldIndexAtNumber(Mid(KeyName, 2), KeyVal)
        ElseIf Left(KeyName, 1) = "@" Then
            DataAccess.Records_OpenFieldIndexAtDate(Mid(KeyName, 2), KeyVal)
        Else
            DataAccess.Records_OpenFieldIndexAt(KeyName, KeyVal)
        End If

        ' Move to the first record if we can, and return success.
        If DataAccess.Records_Available Then
            cDataAccess_GetRecordSet(DataAccess.RS)
            Load = True
        End If
    End Function

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub New()
        CDataAccess_Init
    End Sub

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        mDataAccess.SubClass = Me.mDataAccess
        mDataAccess.DataBase = GetDatabaseAtLocation(1)  ' Kits are only at location 1.
        mDataAccess.Table = TABLE_NAME
        mDataAccess.Index = TABLE_INDEX
    End Sub

    Public Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        Dim I As Integer
        On Error Resume Next
        KitStyleNo = IfNullThenNilString(Trim(RS("KitStyleNo").Value))
        KitSKU = IfNullThenNilString(Trim(RS("KitSKU").Value))
        Heading = IfNullThenNilString(Trim(RS("Heading").Value))
        MemoArea = IfNullThenNilString(Trim(RS("MemoArea").Value))

        For I = 1 To Setup_MaxKitItems
            'For I = 0 To Setup_MaxKitItems - 1
            'Item(I) = IfNullThenNilString(Trim(RS("Item" & I).Value))
            Item(I, IfNullThenNilString(Trim(RS("Item" & I).Value)))
            'ItemRec(I) = IfNullThenZero(RS("Item" & I & "Rec").Value)
            ItemRec(I, IfNullThenZero(RS("Item" & I & "Rec").Value))
            'Quantity(I) = IfNullThenZeroDouble(RS("Quan" & I).Value)
            Quantity(I, IfNullThenZeroDouble(RS("Quan" & I).Value))
        Next

        OrgGM = RS("OrgGM").Value
        Landed = RS("Landed").Value
        OnSale = RS("OnSale").Value
        List = RS("List").Value
        SaleGM = IfNullThenNilString(RS("SaleGM").Value)
        PackPrice = RS("PackPrice").Value
        OptionValue = RS("OptionValue").Value
    End Sub

    Public Function Item(ByVal Index As Integer, Optional ByVal I As String = "Get") As String
        If I = "Get" Then
            'Get property of vb6.0
            'Index = Index + 1
            Item = Choose(Index, Item1, Item2, Item3, Item4, Item5, Item6, Item7, Item8, Item9, Item10)
            Return Item
        Else
            'Let property of vb6.0
            'ChooseSet(Index, I, Item1, Item2, Item3, Item4, Item5, Item6, Item7, Item8, Item9, Item10)
            'Return Nothing
            Select Case Index
                Case 1
                    Item1 = I
                Case 2
                    Item2 = I
                Case 3
                    Item3 = I
                Case 4
                    Item = I
                Case 5
                    Item = I
                Case 6
                    Item = I
                Case 7
                    Item = I
                Case 8
                    Item = I
                Case 9
                    Item = I
                Case 10
                    Item = 10
            End Select
        End If
    End Function

    Public Function ItemRec(ByVal Index As Integer, Optional ByVal R As Integer = -1) As Integer
        If R = -1 Then
            'Get property of vb6.0
            ItemRec = Choose(Index, Item1Rec, Item2Rec, Item3Rec, Item4Rec, Item5Rec, Item6Rec, Item7Rec, Item8Rec, Item9Rec, Item10Rec)
            Return ItemRec
        Else
            'Let property of vb6.0
            'ChooseSet(Index, R, Item1Rec, Item2Rec, Item3Rec, Item4Rec, Item5Rec, Item6Rec, Item7Rec, Item8Rec, Item9Rec, Item10Rec)
            'Return Nothing
            Select Case Index
                Case 1
                    Item1Rec = R
                Case 2
                    Item2Rec = R
                Case 3
                    Item3Rec = R
                Case 4
                    Item4Rec = R
                Case 5
                    Item5Rec = R
                Case 6
                    Item6Rec = R
                Case 7
                    Item7Rec = R
                Case 8
                    Item8Rec = R
                Case 9
                    Item9Rec = R
                Case 10
                    Item10Rec = R
            End Select
        End If
    End Function

    Public Function Quantity(ByVal Index As Integer, Optional ByVal Q As Integer = -1) As Integer
        If Q = -1 Then
            'Get property of vb6.0
            Quantity = Choose(Index, Quan1, Quan2, Quan3, Quan4, Quan5, Quan6, Quan7, Quan8, Quan9, Quan10)
            Return Quantity
        Else
            'Let property of vb6.0
            'ChooseSet(Index, Q, Quan1, Quan2, Quan3, Quan4, Quan5, Quan6, Quan7, Quan8, Quan9, Quan10)
            'Return Nothing
            Select Case Index
                Case 1
                    Quan1 = Q
                Case 2
                    Quan2 = Q
                Case 3
                    Quan3 = Q
                Case 4
                    Quan4 = Q
                Case 5
                    Quan5 = Q
                Case 6
                    Quan6 = Q
                Case 7
                    Quan7 = Q
                Case 8
                    Quan8 = Q
                Case 9
                    Quan9 = Q
                Case 10
                    Quan10 = Q
            End Select
        End If
    End Function
End Class
