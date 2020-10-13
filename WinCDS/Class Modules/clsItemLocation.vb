Public Class clsItemLocation
    ' Field declarations.
    Public ItemLocationID As Integer
    Public StyleNo As String
    Public LocationBarcode As String
    Public StockDate As Date
    Public Cost As Decimal      ' XX
    Public SerialNo As String
    Public Status As ItemLocationStatus
    Public Location As Integer
    Public SaleNo As String
    Public PoNo As Integer
    Public OrderType As String  ' XX
    Public PullDate As Date     ' XX
    Public DelDate As Date

    Public RN As Integer           ' ++
    Public Detail As Integer       ' ++
    Public MarginLine As Integer   ' ++

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME As String = "ItemLocation"
    Private Const TABLE_INDEX As String = "ItemLocationID"

    Public Sub New()
        CDataAccess_Init
    End Sub

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        mDataAccess.SubClass = Me.mDataAccess
        mDataAccess.DataBase = GetDatabaseInventory()
        mDataAccess.Table = TABLE_NAME
        mDataAccess.Index = TABLE_INDEX
    End Sub

    Public Property Bld() As String
        Get
            Dim A As String, B As String, C As String
            DecodeLocation(LocationBarcode, Bld, A, B, C)
        End Get
        Set(value As String)
            Dim A As String, B As String, C As String, D As String
            DecodeLocation(LocationBarcode, A, B, C, D)
            A = Left(value, ItemLocBldMaxLen)
            LocationBarcode = EncodeLocation(A, B, C, D)
        End Set
    End Property

    Public Property Row() As String
        Get
            Dim A As String, B As String, C As String
            DecodeLocation(LocationBarcode, A, Row, B, C)
        End Get
        Set(value As String)
            Dim A As String, B As String, C As String, D As String
            DecodeLocation(LocationBarcode, A, B, C, D)
            B = Left(value, ItemLocRowMaxLen)
            LocationBarcode = EncodeLocation(A, B, C, D)
        End Set
    End Property

    Public Property Lvl() As String
        Get
            Dim A As String, B As String, C As String
            DecodeLocation(LocationBarcode, A, B, Lvl, C)
        End Get
        Set(value As String)
            Dim A As String, B As String, C As String, D As String
            DecodeLocation(LocationBarcode, A, B, C, D)
            C = Left(value, ItemLocLvlMaxLen)
            LocationBarcode = EncodeLocation(A, B, C, D)
        End Set
    End Property

    Public Property Bay() As String
        Get
            Dim A As String, B As String, C As String
            DecodeLocation(LocationBarcode, A, B, C, Bay)
        End Get
        Set(value As String)
            Dim A As String, B As String, C As String, D As String
            DecodeLocation(LocationBarcode, A, B, C, D)
            D = Left(value, ItemLocBayMaxLen)
            LocationBarcode = EncodeLocation(A, B, C, D)
        End Set
    End Property
End Class
