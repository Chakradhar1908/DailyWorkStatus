Public Class InvPull
    Private mPull As String   ' Transplanted from MainMenu.
    Dim StoreTrans
    Dim CC as integer
    Dim NoCost As Boolean
    Dim Billing(0 To Setup_MaxStores - 1) As String
    Dim StoreRec(0 To Setup_MaxStores - 1) As String

    Private Const ExtraFieldCount as integer = 3

    ' Matrix of store transfers.
    ' Values are: 0 = No transfers, 1 = Billing, 2 = Receiving, 3 = Both.
    'Dim TransferList(1 To Setup_MaxStores, 1 To Setup_MaxStores) as integer
    Dim TransferList(0 To Setup_MaxStores - 1, 0 To Setup_MaxStores - 1) as integer

    Public Property Pull() as integer
        Get
            Pull = Val(mPull)
        End Get
        Set(value as integer)
            mPull = value
        End Set
    End Property
End Class