Public Class EditPO
    Private Const VoidValue As String = "V"
    Private Const PreVoidValue As String = "v"
    Private Const CheckValue As String = "x"
    Private Const SavedValue As String = "X"

    Dim PoNo As Integer
    Dim SaleNo As String
    Dim PrintReport As Boolean

    ' Store information - Saved so we can switch between shipping + billing address.
    Dim StoreName As String
    Dim StoreAddress As String
    Dim StoreCity As String
    Dim StorePhone As String
    Dim StoreShipTo As String
    Dim StoreShipAdd As String
    Dim StoreShipCity As String
    Dim StoreShipPhone As String
    Dim StoreOrdDrv As String

    Dim Store1Name As String
    Dim Store1Addr As String
    Dim Store1City As String
    Dim Store1Tele As String

    Dim mCurrentCode As String        ' This is never set!
    Dim invDate As String
    Dim LastSaved As String
    Dim Lines As Integer                 ' Number of grid entries.
    Dim TotCost As Decimal

    Dim ReceivingCancelled As Boolean ' To stop processing mid-PO.
    Dim DefaultTagSize As String      ' Used for printing tags for stock POs
    Dim DefaultTicketPath As String   ' Used for printing tags for stock POs
    Dim LastRecvDate As Date          ' Temporarily global.. needed by loop containing call to ProcessSO.
    Dim ReceiveAll As Boolean         ' Temporarily global.. needed by loop containing call to ProcessSO.
    Dim FreezeRetails As Boolean

    Dim QtyBeforeEdit As Integer         ' Used to temporarily store quantity on edit for automatic adjustment of cost
    Dim KeepDept As String
    Dim PreSoldReceivingLabelsPrinted As Integer

    ' BFH20050202 Used to determine if the data on the page has changed so we don't
    '             carelessly discard data by moving onto a different PO.  Hopefully
    '             this will be used on <, >, <<, >>, Next, and Menu
    Private DataHasChanged As Boolean
    Private EditRow As Integer, EditCol As Integer

    Private Sub EditPO_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    'Public Function QuickViewPO(ByVal PoNo As String) As Boolean
    '    FindPo(PoNo)
    '    EnableFrame(Me, fraAddresses, False)
    '    EnableFrame(Me, fraInfo, False)
    '    EnableFrame(Me, fraGrid, False)
    '    EnableFrame(Me, fraMoveRecords, False)
    '    cmdApply.Enabled = False
    '    cmdNext.Enabled = False
    '    Show(1)
    'End Function

    Public Sub SaveTagPrintingOptions(ByRef TagSize As String, ByRef TicketPath As String)
        DefaultTagSize = TagSize
        DefaultTicketPath = TicketPath
    End Sub

End Class