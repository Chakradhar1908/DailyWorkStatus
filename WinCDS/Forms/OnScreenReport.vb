Public Class OnScreenReport
    Dim PoNo As Integer  ' Saved between calls to MakePO.
    Dim Margin As New CGrossMargin
    Dim MarginNo As Integer
    Dim Row As Integer
    Private balRow As Integer
    Dim Mail As MailNew
    Dim LastName As String
    Dim Tele As String
    Public Index As String
    Dim OrdTotal As Decimal
    Dim TotDue As Decimal
    Dim TaxBackedOut As Boolean
    Dim KitStart As Integer, IsKit As Boolean, KitTotalCost As Decimal

    Private mCurrentLine As Integer 'Current line selected
    Dim Counter As Integer
    Dim mLoading As Boolean
    Dim Lines As Integer

    ' These need to be replaced!  We can do the same thing better with hidden grid columns.
    Dim Quantity(500) As Object
    Dim InvRn(500) As Object
    Dim Cost(500) As Object
    Dim Freight(500) As Object
    Dim Depts(500) As Object
    Dim Vends(500) As Object
    Dim DetailRec(500) As Object

    Public Balance As Decimal, TotTax As Decimal
    Dim SaleNo As String
    Dim Detail As Integer

    'Dim NoOnHand As String             ' Was never used..
    Dim FirstTime As Boolean
    'Private AddedInventory As Boolean  ' Was set but never used..
    Dim LastMfg As String

    Dim LastSale As String                         ' For determining which PO items go on.
    Dim Sales As String
    Dim TaxLoc As Integer
    Dim TaxRate As Integer
    Dim Rate As Object
    Dim SalesTax As Boolean
    Dim PriceChg As String
    Dim SubBalance As Decimal
    Dim PriorBal As Decimal
    Dim NonTaxable As Decimal
    Public LeaveCreditBalance As Boolean

    Private WithEvents MailCheckRef As MailCheck
    Private SaleFound As Boolean

    Dim WasDelSale As Boolean, WriteOutAddedItems As Boolean, WriteOutRemovedAllUndelivered As Boolean
    Dim AskedForTaxRate As Boolean

    Const AllowAdjustDel As Boolean = True
    Const MaxAdjustments As Integer = 30

End Class