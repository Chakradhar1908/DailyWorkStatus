Public Class clsServicePartsOrder
    Public ServicePartsOrderNo As Long
    Public Store As Long
    Public MarginLine As Long
    Public Style As String
    Public Desc As String
    Public Vendor As String
    Public VendorAddress As String
    Public VendorCity As String ' city/state/zip..
    Public VendorTele As String
    Public ServiceOrderNo As Long
    Public DateOfClaim As Date
    Public Status As String
    Public Notes As String

    Public ChargeBackType As Long
    Public ChargeBackAmount As Currency

    Public NoteID As Long

    Public InvoiceNo As String
    Public InvoiceDate As String

    Public Paid As Boolean

    Private WithEvents mDataAccess As CDataAccess
    Implements CDataAccess

    Private Const TABLE_NAME As String = "ServicePartsOrder"
    Private Const TABLE_INDEX As String = "ServicePartsOrderNo"

    Public Function Load(ByVal KeyVal As String, Optional ByVal KeyName As String) As Boolean
        ' Checks the database for a matching LeaseNo.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        ' Search for the Style
        If KeyName = "" Then
            DataAccess.Records_OpenIndexAt KeyVal
  ElseIf Left(KeyName, 1) = "#" Then
            ' This allows searching by AutoNumber - specialized to query by number
            ' since Access is exceptionally picky about quotation marks.
            DataAccess.Records_OpenFieldIndexAtNumber Mid(KeyName, 2), KeyVal
  ElseIf Left(KeyName, 1) = "@" Then
            DataAccess.Records_OpenFieldIndexAtDate Mid(KeyName, 2), KeyVal
  Else
            DataAccess.Records_OpenFieldIndexAt KeyName, KeyVal
  End If

        ' Move to the first record if we can, and return success.
        If DataAccess.Records_Available Then Load = True
    End Function

End Class
