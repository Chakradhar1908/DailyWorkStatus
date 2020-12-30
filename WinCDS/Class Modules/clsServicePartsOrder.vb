Public Class clsServicePartsOrder
    Public ServicePartsOrderNo As Integer
    Public Store As Integer
    Public MarginLine As Integer
    Public Style As String
    Public Desc As String
    Public Vendor As String
    Public VendorAddress As String
    Public VendorCity As String ' city/state/zip..
    Public VendorTele As String
    Public ServiceOrderNo As Integer
    Public DateOfClaim As Date
    Public Status As String
    Public Notes As String

    Public ChargeBackType As Integer
    Public ChargeBackAmount As Decimal

    Public NoteID As Integer

    Public InvoiceNo As String
    Public InvoiceDate As String

    Public Paid As Boolean

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME As String = "ServicePartsOrder"
    Private Const TABLE_INDEX As String = "ServicePartsOrderNo"

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub New()
        CDataAccess_Init
    End Sub

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        mDataAccess.SubClass = Me.mDataAccess
        mDataAccess.DataBase = GetDatabaseAtLocation()
        mDataAccess.Table = TABLE_NAME
        mDataAccess.Index = TABLE_INDEX
    End Sub

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
        If DataAccess.Records_Available Then Load = True
    End Function

    Public Function Save() As Boolean
        On Error GoTo NoSave
        ' This instructs the class (in one simple call) to save its data members to the database.
        If DataAccess.Record_Count = 0 Then
            ' Record not found.  This means we're adding a new one.
            DataAccess.Records_Add()
        End If
        ' Then load our data into the recordset.
        DataAccess.Record_Update()
        ' And finally, tell the class to save the recordset.
        DataAccess.Records_Update()
        Exit Function

NoSave:
        Err.Clear()
        Save = False
    End Function

End Class
