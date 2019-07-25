Public Class clsServiceOrder
    Public ServiceOrderNo as integer
    Public LastName As String
    Public Telphone As String
    Public MailIndex as integer
    Public SaleNo As String
    Public ServiceOnDate As String
    Public DateOfClaim As Date
    Public Status As String
    Public QuickCheck As String
    Public Item As String
    Public Complaint As String
    Public StoreAction As String
    Public SOType As String
    Public Mfg As String
    Public InvoiceNo As String
    Public Detail As String
    Public StopStart As String
    Public StopEnd As String

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME = "Service"
    Private Const TABLE_INDEX = "ServiceOrderID"

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
            DataAccess.Records_Add
        End If
        ' Then load our data into the recordset.
        DataAccess.Record_Update
        ' And finally, tell the class to save the recordset.
        DataAccess.Records_Update
        Exit Function

NoSave:
        Err.Clear()
        Save = False
    End Function
    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function
End Class
