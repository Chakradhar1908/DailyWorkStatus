Public Class cTransaction
    Public ArNo As String
    Private mTransactionID As Integer ' property get only?
    Public LastName As String
    Public TransDate As Date
    Public MailIndex As String
    Public TransType As String
    Public Charges As Decimal
    Public Credits As Decimal
    Public Balance As Decimal
    Public Receipt As String

    Public DataBase As String
    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess
    Private Const TABLE_NAME As String = "Transactions"
    Private Const TABLE_INDEX As String = "TransactionID"

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub Dispose()
        On Error Resume Next
        mDataAccess.Dispose()
    End Sub

    Public Function Load(ByRef KeyVal As String, Optional ByRef KeyName As String = "") As Boolean
        ' Checks the database for a matching TransactionID.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        Load = False
        ' Search for the Style
        If KeyName = "" Then
            DataAccess.Records_OpenFieldIndexAtNumber("TransactionID", KeyVal)
        ElseIf Left(KeyName, 1) = "#" Then
            ' This allows searching by AutoNumber - specialized to query by number
            ' since Access is exceptionally picky about quotation marks.
            DataAccess.Records_OpenFieldIndexAtNumber(Mid(KeyName, 2), KeyVal)
        Else
            DataAccess.Records_OpenFieldIndexAt(KeyName, KeyVal)
        End If

        ' Move to the first record if we can, and return success.
        If DataAccess.Records_Available Then Load = True
    End Function

    Public Function Save(Optional ByRef ErrDesc As String = "") As Boolean
        ErrDesc = ""
        Save = True
        On Error GoTo NoSave
        ' This instructs the class (in one simple call) to save its data members to the database.
        If DataAccess.CurrentIndex <= 0 Then            ' If we're already using the current record,
            DataAccess.Records_OpenFieldIndexAtNumber("TransactionID", TransactionID)  'there's no reason to re-open it.
        End If
        If DataAccess.Record_Count = 0 Then
            DataAccess.Records_Add()      ' Record not found.  This means we're adding a new one.
        End If

        DataAccess.Record_Update()      ' Then load our data into the recordset.
        DataAccess.Records_Update()     ' And finally, tell the class to save the recordset.
        Exit Function

NoSave:
        ErrDesc = Err.Description
        Err.Clear()
        Save = False
    End Function

    Public ReadOnly Property TransactionID() As Integer
        Get
            TransactionID = mTransactionID
            If TransactionID = 0 Then TransactionID = -1
        End Get
    End Property


End Class
