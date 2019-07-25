Public Class cInstallment
    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess
    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function
    Public Sub Dispose()
        On Error Resume Next
        mDataAccess.Dispose
    End Sub

    Public Function Load(ByRef KeyVal As String, Optional ByRef KeyName As String = "") As Boolean
        ' Checks the database for a matching TransactionID.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        Load = False
        ' Search for the Style
        If KeyName = "" Then
            DataAccess.Records_OpenIndexAt(KeyVal)
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


End Class
