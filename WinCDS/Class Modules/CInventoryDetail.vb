Public Class CInventoryDetail
    Private WithEvents mDataAccess As CDataAccess
    Public Style As String
    Public Lease1 As String
    Public Name As String
    Public Misc As String
    Public DDate1 As String
    Public Trans As String
    Public AmtS1 As Single
    Public Ns1 As Single
    Public SO1 As Single
    Public LAW As Single
    Public ItemCost As Decimal  ' BFH20060124
    'Private Loc(1 To Setup_MaxStores_DB)
    Private Loc(0 To Setup_MaxStores_DB - 1)
    Public DetailID as integer  ' MJK Autonumber
    Public MarginRn as integer
    Public InvRn as integer
    Public Store as integer

    'Implements CDataAccess
    Public Function Load(ByVal KeyVal As String, Optional ByRef KeyName As String = "") As Boolean
        ' Checks the database for a matching LeaseNo.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        ' Search for the Style
        Load = False
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
    Public Sub Dispose()
        On Error Resume Next
        mDataAccess.Dispose
    End Sub
    Public Sub SetLocationQuantity(ByVal Location as integer, ByVal Quantity As Double)
        If Not LocationValid(Location) Then Exit Sub
        Loc(Location - 1) = Quantity
    End Sub
    Private Function LocationValid(ByVal Loc as integer) As Boolean
        If Loc <= 0 Or Loc > Setup_MaxStores_DB - 1 Then
            LocationValid = False
            Exit Function
        End If
        LocationValid = True
    End Function

End Class
