Public Class clsAdvertisingType
    Public ID as integer  ' Autonumber
    Public AdType As String
    Public OldTypeID as integer
    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess
    Private Const TABLE_NAME = "AdvertisingType"
    Private Const TABLE_INDEX = "AdvertisingTypeID"

    Public Sub New()
        CDataAccess_Init()
    End Sub

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        With mDataAccess
            .SubClass = Me.mDataAccess
            .DataBase = GetDatabaseAtLocation()
            .Table = TABLE_NAME
            .Index = TABLE_INDEX
        End With
    End Sub

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Private Sub cDataAccess_SetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        RS("AdvertisingType").Value = Trim(AdType)
        RS("OldTypeID").Value = OldTypeID
    End Sub

    Public Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        ID = RS("AdvertisingTypeID").Value
        AdType = IfNullThenNilString(Trim(RS("AdvertisingType").Value))
        OldTypeID = RS("OldTypeID").Value
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
        Else
            DataAccess.Records_OpenFieldIndexAt(KeyName, KeyVal)
        End If

        ' Move to the first record if we can, and return success.
        If DataAccess.Records_Available Then Load = True
    End Function

End Class
