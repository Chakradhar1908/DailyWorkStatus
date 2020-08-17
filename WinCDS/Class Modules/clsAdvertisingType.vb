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
End Class
