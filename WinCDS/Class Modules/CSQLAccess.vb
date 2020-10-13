Public Class CSQLAccess
    Private mDataAccess As CDataAccess
    'Implements CDataAccess

    Public Sub New()
        CDataAccess_Init
    End Sub

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        mDataAccess.SubClass = Me.mDataAccess
        mDataAccess.DataBase = GetDatabaseAtLocation()
        '.Table = TABLE_NAME
        '.Index = TABLE_INDEX
    End Sub
End Class
