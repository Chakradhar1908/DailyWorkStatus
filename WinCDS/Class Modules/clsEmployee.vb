Public Class clsEmployee
    ' Notes:
    '  Since password encryption can only be done in VB, and we need to log a user in
    '  using only a password, most logins will be done by cycling through the user base
    '  until a key/password value matches the input.  as integer as we only have a couple
    '  hundred users per store, this shouldn't be much of an issue.

    ' Field declarations.
    Public ID as integer  ' Autonumber
    Public SalesID As String
    Public CommRate As String
    Public Active As Boolean

    ' Encryption keys and Encrypted fields require property code to maintain.
    Public LastName As String
    Public Password As String
    Public Privs As String

    Public CommTable As String

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME = "Employees"
    Private Const TABLE_INDEX = "ID"

    Public Sub New()
        CDataAccess_Init
    End Sub

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        With mDataAccess
            .SubClass = mDataAccess
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
        RS("LastName").Value = Trim(Left(LastName, 40))
        RS("SalesID").Value = Trim(Left(SalesID, 3))
        RS("CommRate").Value = Trim(Left(CommRate, 8))
        RS("Pwd").Value = Encrypt(EncryptionKey, Password)
        RS("Privs").Value = Encrypt(EncryptionKey, Privs)
        RS("Active").Value = Active
        RS("CommTable").Value = CommTable
    End Sub

    Public Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        ID = RS("ID").Value
        LastName = Trim(IfNullThenNilString(RS("LastName").Value))
        SalesID = Trim(IfNullThenNilString(RS("SalesID").Value))
        CommRate = Trim(IfNullThenNilString(RS("CommRate").Value))
        Password = Decrypt(EncryptionKey, IfNullThenNilString(RS("Pwd").Value))
        Privs = Decrypt(EncryptionKey, IfNullThenNilString(RS("Privs").Value))
        Active = RS("Active").Value
        CommTable = RS("CommTable").Value
    End Sub

    Private Function EncryptionKey() As String
        EncryptionKey = LastName
    End Function

End Class
