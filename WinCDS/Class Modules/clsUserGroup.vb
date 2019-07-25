Public Class clsUserGroup
    ' User Group class object..

    ' Field declarations.
    Public ID as integer  ' Autonumber
    Public Abbrev As String
    Public SystemGroup As Boolean
    Public GroupName As String
    Public Privs As String

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME = "UserGroups"
    Private Const TABLE_INDEX = "ID"

    '  sql = "CREATE TABLE UserGroups " & _
    '        "(ID int identity, " & _
    '        "GroupName varchar(40), " & _
    '        "Abbrev char(1), " & _
    '        "Privs varchar(255), " & _
    '        "SystemGroup YesNo, " & _
    '        "CONSTRAINT UserGroups_PrimaryKey PRIMARY KEY (ID))"

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function
End Class
