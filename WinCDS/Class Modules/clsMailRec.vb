Public Class clsMailRec
    Public Index as integer
    Public Last As String
    Public First As String
    Public Address As String
    Public AddAddress As String
    Public City As String
    Public Zip As String
    Public Tele As String
    Public Tele2 As String
    Public PhoneLabel1 As String
    Public PhoneLabel2 As String
    Public Special As String
    Public MailType As String
    Public CustType As String
    Public TypeIndex as integer
    Public Blank As String
    Public Email As String
    Public Business As Boolean
    Public CreditCard As String
    Public ExpDate As String
    Public TaxZone as integer
    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess
    Private Const TABLE_INDEX = "Index"

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function
    Public Function Load(ByVal KeyVal As String, Optional ByVal KeyName As String = "") As Boolean
        ' Checks the database for a matching LeaseNo.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        ' Search for the Style
        Load = False
        If KeyName = "" Then
            DataAccess.Records_OpenFieldIndexAtNumber(TABLE_INDEX, KeyVal)
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

End Class
