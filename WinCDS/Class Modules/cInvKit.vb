Public Class cInvKit
    Public Item() As String
    Public Quantity() As Double
    Public KitStyleNo As String
    Public KitSKU As String
    Public Heading As String
    Public MemoArea As String
    Public Item1 As String
    Public Item1Rec As Integer
    Public Quan1 As Double
    Public Item2 As String
    Public Item2Rec As Integer
    Public Quan2 As Double
    Public Item3 As String
    Public Item3Rec As Integer
    Public Quan3 As Double
    Public Item4 As String
    Public Item4Rec As Integer
    Public Quan4 As Double
    Public Item5 As String
    Public Item5Rec As Integer
    Public Quan5 As Double
    Public Item6 As String
    Public Item6Rec As Integer
    Public Quan6 As Double
    Public Item7 As String
    Public Item7Rec As Integer
    Public Quan7 As Double
    Public Item8 As String
    Public Item8Rec As Integer
    Public Quan8 As Double
    Public Item9 As String
    Public Item9Rec As Integer
    Public Quan9 As Double
    Public Item10 As String
    Public Item10Rec As Integer
    Public Quan10 As Double
    Public OrgGM As Double
    Public Landed As Decimal
    Public OnSale As Decimal
    Public List As Decimal
    Public SaleGM As Double
    Public PackPrice As Decimal
    Public OptionValue As Integer

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME As String = "InvKit"
    Private Const TABLE_INDEX As String = "KitStyleNo"


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

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function
End Class
