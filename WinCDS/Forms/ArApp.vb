Public Class ArApp
    Dim HousingCode As Integer
    Dim HousingType As Integer
    Dim BankCode As Integer
    Dim PayCode1When As Integer
    Dim PayCode1Type As Integer
    Dim PayCodeType As Integer
    Dim PayCode2When As Integer
    Dim PayCode2Type As Integer
    Dim CustomerLast As String
    Dim MailIndex As String
    Dim ArNo As String
    Dim mArNo As String

    Public SS As String ' Used by ARCard

    Dim FH_NORM As Integer, FH_EXP As Integer

    Private WithEvents mDBAccessArApp As CDbAccessGeneral

    Public Sub GetApp(Optional ByVal mR As Integer = 0, Optional ByVal AN As String = "")
        mArNo = "-1"
        If AN = "" Then ArNo = ArCard.ArNo Else ArNo = AN
        If mR = 0 Then
            MailIndex = ArCard.MailRec
            mDBAccessArApp_Init(MailIndex, True)
        Else
            MailIndex = mR
            ArNo = "#"
            mDBAccessArApp_Init("#" & MailIndex, True)
        End If
        mDBAccessArApp.GetRecord()    ' this gets the record
        mDBAccessArApp.dbClose()
        mDBAccessArApp = Nothing

        If mArNo = "-1" Then 'not found
            Exit Sub
        End If
    End Sub

    Private Sub mDBAccessArApp_Init(ByVal Tid As String, Optional ByVal IsMailIndex As Boolean = False)
        mDBAccessArApp = New CDbAccessGeneral
        mDBAccessArApp.dbOpen(GetDatabaseAtLocation())
        If ArMode("A", "P", "Edit") Then
            mDBAccessArApp.SQL = "SELECT * From ArApp WHERE MailIndex=""" & ProtectSQL(Tid) & """"
        Else 'edit old
            If Microsoft.VisualBasic.Left(ArNo, 1) = "#" Or IsMailIndex Then
                mDBAccessArApp.SQL = "SELECT * From ArApp WHERE MailIndex=""" & ProtectSQL(IIf(Microsoft.VisualBasic.Left(Tid, 1) = "#", Mid(Tid, 2), Tid)) & """"
            Else
                mDBAccessArApp.SQL = "SELECT * From ArApp WHERE ArApp.ArNo=""" & ProtectSQL(Tid) & """"
            End If
        End If
    End Sub

End Class