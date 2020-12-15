Public Class clsServiceNotes
    Public ServiceNoteID As Integer
    Public ServiceCall As Integer
    Public MarginNo As Integer
    Public Note As String
    Public NoteDate As Date
    Public NoteType As Integer
    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess
    Private Const TABLE_NAME = "ServiceNotes"
    Private Const TABLE_INDEX = "ServiceNoteID"

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub New()
        CDataAccess_Init
    End Sub

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        mDataAccess.SubClass = Me.mDataAccess
        mDataAccess.DataBase = GetDatabaseAtLocation()
        mDataAccess.Table = TABLE_NAME
        mDataAccess.Index = TABLE_INDEX
    End Sub

    Public Sub Dispose()
        On Error Resume Next
        mDataAccess.Dispose()
    End Sub

    Public Function cDataAccess_SuperClass() As CDataAccess
        cDataAccess_SuperClass = mDataAccess
    End Function

    Public Sub cDataAccess_SetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        '  rs("ServiceNoteID") = ServiceNoteID
        RS("ServiceCall").Value = IfNullThenZero(ServiceCall)
        RS("MarginNo").Value = IfNullThenZero(MarginNo)
        RS("Note").Value = IfNullThenNilString(Trim(Note))
        RS("NoteDate").Value = Today   ' Always update to date last saved.
        RS("NoteType").Value = IfNullThenZero(NoteType)
    End Sub

    Public Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        ServiceNoteID = RS("ServiceNoteID").Value
        ServiceCall = IfNullThenZero(RS("ServiceCall").Value)
        MarginNo = IfNullThenZero(RS("MarginNo").Value)
        Note = IfNullThenNilString(Trim(RS("Note").Value))
        NoteDate = RS("NoteDate").Value
        NoteType = IfNullThenZero(RS("NoteType").Value)
    End Sub

    Public Function NoteTypeString() As String
        If IsNothing(NoteType) Then NoteType = 0
        Select Case NoteType
            Case 0
                NoteTypeString = "Note"
            Case 1
                NoteTypeString = "Parts Order"
            Case Else
                NoteTypeString = "Strange Note (" & NoteType & ")"
        End Select
    End Function

    Public Function Save() As Boolean
        On Error GoTo NoSave
        ' This instructs the class (in one simple call) to save its data members to the database.
        If DataAccess.Record_Count = 0 Then
            ' Record not found.  This means we're adding a new one.
            DataAccess.Records_Add()
        End If
        ' Then load our data into the recordset.
        DataAccess.Record_Update()
        ' And finally, tell the class to save the recordset.
        DataAccess.Records_Update()
        Exit Function

NoSave:
        Err.Clear()
        Save = False
    End Function

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

    Private Sub mDataAccess_RecordUpdated()
        ServiceNoteID = mDataAccess.Value("ServiceNoteID")
    End Sub

End Class
