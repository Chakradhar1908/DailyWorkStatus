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

End Class
