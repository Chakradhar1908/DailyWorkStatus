Public Class CDbAccessGeneral
    Private WithEvents mRs_External As ADODB.Recordset
    'Private mRs_External As ADODB.Recordset
    Public mConnection As ADODB.Connection
    Public mSQL As String
    Dim Mrs As ADODB.Recordset
    Public Event UpdateFailed(RS As ADODB.Recordset, ByRef Cancel As Boolean)
    Public Event GetRecordNotFound()
    Public Event GetRecordEvent(RS As ADODB.Recordset)
    Public Event SetRecordEvent(RS As ADODB.Recordset)
    Public Event RecordUpdated(RS As ADODB.Recordset)

    ' if 'SetNew:=True' will always create a new record
    Public Function getRecordset(Optional ByVal Always As Boolean = True, Optional ByVal SetNew As Boolean = False, Optional ByVal QuietErrors As Boolean = False, Optional ByVal ErrMsg As String = "", Optional ByVal ProgressForm As Object = False) As ADODB.Recordset
        On Error GoTo AnError
        'Dim rs As New ADODB.Recordset
        mRs_External = New ADODB.Recordset
        'mRs_External.ActiveConnection = mConnection
        mRs_External.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        'mRs_External.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        'mRs_External.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic
        '.Properties("Update Resync") = adResyncAutoIncrement
        'mRs_External.Source = mSQL
        mRs_External.Open(mSQL, mConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic,)

        If Always Then
            If (mRs_External.RecordCount = 0) Or SetNew Then
                mRs_External.AddNew()
            Else
                '.Edit ' for DAO only
            End If
        End If
        mRs_External.ActiveConnection = Nothing
        getRecordset = mRs_External
        mRs_External = Nothing
        Exit Function

AnError:
        ' Hide or change this error message before distributing the program!
        ' or Not...  Might as well let the clients tell us where it broke and how!
        If Not QuietErrors Then
            If ErrMsg = "None" Then
                ' Don't output a message!
            ElseIf ErrMsg <> "" Then
                ErrMsg = Replace(ErrMsg, "$EDESC", Err.Description)
                ErrMsg = Replace(ErrMsg, "$ENO", Err.Number)
                ErrMsg = Replace(ErrMsg, "$ESRC", Err.Source)
                MsgBox("Database Error: " & ErrMsg, vbCritical, "Error")
            Else
                Dim T As String
                T = "getRecordSet Failed:" & vbCrLf
                T = T & vbCrLf
                T = T & mSQL & vbCrLf
                T = T & vbCrLf
                T = T & "ERROR [" & Err.Number & "]:" & Err.Description


                CheckStandardErrors() ' Bookmark/updateable query
            End If
        End If

        DisposeDA(mRs_External, getRecordset)
    End Function

    Public Function SetRecord(Optional ByVal CreateNew As Boolean = False) As Boolean
        On Error GoTo AnError
        Mrs = New ADODB.Recordset

        Mrs.ActiveConnection = mConnection

        Mrs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Mrs.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        '.Mode = adModeReadWrite '?????
        '       .Properties("Update Resync") = adResyncAutoIncrement
        Mrs.LockType = ADODB.LockTypeEnum.adLockOptimistic ' adLockBatchOptimistic
        Mrs.Source = mSQL
        Mrs.Open()

        If ((Mrs.RecordCount = 0) Or CreateNew) Then
            Mrs.AddNew()
        Else
            '.Edit ' for DAO only
        End If
        RaiseEvent SetRecordEvent(Mrs)
        'Recordset_DebugPrint(mrs)
        Mrs.Update() ' .UpdateBatch
        RaiseEvent RecordUpdated(Mrs)
        Mrs.ActiveConnection = Nothing
        Mrs.Close()
        Exit Function

AnError:
        ' If (GetVendorNameSucceeded = True) Then
        MsgBox("SetRecord Failed [" & Err.Number & "]:  Error Updating Database" & Err.Description)
        SetRecord = False
        Exit Function
    End Function


    Public Sub dbClose()
        On Error Resume Next
        mConnection.Close()
        mConnection = Nothing
        Mrs.Close
        Mrs = Nothing
        Err.Clear()
    End Sub
    Public Function dbOpen(Optional ByVal DBName As String = "") As Boolean
        ' --> Below databse connection code has been commented. It will be implemented in app.config file. <--

        Dim TryCount as integer, T As Date
        'TrackDataAccess
        On Error GoTo AnError
Retry:
        Dim Rst As ADODB.Recordset, ConnString As String

        If DBName = "" Then DBName = GetDatabaseAtLocation()
        ConnString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName '& PasswordProtectedDatabaseString & ";"

        Dim BackingUpAlert As Date
        BackingUpAlert = DateAdd("s", 25, Now)
        Do While BackupSemaforeFile(ItExists:=True)
            Application.DoEvents()

            If DateAfter(Now, BackingUpAlert, True, "s") Then
                MsgBox("Please wait a moment..." & vbCrLf2 & "An Administrator is currently backing up the database..." & vbCrLf2 & "This should only take an additional moment.", vbOKOnly + vbCritical, "Operation Delayed")
            End If
        Loop


        mConnection = New ADODB.Connection
        mConnection.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        mConnection.Open(ConnString)

        dbOpen = True

        Exit Function
AnError:
        Dim N as integer, D As String
        N = Err.Number
        D = Err.Description

        Select Case N
        'BFH20060804 - added handler for -2147467259
        ' Cannot open database:  it may not be a dat base that your application recognizes, or the file may be corrupt
            Case -2147467259
                If InStr(D, "installable ISAM") <> 0 Then
                    ErrMsg("Database Password Problem.")
                End If

                'WAIT 2 SECONDS, TRY AGAItN.. up to 5 times
                If TryCount <= 15 Then
                    T = DateAdd("s", 2, Now)
                    Do While DateBefore(Now, T, True, "s") : Application.DoEvents() : Loop
                    TryCount = TryCount + 1
                    Resume Retry
                End If

                ' Add handler here for other errors a store may encounter.
        End Select
        ' This handles all unhandled and un-Resume-d errors

        Dim RR As String
        RR = "CDbAccessGeneral.dbOpen(DBName=" & DBName & ")" & vbCrLf2

        If Not FileExists(DBName) Then
            MsgBox(RR & "Database Not Found: " & DBName & vbCrLf & "Error number " & N & ": " & D)
            ReportError("CDbAccessGeneral.dbOpen - Database not found")
        Else
            MsgBox(RR & "Error Opening Database: " & DBName & vbCrLf & "Error number " & N & ": " & D)
            ReportError("CDbAccessGeneral.dbOpen - Error Opening Database")
        End If
    End Function

    Public Function UpdateRecordSet(ByRef RS As ADODB.Recordset, Optional ByVal SilentErrors As Boolean = False) As ADODB.Recordset
        On Error GoTo AnError
        RS.ActiveConnection = mConnection
TryAgain:
        mConnection.Errors.Clear()
        Err.Clear()
        RS.UpdateBatch()
        RS.ActiveConnection = Nothing
        Exit Function

AnError:
        If SilentErrors Then Exit Function
        Dim objError As ADODB.Error
        Dim strError As String
        If mConnection.Errors.Count > 0 Then
            For Each objError In mConnection.Errors
                If objError.NativeError = 66913278 Then
                    ' Disk or Network error
                    ' This should allow retry/fail.
                    ' Unless it causes double updates..... maybe the filter will help.
                    If MsgBox("Disk or network error.  Try again?", vbCritical + vbYesNo, "Error") = vbYes Then
                        'RS.Filter = adFilterConflictingRecords
                        RS.Filter = ADODB.FilterGroupEnum.adFilterConflictingRecords
                        Resume TryAgain
                    End If
                End If

                If Allow_ADODB_Errors And objError.Number = -2147467259 Then
                    ' Don't do anything!  We expected this error.
                    ' This is used during v8-9 conversion to ignore errors
                    ' while inserting invalid records - usually due to blank
                    ' records or other invalid structure issues.
                    ' This same error number appears to catch Disk or Network Errors.
                Else
                    strError = strError & "Error #" & objError.Number &
                " " & objError.Description & vbCrLf &
                "NativeError: " & objError.NativeError & vbCrLf &
                "SQLState: " & objError.SQLState & vbCrLf &
                "Reported by: " & objError.Source & vbCrLf &
                "Help file: " & objError.HelpFile & vbCrLf &
                "Help Context ID: " & objError.HelpContext
                End If
            Next
        End If
        'RS.Filter = adFilterConflictingRecords  ' This may have to change to PendingRecords.
        RS.Filter = ADODB.FilterGroupEnum.adFilterConflictingRecords
        Dim Cancel As Boolean
        RaiseEvent UpdateFailed(RS, Cancel)
        If Cancel = True Then Exit Function
        If strError = "" Then Resume Next
        MsgBox("UpdateRecordset Failed: " & strError, vbCritical, "Error")
        Do While Not RS.EOF
            Debug.Print(RS.Status) 'check the status

            RS.MoveNext() 'fix error
        Loop
    End Function

    Public Property SQL As String
        Get
            SQL = mSQL
        End Get
        Set(value As String)
            mSQL = value
        End Set
    End Property
    Public Function GetRecord() As Boolean
        Mrs = New ADODB.Recordset
        Mrs.ActiveConnection = mConnection
        Mrs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Mrs.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        Mrs.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic
        '        .Properties("Update Resync") = adResyncAutoIncrement
        '        .ActiveConnection = mConnection
        '.CursorLocation = adUseClient
        '.CursorLocation = adUseServer
        '        .CursorType = adOpenStatic
        '        .LockType = adLockReadOnly

        Mrs.Source = mSQL
        Mrs.Open()
        Mrs.ActiveConnection = Nothing

        If (Mrs.RecordCount = 0) Then
            RaiseEvent GetRecordNotFound()
            GetRecord = False
            Mrs.Close()
            Mrs = Nothing
            Exit Function
        End If

        Dim NewArNo As String
        ' GDM BEGINNING OF CHANGE 3/29/2001
        ' Now move to the correct record for display
        ' And clear out the f_strDirection request
        ' This is actually still used....
        If ArCard.UsingMoveButton Then
            Dim Current As String
            Select Case ArCard.f_strDirection
                Case "First"    'Get to the first record
                    RS.MoveFirst
                    ArCard.f_strDirection = ""
                Case "Last"     'Move to the last(right) record
                    RS.MoveLast
                    ArCard.f_strDirection = ""
                Case "Previous" 'Move to the previous record
                    RS.MoveLast
                    'Current = RS!ArNo
                    Current = RS.Fields("ArNo").Value
                    RS.MovePrevious
                    Do While Not RS.BOF
                        If RS.Fields("ArNo").Value <> Current Then Exit Do
                        RS.MovePrevious
                    Loop
                    If RS.BOF Then RS.MoveFirst
                    ArCard.f_strDirection = ""
                Case "Next"     'Move to the next record
                    RS.MoveFirst
                    Current = RS.Fields("ArNo").Value
                    RS.MoveNext
                    Do While Not RS.EOF
                        If RS.Fields("ArNo").Value <> Current Then Exit Do
                        RS.MoveNext
                    Loop
                    If RS.EOF Then RS.MoveLast
                    ArCard.f_strDirection = ""
            End Select
            ' Get the new arno and save it
            NewArNo = RS.Fields("ArNo").Value
            ' create a new sql to that record
            mSQL =
              "SELECT InstallmentInfo.*" _
              & " From InstallmentInfo" _
              & " WHERE InstallmentInfo.Status <> '" & arST_Void & "'" _
              & " AND (((InstallmentInfo.ArNo)  =""" & ProtectSQL(NewArNo) & """))"

            ArCard.UsingMoveButton = False
        End If

        ' GDM END OF CHANGE
        GetRecord = True
        RaiseEvent GetRecordEvent(RS)

        On Error Resume Next ' In case GetRecordEvent closed the recordset.
        Mrs.Close()
        GetRecord = True
        Exit Function

AnError:
        ' If (GetVendorNameSucceeded = True) Then
        MsgBox("GetRecord Failed:  database not found")
        GetRecord = False
        Exit Function
    End Function
    Public ReadOnly Property RS() As ADODB.Recordset
        Get
            RS = Mrs
        End Get
    End Property
End Class
