Module modConfigTable
    Public Function GetConfigTableValue(ByVal FieldName As String, Optional ByVal DefaultValue As String = "") As String
        '::::GetConfigTableValue
        ':::SUMMARY
        ': Return global config settings
        ':::DESCRIPTION
        ': Return values based on key from config table.  Returns default value if not found.
        ':::PARAMETERS
        ': - FieldName - Indicates the String value.
        ': - DefaultValue - Indicates the String value.
        ':::RETURN
        ': String
        Dim SQL As String, RS As ADODB.Recordset
        SQL = "SELECT * FROM Config WHERE [FieldName] = '" & FieldName & "'"
        RS = GetRecordsetBySQL(SQL, , GetDatabaseInventory, True)
        On Error GoTo BadConfigTable
        If Not RS.EOF Then
            GetConfigTableValue = IfNullThenNilString(RS("Value").Value)
        Else
            GetConfigTableValue = DefaultValue
        End If
        RS = Nothing
BadConfigTable:
    End Function
    Public Function SetConfigTableValue(ByVal FieldName As String, ByVal Value As String) As Boolean
        '::::SetConfigTableValue
        ':::SUMMARY
        ': Sets global config value.
        ':::DESCRIPTION
        ': Sets the config key to the specified value.  Removes the key if value is empty.
        ':::PARAMETERS
        ': - FieldName
        ': - Value
        ':::RETURN
        ': Boolean - Returns True
        Dim SQL As String
        SQL = "DELETE * FROM [Config] WHERE [FieldName]='" & FieldName & "'"
        ExecuteRecordsetBySQL(SQL, , GetDatabaseInventory, True)
        If Value <> "" Then
            SQL = "INSERT INTO Config ([FieldName], [Value]) VALUES ('" & FieldName & "', '" & Value & "')"
            ExecuteRecordsetBySQL(SQL, , GetDatabaseInventory, True)
        End If
        SetConfigTableValue = True
    End Function
    Public Function GetConfigAutoNumber(ByVal AN_Name As String, Optional ByVal MinValue as integer = 0, Optional ByVal ConvertValue as integer = -1) as integer
        '::::GetConfigAutoNumber
        ':::SUMMARY
        ': Return Config Auto-Number Value
        ':::DESCRIPTION
        ': Increments and returns an auto-number value (from the config table).
        ': -  Value name is prefixed, preventing overrun.
        ':::PARAMETERS
        ': - AN_Name - AutoNumber Field Name
        ': - [MinValue] - Optional.  Minimum value.  If returned value is less, minimum will be given.
        ': - [ConvertValue] - Optional. Used to convert from file-based AN to config field.
        ':::RETURN
        ': Long - Returns the result as a long value.
        Dim T As String, N as integer
        AN_Name = "AutoNumber_" & AN_Name
        T = GetConfigTableValue(AN_Name)

        ' ConvertValue is for changing from a different mode to this one.
        ' If there is no record in the DB for this autonumber, and ConverValue is supplied,
        ' It will be used instead... This should handle the transition from any other form of
        ' Autonumber to this one seamlessly.  After the first call, a record will be in the DB
        ' with the incremented value, and it should no longer use ConvertValue
        If T = "" And ConvertValue <> -1 Then
            T = "" & ConvertValue
        End If

        N = Val(T) ' (covers "")
        N = N + 1
        If MinValue <> 0 And N < MinValue Then N = MinValue
        SetConfigTableValue(AN_Name, N)
        GetConfigAutoNumber = N
    End Function
    Public Function SetConfigAutoNumber(ByVal AN_Name As String, ByVal Value As String) As Boolean
        '::::SetConfigAutoNumber
        ':::SUMMARY
        ': Set an Auto Number field.
        ':::DESCRIPTION
        ': This function is used to update the Auto Number value.
        ':::PARAMETERS
        ': - AN_Name - Auto Number Field Name
        ': - Value1
        ':::RETURN
        ': Boolean - Returns True
        AN_Name = "AutoNumber_" & AN_Name
        SetConfigAutoNumber = SetConfigTableValue(AN_Name, Value)
    End Function

    Public Function NextReceiptNumber() As Integer
        '::::NextReceiptNumber
        ':::SUMMARY
        ': Get next receipt number
        ':::DESCRIPTION
        ': Allows for a specialized case of the AutoNumber functions.
        ': - Adds a sub-prefix to the actual field name stored.
        ': - Increments field value and returns next available Receipt number.
        ':::PARAMETERS
        ':::RETURN
        ': Long - The next available Receipt Number
        NextReceiptNumber = GetConfigAutoNumber("CR_ReceiptNo", 1000)
    End Function

    Public Function AllowRunOnce(ByVal Check As String) As Boolean
        '::::AllowRunOnce
        ':::SUMMARY
        ': Semafore to keep a feature to run only once.
        ':::DESCRIPTION
        ': When passed a feature name to check, returns True if the operation is allowed or false if not.
        ':
        ': The function will store a datestamp if the operation has been run, false otherwise.
        ':::PARAMETERS
        ': - Check - Feature name to check
        ':::RETURN
        ': Boolean - Returns True if the operation can proceed
        Dim X As String, K As String
        K = "RunOnce_" & Check
        X = GetConfigTableValue(K)
        If X = "" Then
            AllowRunOnce = True
            SetConfigTableValue(K, DateTimeStamp)
        End If
    End Function

End Module
