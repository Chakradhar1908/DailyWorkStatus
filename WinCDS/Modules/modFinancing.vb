Module modFinancing
    Public Const ArNo_AddOnRecordSeparator As String = "-"
    Public Const ArNo_AddOnRecordToken As String = "AddOnRecord"
    Public Const ArNo_AddOnRecordIndicator As String = ArNo_AddOnRecordSeparator & ArNo_AddOnRecordToken & ArNo_AddOnRecordSeparator
    Public Const ArNo_AddOnRecordPattern_LIKE As String = "*" & ArNo_AddOnRecordToken & "*"
    Public Const ArNo_AddOnRecordPattern_SQL As String = "%" & ArNo_AddOnRecordToken & "%" ' MS Access requires % for wildcard

    Public Function ArNoIsAddOnRecord(ByVal ArNo As String) As Boolean
        ArNoIsAddOnRecord = ArNo Like ArNo_AddOnRecordPattern_LIKE
    End Function

End Module
