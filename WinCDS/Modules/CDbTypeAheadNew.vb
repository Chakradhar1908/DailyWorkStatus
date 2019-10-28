Module CDbTypeAheadNew
    '::::CDbTypeAheadNew
    ':::SUMMARY
    ':Standard WinCDS TypeAhead Module
    ':::DESCRIPTION
    ':Type ahead module for pulling a recordset and handling type-ahead.

    Public Function New_CDbTypeAhead(ByVal Table As String, ByVal Field As String, Optional ByVal Value As String = "", Optional ByVal Match As Integer = 1, Optional ByVal MinLength As Integer = 0, Optional ByVal ExtraCondition As String = "", Optional ByVal ExtraSort As String = "") As CDbTypeAhead
        '::::New_CDbTypeAhead
        ':::SUMMARY
        ': Create a new DB Type Ahead
        ':::DESCRIPTION
        ': This function is used to get recordset and initialize Table, Field, Value.
        ': This will not call event, since there is not withevents association.
        ':::PARAMETERS
        ': - Table
        ': - Field
        ': - Value
        ': - Match
        ': - MinLength
        ': - ExtraCondition
        ': - ExtraSort
        ':::RETURN
        ': - CDbTypeAhead


        New_CDbTypeAhead = New CDbTypeAhead
        New_CDbTypeAhead.Table = Table
        New_CDbTypeAhead.Field = Field
        New_CDbTypeAhead.Value = Value
        New_CDbTypeAhead.Match = Match
        New_CDbTypeAhead.MinLength = MinLength
        New_CDbTypeAhead.ExtraCondition = ExtraCondition
        New_CDbTypeAhead.ExtraSort = ExtraSort
        '.Initialize Table, Field, Value
        '.getRecordset
        '.Refresh ' This will not call event, since there is not withevents association
    End Function
End Module
