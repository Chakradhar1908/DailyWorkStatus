Module dbSQL
    ' get number of records
    Public Function getRecordsetCountByTable(ByRef Table As String, Optional ByRef File As String = "") as integer
        '::::getRecordsetCountByTable
        ':::SUMMARY
        ': Used to get the Recordset Count.
        ':::DESCRIPTION
        ': This function is used to get the Recordset Count by using Table Name.
        ':::PARAMETERS
        ': - Table - Indicates the Table value.
        ': - File - Indicates the file name.
        ':::RETURN
        ': Long - Returns the result as a long.
        getRecordsetCountByTable = getRecordsetCountBySQL(SQL:="SELECT [" & Table & "].*" & " FROM " & Table, File:=File)
    End Function
    ' get number of records
    Public Function getRecordsetCountBySQL(ByRef SQL As String, Optional ByRef File As String = "") as integer
        '::::getRecordsetCountBySQL
        ':::SUMMARY
        ': Used to get Number of Records.
        ':::DESCRIPTION
        ': This function is used to get the Recordset Count by using SQL value.
        ': - This is a horribly wasteful routine and should likely never be used.  Use an aggregate or check the RecordCount yourself.
        ':::PARAMETERS
        ': - SQL - Indicates the String value.
        ': - File - Indicates the String value.
        ':::RETURN
        ': Long - Returns the result as a long.
        Dim cDB As CDbAccessGeneral
        getRecordsetCountBySQL = cDB.getRecordset(Always:=False).RecordCount  'OVERFLOW
        cDB.dbClose
    End Function
    ' get mail recordset with sql value
    Public Function getRecordsetByTableLabelIndex(ByRef Table As String, ByRef Label As String, ByRef Index As String, Optional ByRef Always As Boolean = False, Optional ByRef File As String = "", Optional ByRef OrderBy As String = "", Optional ByRef QuietErrors As Boolean = False, Optional ByRef ErrMsg As String = "") As ADODB.Recordset
        '::::getRecordsetByTableLabelIndex
        ':::SUMMARY
        ': Used to get the Mail Recordset using SQL value.
        ':::DESCRIPTION
        ': This function is used to get the Recordset with Sql value.
        ':::PARAMETERS
        ':::RETURN
        getRecordsetByTableLabelIndex = GetRecordsetBySQL(getSQLByTableLabelIndex(Table, Label, Index, , , OrderBy), Always, File, QuietErrors, ErrMsg)
    End Function
    ' get mail recordset by general
    Public Function getSQLByTableLabelIndex(
          ByVal Table As String _
        , Optional ByVal Label As String = "" _
        , Optional ByVal Index As String = "" _
        , Optional ByVal Operators As String = "=" _
        , Optional ByVal Operation As String = "SELECT" _
        , Optional ByVal OrderBy As String = ""
        ) As String
        '::::getSQLByTableLabelIndex
        ':::SUMMARY
        ': Build a SQL statement for a table and index.
        ':::DESCRIPTION
        ': Return a SQL query for any table given a field and index value.  Optionally supply an operator and a table sort.
        ':::PARAMETERS
        ': - Table - Indicates the String value.
        ': - Lable - Indicates the String value.
        ': - Operator - Indicates the String value.
        ': - Operation - Indicates the String value.
        ': - OrderBy - Indicates the String value.
        ':::RETURN
        ': String - Returns the result as a String.
        Dim tStr As String
        If Len(Index) > 0 Then
            If Asc(Mid(Index, 1, 1)) = 0 Then Index = ""
        End If
        tStr = IIf((Index = "") _
            , "" _
            , " WHERE " & Label & " " & Operators & "'" & ProtectSQL(Index) & "'")
        ' [" & Table & "].  ' removed 20030904
        If InStr(Table, " join ") < 1 Then Table = "[" & Table & "]"
        getSQLByTableLabelIndex = Operation & " *" _
                & " FROM " & Table _
                & tStr
        If OrderBy <> "" Then getSQLByTableLabelIndex = getSQLByTableLabelIndex & " Order By " & ProtectSQL(OrderBy)
    End Function

    ' Save (and return) the recordset
    Public Sub SetRecordsetByTableLabelIndex(ByRef RS As ADODB.Recordset, ByRef Table As String, ByRef Label As String, ByRef Index As String, Optional ByRef File As String = "")
        '::::SetRecordsetByTableLabelIndex
        ':::SUMMARY
        ': Used to Save
        ':::DESCRIPTION
        ': This function is used to Save and Return the Recordset.
        ': Its is also used to update the DataBase.
        ':::PARAMETRS
        ':::RETURN
        Dim cDB As CDbAccessGeneral
        cDB = DbAccessGeneral(SQL:=getSQLByTableLabelIndex(Table, Label, Index), File:=File)
        cDB.UpdateRecordSet(RS)   ' This must be called to update the database.
        cDB.dbClose()              ' used to close recordset
    End Sub

    Public Function GetField_BlankDefault(ByRef RS As ADODB.Recordset, ByRef Field As String) As String
        '::::GetField_BlankDefault
        ':::SUMMARY
        ': Return Field Value or empty string if null
        ':::DESCRIPTION
        ': This function is used to display the Blank value to any field that is null, otherwise actual value.
        ':::PARAMETERS
        ': - RS - Indicates the Recordset.
        ': - Field - Indicates the name of the field to query from the RS
        ':::RETURN
        ': String - Returns the result as a String.
        Dim Result As String = ""
        GetField_BlankDefault = IIf(IsNothing(RS(Field).Value), "", RS(Field).Value)
    End Function

    Public Function GetEmptyRecordsetByTable(ByRef Table As String, Optional ByRef Always As Boolean = True, Optional ByRef File As String = "", Optional ByRef QuietErrors As Boolean = False) As ADODB.Recordset
        '::::GetEmptyRecordsetByTable
        ':::SUMMARY
        ': Used to get the Empty Recordset by using SQL value.
        ':::DESCRIPTION
        ': This function is used to get the Empty Recordset by using parameters.
        ':::PARAMETERS
        ':::RETURN
        GetEmptyRecordsetByTable = GetRecordsetBySQL(GetEmptySQLByTable(Table), Always, File, QuietErrors)
    End Function
    Public Function GetEmptySQLByTable(ByVal Table As String) As String
        '::::GetEmptySQLByTable
        ':::SUMMARY
        ': Used to get the blank records for inserts.
        ':::DESCRIPTION
        ': This function is intended to be a quicker way to get a blank record for inserts.
        ':::PARAMETERS
        ': - Table - Indicates the String value.
        ':::RETURN

        ' Intended to be a quicker way to get a blank record for inserts.
        GetEmptySQLByTable = "SELECT * FROM [" & Table & "] WHERE False=True"
    End Function

    ' get mail recordset with sql  value
    Public Function GetRecordsetBySQL(ByVal SQL As String, Optional ByVal Always As Boolean = False, Optional ByVal File As String = "", Optional ByVal QuietErrors As Boolean = False, Optional ByVal ErrMsg As String = "", Optional ByVal ProgressForm As Object = False) As ADODB.Recordset
        '::::GetRecordsetBySQL
        ':::SUMMARY
        ': Used to get a Record Set.
        ':::DESCRIPTION
        ': This function is used to get a Recordset by using SQL.
        ':::PARAMETERS
        ': - SQL - String
        ': - Always - Boolean
        ': - File - Database Filename
        ': - QuietErrors - Boolean
        ': - ErrMsg - ByRef.  String
        ': - ProgressForm - Boolean
        ':::RETURN
        ': ADODB.RecordSet
        Dim dbGetRec As CDbAccessGeneral

        GetRecordsetBySQL = Nothing
        dbGetRec = DbAccessGeneral(File:=File, SQL:=SQL)
        If Not dbGetRec Is Nothing Then
            GetRecordsetBySQL = dbGetRec.getRecordset(Always:=Always, QuietErrors:=QuietErrors, ErrMsg:=ErrMsg, ProgressForm:=ProgressForm) ' if 'SetNew:=False' by default
            dbGetRec.dbClose()
        End If
        dbGetRec = Nothing
    End Function

    Public Sub ExecuteRecordsetBySQL(ByVal SQL As String, Optional ByVal Always As Boolean = False, Optional ByVal File As String = "", Optional ByVal QuietErrors As Boolean = False, Optional ByVal ErrMsg As String = "")
        '::::GetRecordsetBySQL
        ':::SUMMARY
        ': Used to execute a SQL statement without caring about the result.
        ':::DESCRIPTION
        ': Executes a SQL but does not return any result (useful for DELETE, ALTER, CREATE, etc).
        ':::PARAMETERS
        ': - SQL - String
        ': - Always - Boolean
        ': - File - Database Filename
        ': - QuietErrors - Boolean
        ': - ErrMsg - ByRef.  String
        ': - ProgressForm - Boolean
        Dim dbGetRec As CDbAccessGeneral
        dbGetRec = DbAccessGeneral(File:=File, SQL:=SQL)
        If Not dbGetRec Is Nothing Then
            dbGetRec.getRecordset(Always, , QuietErrors, ErrMsg)
            dbGetRec.dbClose()
        End If
        dbGetRec = Nothing
    End Sub

    Public Function getRecordsetByTableLabelIndexNumber(ByVal Table As String, ByVal Label As String, ByVal Index As String, Optional ByVal Always As Boolean = False, Optional ByVal File As String = "") As ADODB.Recordset
        '::::getRecordsetByTableLabelIndexNumber
        ':::SUMMARY
        ': Used to get Mail Recordset with Sql value.
        ':::DESCRIPTION
        ': This function is used to get the Recordset with Sql value.
        ':::PARAMETERS
        ':::RETURN
        getRecordsetByTableLabelIndexNumber = GetRecordsetBySQL(getSQLByTableLabelIndexNumber(Table, Label, Index), Always, File)
    End Function

    ' Save the recordset
    Public Sub SetMailRecordsetByTableLabelIndex(ByRef RS As ADODB.Recordset, ByRef Table As String, ByRef Label As String, ByRef Index As String, Optional ByRef File As String = "")
        '::::SetMailRecordsetByTableLabelIndex
        ':::SUMMARY
        ': Used to Save the Recordset.
        ':::DESCRIPTION
        ': This function is used to Save the Recordset by using parameters.
        ':::PARAMETERS
        ':::RETURN
        Dim cDB As CDbAccessGeneral
        cDB = DbAccessGeneral(SQL:=getSQLByTableLabelIndex(Table, Label, Index), File:=File)
        cDB.UpdateRecordSet(RS)   ' This must be called to update the database
        cDB.dbClose()              ' used to close recordset
    End Sub

    Public Function getSQLByTableLabelIndexNumber(
          ByVal Table As String _
        , Optional ByVal Label As String = "" _
        , Optional ByVal Index As String = "" _
        , Optional ByVal Operatorx As String = "="
        ) As String
        '::::getSQLByTableLabelIndexNumber
        ':::SUMMARY
        ': Used to get the Recordset by using Index Number.
        ':::DESCRIPTION
        ': This function is used to get the Recordset using Index Number through Sql statement.
        ':::PARAMTERS
        ': - Lable - Indicates the String value.
        ': - Index - Indicates the String value.
        ': - Operator - Indicates the String value.
        ':::RETURN
        ': String - Returns the result as a String.
        Dim tStr As String
        tStr = IIf((Index = ""), "", " WHERE " & Label & " " & Operatorx & "" & Index & "")
        If InStr(Table, " join ") < 1 Then Table = "[" & Table & "]"

        getSQLByTableLabelIndexNumber = "SELECT *" & " FROM " & Table & "" & tStr
    End Function

    Public Function GetSQLByTableLabelIndexNextPreviousCommon(Table As String, Optional Field As String = "", Optional Value As String = "", Optional Direction As Integer = 1) As String
        '::::getSQLByTableLabelIndexNextPreviousCommon
        ':::SUMMARY
        ': Used to get the Next or Previous record.
        ':::DESCRIPTION
        ': This function is used to get the Next or Previous records.
        ':::PARAMETERS
        ': - Table - Indicates the String value.
        ': - Field - Indicates the String value.
        ': - Value - Indicates the String value.
        ': - Direction - Indicates the Integer value.
        ':::RETURN
        ': String - Returns the result as a String.

        Dim Operation As String
        Dim Order As String
        Dim Operatorx As String

        Operation = "SELECT"
        Order = " ORDER BY  " & Table & "." & Field
        Order = Order & IIf((Direction = 1), "", " DESC")
        Operatorx = IIf((Direction = 1), ">", "<")

        GetSQLByTableLabelIndexNextPreviousCommon = Operation & " TOP 1 [" & Table & "].*" _
                & " FROM [" & Table & "]" _
                & " WHERE [" & Table & "]." & Field & " " & Operatorx & """" & ProtectSQL(Value) & """" _
                & Order

    End Function
    Public Function SQLCurrency(ByVal Amount As Decimal) As String
        '::::SQLCurrency
        ':::SUMMARY
        ': Used to format the Currency values for SQL statements
        ':::DESCRIPTION
        ': This function is used to format the Currency Amounts.
        ': Prevents breaking sql statements with $ and ,
        ':::PARAMETERS
        ': - Amount - Indicates the Currency Amount.
        ':::RETURN
        ': String - Returns the Result as a String.
        SQLCurrency = CurrencyFormat(Amount, , , True) ' both $ and , break sql statements w/ currencies
    End Function
    Public Function GetTableRecordMax(
          Table As String _
        , Field As String _
        , Optional File As String = "" _
        , Optional fieldType As String = ""
    ) As Integer
        '::::GetTableRecordMax
        ':::SUMMARY
        ': The Max used index of a table.
        ':::DESCRIPTION
        ': This function is used to get the Maximum Index currently used in a table.
        ':::PARAMETERS
        ': - Table - Indicates the Table Name.
        ': - Field - Indicates the Field String.
        ': - File - Indicates the File String.
        ': - fieldType - Indicates the Field Type.
        ':::RETURN
        ': Long - The max Index
        GetTableRecordMax = GetTableRecordFunction(functionType:="Max", Table:=Table, Field:=Field, fieldType:=fieldType, File:=File)
    End Function
    Public Function GetTableRecordFunction(
          ByVal Table As String _
        , ByVal Field As String _
        , Optional ByVal functionType As String = "Max" _
        , Optional ByVal File As String = "" _
        , Optional ByVal fieldType As String = ""
    ) As Integer

        '::::GetTableRecordFunction
        ':::SUMMARY
        ': Perform a aggregate function query
        ':::DESCRIPTION
        ': Run a query with an aggregate function on a table.
        ': This function is also used to handle errors.
        ':::PARANETERS
        ': - Table - Indicates the Table Name.
        ': - Field - Indicates the Field String.
        ': - functionType - Indicates the function type.
        ': - File - Indicates the File String.
        ': - fieldType - Indicates the field Type.
        ':::RETURN
        ': Long - Returns the result of the aggregate function


        Dim RS As ADODB.Recordset
        Dim SQL As String
        Dim fieldInfo As String
        fieldInfo = IIf((fieldType = "Text"), "( (" & Field & "))", "(clng(" & Field & "))")

        SQL = "SELECT " & functionType & fieldInfo & "  AS RESULT  FROM [" & Table & "];"

        On Error GoTo HandleErr

        Dim cDB As CDbAccessGeneral
        cDB = DbAccessGeneral(SQL:=SQL, File:=File)

        RS = cDB.getRecordset(Always:=False)
        GetTableRecordFunction = RS("RESULT").Value
        cDB.dbClose()

        Exit Function
HandleErr:
        GetTableRecordFunction = 0
        Resume Next

    End Function

    Public Function GetValueBySQLLong(ByVal SQL As String, Optional ByVal Always As Boolean = False, Optional ByVal File As String = "", Optional ByVal QuietErrors As Boolean = False, Optional ByVal ErrMsg As String = "", Optional ByVal DefaultValue As String = "") As Integer
        '::::GetValueBySQLLong
        ':::SUMMARY
        ': Used to return a single value via a SQL statement (as a Long)
        ':::DESCRIPTION
        ': Executes a query and returns the first column from the first row.
        ':::PARAMETERS
        ': - SQL - String
        ': - Always - Boolean
        ': - File - Database Filename
        ': - QuietErrors - Boolean
        ': - ErrMsg - ByRef.  String
        ': - ProgressForm - Boolean
        ':::SEE ALSO
        ': - GetValueBySQL, GetValueBySQLLong, GetValueBySQLString, GetValueBySQLDate, GetValueBySQLCurrency, GetValueBySQLDouble
        ':::RETURN
        ': - The value from the executed SQL
        On Error Resume Next
        GetValueBySQLLong = Val(IfNullThenZero(GetValueBySQL(SQL, Always, File, QuietErrors, ErrMsg, DefaultValue)))
    End Function

    Public Function GetValueBySQL(ByVal SQL As String, Optional ByVal Always As Boolean = False, Optional ByVal File As String = "", Optional ByVal QuietErrors As Boolean = False, Optional ByVal ErrMsg As String = "", Optional ByVal DefaultValue As String = "") As Object
        '::::GetValueBySQL
        ':::SUMMARY
        ': Used to return a single value via a SQL statement.
        ':::DESCRIPTION
        ': Executes a query and returns the first column from the first row.
        ':::PARAMETERS
        ': - SQL - String
        ': - Always - Boolean
        ': - File - Database Filename
        ': - QuietErrors - Boolean
        ': - ErrMsg - ByRef.  String
        ': - ProgressForm - Boolean
        ':::SEE ALSO
        ': - GetValueBySQL, GetValueBySQLLong, GetValueBySQLString, GetValueBySQLDate, GetValueBySQLCurrency, GetValueBySQLDouble
        ':::RETURN
        ': - The value from the executed SQL
        Dim R As ADODB.Recordset
        On Error Resume Next
        R = GetRecordsetBySQL(SQL, Always, File, QuietErrors)
        If Not R.EOF Then
            GetValueBySQL = R.Fields(0).Value
        End If
        If IsNothing(GetValueBySQL) Then
            GetValueBySQL = DefaultValue
        ElseIf IsNothing(GetValueBySQL) Then
            GetValueBySQL = ""
        End If
        DisposeDA(R)
    End Function

    Public Function SQLDate(ByVal vDate As Date, Optional ByVal Delimiter As String = "#") As String
        '::::SQLDate
        ':::SUMMARY
        ':::DESCRIPTION
        ': This function is used to format the Date
        ':::PARAMETERS
        ': - vDate - Indicates the Date value.
        ': - Delimiter - Indicates the date delimiter ('#' in MSJET)
        ':::RETURN
        ': String - Returns the Result as a String.

        SQLDate = Delimiter & DateFormat(DateValue(vDate), "/") & Delimiter
    End Function

End Module
