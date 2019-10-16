Module modDBFunctions
    Public Function IfNullThenNilString(ByVal T As Object) As String
        '::::IfNullThenNilString
        ':::SUMMARY
        ': Null Handle String Fields
        ':::DESCRIPTION
        ': Return a string result from field, processing DBNull as ""
        ':::PARAMETERS
        ': - T - Typically a Recordset Field or value.
        ':::RETURN
        ': String - Returns field value as string or ""
        ':::ALIASES
        ': - INTNS
        ':::SEE ALSO
        ': - IfNullThenBoolean, IfNullThenZeroCurrency, IfNullThenZeroLong, IfNullThenZeroDouble

        'If IsMissing(T) Then T = ""
        If IsNothing(T) Then T = ""
        'IfNullThenNilString = IIf(IsNull(T), "", T)
        IfNullThenNilString = IIf(IsNothing(T), "", T)
    End Function
    Public Function IfNullThenZero(ByVal T As Object) as integer
        '::::IfNullThenZero
        ':::SUMMARY
        ': Null Handle Number Fields
        ':::DESCRIPTION
        ': Return a Long result from field, processing DBNull as 0
        ':::PARAMETERS
        ': - T - Typically a Recordset Field or value.
        ':::RETURN
        ': - Long
        ':::ALIASES
        ': - INTZ, INTZL, IfNullThenZeroLong
        ':::SEE ALSO
        ': - IfNullThenNilString, IfNullThenZeroCurrency, IfNullThenZeroDouble
        'If IsMissing(T) Then T = 0
        If IsNothing(T) Then T = 0
        'If IsNull(T) Then Exit Function '0
        If IsNothing(T) Then IfNullThenZero = 0 : Exit Function '0

        If Val(T) < -2147483648.0# Then
            IfNullThenZero = -2147483648.0#
        ElseIf Val(T) > 2147483647.0# Then
            IfNullThenZero = 2147483647.0#
        Else
            IfNullThenZero = Val(T)
        End If
    End Function
    Public Function IfNullThenZeroCurrency(ByVal T As Object) As Decimal
        '::::IfNullThenZeroCurrency
        ':::SUMMARY
        ': Null Handle Currency Fields
        ':::DESCRIPTION
        ': Return a Currency result from field, processing DBNull as 0
        ':::PARAMETERS
        ': - T - Typically a Recordset Field or value.
        ':::RETURN
        ': - Currency
        ':::ALIASES
        ': - INTZC
        ':::SEE ALSO
        ': - IfNullThenNilString, IfNullThenZeroLong
        'If IsMissing(T) Then T = 0
        If IsNothing(T) Then T = 0
        'If IsNull(T) Then T = 0
        If IsNothing(T) Then T = 0
        If Not IsNumeric(T) Then T = 0
        'IfNullThenZeroCurrency = IIf(IsNull(T), 0#, T)
        IfNullThenZeroCurrency = IIf(IsNothing(T), 0#, T)
    End Function
    Public Function IfNullThenZeroDouble(ByVal T As Object) As Double
        '::::IfNullThenZeroDouble
        ':::SUMMARY
        ': Null Handle Number Fields
        ':::DESCRIPTION
        ': Return a Long result from field, processing DBNull as 0
        ':::PARAMETERS
        ': - T - Typically a Recordset Field or value.
        ':::RETURN
        ': - Double
        ':::ALIASES
        ': - INTZD
        ':::SEE ALSO
        ': - IfNullThenNilString, IfNullThenZeroCurrency, IfNullThenZero
        'If IsMissing(T) Then T = 0
        If IsNothing(T) Then T = 0
        'IfNullThenZeroDouble = IIf(IsNull(T), 0, T)
        IfNullThenZeroDouble = IIf(IsNothing(T), 0, T)
    End Function
    Public Function ProtectSQL(ByVal Str As String, Optional ByVal UseDoubleQuotes As Boolean = True) As String
        '::::ProtectSQL
        ':::SUMMARY
        ': Protect SQL statements from SQL injection attacks.
        ':::DESCRIPTION
        ': Parse out quotes and other breaking characters from database queries.
        ':::PARAMETERS
        ': - Str - Input Query
        ': - UseDoubleQuotes - Use Double quotes (True) or Single Quotes (False)
        ':::RETURN
        ': String - The protected SQL query
        If UseDoubleQuotes Then
            Str = Replace(Str, """", """""")
        Else
            Str = Replace(Str, "'", "''")   ' This is bad in Access..
        End If
        ProtectSQL = Str
    End Function
    Public Function IfNullThenZeroDate(ByVal T As Object) As Date
        '::::IfNullThenZeroDate
        ':::SUMMARY
        ': Null Handle Date Fields
        ':::DESCRIPTION
        ': Return a Date result from field, processing DBNull as Zero Date (1/1/1900, 12:00:00 AM, CDate(0))
        ':::PARAMETERS
        ': - T - Typically a Recordset Field or value.
        ':::RETURN
        ': - Date
        ':::ALIASES
        ': - INTZDt
        ':::SEE ALSO
        ': - IfNullThenNilString, IfNullThenZeroCurrency, IfNullThenZero, IfNullThenZeroDouble
        On Error Resume Next
        If TypeName(T) = "Field" Then T = T.Value
        If TypeName(T) = "String" Then
            If Not IsDate(T) Then IfNullThenZeroDate = CDate("12:00:00 AM") : Exit Function
            IfNullThenZeroDate = DateValue(T)
        End If
        IfNullThenZeroDate = CDate(IIf(IsNothing(T), 0, T))

    End Function
    Public Function IfZeroThenNilString(ByVal T As Object) As String
        '::::IfZeroThenNilString
        ':::SUMMARY
        ': Null And Zero Handle String Fields
        ':::DESCRIPTION
        ': Return a string result from field, processing DBNull AND Zero as ""
        ':::PARAMETERS
        ': - T - Typically a Recordset Field or value.
        ':::RETURN
        ': String - Returns field value as string or ""
        ':::ALIASES
        ': - IZTNS
        ':::SEE ALSO
        ': - IfNullThenNilString
        ': - IfNullThenBoolean, IfNullThenZeroCurrency, IfNullThenZeroLong, IfNullThenZeroDouble
        'If IsMissing(T) Then T = ""
        If IsNothing(T) Then T = ""
        IfZeroThenNilString = IIf(T = 0, "", CStr(T))
    End Function

    Public Function IfNullThenNullDate(ByVal T As Object) As Date
        '::::IfNullThenNullDate
        ':::SUMMARY
        ': Null Handle Date Fields
        ':::DESCRIPTION
        ': Return a Date result from field, processing DBNull as NullDate() (1 /1 /2001, NullDate)
        ':::PARAMETERS
        ': - T - Typically a Recordset Field or value.
        ':::RETURN
        ': - Date
        ':::ALIASES
        ': - INTNDt
        ':::SEE ALSO
        ': - IfNullThenNilString, IfNullThenZeroCurrency, IfNullThenZero, IfNullThenZeroDouble
        On Error Resume Next
        If TypeName(T) = "Field" Then T = T.Value
        If TypeName(T) = "String" Then
            If Not IsDate(T) Then IfNullThenNullDate = NullDate : Exit Function
            IfNullThenNullDate = DateValue(T)
        End If
        IfNullThenNullDate = IIf(IsNothing(T), NullDate, T)
    End Function

    Public Function IfNegativeThenZero(ByVal T As Object) As Double
        '::::IfNegativeThenZero
        ':::SUMMARY
        ': Null And Negative Handle Number Fields
        ':::DESCRIPTION
        ': Return a string result from field, processing DBNull AND < 0 as 0 (Double)
        ':::PARAMETERS
        ': - T - Typically a Recordset Field or value.
        ':::RETURN
        ': Double - Returns field value as string or ""
        ':::SEE ALSO
        ': - IfNullThenNilString
        ': - IfNullThenBoolean, IfNullThenZeroCurrency, IfNullThenZeroLong, IfNullThenZeroDouble
        On Error GoTo BadNumber
        If IsNothing(T) Then IfNegativeThenZero = 0 : Exit Function
        If T < 0 Then IfNegativeThenZero = 0 : Exit Function
        IfNegativeThenZero = T
        Exit Function
BadNumber:
        IfNegativeThenZero = 0
    End Function

    Public Function IfNullThenBoolean(ByVal T As Object, Optional ByVal DefaultValue As Boolean = False) As Boolean
        '::::IfNullThenBoolean
        ':::SUMMARY
        ': Null Handle Boolean Fields
        ':::DESCRIPTION
        ': Return a Boolean result from field, processing DBNull as False
        ': - NOTE: If the value cannot be parsed as a boolean, error handling will return the default value.
        ':::PARAMETERS
        ': - T - Typically a Recordset Field or value.
        ': - DefaultValue - What is returned if is null or Boolean conversion fails
        ':::RETURN
        ': Boolean
        ':::SEE ALSO
        ': - IfNullThenNilString, IfNullThenZeroCurrency, IfNullThenZeroLong, IfNullThenZeroDouble
        On Error Resume Next
        IfNullThenBoolean = DefaultValue
        IfNullThenBoolean = IIf(IsNothing(T), DefaultValue, T)
    End Function

End Module
