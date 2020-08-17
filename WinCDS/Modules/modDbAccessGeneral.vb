Module modDbAccessGeneral
    ' Return a CDbAccessGeneral object based on sql argument
    Public Function DbAccessGeneral(Optional ByVal File As String = "", Optional ByVal SQL As String = "") As CDbAccessGeneral
        '::::DbAccessGeneral
        ':::SUMMARY
        ': DB Access General Object
        ':::DESCRIPTION
        ': Returns a CDbAccessGeenral object populated with database and sql
        ':::PARAMETERS
        ': - File
        ': - SQL
        ':::RETURN
        ': - CDbAccessGeneral
        DbAccessGeneral = New CDbAccessGeneral
        If DbAccessGeneral.dbOpen(File) Then
            DbAccessGeneral.SQL = SQL  ' This doesn't seem to do much for us, but oh well.
        Else
            DbAccessGeneral = Nothing
        End If
    End Function

    Public Function GetNextFieldValue(ByVal Table As String, ByVal Field As String, Optional ByVal ZeroValue As Integer = 0) As Integer
        '::::GetNextFieldValue
        ':::SUMMARY
        ': Used to get Next Field Value.
        ':::DESCRIPTION
        ': This function is used to get Next Field Value from required table.
        ':::PARAMETERS
        ': - Table
        ': - Field
        ': - ZeroValue - Default is ZERO since we are looking for MAX.Next value would be 1.
        ':::RETURN
        ': Long - Returns a result as a Long.
        GetNextFieldValue = GetMaxFieldValue(Table, Field, ZeroValue) + 1
    End Function

    Public Function GetMaxFieldValue(ByVal Table As String, ByVal Field As String, Optional ByVal ZeroValue As Integer = 0) As Integer
        '::::GetMaxFieldValue
        ':::SUMMARY
        ': Used to get Maximum Index Value of a table.
        ':::DESCRIPTION
        ': This function is used to get maximum index value from a table.
        ':::PARAMETERS
        ': - Table
        ': - Field
        ': - ZeroValue - Default is ZERO since we are looking for MAX.Next value would be 1.
        ':::RETURN
        ': Long - Returns a result as a Long.

        Dim RS As ADODB.Recordset
        Dim DG As CDbAccessGeneral

        On Error GoTo NoRecords
        DG = DbAccessGeneral(SQL:="SELECT Count(" & Field & ") AS C, Max(" & Field & ") AS M FROM " + Table)
        RS = DG.getRecordset(Always:=False)
        GetMaxFieldValue = IIf(RS("C").Value = 0, ZeroValue, RS("M"))
        DG.dbClose()

        Exit Function
NoRecords:
        GetMaxFieldValue = ZeroValue ' Default is ZERO since we are looking for MAX.  Next value would be 1.
    End Function

End Module
