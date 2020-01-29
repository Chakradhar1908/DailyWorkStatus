Module modWebIntegration
    Public Sub Generate2DataCSV(Optional ByRef FN As String = "", Optional ByRef WebSite As String = "")
        Dim SQL As String, RS As Recordset, FNum As Long, Line As String

        On Error Resume Next
        Kill FN
  On Error GoTo 0

        Count = 0
        dL = GetDepartmentList()

        FNum = FreeFile()
        Open FN For Output As #FNum

  Print #FNum, GenerateCSVHeader()

  SQL = "SELECT [2Data].*, iif(isnull(Search.RN),0,1) AS Orderable FROM [2DATA] LEFT JOIN [Search] ON [2Data].Style = [Search].Style"
'  SQL = SQL & " WHERE Trim([2Data].Dept)='1'"
  Set RS = GetRecordsetBySQL(SQL, , GetDatabaseInventory)
  Do Until RS.EOF
            Line = GenerateCSVLine(RS, WebSite)
            If Len(Trim(Line)) > 0 Then Print #FNum, Line
    RS.MoveNext
        Loop

        Close #FNum
End Sub

End Module
