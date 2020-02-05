Module modRaycom
    Private Const RaycomHeader As String = "Customer-Name,Phone,Email,Product,Sales-Amount,Sales-Person,Delivery-Date,Store"
    ' ****  Access (ONLY ever use this function to determine access)
    Public Function AllowRaycom() As Boolean
        ' NEVER EVER EVER use ANY other method to check store name other than a built-in store handler, or CheckStoreName.  Ever.
        AllowRaycom = IsMattressWarehouse And StoresSld = 1 Or IsDevelopment()
    End Function

    ' **** Report - Other than external use of the AllowRaycom and RaycomOutputFile, this is the only public interface.  It only takes optional start and end dates.
    Public Function RaycomNightlyReport(Optional ByRef StartDate As String = "", Optional ByRef EndDate As String = "") As Boolean
        Dim I As Integer

        RaycomNightlyReport = True
        If Not AllowRaycom() Then Exit Function

        ' This uses the standard WinCDS handlers for all these operations.  Delete file, write header, and begin.
        DeleteFileIfExists(RaycomOutputFile)
        WriteFile(RaycomOutputFile, RaycomHeader)

        For I = 1 To LicensedNoOfStores()
            GenerateRaycomCSV(I, StartDate, EndDate)
        Next
    End Function

    Public Function RaycomOutputFile() As String
        ' Unless there's an explicit request for an alternate date format, ALWAYS use WinCDS standard: DateStamp
        '  RaycomOutputFile = Format(Date, "MM_dd_yyyy") & "_advertising.csv"
        RaycomOutputFile = RaycomOutputFolder() & DateStamp() & "_advertising.csv"
    End Function


    ' **** Configuration (ONLY ever use these functions for config
    Private Function RaycomOutputFolder() As String
        On Error Resume Next
        RaycomOutputFolder = UIOutputFolder() & "Raycom\"                       'Put the file on the Users desktop
        If Not FolderExists(RaycomOutputFolder) Then MkDir(RaycomOutputFolder)
        If Not FolderExists(RaycomOutputFolder) Then RaycomOutputFolder = UIOutputFolder()
    End Function

    ' **** Support - Functions needed to generate the CSV files.  Only ever called from the main report function
    Private Sub GenerateRaycomCSV(ByVal StoreNum As Integer, Optional ByVal StartDate As String = "", Optional ByVal EndDate As String = "")
        Dim FN As String
        Dim SQL As String, RS As ADODB.Recordset, Line As String
        FN = RaycomOutputFile()     ' For Speeed.  Don't recalculate FN each iteration.

        SQL = ""
        SQL = SQL & "SELECT Trim([Mail].[First] + ' ' + [Mail].[Last]) as [FullName], [Mail].[Tele], [Mail].[Email], [GrossMargin].[Style], "
        SQL = SQL & "[GrossMargin].[SellPrice] as [SalesAmount], [Employees].[LastName] as [SalesmanName], [GrossMargin].[DelDate] , [GrossMargin].[Store] "
        SQL = SQL & " FROM  "
        SQL = SQL & "( ([GrossMargin] "
        SQL = SQL & "   INNER JOIN [Mail] on [Mail].[Index] = [GrossMargin].[MailIndex] "
        SQL = SQL & "  ) INNER JOIN [Employees] on [Employees].[SalesID] = [GrossMargin].[Salesman]) "
        SQL = SQL & " WHERE 1=1 "

        If StartDate = "" Or Not IsDate(StartDate) Or Not IsDate(EndDate) Then
            SQL = SQL & "AND DateDiff(""d"", grossMargin.sellDate, Now) <= 7 """
        Else
            SQL = SQL & SQLDateRange("GrossMargin.SellDate", StartDate, EndDate, True)
        End If

        SQL = SQL & " ORDER BY trim([Mail].[First] + ' ' + [Mail].[Last])"

        On Error Resume Next
        RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation(StoreNum))
        Do Until RS.EOF
            Line = ""
            Line = CSVLine(
                    IfNullThenNilString(RS("FullName").Value),
                    IfNullThenNilString(RS("Tele").Value),
                    IfNullThenNilString(RS("Email").Value),
                    IfNullThenNilString(RS("Style").Value),
                    SQLCurrency(IfNullThenZeroCurrency(RS("SalesAmount").Value)),
                    IfNullThenNilString(RS("SalesmanName").Value),
                    IfNullThenNilString(RS("DelDate").Value),
                    RS("Store").Value
                    )
            If Len(Line) > 0 Then WriteFile(FN, Line)
            RS.MoveNext()
        Loop
    End Sub

End Module
