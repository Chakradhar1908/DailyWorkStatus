Module modDepartments
    Private Const tblDepartments As String = "Departments"
    Private Const tblDepartments_Deleted As String = "[DELETED]"

    Public Const dpstDeleted As Integer = 0
    Public Const dpstActive As Integer = 1

    Private Const DefDepts As String = "Occasional/Dining Room/Bedroom/Entertain & Office/Upholstery/Motion & Sleepers/Bedding/Lamps/Pictures/Accessories"

    Public Structure Department
        Dim ID As Integer
        Dim Name As String
        Dim Status As Integer
    End Structure

    Private DepartmentList() As Department, DeptLoaded As Integer

    Public Sub LoadDeptNamesIntoComboBox(ByRef Cbo As ComboBox, Optional ByVal Dept As String = "", Optional ByVal Separated As Boolean = False, Optional ByVal OnlyName As Boolean = False)
        '::::LoadDeptNamesIntoComboBox
        ':::SUMMARY
        ':Loads department names
        ':::DESCRIPTION
        ':Fetches the department names from the Departments table and loads it into Combo box.
        ':::SEE ALSO
        ':LoadDeptNamesIntoListBox
        Dim I As Integer

        Cbo.Items.Clear()
        If CountDepartments() <= 0 Then Exit Sub

        For I = LBound(DepartmentList) To UBound(DepartmentList)
            If Not DepartmentList(I).Status = dpstActive Then GoTo SkipItem
            If Separated Then
                'Cbo.AddItem DepartmentList(I).Name
                'Cbo.itemData(Cbo.NewIndex) = DepartmentList(I).ID
                Cbo.Items.Add(New ItemDataClass(DepartmentList(I).Name, DepartmentList(I).ID))
            Else
                ' Don't worry about alignment or spacing while there are only 10 departments.
                ' There can be more than 10 now, so we have to decide if spacing matters soon...
                'Cbo.AddItem DepartmentList(I).ID & "   " & DepartmentList(I).Name
                'Cbo.itemData(Cbo.NewIndex) = DepartmentList(I).ID
                Cbo.Items.Add(New ItemDataClass(DepartmentList(I).ID & "   " & DepartmentList(I).Name, DepartmentList(I).ID))
            End If
SkipItem:
        Next
    End Sub

    Public Function CountDepartments(Optional ByVal StoreNo As Integer = 0) As Integer
        '::::CountDepartments
        ':::SUMMARY
        ':Counts the total no. of departments
        ':::DESCRIPTION
        ':Counts the total no. of departments of Departments Table
        ':
        ':::RETURN
        ':  Long -  Returns total no. of departments
        ':::SEE ALSO
        ':
        On Error Resume Next
        LoadDepartments(StoreNo)
        CountDepartments = UBound(DepartmentList) - LBound(DepartmentList) + 1
    End Function

    Private Function LoadDepartments(Optional ByVal StoreNo As Integer = 0) As Boolean
        GetDepartmentList(StoreNo)
    End Function

    Public Function GetDepartmentList(Optional ByVal StoreNo As Integer = 0) As String()
        '::::GetDepartmentList
        ':::SUMMARY
        ':Fetches the Department names from the database.
        ':::DESCRIPTION
        ':Fetches the Department names from the database and returns as a string array.
        ':Also stores Department ID, Name, Status.
        ':::RETURN
        ':  String - Array of Department names
        ':::SEE ALSO
        ':GetDepartmentName, GetDepartmentNo

        Dim sSql As String, RS As ADODB.Recordset
        Dim N As Integer
        Dim RET() As String

        If StoreNo = 0 Then StoreNo = StoresSld

        EnsureDepartmentTableExists(StoreNo)
        ImportDeptFile(StoreNo)

        If DeptLoaded <> StoreNo Then                   ' We can only have one store cached at a time..
            sSql = "SELECT * FROM [" & tblDepartments & "] ORDER BY [Id]"
            RS = GetRecordsetBySQL(sSql, True, GetDatabaseAtLocation(), True)

            N = 0
            Do Until RS.EOF
                ReDim Preserve DepartmentList(0 To N)       ' Add active elements.
                With DepartmentList(N)
                    .ID = IfNullThenZero(RS("id").Value)
                    .Name = IfNullThenNilString(RS("name").Value)
                    .Status = IfNullThenZero(RS("status").Value)
                End With
                N = N + 1
                RS.MoveNext()
            Loop

            DeptLoaded = StoreNo                          ' Mark which store's dept list is cached.
        End If

        ReDim RET(0 To UBound(DepartmentList))          ' Now we know the data is in the cache -- read from cache.
        For N = LBound(DepartmentList) To UBound(DepartmentList)
            RET(N) = IIf(DepartmentList(N).Status = dpstDeleted, "", DepartmentList(N).Name)
        Next

        GetDepartmentList = RET                         ' Return the string list of names only.
    End Function

    Private Function EnsureDepartmentTableExists(Optional ByVal StoreNo As Integer = 0) As Boolean
        Dim sSql As String, RS As ADODB.Recordset
        If StoreNo <= 0 Then StoreNo = StoresSld

        If TableExists(StoreNo, tblDepartments) Then
            RS = GetRecordsetBySQL("SELECT * FROM " & tblDepartments & " WHERE FALSE=TRUE", , GetDatabaseAtLocation(StoreNo))
            If RS.Fields(1).Name = "DepartmentName" Then
                ExecuteRecordsetBySQL("DROP TABLE " & tblDepartments)
            End If
        End If

        If Not TableExists(StoreNo, tblDepartments) Then
            sSql = "CREATE TABLE [" & tblDepartments & "] (ID INTEGER, Name VARCHAR(100), Status INTEGER)"
            ExecuteRecordsetBySQL(sSql, False, GetDatabaseAtLocation(StoreNo), False, "Could not build table [" & tblDepartments & "] in " & GetDatabaseAtLocation(StoreNo) & "!")
        End If

        EnsureDepartmentTableExists = TableExists(StoreNo, tblDepartments)

        If Not EnsureDepartmentTableExists Then MsgBox("Unable to create [Departments] table in " & GetDatabaseAtLocation(StoreNo) & ".")
    End Function

    Private Function ImportDeptFile(ByVal StoreNo As Integer) As Boolean
        Const Ignore As String = "#IGNORE"
        Dim dFile As String, dName As String
        Dim N As Integer, I As Integer
        Dim RS As ADODB.Recordset
        Dim Status As Integer

        dFile = DepartmentFile(StoreNo)
        If Not FileExists(dFile) Then ImportDeptFile = True : Exit Function
        If ReadFile(dFile, 1, 1) = Ignore Then Exit Function

        ImportDeptFile = True

        ClearDepartments()

        N = CountFileLines(dFile)
        For I = 1 To N
            dName = Trim(ReadFile(dFile, I, 1))
            If dName = "" Then
                If I = N Then Exit For            ' Don't include a trailing empty...
                dName = tblDepartments_Deleted
            End If
            AddDepartment(dName)
        Next


        'Name dFile As GetFileBase(dFile, True, True) & "-" & DateTimeStamp & "." & GetFileExt(dFile, True)
        My.Computer.FileSystem.RenameFile(dFile, GetFileBase(dFile, True, True) & "-" & DateTimeStamp() & "." & GetFileExt(dFile, True))
        WriteFile(dFile, Ignore, True)

        ImportDeptFile = True
    End Function

    Private Function ClearDepartments(Optional ByVal StoreNo As Integer = 0) As Boolean
        EnsureDepartmentTableExists()
        ExecuteRecordsetBySQL("DELETE FROM " & tblDepartments, , GetDatabaseAtLocation(StoreNo))
        ClearDepartments = True
    End Function

    Public Function AddDepartment(ByVal DeptName As String) As Boolean
        '::::AddDepartment
        ':::SUMMARY
        ':Adds the department record
        ':::DESCRIPTION
        ':Checks the department exists in the database.
        ':If not exists, adds the new department(Deptno, Department Name and Status) into Departments table
        ':If Department exists, returns the Department No.
        ':If Department name will be passed without Department No, fetches the max department no, add 1 to it
        ':and then add the Department details to the database (Departments table)
        ':::RETURN
        ':  Boolean -  Returns the department no.
        ':::SEE ALSO
        ':AddDepartmentData, ImportDeptFile

        Dim Status As Integer, ID As Integer, nSt As Integer
        Dim sSql As String
        Dim RS As ADODB.Recordset

        EnsureDepartmentTableExists()
        ' Read directly to prevent recursion.
        ID = CheckDepartmentExistsInternal(DeptName)
        If ID >= 0 Then
            nSt = GetValueBySQLLong("SELECT Status FROM " & tblDepartments & " WHERE ID=" & ID, , GetDatabaseAtLocation)
            If nSt = dpstDeleted Then
                ExecuteRecordsetBySQL("UPDATE " & tblDepartments & " SET Status=" & dpstActive & " WHERE ID=" & ID, , GetDatabaseAtLocation)
            End If
            Exit Function
        End If

        ID = GetNextFieldValue(tblDepartments, "Id", -1)
        Status = IIf(DeptName = "" Or DeptName = tblDepartments_Deleted, dpstDeleted, dpstActive)
        DeptName = Trim(DeptName)

        sSql = "INSERT INTO [" & tblDepartments & "] (ID, Name, Status) VALUES (" & ID & ", '" & ProtectSQL(DeptName) & "', " & Status & ")"
        ExecuteRecordsetBySQL(sSql, , GetDatabaseAtLocation(), False, "Error creating department: " & DeptName)

        If CheckDepartmentExistsInternal(DeptName) = -1 Then
            MessageBox.Show("Failed to create department [" & DeptName & "].", "Department Creation Failure")
        End If

        ReloadDepartments()         ' Mark cache as invalid.
        AddDepartment = True
    End Function

    Private Function CheckDepartmentExistsInternal(ByVal DeptName As String) As Integer
        Dim sSql As String
        sSql = "SELECT ID FROM " & tblDepartments & " WHERE name='" & ProtectSQL(DeptName) & "'"
        CheckDepartmentExistsInternal = GetValueBySQLLong(sSql, , GetDatabaseAtLocation(), True, , -1)
    End Function

    Public Function ReloadDepartments()
        '::::ReoadDepartments
        ':::SUMMARY
        ':Refresh department list cache.
        ':::DESCRIPTION
        ':Marks the department list cache as dirty so it will reload on next query.
        ':
        ':::PARAMETERS
        ':- StoreNo - The StoreNo for which Table exists will check
        ':- Reload - Specify True to re-read from the database, even if this store is already cached.
        ':::RETURN
        ':  Boolean - True, If Table exists. Else False
        ':::SEE ALSO
        ':
        DeptLoaded = 0
    End Function

End Module
