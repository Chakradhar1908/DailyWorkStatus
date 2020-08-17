Module modGetSales
    Private Sm As Object, SmLoc As Integer    ' cache

    Public Function getSalesNumber(ByVal EmployeeName As String, Optional ByVal Store as integer = 0, Optional ByVal OnEmpty As String = "") As String
        '::::getSalesNumber
        ':::SUMMARY
        ': Get Salesman Employee Number
        ':::DESCRIPTION
        ': Return the employee number by name
        ':::PARAMETERS
        ': - EmployeeName
        ': - Store
        ': - OnEmpty
        ':::RETURN
        ': String

        Dim I as integer
        On Error Resume Next
        If Store = 0 Then Store = StoresSld
        If SmLoc <> Store Then GetSalesmanDatabase(Store)
        If EmployeeName = "HOUSE" Then getSalesNumber = "99" : Exit Function
        Sm = GetSalesmanDatabase(Store)

        For I = LBound(Sm, 1) To UBound(Sm, 1)
            'If Sm(I, 1) = EmployeeName Then
            If Sm(I, 0) = EmployeeName Then
                'getSalesNumber = Sm(I, 2)
                getSalesNumber = Sm(I, 1)
                Exit Function
            End If
        Next
        getSalesNumber = OnEmpty
    End Function

    Public Function getSalesName(ByVal EmployeeNumber As String, Optional ByVal Store As Integer = 0) As String
        '::::getSalesName
        ':::SUMMARY
        ': Query The Sales Name
        ':::DESCRIPTION
        ': Returns the salesman name by employee number.
        ':::PARAMETERS
        ': - EmployeeNumber
        ': - Store
        ':::RETURN
        ': String

        Dim I As Integer
        On Error Resume Next
        If Store = 0 Then Store = StoresSld
        If SmLoc <> Store Then GetSalesmanDatabase(Store)
        Sm = GetSalesmanDatabase(Store)

        For I = LBound(Sm, 1) To UBound(Sm, 1)
            'If Sm(I, 2) = EmployeeNumber Then
            If Sm(I, 1) = EmployeeNumber Then
                'getSalesName = Sm(I, 1) & " "
                getSalesName = Sm(I, 0) & " "
                Exit Function
            End If
        Next
        getSalesName = ""
    End Function

    ' This loads the Cache, and always loads the specified store
    Public Function GetSalesmanDatabase(Optional ByVal Store as integer = 0, Optional ByVal SalesOnly As Boolean = False, Optional ByVal IncludeDisabled As Boolean = False) As Object
        '::::GetSalesmanDatabase
        ':::SUMMARY
        ': Loads the cache on specified store
        ':::DESCRIPTION
        ': This function is used to loads the Cache and alsways loads the specified store.
        ':::PARAMETERS
        ': - Store
        ': - SalesOnly
        ': - IncludeDisabled
        Const FCount as integer = 8
        Dim Emp As clsEmployee, E As Integer, Tmp(0, 0) As String, Tmp2(0, 0) As String

        If Store < 0 Then Store = 1
        If Store < 1 Or Store > Setup_MaxStores Then Store = StoresSld
        If Store < 1 Or Store > Setup_MaxStores Then Store = 1
        '  If Not FileExists(GetDatabaseAtLocation(Store)) Then GoTo HandleErr   ' No database

        On Error GoTo HandleErr

        Emp = New clsEmployee
        Emp.DataAccess.DataBase = GetDatabaseAtLocation(Store)
        Emp.DataAccess.Records_Open("ID", "None")
        If Emp.DataAccess.Record_EOF Then GoTo HandleErr

        E = 0
        ReDim Tmp(0 To FCount - 1, 1)

        Do While Emp.DataAccess.Records_Available
            Emp.cDataAccess_GetRecordSet(Emp.DataAccess.RS)
            If (IncludeDisabled Or Emp.Active) And ((Not SalesOnly) Or Trim(Emp.SalesID) <> "") Then
                ReDim Preserve Tmp(0 To FCount - 1, E)
                Tmp(0, E) = Trim(Emp.LastName)
                Tmp(1, E) = Emp.SalesID
                Tmp(2, E) = Emp.CommRate
                Tmp(3, E) = Emp.Active
                Tmp(4, E) = Emp.ID
                Tmp(5, E) = Emp.Privs
                Tmp(6, E) = Emp.Password
                Tmp(7, E) = Emp.CommTable
                E = E + 1
            End If
        Loop
        Err.Clear()    ' Records_Available is leaving an error object open?
        If E = 0 Then GoTo HandleErr

        'ReDim Tmp2(LBound(Tmp, 2) To UBound(Tmp, 2), 1 To FCount)
        'ReDim Tmp2(0 To UBound(Tmp, 2), 0 To FCount - 1)
        ReDim Tmp2(0 To UBound(Tmp, 2), 0 To FCount - 1)
        'For E = LBound(Tmp, 2) To UBound(Tmp, 2)
        For E = LBound(Tmp, 2) To UBound(Tmp, 2)
            Tmp2(E, 0) = Tmp(0, E)
            Tmp2(E, 1) = Tmp(1, E)
            Tmp2(E, 2) = Tmp(2, E)
            Tmp2(E, 3) = Tmp(3, E)
            Tmp2(E, 4) = Tmp(4, E)
            Tmp2(E, 5) = Tmp(5, E)
            Tmp2(E, 6) = Tmp(6, E)
            Tmp2(E, 7) = Tmp(7, E)
        Next

        GetSalesmanDatabase = Tmp2
        Sm = GetSalesmanDatabase          ' Private Cache
        SmLoc = Store                     ' Private Cache
        Exit Function

HandleErr:
        'MsgBox "Can't load Salesman database. Try restarting the program." & vbCrLf & "If this problem persists, please contact CDS.", vbCritical, "Error"
        ReDim Tmp(0, 7)
        Tmp(0, 1) = "HOUSE"
        Tmp(0, 2) = "99"
        Tmp(0, 3) = ""
        Tmp(0, 4) = False
        Tmp(0, 5) = ""
        Tmp(0, 6) = ""
        Tmp(0, 7) = ""
        GetSalesmanDatabase = Tmp
        'Sm = GetSalesmanDatabase
    End Function

    Public Function TranslateSalesman(ByVal Salesman As String, Optional ByVal Num As Integer = -1, Optional ByVal Store As Integer = 0) As String
        '::::TranslateSalesman
        ':::SUMMARY
        ': Translate from Salesman Number List to Salesman Name (by Index)
        ':::DESCRIPTION
        ': Given a list of (space-separated, up to three) employee numbers, return the specified item employee name.
        ':::PARAMETERS
        ': - Salesman
        ': - Num
        ': - Store
        ':::RETURN
        ': String
        ':::SEE ALSO
        ': TranslateSalesmen
        Dim E() As String
        If (Num < 0) Then TranslateSalesman = TranslateSalesmen(Salesman, Store) : Exit Function
        If Trim(Salesman) = "" And Num = 0 Then
            TranslateSalesman = "HOUSE"
            Exit Function
        End If

        E = Split(Trim(Salesman), " ")
        If Num < LBound(E) Or Num > UBound(E) Then Exit Function
        TranslateSalesman = getSalesName(CStr(E(Num)), Store)
    End Function

    Public Function TranslateSalesmen(ByVal Salesmen As String, Optional ByVal Store As Integer = 0) As String
        '::::TranslateSalesmen
        ':::SUMMARY
        ': Translate an entire list of employee numbers to salesmen
        ':::DESCRIPTION
        ': This function is used to specify available salesmen for any sale.So,Customer can select among the available salesmen.
        ': Because sometimes when customer made any new sale, its directly goes to STORE instead of specifying any salesman.
        ':::PARAMETERS
        ': - Salesman
        ': - Num
        ': - Store
        ':::RETURN
        ': String

        Dim L As Object

        TranslateSalesmen = String.Empty
        If Store = 0 Then Store = StoresSld
        If Trim(Salesmen) = "" Then TranslateSalesmen = "HOUSE" : Exit Function
        For Each L In Split(Trim(Salesmen), " ")
            TranslateSalesmen = TranslateSalesmen & IIf(Len(TranslateSalesmen) > 0, ", ", "") & Trim(getSalesName(CStr(L), Store))
        Next
    End Function

    Public Function QueryUserGroupPrivString(ByVal StoreNum As String, ByVal GroupAbbr As String) As String
        '::::QueryUserGroupPrivString
        ':::SUMMARY
        ': Group Priv String
        ':::DESCRIPTION
        ': Custom Design Software can allow you to customize each user group. Each user group is preconfigured to allow permissions for each group.
        ': This function is used to go through all records present in UserGroups table and can provide query for UserGroups Privs using parameters.
        ':::PARAMETERS
        ': - StoreNum
        ': - GroupAbbr
        ':::RETURN
        ': String

        Dim objGroup As clsUserGroup
        objGroup = New clsUserGroup

        objGroup.DataAccess.Records_Open("GroupName", "None")
        Do While objGroup.DataAccess.Records_Available
            If objGroup.Abbrev = GroupAbbr Then QueryUserGroupPrivString = objGroup.Privs
        Loop
        DisposeDA(objGroup)
    End Function

End Module
