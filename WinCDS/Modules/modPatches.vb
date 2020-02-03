Module modPatches
    Private PatchTables As Collection
    Public AutoPatching As Boolean
    Private Enum PatchStates
        PatchStatus_NeverRun = -1
        PatchStatus_Running = 0
        PatchStatus_Complete = 1
        PatchStatus_Unknown = -2
    End Enum

    Public Function IsPatchApplied(ByVal PatchName As String, Optional ByVal StoreNo As Integer = -1) As Boolean
        Dim N As String
        Static X As String, Y As Boolean
        If PatchName = X Then IsPatchApplied = Y : Exit Function

        Select Case StoreNo
            Case -1 : N = GetDatabaseAtLocation()
            Case 0 : N = GetDatabaseInventory()
            Case Else : N = GetDatabaseAtLocation(StoreNo)
        End Select

        IsPatchApplied = PatchStatus(N, PatchName) = PatchStates.PatchStatus_Complete
        X = PatchName
        Y = IsPatchApplied
    End Function

    Private Function PatchStatus(ByVal DBName As String, ByVal PatchName As String) As PatchStates
        ' Assume the PatchHistory table exists.
        Dim SQL As String, RS As ADODB.Recordset
        '  SQL = "SELECT Top 1 Status, ApplyDate FROM PatchHistory WHERE Name='" & ProtectSQL(PatchName) & "' ORDER BY ID Desc"
        '  Set RS = GetRecordsetBySQL(SQL, False, DBName)

        OpenPatchTables()
        On Error Resume Next
        RS = PatchTables(DBName)
        RS.Filter = "Name='" & ProtectSQL(PatchName) & "'"

        '  If Dir(DBName) = "" Then PatchStatus = 1: Exit Function

        If RS Is Nothing Then
            PatchStatus = PatchStates.PatchStatus_Unknown  ' Error determining patch status
        ElseIf RS.EOF Then
            PatchStatus = PatchStates.PatchStatus_NeverRun ' Never been patched
            RS = Nothing
        Else
            PatchStatus = FitRange(PatchStates.PatchStatus_NeverRun, IfNullThenZero(RS("Status").Value), PatchStates.PatchStatus_Complete) ' Patched, or being patched.
            If PatchStatus = PatchStates.PatchStatus_Running And DateDiff("d", Today, RS("ApplyDate").Value) > 0 Then PatchStatus = PatchStates.PatchStatus_NeverRun ' Patch broke?
            RS = Nothing
        End If
    End Function

    Private Sub OpenPatchTables()
        Dim I As Integer, RS As ADODB.Recordset, SQL As String, DBName As String

        If Not PatchTables Is Nothing Then Exit Sub


        PatchTables = New Collection
        SQL = "SELECT Name, Status, ApplyDate FROM PatchHistory ORDER BY ID DESC"
        DBName = GetDatabaseInventory()
        If Not PatchTableExists(DBName) Then BuildPatchTable(DBName)
        RS = GetRecordsetBySQL(SQL, , DBName)
        PatchTables.Add(RS, DBName)

        On Error GoTo NoMoreDBs
        For I = 1 To MaxPatchStore
            DBName = GetDatabaseAtLocation(I)
            If Dir(DBName) <> "" Then
                If Not PatchTableExists(DBName) Then BuildPatchTable(DBName)
                RS = GetRecordsetBySQL(SQL, , DBName, True)
                PatchTables.Add(RS, DBName)
            End If
        Next
NoMoreDBs:
    End Sub

    Private Function PatchTableExists(ByVal DBName As String) As Boolean
        Dim SQL As String, RS As ADODB.Recordset
        SQL = "SELECT COUNT(*) FROM PatchHistory"
        RS = GetRecordsetBySQL(SQL, False, DBName, True)
        If Not RS Is Nothing Then
            PatchTableExists = True
            RS.Close()
            RS = Nothing
        End If
    End Function

    Private Sub BuildPatchTable(ByVal DBName As String)
        Dim SQL As String
        SQL = "CREATE TABLE PatchHistory " &
        "(ID int identity, Name varchar(40), ApplyDate datetime, Status int, Duration int, " &
        "CONSTRAINT AutoIncrementTest_PrimaryKey PRIMARY KEY (ID))"
        ExecuteRecordsetBySQL(SQL, False, DBName, True, "Could not build Patch History table in " & DBName & "!")
    End Sub

    Private ReadOnly Property MaxPatchStore() As Integer
        Get
            MaxPatchStore = ActiveNoOfLocations '  Setup_MaxStores
        End Get
    End Property

    'called from frmSplash
    Public Function MoveUserRegistryToSystem() As Boolean
        ' If the registry's already moved, don't do it again.
        If GetSetting(RegistrySection, RegistryAppName, "IsServer") = "" Then Exit Function

        MoveUserRegistryKeyToSystem RegistryAppName, "IsServer"
  MoveUserRegistryKeyToSystem RegistryAppName, "Location"
  MoveUserRegistryKeyToSystem RegistryAppName, "Label Printer"
  MoveUserRegistryKeyToSystem RegistryAppName, "Cash Register Printer"
  MoveUserRegistryKeyToSystem RegistryAppName & "\BarCode", "COM Port"
  MoveUserRegistryKeyToSystem RegistryAppName & "\Cash Drawer", "COM Port"
  MoveUserRegistryKeyToSystem RegistryAppName & "\ScanPal 2", "Settings"
  MoveUserRegistryKeyToSystem RegistryAppName & "\ScanPal 2", "COM Port"
  On Error Resume Next
        DeleteSetting RegistrySection, RegistryAppName & "\BarCode"
  DeleteSetting RegistrySection, RegistryAppName & "\Cash Drawer"
  DeleteSetting RegistrySection, RegistryAppName & "\ScanPal 2"
  DeleteSetting RegistrySection, RegistryAppName
End Function

    Public Sub AutoPatch()
        ' This function should be called when the program loads (and after restore).
        ' It will check the databases for each required patch, and apply changes.
        BuildPatchList                  ' Create the array of patch definitions.
        If DateBefore(MostRecentPatchDate, GetLastPatchDate, False) Then Exit Sub

        OpenPatchTables() ' would do it automatically, but we'll do it here too
        If Dir(GetDatabaseInventory) = "" Then
            MsgBox "The database [" & GetDatabaseInventory() & "] could not be found..." & vbCrLf & "This is a critical error and you should shut down immediately.", vbCritical, "Database Not Found!"
    Exit Sub
        End If

        AutoPatching = True
        CheckRequiredPatches GetDatabaseInventory, False
  Dim I As Long
        For I = 1 To MaxPatchStore
            If Dir(GetDatabaseAtLocation(I)) <> "" Then CheckRequiredPatches GetDatabaseAtLocation(I), True
  Next
        ClosePatchTables
        SetLastPatchDate MostRecentPatchDate

  AutoPatching = False

    End Sub

End Module
