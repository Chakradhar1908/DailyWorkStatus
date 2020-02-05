Module modPatches
    Private PatchTables As Collection
    Public AutoPatching As Boolean
    Private MostRecentPatchDate As Date
    Public PatchList() As PatchDef
    Public Practicing As Boolean
    Private Enum PatchStates
        PatchStatus_NeverRun = -1
        PatchStatus_Running = 0
        PatchStatus_Complete = 1
        PatchStatus_Unknown = -2
    End Enum
    Public Structure PatchDef
        Dim PatchDate As Date
        Dim Name As String
        Dim Desc As String
        Dim Required As Boolean
        Dim Repeatable As Boolean
        Dim Inventory As Boolean
        Dim Stores As Boolean
        Dim AllowPractice As Boolean
    End Structure


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

        MoveUserRegistryKeyToSystem(RegistryAppName, "IsServer")
        MoveUserRegistryKeyToSystem(RegistryAppName, "Location")
        MoveUserRegistryKeyToSystem(RegistryAppName, "Label Printer")
        MoveUserRegistryKeyToSystem(RegistryAppName, "Cash Register Printer")
        MoveUserRegistryKeyToSystem(RegistryAppName & "\BarCode", "COM Port")
        MoveUserRegistryKeyToSystem(RegistryAppName & "\Cash Drawer", "COM Port")
        MoveUserRegistryKeyToSystem(RegistryAppName & "\ScanPal 2", "Settings")
        MoveUserRegistryKeyToSystem(RegistryAppName & "\ScanPal 2", "COM Port")
        On Error Resume Next
        DeleteSetting(RegistrySection, RegistryAppName & "\BarCode")
        DeleteSetting(RegistrySection, RegistryAppName & "\Cash Drawer")
        DeleteSetting(RegistrySection, RegistryAppName & "\ScanPal 2")
        DeleteSetting(RegistrySection, RegistryAppName)
    End Function

    Private Function MoveUserRegistryKeyToSystem(ByVal Section As String, ByVal KeyName As String) As Boolean
        SaveSystemSetting(RegistrySection, Section, KeyName, GetSetting(RegistrySection, Section, KeyName))
        On Error Resume Next
        DeleteSetting(RegistrySection, Section, KeyName)
    End Function

    Public Sub AutoPatch()
        ' This function should be called when the program loads (and after restore).
        ' It will check the databases for each required patch, and apply changes.
        BuildPatchList                  ' Create the array of patch definitions.
        If DateBefore(MostRecentPatchDate, GetLastPatchDate, False) Then Exit Sub

        OpenPatchTables() ' would do it automatically, but we'll do it here too
        If Dir(GetDatabaseInventory) = "" Then
            MessageBox.Show("The database [" & GetDatabaseInventory() & "] could not be found..." & vbCrLf & "This is a critical error and you should shut down immediately.", "Database Not Found!")
            Exit Sub
        End If

        AutoPatching = True
        CheckRequiredPatches(GetDatabaseInventory, False)
        Dim I As Integer
        For I = 1 To MaxPatchStore
            If Dir(GetDatabaseAtLocation(I)) <> "" Then CheckRequiredPatches(GetDatabaseAtLocation(I), True)
        Next
        ClosePatchTables()
        SetLastPatchDate(MostRecentPatchDate)

        AutoPatching = False

    End Sub

    Private Sub ClosePatchTables()
        On Error Resume Next
        Dim L As Object
        If PatchTables Is Nothing Then Exit Sub
        For Each L In PatchTables
            L.Close
        Next
        PatchTables = Nothing
    End Sub

    Private Sub CheckRequiredPatches(ByVal DBName As String, ByVal IsStore As Boolean)
        Dim I As Integer

        If Not PatchTableExists(DBName) Then BuildPatchTable(DBName)
        For I = LBound(PatchList) To UBound(PatchList)
            If Not PatchList(I).Required Then GoTo Skip
            If ((IsStore And PatchList(I).Stores) Or (PatchList(I).Inventory And Not IsStore)) Then
                If IsFormLoaded("frmBackUpGeneric") Then
                    Dim SN As String
                    SN = IIf(GetStoreNumber(DBName) > 0, "Store #" & GetStoreNumber(DBName) & " DB", "Inventory DB")
                    frmBackUpGeneric.Status = "Patching " & SN & IIf(IsDevelopment, ": " & PatchList(I).Name, "")
                End If
                ApplyPatch(PatchList(I).Name, DBName, , True, True)  ' This filters out pre-applied patches.
            End If
Skip:
        Next
    End Sub

    Private Function FindPatch(ByVal PatchName As String) As PatchDef
        Dim I As Integer
        For I = LBound(PatchList) To UBound(PatchList)
            If PatchList(I).Name = PatchName Then
                FindPatch = PatchList(I)
                Exit Function
            End If
        Next
    End Function

    Private Function UpdateInvRecTelephoneRecords() As Boolean
        CleanTelephoneRecords("PO", "POID", "ShipToTele", GetDatabaseInventory)
        UpdateInvRecTelephoneRecords = True
    End Function

    Private Sub CleanTelephoneRecords(ByVal TableName As String, ByVal IDField As String, ByRef FieldList As Object, ByVal DataBase As String)
        On Error GoTo HandleErr
        'If Not IsArray(FieldList) Then FieldList = Array(FieldList)
        If Not IsArray(FieldList) Then FieldList = {FieldList}
        Dim SQL As String, FieldName As Object
        SQL = "[" & IDField & "], "
        For Each FieldName In FieldList
            SQL = SQL & "[" & FieldName & "], "
        Next
        SQL = Left(SQL, Len(SQL) - 2)
        SQL = "Select " & SQL & " From [" & TableName & "]"

        Dim RS As ADODB.Recordset
        RS = GetRecordsetBySQL(SQL, False, DataBase)
        Do Until RS.EOF
            For Each FieldName In FieldList
                If CleanAni("" & RS(FieldName).Value) <> "" Then RS(FieldName).Value = CleanAni("" & RS(FieldName).Value)
            Next
            RS.Update()
            RS.MoveNext()
        Loop
        Dim Fred As New CDbAccessGeneral
        Fred.dbOpen(DataBase)
        RS.ActiveConnection = Fred.mConnection
        RS.UpdateBatch()
        DisposeDA(RS, Fred)
        Fred.dbClose()
        Exit Sub

HandleErr:
        MessageBox.Show("Error converting telephone records in " & TableName & ".", "Error!")
    End Sub

    Private Function UpdateTelephoneRecords(ByVal DBName As String, ByVal Complete As Boolean) As Boolean
        ' These tables are in the old databases.
        CleanTelephoneRecords("InstallmentInfo", "ArNo", "Telephone", DBName)
        CleanTelephoneRecords("Service", "ServiceOrderNo", {"Telephone"}, DBName)

        ' These are only in the new, and will be handled by the conversion code.
        If Complete Then
            CleanTelephoneRecords("Mail", "Index", {"Tele", "Tele2"}, DBName)
            CleanTelephoneRecords("MailShipTo", "Index", "Tele", DBName)
            CleanTelephoneRecords("GrossMargin", "MarginLine", "Tele", DBName)
        End If
    End Function

    Private Function PatchNonItemLocations(ByVal DBName As String) As Boolean
        ' SUB, TAX*, PAYMENT, --- Adj ---, etc must have Location=0!
        ExecuteRecordsetBySQL _
    ("Update GrossMargin set Location=0 where Location<>0 and Trim(Style) in ('SUB','TAX1','TAX2','PAYMENT','--- Adj ---','STAIN','NOTES','DEL','LAB')",
    False, DBName, True)
        PatchNonItemLocations = True
    End Function

    Private Function PatchSalesTaxReceived(ByVal DBName As String) As Boolean
        Dim TaxType As String, I As Integer, TaxList As Object, NewType As Integer
        ExecuteRecordsetBySQL("Update GrossMargin set Quantity=1 where Quantity=0 and Style='TAX2'", False, DBName, True)
        ExecuteRecordsetBySQL("Update Audit set TaxCode=1 where TaxCode=0", False, DBName, True)
        TaxList = QuerySalesTax2List()
        NewType = 0
        For I = LBound(TaxList) To UBound(TaxList)
            TaxType = TaxList(I)
            NewType = NewType + 1
            ExecuteRecordsetBySQL _
      ("Update GrossMargin set Quantity=" & NewType & " where Style='TAX2' and Desc like 'SALES TAX%" & ProtectSQL(TaxType) & " ='",
      False, DBName, True)
        Next
        ExecuteRecordsetBySQL("UPDATE GrossMargin,Audit SET Audit.TaxCode = GrossMargin.Quantity WHERE GrossMargin.Style Like ""TAX_"" and Trim(GrossMargin.SaleNo) = Trim(Audit.SaleNo)", False, DBName, True)

        ' Update Tax Received - sales were recording this as part of Delivered Sales.
        Dim RS As ADODB.Recordset
        RS = GetRecordsetBySQL("SELECT SaleNo, (select sum(taxcharged1) from audit as audit2 where audit2.saleno=audit.saleno) AS Charged From Audit Where Audit.TaxRec1 = 0 And Audit.Name1 Like ""DS%""",
    False, DBName, True)
        Do Until RS.EOF
            ExecuteRecordsetBySQL _
      ("UPDATE Audit SET DelSls=DelSls-" & RS("Charged").Value & ", TaxRec1=TaxRec1+" & RS("Charged").Value & " where SaleNo='" & RS("SaleNo").Value & "' and Name1 like 'DS %' And TaxRec1=0",
      False, DBName, True)
            RS.MoveNext()
            Application.DoEvents()
        Loop
        PatchSalesTaxReceived = True
    End Function

    Private Function PatchNoteDates(ByVal DBName As String) As Boolean
        ExecuteRecordsetBySQL("ALTER TABLE ArNotes ADD COLUMN NoteDate DateTime", False, DBName, True)
        ExecuteRecordsetBySQL("ALTER TABLE SaleNotes ADD COLUMN NoteDate DateTime", False, DBName, True)
        PatchNoteDates = True
    End Function

    Private Function PatchDepartmentNumbers(ByVal DBName As String) As Boolean
        ' Department was text(1), needs to be text(3).
        If DBName = GetDatabaseInventory() Then
            ExecuteRecordsetBySQL("ALTER TABLE Search ALTER COLUMN Dept TEXT(3)", False, DBName)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Dept TEXT(3)", False, DBName)
        Else
            ExecuteRecordsetBySQL("ALTER TABLE GrossMargin ALTER COLUMN DeptNo TEXT(3)", False, DBName)
            ExecuteRecordsetBySQL("ALTER TABLE GrossMarginTmp ALTER COLUMN DeptNo TEXT(3)", False, DBName)
        End If
        PatchDepartmentNumbers = True
    End Function

    Private Function PatchServiceTable(ByVal DBName As String) As Boolean
        On Error GoTo DetailAdded
        ExecuteRecordsetBySQL("alter table Service alter column [Type] TEXT(5);", False, DBName, False)
        ExecuteRecordsetBySQL("alter table Service drop column [Setail];", False, DBName, True)
        ExecuteRecordsetBySQL("alter table Service add column [Detail] TEXT(50);", False, DBName, True)
DetailAdded:
        PatchServiceTable = True
    End Function

    Private Function PatchKitSKU2(ByVal DBName As String) As Boolean
        If GetStoreNumber(DBName) = 1 Then
            'Take the item1rec than find the Vendor
            Dim RS As ADODB.Recordset, RS2 As ADODB.Recordset
            RS = GetRecordsetBySQL("SELECT [Item1Rec] FROM [InvKit]", False, DBName)
            Do Until RS.EOF
                RS2 = GetRecordsetBySQL("SELECT [Vendor] FROM [2Data] WHERE [RN] = " & IfNullThenZero(RS("Item1Rec").Value), False, GetDatabaseInventory)
                If Not RS2.EOF Then
                    ExecuteRecordsetBySQL("UPDATE [InvKit] SET [KitSKU]='" & ProtectSQL(IfNullThenNilString(RS2("vendor").Value)) & "' WHERE [Item1Rec]=" & IfNullThenZero(RS("Item1Rec").Value), False, DBName)
                End If
                RS.MoveNext()
            Loop
        End If
        PatchKitSKU2 = True
    End Function

    Private Function PatchNonTaxableAmount(ByVal DBName As String) As Boolean
        ExecuteRecordsetBySQL("update Holding set nontaxable=sale where leaseno not in (select saleno from grossmargin where left(style,3)=""TAX"")", False, DBName)
        ExecuteRecordsetBySQL("update Audit set TaxCode=1 where TaxCode=0", False, DBName)
        PatchNonTaxableAmount = True
    End Function

    Private Function FixDecimalQuantityDatabase(ByVal dbname As String) As Boolean
        Dim RS As ADODB.Recordset

        If dbname = GetDatabaseInventory() Then
            ' Update the PO table so we can have decimal quantities.
            ExecuteRecordsetBySQL("ALTER TABLE PO ALTER COLUMN InitialQuantity Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE PO ALTER COLUMN Quantity Single DEFAULT 0", , dbname)

            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN AmtSold Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN NewStock Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN SpecOrd Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN LAW Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc1 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc2 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc3 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc4 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc5 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc6 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc7 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc8 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc9 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc10 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc11 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc12 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc13 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc14 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc15 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE Detail ALTER COLUMN Loc16 Single DEFAULT 0", , dbname)

            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnHand Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Available Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN PoSold Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Sales1 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Sales2 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Sales3 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Sales4 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN PSales1 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN PSales2 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN PSales3 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN PSales4 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc1Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc2Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc3Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc4Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc5Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc6Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc7Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc8Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc9Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc10Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc11Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc12Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc13Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc14Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc15Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN Loc16Bal Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder1 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder2 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder3 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder4 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder5 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder6 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder7 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder8 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder9 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder10 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder11 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder12 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder13 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder14 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder15 Single DEFAULT 0", , dbname)
            ExecuteRecordsetBySQL("ALTER TABLE [2Data] ALTER COLUMN OnOrder16 Single DEFAULT 0", , dbname)
        End If
        FixDecimalQuantityDatabase = True
    End Function

    Private Function FixDetailHistory(ByVal DBName As String, Optional ByRef ParentForm As Form = Nothing) As Boolean
        Dim RS As ADODB.Recordset
        '  Dim Store as integer
        Dim PFOK As Boolean
        If Not ParentForm Is Nothing Then
            If ParentForm.Name = "Practice" Then PFOK = True
        End If

        '  If PFOK Then Practice.pbarStore.Visible = True
        If PFOK Then Practice.pbarMargin.Visible = True
        '  For Store = 1 To 8
        '    DBName = GetDatabaseAtLocation(Store)
        ' We need to update all the margin records to match their right 2Data.
        ' This is difficult because 2Data is in a different database.  Ugh.
        ' But we know how to run queries on other databases, so we'll do it.
        If Dir(DBName) = "" Then Exit Function ' For
        ExecuteRecordsetBySQL("UPDATE GrossMargin SET Style=""SUB"", Rn=0 WHERE Desc=""Sub Total =""", , DBName)
        ExecuteRecordsetBySQL("UPDATE GrossMargin SET Style=""TAX1"", Rn=0 WHERE left(Desc,14)=""SALES TAX DIFF""", , DBName)
        ExecuteRecordsetBySQL("UPDATE GrossMargin SET Style=""--- Adj ---"", Rn=0 WHERE Desc LIKE ""*Adjustments*""", , DBName)
        ExecuteRecordsetBySQL("UPDATE GrossMargin INNER JOIN [" & GetDatabaseInventory() & "].2Data ON GrossMargin.Style = [2Data].Style SET GrossMargin.Rn = [2Data].[rn]", , DBName)

        Dim LastSale As String
        Dim RowsForward As Integer, SearchStyle As String, SearchQuan As Double, SearchStat As String, SearchPrice As Double
        Dim OrigDetail As Integer
        Dim GM As CGrossMargin
        LastSale = ""
        RS = GetRecordsetBySQL("SELECT SaleNo, Detail AS DetailLine, Count(MarginLine) AS MarginLines FROM GrossMargin GROUP BY SaleNo, Detail HAVING Detail<>0 AND Count(MarginLine)>1 ORDER BY SaleNo, Detail;", False, DBName, True)
        If PFOK Then Practice.pbarMargin.Value = 0
        If PFOK And RS.RecordCount > 0 Then Practice.pbarMargin.Maximum = RS.RecordCount
        Do Until RS.EOF
            If Trim(RS("SaleNo").Value) <> "" And Mid(RS("SaleNo").Value, 1, 1) <> Chr(0) Then

                DisposeDA(GM)
                GM = New CGrossMargin
                GM.DataAccess.DataBase = DBName
                GM.DataAccess.Records_OpenFieldIndexAt("SaleNo", RS("SaleNo").Value, "MarginLine")

                If GM.DataAccess.Records_Available Then
                    ' We are now on the first Margin record.
                    If LastSale <> GM.SaleNo Then
                        '          If GM.SaleNo = "1-0656" Then Stop
                        '          If GM.SaleNo = "1-10332" Then Stop

                        ' Gather the matching detail records.
                        Dim Detail As CInventoryDetail
                        Detail = New CInventoryDetail
                        Detail.DataAccess.Records_OpenSQL("SELECT * FROM Detail WHERE SaleNo=""" & ProtectSQL(GM.SaleNo) & """ ORDER BY DetailID")

                        Do Until GM.DataAccess.Record_EOF
                            If IsNonItemStyle(GM.Style) Or GM.Status Like "*SS*" Or GM.Quantity = 0 Or Left(Trim(GM.Desc), 14) = "SALES TAX DIFF" Or Trim(GM.Desc) = "Sub Total =" Then
                                ' Clear detail pointers on non-item styles and SS records.
                                GM.Detail = 0
                                GM.RN = 0
                                GM.Save()
                            Else
                                ' Update detail pointers and detail records on item styles.
                                If GM.Status Like "VD*" Or GM.Status = "VOID" Then
                                    ' If it's a negative quantity, we should go backward to find the matching row...
                                    If GM.Quantity < 0 Then
                                        SearchStyle = GM.Style
                                        SearchQuan = GM.Quantity
                                        SearchStat = GM.Status
                                        SearchPrice = GM.SellPrice
                                        Do Until GM.DataAccess.Record_BOF
                                            GM.DataAccess.Records_MovePrevious()
                                            RowsForward = RowsForward + 1
                                            If GM.Style = SearchStyle And SameStatus(GM.Status, SearchStat) And GM.Quantity = -SearchQuan And GM.SellPrice = -SearchPrice Then
                                                ' This record matches all we could hope for...
                                                OrigDetail = GM.Detail
                                                Exit Do
                                            End If
                                        Loop
                                        Do Until RowsForward = 0
                                            GM.DataAccess.Records_MoveNext()
                                            RowsForward = RowsForward - 1
                                        Loop
                                        GM.Detail = OrigDetail
                                        GM.Save()

                                        If GM.Status = "VDxSO" Then
                                            LinkMarginAndDetail(GM, Detail, "IN")
                                        End If
                                    Else
                                        LinkMarginAndDetail(GM, Detail, "VD")
                                    End If
                                ElseIf GM.Status Like "x*" Then
                                    ' Void record.
                                    If GM.Quantity > 0 Then
                                        ' This is an original sale record, which should point to a "VD" detail record.
                                        ' The detail record should point back to this one.
                                        LinkMarginAndDetail(GM, Detail, "VD")
                                        ' Loop forward to the matching void line with a negative quantity, mark its detail negative.
                                        ' Then loop back to this row.
                                        SearchStyle = GM.Style
                                        SearchQuan = GM.Quantity
                                        SearchStat = GM.Status
                                        SearchPrice = GM.SellPrice
                                        OrigDetail = GM.Detail
                                        Do While GM.DataAccess.Records_Available
                                            RowsForward = RowsForward + 1
                                            If GM.Style = SearchStyle And SameStatus(GM.Status, SearchStat) And GM.Quantity = -SearchQuan And GM.SellPrice = -SearchPrice Then
                                                ' This record matches all we could hope for...
                                                GM.Detail = -OrigDetail
                                                GM.Save()
                                                Exit Do
                                            End If
                                        Loop
                                        Do Until RowsForward = 0
                                            GM.DataAccess.Records_MovePrevious()
                                            RowsForward = RowsForward - 1
                                        Loop
                                    Else
                                        ' This is a void record, which should already be pointing to the right detail.
                                        ' The detail does not point back to this record.  That's fine.
                                        '  If this is a *REC* or *RC*, the next line should be "IN".
                                        GM.Detail = -GM.Detail
                                        GM.Save()

                                        If GM.Status Like "*REC*" Or GM.Status Like "*RC*" Or GM.Status Like "*SORE*" Then
                                            LinkMarginAndDetail(GM, Detail, "IN")
                                        End If
                                    End If
                                ElseIf GM.Status Like "*DEL*" Then
                                    ' Delivered record.
                                    LinkMarginAndDetail(GM, Detail, "DS")
                                Else
                                    LinkMarginAndDetail(GM, Detail, "NS")
                                End If
                            End If

                            GM.DataAccess.Records_MoveNext()
                        Loop

                        Do While Detail.DataAccess.Records_Available
                            ' If there are more detail records available, they're bad!
                            Detail.MarginRn = 0
                            Detail.Trans = "VD"
                            Detail.Save()
                        Loop
                        Detail.DataAccess.Records_Close()
                        Detail = Nothing
                    End If
                    LastSale = GM.SaleNo
                    GM.DataAccess.Records_Close()
                    GM = Nothing
                End If
            End If

            If PFOK Then Practice.pbarMargin.Value = Practice.pbarMargin.Value + 1
            Application.DoEvents()
            RS.MoveNext()
        Loop
        '    If PFOK Then Practice.pbarStore.value = Practice.pbarStore.value + 1
        '  Next Store
        '  If PFOK Then Practice.pbarStore.Visible = False
        If PFOK Then Practice.pbarMargin.Visible = False
        DisposeDA(GM)
        FixDetailHistory = True
    End Function

    Private Sub LinkMarginAndDetail(ByRef GM As CGrossMargin, ByRef Detail As CInventoryDetail, ByVal Status As String)
        If Detail.DataAccess.Records_Available Then
            ' Use the next detail record.
        Else
            ' Create a new detail record!
            ' Set sale variables and such.. this shouldn't happen, but has to be accounted for.
            Detail.Lease1 = GM.SaleNo
            Detail.Name = GM.Name
            Detail.Misc = "Auto"
            Detail.DataAccess.Records_Add()    ' Save the new detail and get a DetailID.
        End If
        Detail.Style = GM.Style
        Detail.DDate1 = GM.SellDte
        GM.Detail = Detail.DetailID
        Detail.Trans = Status
        Detail.InvRn = GM.RN
        Detail.MarginRn = GM.MarginLine
        Detail.Store = GM.Location

        Select Case Status
            Case "DS", "NS", "VD"
                Detail.AmtS1 = GM.Quantity
                Detail.Ns1 = 0
                Detail.SO1 = 0
                Detail.LAW = 0
            Case "IN"
                Detail.AmtS1 = 0
                Detail.Ns1 = GM.Quantity
                Detail.SO1 = 0
                Detail.LAW = 0
            Case Else
                ' Do nothing.
        End Select

        ' This patch is and should be limited to 8 stores..
        ' no need to upgrade old patches because they should have already been done!
        Dim I As Integer
        For I = 1 To 8
            Detail.SetLocationQuantity(I, IIf(Detail.Store = I, GM.Quantity, 0))
        Next

        Detail.Save()
        GM.Save()
    End Sub

    Public Sub ApplyPatch(ByVal PatchName As String, ByVal DBName As String, Optional ByRef ParentForm As Form = Nothing, Optional ByVal Silent As Boolean = False, Optional ByVal FilterPreApplied As Boolean = False)
        ' Apply a single patch to the specified database..
        Dim StartTime As Date, Status As PatchStates
        Dim Result As Boolean, DeferPatch As Boolean
        StartTime = Now

        ' Stop if it's currently being applied, or has been applied and shouldn't be repeated.
        ' Also stop if the patch doesn't apply to this kind of database.
        Dim PatchObj As PatchDef
        PatchObj = FindPatch(PatchName)
        If PatchObj.Name = "" Then Exit Sub
        If DBName = GetDatabaseInventory() And Not PatchObj.Inventory Then Exit Sub
        If DBName <> GetDatabaseInventory() And Not PatchObj.Stores Then Exit Sub
        Status = PatchStatus(DBName, PatchName)
        If Status = PatchStates.PatchStatus_Unknown Then
            MessageBox.Show("There was a problem determining the status of this patch:  " & PatchName & vbCrLf & "DBName: " & DBName & vbCrLf &
           "If you encounter any errors, please restart the the software.", "Warning -- Problem Executing Patches!")
            Exit Sub
        End If
        If Status = PatchStates.PatchStatus_Running And Not Practicing Then Exit Sub
        If Status = PatchStates.PatchStatus_Complete And Not PatchObj.Repeatable Then Exit Sub ' Already been patched.
        If FilterPreApplied And Status = PatchStates.PatchStatus_Complete Then Exit Sub        ' AutoPatch ignores repeatable required patches, if they're been applied.

        Dim SQL As String
        SQL = "INSERT INTO PatchHistory (Name, ApplyDate, Status, Duration) VALUES ('" & ProtectSQL(PatchName) & "', #" & DateValue(StartTime) & "#, " & PatchStates.PatchStatus_Running & ", 0)"
        ExecuteRecordsetBySQL(SQL, , DBName)

        On Error Resume Next
        Select Case PatchName
            Case "Clean Telephone Numbers"
                If DBName = GetDatabaseInventory() Then
                    Result = UpdateInvRecTelephoneRecords()
                Else
                    Result = UpdateTelephoneRecords(DBName, True)
                End If
            Case "Non-Item Locations" : Result = PatchNonItemLocations(DBName)
            Case "Sales Tax Received" : Result = PatchSalesTaxReceived(DBName)
            Case "Note Dates" : Result = PatchNoteDates(DBName)
            Case "Department Numbers" : Result = PatchDepartmentNumbers(DBName)
            Case "Update Service Table Structure" : Result = PatchServiceTable(DBName)
            Case "Update KitSKU with Vendor Name" : Result = PatchKitSKU2(DBName)
            Case "Correct Nontaxable Sales" : Result = PatchNonTaxableAmount(DBName)
            Case "Decimal Quantity Database" : Result = FixDecimalQuantityDatabase(DBName)
            Case "Rebuild Detail" : Result = FixDetailHistory(DBName, ParentForm)
            Case "Add Employees to Database" : Result = CreateEmployeesTable(DBName)
                '            Case "Add User Groups to Database" : Result = CreateUserGroupsTable(DBName)
                '            Case "General Ledger Tracking Features" : Result = PatchGLTracking(DBName)
                '            Case "General Ledger Tracking Features II" : Result = PatchGLTracking2(DBName)
                '            Case "Freight Type" : Result = PatchFreightType(DBName)
                '            Case "ServicePartsOrder Table" : Result = PatchServicePartsOrder(DBName)
                '            Case "Config Table" : Result = PatchConfigTable(DBName)
                '            Case "ItemLocation Table" : Result = PatchItemLocation(DBName)
                '            Case "PoSold Rectification" : Result = PatchPoSold(DBName)
                '            Case "Update ItemLocation" : Result = PatchItemLocationUpdate(DBName)
                '            Case "Sales Tax Received Rectification" : Result = SalesTaxReceivedPatch(DBName)
                '            Case "ArApp Note Field" : Result = PatchArAppNote(DBName)
                '            Case "Cost Tracking System" : Result = CreateCostTrackingSystem(DBName)
                '            Case "InstallmentInfo.LastMetro426Status" : Result = Metro426ToInstallment(DBName)
                '            Case "Cost in Detail" : Result = PatchCostInDetail(DBName)
                '            Case "Separate Comm Table" : Result = PatchSeparateCommTables(DBName)
                '            Case "Decimal Kit Quantities" : Result = PatchDecimalKitQuantities(DBName)
                '            Case "Add Non-Taxable To Audit" : Result = PatchNonTaxableToAudit(DBName)
                '            Case "Re-number vendors" : Result = PatchRenumberVendors(DBName)
                '            Case "Extend Descriptions" : Result = PatchExtendDesc(DBName)
                '            Case "Commission Spiff To GM" : Result = PatchSpiffGM(DBName)
                '            Case "Extend Sale Notes" : Result = PatchSaleNotes(DBName)
                '            Case "Rectify Service Calls" : Result = PatchRecitfyServiceCalls(DBName)
                '            Case "Finance Charge Sales Tax" : Result = PatchFinanceChargeSalesTax(DBName)
                '            Case "Add Sales Split" : Result = PatchAddSalesSplit(DBName)
                '            Case "Correct Installment Rate" : Result = PatchInstallmentRate(DBName)
                '            Case "Fix Jerrys Connect" : Result = PatchJerrysConnect()
                '            Case "PoNotes Table" : Result = PatchPoNotesTable(DBName)
                '            Case "Accomodate Time Stops" : Result = PatchTimeStops(DBName)
                '            Case "Accomodate Time Stops II" : Result = PatchTimeStops2(DBName)
                '            Case "Accomodate Time Stops III" : Result = PatchTimeStops3(DBName)
                '            Case "Extend Comments" : Result = PatchExtendComments(DBName)
                '            Case "Pictures Table" : Result = PatchPicturesTable(DBName)
                '            Case "Add Spiff Field" : Result = PatchSpiff2Data(DBName)
                '            Case "Add Weekly Installments" : Result = PatchWeeklyInstallments(DBName)
                '            Case "Add Cubes" : Result = PatchAddCubes(DBName)
                '            Case "Clear ItemCost" : Result = PatchClearItemCost(DBName)
                '            Case "Fix Tax-Included Sales" : Result = PatchTaxIncludedSales(DBName)
                '            Case "Add Transfer Notes" : Result = PatchDetailNotes(DBName)
                '            Case "Life Type in Installment Info" : Result = PatchLifeType(DBName)
                '            Case "Patch DDelDat in DELTW Sales." : Result = PatchDDelDatInDELTW(DBName)
                '            Case "Tibbees Service" : Result = PatchTibbeesService(DBName)
                '            Case "InstallmentInfo.Satisfied" : Result = PatchInstallmentInfoSatisfied(DBName)
                '            Case "Turn Off Name AutoCorrect" : Result = PatchAutoCorrectName(DBName)
                '            Case "Set SubDatasheet to None" : Result = PatchSubDataSheetToNone(DBName)
                '            Case "Fix Installment TotPaid" : Result = PatchFixInstallmentTotPaid(DBName)
                '            Case "Fix Negative GMs" : Result = PatchFixNegativeGMs(DBName)
                '            Case "SIP Add" : Result = PatchSIPAdd(DBName)
                '            Case "IUI To InstallmentInfo" : Result = PatchIUI(DBName)
                '            Case "GM Indexes" : Result = PatchGMIndexes(DBName)
                '            Case "More Indexes" : Result = PatchMoreIndexes(DBName)
                '            Case "Package Fields" : Result = PatchPackageFields(DBName)
                '            Case "Update 2Data GM" : Result = Patch2DataGM(DBName)
                '            Case "Add [Cashier] to [Audit]" : Result = PatchCashierToAudit(DBName)
                ''    Case "Calculate Packages"
                '                                        ' BFH20100829
                '                                        ' Stepped Release...  Over course of a month.
                ''                                        Dim X as integer, F As Date
                ''                                        X = Asc(LCase(StoreSettings.Name)) - Asc("a")
                ''                                        F = DateAdd("d", X, #9/1/2010#)
                ''                                        If Practicing And DateBefore(Date, F) Then
                ''                                          If MsgBox(StoreSettings.Name & vbCrLf & "This store is not set to receive this patch until " & F & "." & vbCrLf2 & "Run Anyway?", vbQuestion + vbOKCancel, "Developer") = vbCancel Then
                ''                                            If DateBefore(Date, F) Then DeferPatch = True: GoTo EndPatches
                ''                                          End If
                ''                                        End If
                ''                                                  Result = PatchCalculatePackages(DBName)
                '            Case "Calculate Packages" : Result = PatchCalculatePackages(DBName)
                ''    Case "AP Check Name Field":                   Result = PatchAPCheckNameField(DBName)
                '            Case "TransID to GM" : Result = PatchTransIDToGM(DBName)
                '            Case "CreateOnlineOrderRecordTable" : Result = PatchCreateOnlineOrderRecordTable(DBName)
                '            Case "StoreCount32" : Result = PatchStoreCount32(DBName)
                '            Case "KitSKU" : Result = PatchKitSKU(DBName)
                '            Case "InstallmentInfo Indexes" : Result = PatchInstallmentInfoIndexes(DBName)
                '            Case "Adjustment TAX2 Loc" : Result = PatchAdjustmentTAX2Loc(DBName)
                '            Case "Patch Sales Notes Taxable+++" : Result = PatchSalesNotesTaxable(DBName)
                '            Case "Sale Mail Index" : Result = PatchSaleMailIndex(DBName)
                '            Case "Add Distributors" : Result = CreateDistributorsTable(DBName)
                '            Case "Add Perm Order Status" : Result = PatchPermOrderStatus(DBName)
                '            Case "Add Telephone Labels" : Result = PatchAddTelephoneLabels(DBName)
                '            Case "Add ArNo to Holding" : Result = PatchAddArNoToHolding(DBName)
                '            Case "BFMyer Commissions" : Result = PatchBFMyerCommissions(DBName)
                '            Case "Fix Short Vendor Numbers" : Result = PatchFormatVendorNo(DBName)
                '            Case "Fix Missing Vendor Numbers on Returns" : Result = PatchGMMissingVendorNo(DBName)
                '            Case "Store Setup to INI" : Result = PatchStoreSetupUpgrade(DBName)
                '            Case "Revolving Update Delivery Day" : Result = PatchRevolvingUpdateDeliveryDay(DBName)
                '            Case "Activate UseScheduledTask" : Result = PatchUseScheduledTask(DBName)
                '            Case "Copy Ashley Credentials" : Result = PatchCopyAshleyCredentials(DBName)
                '            Case "Fix Weekly Monthly Installment" : Result = PatchWeeklyMonthlyInstallmentInfo(DBName)
                '            Case "DeliveryTicketMessageFile" : Result = PatchDeliveryTicketMessageFile(DBName)
                '            Case "Fix DISCOUNT Sales+" : Result = PatchFixDiscountSales(DBName)
                '            Case "PatchConnectCmdv2" : Result = PatchConnectCmdv2(DBName)
                '            Case "FXFolder" : Result = PatchFxFolder(DBName)
                '            Case "FXFolder2" : Result = PatchFxFolder2(DBName)
                '            Case "Michaels Last Notice" : Result = PatchMichaelsLastNotice(DBName)
                '            Case "LastLateCharge Added" : Result = PatchLastLateCharge(DBName)
                '            Case "McClure Old Account Fix" : Result = PatchMcClureOldAccounts(DBName)
                '            Case "LastLateCharge Fix" : Result = PatchLastLateChargeFix(DBName)
                '            Case "AddTransactionsfldPosted***" : Result = PatchAddTransactionsfldPosted(DBName)
                '            Case "AddTerminalTracking" : Result = PatchAddTerminalTracking(DBName)
                '            Case "MarkPastInstallments" : Result = PatchMarkPastInstallments(DBName)
                '            Case "ExportInstallmentAccounts" : Result = PatchExportInstallmentAccounts(DBName)
                '            Case "Payment History Profile Keeping" : Result = PatchPaymentHistoryProfileKeeping(DBName)
                '            Case "Date Past Installments" : Result = PatchDatePastInstallments(DBName)
                '            Case "ExtendStyle" : Result = PatchExtendStyle(DBName)
                '            Case "Fix Warehouse Furniture 2.15.2018" : Result = PatchFixWarehouseFurniture02152018(DBName)
                '            Case "ClassicInteriorsAPVendors" : Result = PatchClassicSetVendorList(DBName)
                '            Case "HomecraftersRebuildSearch" : Result = PatchHomecraftersRebuildSearch(DBName)
                '            Case "Fix Homecrafters II" : Result = PatchFixHomecrafters(DBName)
                '            Case "Fix Lindsay 3336B" : Result = PatchFixLindsay3336B(DBName)
                '            Case "NewAgeChicagoFurniture20180329" : Result = NewAgeChicagoFurniture20180329(DBName)
                '            Case "UnloadGilsLoc20Bal20180521" : Result = UnloadGilsLoc20Bal20180521(DBName)
                '            Case "BudgetRemoveDuplicateVoid20180911-b" : Result = BudgetRemoveDuplicateVoid20180911(DBName)
                '            Case "CasaBellaRemoveAddOns20180915" : Result = CasaBellaRemoveAddOns20180915(DBName)
                '            Case "HomecraftersFixItem-20181022" : Result = HomecraftersFixItem20181022(DBName)

                '            Case "CleanConfig" : Result = PatchCleanConfigTable(DBName)
                '            Case "Michaels20190201" : Result = Michaels20190201(DBName)
                '            Case "Budget20190212-2" : Result = Budget20190212(DBName)
                '            Case "JAllen20190221" : Result = JAllen20190221(DBName)
                '            Case "Budget20190514" : Result = Budget20190514(DBName)

                ' Unknown patch?
            Case Else : DevErr("Unknown Patch: " & PatchName) ' :exit sub
        End Select


EndPatches:
        If DeferPatch Then
            SQL = "DELETE FROM PatchHistory WHERE [Name]='" & ProtectSQL(PatchName) & "' AND ApplyDate=#" & DateValue(StartTime) & "# AND Status=" & PatchStates.PatchStatus_Running
        ElseIf Not Result Or Err.Number <> 0 Then
            ' Returned false or an error occurred during the patch.
            ' Delete record to re-run again later.
            MessageBox.Show("Failed to apply patch [" & PatchName & "] to database [" & DBName & "]." & vbCrLf & "[" & Err.Number & "] " & Err.Description)
            Err.Clear()
            SQL = "DELETE FROM PatchHistory WHERE [Name]='" & ProtectSQL(PatchName) & "' AND ApplyDate=#" & DateValue(StartTime) & "# AND Status=" & PatchStates.PatchStatus_Running
        Else
            ' Successful.  Mark as done.
            SQL = "UPDATE PatchHistory SET Status=" & PatchStates.PatchStatus_Complete & ", Duration=" & DateDiff("s", StartTime, Now) & " WHERE Name='" & ProtectSQL(PatchName) & "' AND ApplyDate=#" & DateValue(StartTime) & "#"
        End If

        On Error GoTo 0
        If SQL <> "" Then ExecuteRecordsetBySQL(SQL, , DBName)

ExitApplyPatch:
        Application.DoEvents()
    End Sub

    Private Function CreateEmployeesTable(ByVal DBName As String) As Boolean
        CreateEmployeesTable = True
        '  Dim SQL As String
        '  SQL = "CREATE TABLE Employees " & _
        '        "(ID int identity, " & _
        '        "LastName varchar(40), " & _
        '        "SalesID varchar(3), " & _
        '        "CommRate varchar(8), " & _
        '        "Pwd varchar(20), " & _
        '        "Privs varchar(255), " & _
        '        "Active YesNo, " & _
        '        "CONSTRAINT Employees_PrimaryKey PRIMARY KEY (ID))"
        '  ExecuteRecordsetBySQL SQL, False, DBName, True, "Could not build Employees table in " & DBName & "!"
        '
        '  ' Insert the salesman file as active users, no passwords or privs.
        '  Dim Emp As clsEmployee, f1 As String, f2 As String, f3 As String, FNum as integer
        '  FNum = FreeFile
        '  On Error GoTo NoSalesmen
        '  Open SalesmanFile(GetStoreNumber(DBName)) For Input As #FNum
        '  Do While Not EOF(FNum)
        '    Set Emp = New clsEmployee
        '    Emp.DataAccess.DataBase = DBName
        '    Input #FNum, f1, f2, f3
        '    Emp.LastName = f1
        '    Emp.SalesID = f2
        '    Emp.CommRate = f3
        '    Emp.Active = True
        '    Emp.Privs = "ES" ' Everybody + Sales
        '    Emp.Save
        '  Loop
        '  Close #FNum
        'NoSalesmen:
        '
        '  ' Refresh the salesman list on frmSetup, if this is the active store.
        '  If GetStoreNumber(DBName) = StoresSld Then
        '    MsgBox "Alert: The user administration and password system has been updated." & vbCrLf & _
        '      "This has invalidated any existing passwords, and will temporarily allow" & vbCrLf & _
        '      "unrestricted access to many features.  Please adjust the Security Level" & vbCrLf & _
        '      "setting in Store Setup **for each computer**, and use the Sales Staff " & vbCrLf & _
        '      "editor to create a new Administrator password.  It is strongly recommended" & vbCrLf & _
        '      "that you create an Administrator account before changing this computer's" & vbCrLf & _
        '      "Security Level.", vbCritical, "Store " & GetStoreNumber(DBName) & " - WinCDS"
        '  End If
        '
        '  modStores.SecurityLevel = seclevNoPasswords  ' Security Level = DEMO
    End Function

    Private Sub ClearPatchList()
        Dim nPatchList() As PatchDef
        PatchList = nPatchList
    End Sub

    Private Sub AddPatchDef(ByVal vDate As Date, ByVal PatchName As String, ByVal PatchDesc As String, ByVal IsRequired As Boolean, ByVal IsRepeatable As Boolean, ByVal AffectsInventory As Boolean, ByVal AffectsStores As Boolean, Optional ByVal AllowPractice As Boolean = True, Optional ByVal ExpireAfter As Integer = 365)
        Dim X As Integer

        If ExpireAfter > 0 And DateAfter(Today, DayAdd(vDate, ExpireAfter)) Then
            Exit Sub
        End If

        On Error Resume Next
        X = UBound(PatchList)
        X = X + 1
        'ReDim Preserve PatchList(1 To X)
        ReDim Preserve PatchList(0 To X - 1)
        On Error GoTo 0

        PatchList(X - 1).PatchDate = vDate
        PatchList(X - 1).Name = PatchName
        PatchList(X - 1).Desc = PatchDesc
        PatchList(X - 1).Required = IsRequired
        PatchList(X - 1).Repeatable = IsRepeatable
        PatchList(X - 1).Inventory = AffectsInventory
        PatchList(X - 1).Stores = AffectsStores
        PatchList(X - 1).AllowPractice = AllowPractice

        If DateAfter(vDate, MostRecentPatchDate) Then MostRecentPatchDate = vDate
    End Sub

    Private Sub BuildPatchList()
        ClearPatchList()
        AddPatchDef(#1/1/2005#, "Clean Telephone Numbers", "Removes non-numeric characters from telephone numbers in the databases.", False, True, True, True)
        AddPatchDef(#1/1/2005#, "Rebuild Detail", "Reconstructs Detail table from sales history.", False, True, False, True)
        AddPatchDef(#1/1/2005#, "Correct Nontaxable Sales", "Adjusts non-taxable sales records to make them report as such.", False, True, False, True)
        AddPatchDef(#1/1/2005#, "Update Service Table Structure", "Updates the structure of the Service Order tables.", True, False, False, True)
        AddPatchDef(#1/1/2005#, "Decimal Quantity Database", "Updates the database structure, allowing decimal quantities to be used.", True, False, True, False)
        AddPatchDef(#1/1/2005#, "Department Numbers", "Allows three-digit department numbers.", True, False, True, True)
        AddPatchDef(#1/1/2005#, "Note Dates", "Corrects issues regarding note dates on sales and service orders.", True, False, False, True)
        AddPatchDef(#1/1/2005#, "Sales Tax Received", "Corrects sales tax figures for stores with multiple tax rates.", False, False, False, True)
        AddPatchDef(#1/1/2005#, "Non-Item Locations", "Addresses an issue regarding subtotals, tax, etc being included in Item reports.", True, True, False, True)
        AddPatchDef(#1/1/2005#, "Add Employees to Database", "Creates database structure for employee data.", True, False, False, True)
        AddPatchDef(#1/1/2005#, "Add User Groups to Database", "Creates database structure for user group data.", True, False, True, False)
        AddPatchDef(#1/1/2005#, "General Ledger Tracking Features", "Adds tracking fields to Cash and Audit tables for import into General Ledger reports.", True, True, False, True)
        AddPatchDef(#1/1/2005#, "General Ledger Tracking Features II", "Adds tracking fields to Holding and GrossMargin tables for import into General Ledger reports.", True, True, False, True)
        AddPatchDef(#1/1/2005#, "Freight Type", "Adds the column 'FreightType' to the 2Data table for determining % or $ amount.", True, True, True, False)
        AddPatchDef(#1/1/2005#, "ServicePartsOrder Table", "Creates the table [ServicePartsOrder] in all store databases for parts ordering.", True, True, False, True)
        AddPatchDef(#1/1/2005#, "Config Table", "Creates the table [Config] in Invent Database", True, True, True, False)
        AddPatchDef(#1/1/2005#, "ItemLocation Table", "Creates the table [ItemLocation] in Invent Database", True, True, True, False)
        AddPatchDef(#1/1/2005#, "PoSold Rectification", "Updates the PoSold value (Pre-sold) in the [2Data] table of the inventory database", False, True, True, False)
        AddPatchDef(#1/1/2005#, "Update ItemLocation", "Updates the [ItemLocation] table in the Invent DB", True, True, True, False)
        AddPatchDef(#10/10/2005#, "Sales Tax Received Rectification", "Updates the [Audit].[TaxRec1] field for delivered sales that have gone awry.", False, True, False, True)
        AddPatchDef(#10/11/2005#, "ArApp Note Field", "Adds a [Notes] field to table [ArApp] in stores", True, False, False, True)
        AddPatchDef(#12/20/2005#, "Cost Tracking System", "Makes database changes to handle cost tracking.  Creation and modification of tables.", False, True, True, False)
        AddPatchDef(#1/16/2006#, "InstallmentInfo.LastMetro426Status", "Add [LastMetro426Status] to [InstallmentInfo] table.", True, False, False, True)
        AddPatchDef(#1/24/2006#, "Cost in Detail", "Add a cost field to the Detail Table", True, False, True, False)
        AddPatchDef(#2/13/2006#, "Separate Comm Table", "Add [CommTable] Field to [Employees] for separate comm tables per salesman.", True, False, False, True)
        AddPatchDef(#3/9/2006#, "Decimal Kit Quantities", "Change Store1's GM DB to allow decimal quantities for kits.", True, False, False, True)
        AddPatchDef(#5/19/2006#, "Add Non-Taxable To Audit", "Adds a [NonTaxable] field to audit, mainly for the Sales Tax Report (written)", True, False, False, True)
        AddPatchDef(#5/25/2006#, "Re-number vendors", "Renumbers the vendors in [2Data] and all affected tables.", False, True, True, False)
        AddPatchDef(#6/2/2006#, "Extend Descriptions", "Allow Descriptions to be extended from 46 to 138", True, True, True, True)
        AddPatchDef(#7/10/2006#, "Commission Spiff To GM", "Add [Spiff] Field to [GrossMargin] and [GrossMarginTmp] for Commissions", True, False, False, True)
        AddPatchDef(#9/22/2006#, "Extend Sale Notes", "Add MailIndex to SaleNotes Table", True, False, False, True)
        AddPatchDef(#10/6/2006#, "Rectify Service Calls", "Fix Broken Service Calls", False, True, False, True)
        AddPatchDef(#10/6/2006#, "Finance Charge Sales Tax", "Fix Broken Service Calls", True, False, False, True)
        AddPatchDef(#2/5/2007#, "Add Sales Split", "Add Sales Split to Gross Margin", True, False, False, True)
        AddPatchDef(#3/1/2007#, "Correct Installment Rate", "Change InstallmentInfo.Rate a Double and Add APR", True, False, False, True)
        AddPatchDef(#3/1/2007#, "Fix Jerrys Connect", "Make the desktop icon point to the client instead of server program.", True, True, True, False)
        AddPatchDef(#4/2/2007#, "PoNotes Table", "Add Table [PoNotes] to Inventory.", True, True, True, False)
        AddPatchDef(#4/6/2007#, "Accomodate Time Stops", "Add fields to Gross Margin to allow for time stops.", True, True, False, True)
        AddPatchDef(#4/16/2007#, "Accomodate Time Stops II", "Add fields to Gross Margin to allow for time stops.", True, True, False, True)
        AddPatchDef(#4/24/2007#, "Accomodate Time Stops III", "Add fields to Service Tables to allow for time stops.", True, True, False, True)
        AddPatchDef(#4/27/2007#, "Extend Comments", "Extends the [Comments] Field in [2Data] from 50 to 138 characters.", True, False, True, True)
        AddPatchDef(#5/15/2007#, "Pictures Table", "Add Pictures Table", True, False, False, True)
        AddPatchDef(#8/23/2007#, "Add Spiff Field", "Add [Spiff] Field to [2Data]", True, False, True, False)
        AddPatchDef(#9/13/2007#, "Add Weekly Installments", "Add [Period] Field to [InstallmentInfo]", True, False, False, True)
        AddPatchDef(#10/14/2007#, "Add Cubes", "Add [Cubes] to [2Data]", True, False, True, False)
        AddPatchDef(#10/14/2007#, "Clear ItemCost", "Delete all rows from [ItemCost]", True, True, True, False)
        AddPatchDef(#11/30/2007#, "Fix Tax-Included Sales", "Fix sales that had no Quantity Value in TAX1 lines in the [GrossMargin] table due to the -Tax option.", True, True, False, True)
        AddPatchDef(#2/12/2008#, "Add Transfer Notes", "Add [Notes] to [Detail].", True, False, True, False)
        AddPatchDef(#4/22/2008#, "Life Type in Installment Info", "Add [LifeType] to [Installmentinfo].", True, False, False, True)
        AddPatchDef(#5/5/2008#, "Patch DDelDat in DELTW Sales.", "Put in a delivery date into DELTW sales.", True, False, False, True)
        AddPatchDef(#7/11/2008#, "Tibbees Service", "Fix Service Mail Index for Tibbees", True, True, False, True)
        AddPatchDef(#11/7/2008#, "InstallmentInfo.Satisfied", "Add a [Satisfied} field to [InstallmentInfo] table.", True, False, False, True)
        AddPatchDef(#3/24/2009#, "Turn Off Name AutoCorrect", "Turn off Name AutoCorrect Feature (Performance Issue)", True, False, True, True)
        AddPatchDef(#3/24/2009#, "Set SubDatasheet to None", "Set all sub datasheets to none (Performance Issue)", True, False, True, True)
        AddPatchDef(#4/4/2009#, "Fix Installment TotPaid", "Add-on Accounts were failing to set TotPaid, Life, Prop, and Acc.  Uses Detail to repair these records.", True, True, False, True)
        AddPatchDef(#4/24/2009#, "Fix Negative GMs", "Fix Random Negative GM on Non-commissioned Items", True, True, False, True)
        AddPatchDef(#5/27/2009#, "SIP Add", "Add fields to ServiceItemParts", True, True, False, True)
        AddPatchDef(#7/22/2009#, "IUI To InstallmentInfo", "Add field [IUI] to [InstallmentInfo]", True, False, False, True)
        AddPatchDef(#8/27/2009#, "GM Indexes", "Add some more indexes to GrossMargin table.", True, True, False, True)
        AddPatchDef(#8/28/2009#, "More Indexes", "Add some more indexes to the database in general.", True, True, True, True)
        AddPatchDef(#4/1/2010#, "Package Fields", "Add some more Package related fields to GM and Temp.", True, False, False, True)
        AddPatchDef(#5/5/2010#, "Update 2Data GM", "Recalculate all GMs in [2Data]", True, True, True, False)
        AddPatchDef(#7/27/2010#, "Add [Cashier] to [Audit]", "Adds [Cashier] field to [Audit] table.", True, False, False, True)
        AddPatchDef(#9/1/2010#, "Calculate Packages", "Populates Package Fields in GM Table.", True, True, False, True)
        '  AddPatchDef( #11/18/2010#, "AP Check Name Field", "Add [CheckName] to APDB.[tblChecks]", True, True, True, False
        AddPatchDef(#3/9/2011#, "TransID to GM", "Adds [TransID] to GM and GMtmp", True, False, False, True)
        AddPatchDef(#5/1/2011#, "CreateOnlineOrderRecordTable", "Create OnlineOrderRecord Table", True, False, True, False)
        AddPatchDef(#7/7/2011#, "StoreCount32", "Go from 16 to 32 stores (DB).", True, False, True, False)
        AddPatchDef(#9/9/2011#, "KitSKU", "Add [KitSKU] to [InvKit] (DB1 only)", True, False, False, True)
        AddPatchDef(#12/4/2011#, "InstallmentInfo Indexes", "Add Indexes to InstallmentInfo", True, False, False, True)
        AddPatchDef(#1/12/2012#, "Adjustment TAX2 Loc", "Fix Lost TAX2 Location", True, False, False, True)
        AddPatchDef(#5/1/2012#, "Patch Sales Notes Taxable+++", "Make Sales Notes from 1/26 to [...] Taxable [+reset, run again]", True, True, False, True)
        AddPatchDef(#4/1/2012#, "Sale Mail Index", "Make sure MailIndex is carried down correctly.", True, True, False, True)
        AddPatchDef(#6/6/2012#, "Add Distributors", "Add Table Field [Distributors] to Invent", True, True, True, False)
        AddPatchDef(#1/30/2013#, "Add Perm Order Status", "Add Permission 'Order Status' = 37", True, True, True, False)
        AddPatchDef(#11/23/2013#, "Add Telephone Labels", "Add Telephone Labels", True, False, False, True)
        AddPatchDef(#2/23/2014#, "Add ArNo to Holding", "Adds ArNo to Holding table", True, True, False, True)
        AddPatchDef(#2/23/2014#, "BFMyer Commissions", "Patch BF Myers Commissions Error", False, True, False, True)
        AddPatchDef(#5/24/2014#, "Fix Short Vendor Numbers", "Repair damaged vendor numbers in GM", True, True, False, True, True)
        AddPatchDef(#6/23/2014#, "Fix Missing Vendor Numbers on Returns", "Repair missing vendor numbers in GM", True, True, False, True, True)
        AddPatchDef(#9/18/2014#, "Store Setup to INI", "Convert Store Setup files to INI Format", True, True, False, True)
        AddPatchDef(#1/27/2015#, "Revolving Update Delivery Day", "Convert Store Setup files to INI Format", False, True, False, True)
        AddPatchDef(#4/17/2015#, "Activate UseScheduledTask", "Set UseScheduledTask=1", True, True, True, False)
        AddPatchDef(#4/30/2015#, "Copy Ashley Credentials", "Move Ashley ATP Credentials to INI File", True, True, True, False)
        AddPatchDef(#9/24/2015#, "Fix Weekly Monthly Installment", "Update [InstallmentInfo] where Period is Null.", False, True, False, True)
        AddPatchDef(#10/31/2015#, "Delivery Ticket Message File", "DeliveryTicketMessageFile", False, True, False, True)
        AddPatchDef(#2/19/2016#, "Fix DISCOUNT Sales+", "Fix Sales Created With Discount Showing Non-Taxable", True, True, False, True)
        AddPatchDef(#6/6/2016#, "PatchConnectCmdv2", "Upgrade Connect.cmd to allow full admin control with UAC workaround", True, True, True, False)
        AddPatchDef(#8/4/2016#, "FXFolder", "Separate PX and FX.", True, True, True, False)
        AddPatchDef(#8/4/2016#, "FXFolder2", "Separate PX and FX - Patch #2.", True, True, True, False)
        AddPatchDef(#9/7/2016#, "Michaels Last Notice", "For Michael's only.  Add [LastNotice] to [InstallmentInfo].", True, True, False, True)
        AddPatchDef(#1/7/2017#, "LastLateCharge Added", "For Michael's.  Add [LastLateCharge] to [InstallmentInfo].", True, True, False, True)
        AddPatchDef(#2/13/2017#, "McClure Old Account Fix", "McClure Only - Fix Already Created Old Installment Accounts", True, True, False, True)
        AddPatchDef(#2/16/2017#, "LastLateCharge Fix", "Update LastLateCharge, fix value", True, True, False, True)
        AddPatchDef(#4/3/2017#, "Update KitSKU with Vendor Name", "KITSKU2", True, True, False, True, False)
        AddPatchDef(#4/9/2017#, "AddTransactionsfldPosted***", "Add [Transactions].[fldPosted]", True, False, False, True)
        AddPatchDef(#5/22/2017#, "AddTerminalTracking", "Add *.[Terminal] Wherever *.[Cashier] exists.", True, False, False, True)
        AddPatchDef(#5/22/2017#, "ExportInstallmentAccounts", "ExportInstallmentAccounts (JOHNSONS)", False, True, True, False, True)
        AddPatchDef(#5/30/2017#, "MarkPastInstallments", "Add Marker to each existing Installment Account.", False, True, False, True, True)
        AddPatchDef(#6/12/2017#, "Payment History Profile Keeping", "Add field to [InstallmentInfo] to keep last Payment History Profile for Add Ons.", True, True, False, True, True)
        AddPatchDef(#9/13/2017#, "Date Past Installments", "Fix date on Marker to each existing Installment Account (with AutoPatch).", False, True, False, True, True)
        AddPatchDef(#1/1/2018#, "ExtendStyle", "Extend Style Length", False, True, True, True, True)
        AddPatchDef(#2/15/2018#, "Fix Warehouse Furniture 2.15.2018", "Store Specific Patch", True, False, False, True, False, 15)
        AddPatchDef(#2/27/2018#, "ClassicInteriorsAPVendors", "Patch Classic Interiors AP Vendors", True, False, True, False, False, 15)
        AddPatchDef(#2/27/2018#, "HomecraftersRebuildSearch", "Rebuild Search for Homecrafters", True, False, True, False, False, 15)
        AddPatchDef(#3/10/2018#, "Fix Homecrafters II", "Fix Homecrafter's Search Index", True, True, True, False, True, 7)
        AddPatchDef(#3/15/2018#, "Fix Lindsay 3336B", "Void account #3336B for Lindsays", True, True, False, True, True, 7)
        AddPatchDef(#3/29/2018#, "NewAgeChicagoFurniture20180329", "NewAgeChicagoFurniture20180329", True, False, False, True, False, 7)
        AddPatchDef(#5/21/2018#, "UnloadGilsLoc20Bal20180521", "UnloadGilsLoc20Bal20180521", True, False, True, False, False)
        AddPatchDef(#9/15/2018#, "CasaBellaRemoveAddOns20180915", "CasaBellaRemoveAddOns20180915", True, True, False, True, False, 15)
        AddPatchDef(#10/11/2018#, "BudgetRemoveDuplicateVoid20180911-b", "BudgetRemoveDuplicateVoid20180911-b", True, True, False, True, False, 30)
        AddPatchDef(#10/22/2018#, "HomecraftersFixItem-20181022", "HomecraftersFixItem-20181022", True, True, True, False, False, 10)
        AddPatchDef(#12/19/2018#, "CleanConfig", "Clean Config Table", True, True, True, False, True)
        AddPatchDef(#2/1/2019#, "Michaels20190201", "Michaels20190201", True, False, False, True, False, 7)
        AddPatchDef(#2/18/2019#, "Budget20190212-2", "Budget20190212-2", True, False, False, True, False, 7)
        AddPatchDef(#2/21/2019#, "JAllen20190221", "JAllen20190221-2", True, False, False, True, False, 7)
        AddPatchDef(#5/14/2019#, "Budget20190514", "Budget20190514", True, False, False, True, False, 7)

        ' For Adding Packages...
        ' after you add it here, you need to add it to the ApplyPatch Select Case below using the exact short desc used above
        ' Finally, make sure you make the function
    End Sub

End Module
