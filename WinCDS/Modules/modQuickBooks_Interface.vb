Imports QBFC5Lib
Module modQuickBooks_Interface
    Private mUseQB As Boolean
    Private mQBTested As Boolean
    Private mQBMfgLst As Object

    Public Function UseQB(Optional ByRef FailReason As String = "", Optional ByVal NotifyWantedButNotSetup As Boolean = False, Optional ByVal Loc As Long = 0) As Boolean
        If Not QBConnect(FailReason, Loc) Then Exit Function

        If Not QBTested() Then
            mUseQB = IsQBSetup(FailReason)
            If mUseQB Then mQBTested = True
        End If
        UseQB = mUseQB

        If QBWanted() And NotifyWantedButNotSetup And Not UseQB Then
            MessageBox.Show("Quickbooks is selected in store setup, but is not fully configured." & vbCrLf & "Please enter the Quickbooks Interface Panel to inspect and complete your configuration." & vbCrLf2 & FailReason, "QuickBooks Is Selected, But Is Not Fully Configured", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Function

    Public Function QBGetVendorName(ByVal POName As String, ByRef completeName As String, Optional ByRef Address As String = "", Optional ByRef Address2 As String = "", Optional ByRef Address3 As String = "", Optional ByRef Zip As String = "", Optional ByRef Phone As String = "", Optional ByRef Fax As String = "", Optional ByRef CompleteCode As String = "", Optional ByRef EmailAddress As String = "") As Boolean

        Dim E As Long, S As String
        Dim X As Object 'As QBFC13Lib.IVendorRet
        Dim M1 As Boolean, M2 As Boolean, M3 As Boolean

        POName = Trim(POName)
        If POName = "" Then Exit Function

        On Error Resume Next
        Dim T As Object, L As Object, SS As String
        Dim K As String
        K = POName

        T = QBMfgLst()

        For Each L In T
            SS = ""
            SS = L.Name.GetValue
            If SS <> "" Then
                If UCase(Left(SS, Len(K))) = UCase(K) Or UCase(SS) = UCase(Left(K, Len(SS))) Then
                    K = SS
                    GoTo FoundCloseMatch
                End If
            End If
        Next

NoCloseMatch:
        For Each L In T
            SS = ""
            SS = L.Name.GetValue
            If SS <> "" Then
                ' We have three close match functions we could try:  Metaphone, SoundEx, and FuzzyStringMatch()..
                M1 = SoundEx(SS) = SoundEx(K)
                M2 = Metaphone(SS) = Metaphone(K)
                M3 = FuzzyStringMatch(SS, K)
                ' For now, we'll just try the last 2...
                If M2 Or M3 Then
                    K = SS
                    GoTo FoundCloseEnoughMatch
                End If
            End If
        Next

        Exit Function

FoundCloseMatch:
        ' We could differentiate here if we wanted..
FoundCloseEnoughMatch:
        X = QB_VendorQuery_Vendor(K, E, S)
        If E = 0 Then
            completeName = IfNNGetValue(X.Name)
            Address = IfNNGetValue(X.VendorAddress.addr1)
            Address2 = IfNNGetValue(X.VendorAddress.addr2)
            Address3 = Trim(IfNNGetValue(X.VendorAddress.City) & " " & IfNNGetValue(X.VendorAddress.State))
            Zip = IfNNGetValue(X.VendorAddress.PostalCode)
            Phone = IfNNGetValue(X.Phone)
            Fax = IfNNGetValue(X.Fax)
            CompleteCode = IfNNGetValue(X.CompanyName)
            EmailAddress = IfNNGetValue(X.Email)

            '      If QBFCVersion >= 12 Then
            '        If EmailAddress = "" Then
            '          Dim IX As Long
            '          For IX = 0 To X.AdditionalContactRefList.Count - 1
            '            If ValidEmailAddress(X.AdditionalContactRefList.GetAt(IX).FullName.GetValue) Then EmailAddress = X.AdditionalContactRefList.GetAt(IX).FullName.GetValue
            '          Next
            '        End If
            '      End If

            QBGetVendorName = True
        End If
        X = Nothing
    End Function

    Public Function QBConnect(Optional ByRef FailReason As String = "", Optional ByVal Loc As Long = 0) As Boolean
        If Not QBWanted() Then FailReason = "QB not selected in options page in Store Setup for Loc " & StoresSld & "." : Exit Function

        If QBUseRDS() Then

        Else
            If QB_File(Loc) = "" Then FailReason = "No quickbooks file is specified for this location." : Exit Function
            If Not FileExists(QB_File(Loc)) Then FailReason = "Quickbooks file does not exist." : Exit Function
        End If
        If Not QBObjectsExist(FailReason) Then Exit Function
        QBConnect = True
    End Function

    Public Function QBTested() As Boolean
        QBTested = mQBTested
    End Function

    Public Function IsQBSetup(Optional ByRef FailReason As String) As Boolean
        Dim RET As Long, RetMsg As String
        Dim I As Long, T As String, lst() As Object

        ' these 2 checks now in QBConnect, b/c we cant even connect w/o them
        '  If GetQBSetupValue("file") = "" Then
        '    FailReason = "No quickbooks file is specified."
        '    Exit Function
        '  End If
        '
        '  If Not QBObjectsExist(FailReason) Then Exit Function

        If Not QBCheckPreferences(False, FailReason) Then Exit Function

        FailReason = "Location Classes Not Created"
        If Not QB_ClassExists(QBLocationClassID(0)) Then Exit Function
        For I = 1 To LicensedNoOfStores()
            If Not QB_ClassExists(QBLocationClassID(I, True)) Then Exit Function
        Next

        FailReason = "Customer Deposits Account Must Exist.  Click 'Make Special Custs' from the QB Setup page."
        If Not QB_CustomerExistsByName(QBCustomerDepositsName) Then Exit Function

        FailReason = "All G/L accounts must be mapped to existing QB accounts."
        Dim AList() As GLAccount, N As Long
        AList = GLAccountList(N)

        On Error GoTo FailOnAccountMaps
        lst = QB_AccountQuery_All()

        For I = LBound(lst) To UBound(lst)
            lst(I) = lst(I).Name.GetValue
        Next

        For I = 0 To N
            T = QueryGLQBAccountMap(AList(I).Account)
            If T = "" Or Not IsInArray(T, lst) Then
                FailReason = "All G/L accounts must be mapped to existing QB accounts: " & AList(I).Account & ", " & T
                Exit Function
            End If
        Next

        FailReason = ""
        IsQBSetup = True
        Exit Function
FailOnAccountMaps:
        FailReason = FailReason & " (Err=" & Err.Description & ")"
        Resume Next
    End Function

    Public Function QBWanted() As Boolean
        QBWanted = (GetQBSetupValue("useqb") = "True")
    End Function

    Public Function QBMfgLst(Optional ByVal Reset As Boolean = False) As Object
        If Reset Then mQBMfgLst = Nothing
        If IsNothing(mQBMfgLst) Then
            mQBMfgLst = QB_VendorQuery_All()
        End If
        QBMfgLst = mQBMfgLst
    End Function

    Public Function QBCheckPreferences(Optional ByVal Notify As Boolean = True, Optional ByRef FailReason As String = "") As Boolean
        Dim IP As IPreferencesRet, M As String, RET As Long
        On Error GoTo NoQB


        QBCheckPreferences = True
        Exit Function    ' for right now, these are no preferences we really check

        '  Set IP = QB5_PreferencesQuery(ret, M)
        '  If ret < 0 Then GoTo NoQB


        '  If Not QBPrefClassTracking(IP) Then
        '    QBCheckPreferences = False
        '    M = M & IIf(M = "", "", vbCrLf) & "You do not have Class Tracking turned on in QuickBooks."
        '  End If

        '  If Not QBPrefAccountNumbers(IP) Then
        '    QBCheckPreferences = False
        '    M = M & IIf(M = "", "", vbCrLf) & "You do not have Account Numbers turned on in QuickBooks."
        '  End If
        '  If Not QBPrefInventory(IP) Then
        '    QBCheckPreferences = False
        '    M = M & IIf(M = "", "", vbCrLf) & "You do not have Inventory Management turned on in QuickBooks."
        '  End If
        If M <> "Status OK" Then FailReason = M
        If Notify And M <> "" And M <> "Status OK" Then MessageBox.Show("Your QuickBooks company must have the correct preferences set." & vbCrLf & "Please see below to see which preferences you must set " & vbCrLf & "in order to use the WinCDS QuickBooks interface:" & vbCrLf2 & M, "QuickBooks Incorrectly Configured", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        IP = Nothing
        Exit Function
NoQB:
        FailReason = "Could not connect to QuickBooks: " & M
        If Notify Then MessageBox.Show(FailReason, "Start QB", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Function

    Public Function QBLocationClassID(ByVal Loc As Long, Optional ByVal WithParentRef As Boolean = False) As String
        If Loc > 0 And WithParentRef Then
            QBLocationClassID = QBLocationClassID(0) & ":"
        Else
            QBLocationClassID = ""
        End If
        QBLocationClassID = QBLocationClassID & IIf(Loc > 0, "Loc " & Loc, "Locations")
    End Function

End Module
