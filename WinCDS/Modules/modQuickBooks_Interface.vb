Module modQuickBooks_Interface
    Public Function UseQB(Optional ByRef FailReason As String = "", Optional ByVal NotifyWantedButNotSetup As Boolean = False, Optional ByVal Loc As Long = 0) As Boolean
        If Not QBConnect(FailReason, Loc) Then Exit Function

        If Not QBTested Then
            mUseQB = IsQBSetup(FailReason)
            If mUseQB Then mQBTested = True
        End If
        UseQB = mUseQB

        If QBWanted And NotifyWantedButNotSetup And Not UseQB Then
            MsgBox "Quickbooks is selected in store setup, but is not fully configured." & vbCrLf & "Please enter the Quickbooks Interface Panel to inspect and complete your configuration." & vbCrLf2 & FailReason, vbExclamation, "QuickBooks Is Selected, But Is Not Fully Configured"
  End If
    End Function

    Public Function QBGetVendorName(ByVal POName As String, ByRef completeName As String, Optional ByRef Address As String = "", Optional ByRef Address2 As String = "", Optional ByRef Address3 As String = "", Optional ByRef Zip As String = "", Optional ByRef Phone As String = "", Optional ByRef Fax As String = "", Optional ByRef CompleteCode As String = "", Optional ByRef EmailAddress As String = "") As Boolean

        Dim E As Long, S As String
        Dim X As Object 'As QBFC13Lib.IVendorRet
        Dim M1 As Boolean, M2 As Boolean, M3 As Boolean

        POName = Trim(POName)
        If POName = "" Then Exit Function

        On Error Resume Next
        Dim T As Variant, L As Variant, SS As String
        Dim K As String
        K = POName

        T = QBMfgLst
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
  Set X = QB_VendorQuery_Vendor(K, E, S)
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
  Set X = Nothing
End Function

End Module
