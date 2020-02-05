Module modPractice
    Public Function DoCompactAndRepairDAO(Optional ByVal AllDbs As Boolean = False) As Boolean
        Dim X() As Object, I As Integer, N As Integer, dB As String
        If AllDbs Then
            N = 2
        Else
            ArrAdd(X, "Enter Pathname")
            ArrAdd(X, "All WinCDS Databases")
            ArrAdd(X, "Inventory")
            For I = 1 To LicensedNoOfStores()
                ArrAdd(X, "Store #" & I)
            Next
            N = SelectOptionArray("Select Database", frmSelectOption.ESelOpts.SelOpt_List, X, "Compact / Repair")
        End If
        Debug.Print(N)


        If N <= 0 Then Exit Function
        If N = 1 Then
            dB = InputBox("Path:", "Enter Database Name", InventFolder)
            If dB = "" Then Exit Function
            If Not FileExists(dB) Then
                MessageBox.Show("File does not exist:" & vbCrLf & dB)
                Exit Function
            End If
        ElseIf N = 2 Then
            For I = 1 To NoOfActiveLocations
                CompactRepairAccessDB(GetDatabaseAtLocation(I))
            Next
            dB = GetDatabaseInventory()
        Else
            dB = IIf(N = 3, GetDatabaseInventory, GetDatabaseAtLocation(N - 3))
        End If
        '  If MsgBox("Compact and repair the following database?" & vbCrLf & DB, vbQuestion + vbOKCancel, "Confirm Compact and Repair") = vbCancel Then
        '    Exit Function
        '  End If

        CompactRepairAccessDB(dB)

        MessageBox.Show("Complete.", "Compact And Repair")
    End Function

End Module
