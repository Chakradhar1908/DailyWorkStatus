Imports QBFC5Lib

Module modQuickBooks_Constants
    Public Const QBCustomerDepositsName As String = "Customer Deposits"
    Public Structure GLAccount
        Dim Account As String
        Dim Desc As String
        Dim eType As ENAccountType
    End Structure

    Public Function QBUseRDS() As Boolean
        QBUseRDS = Val(GetQBSetupValue("qbrds")) <> 0
    End Function

    Public Function GLAccountList(ByRef Count As Long) As GLAccount()
        Dim L() As GLAccount, N As Long, C As Long
        C = -1

        '  C = C + 1: ReDim Preserve L(C): L(C) = CreateGLAccountDef("00001", "WinCDS Overflow", atOtherCurrentAsset)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("01200", "Accounts Receivable", atAccountsReceivable)

        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("10001", "Accounts Payable", atAccountsPayable)

        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("10100", "Petty Cash", atOtherCurrentAsset)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("10200", "Checking Account", atBank)

        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("11000", "Ar Cash Sales", atFixedAsset)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("11100", "Less Und Sales", atFixedAsset)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("11200", "Back Orders", atFixedAsset)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("11300", "A/R Principal Pay", atOtherAsset)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("11500", "Inventory Asset", atFixedAsset)

        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("21400", "Customer Dep", atLongTermLiability)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("21500", "Exchange-Refunds", atLongTermLiability)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("21600", "State Tax Payable", atLongTermLiability)

        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("40200", "Written Sales", atIncome)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("40500", "A/R Late Charges", atIncome)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("40600", "Und Sales", atIncome)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("41500", "Forfeit Deposits", atIncome)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("41700", "Sales Tax Rec", atIncome)

        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("50100", "Cost of Goods Sold", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("50200", "Purchases COD", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("50500", "Freight", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("50600", "Discount/Finan", atExpense)

        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("52000", "Cash Over/Short", atExpense)

        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("60100", "Gas & Oil", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("60500", "Disc Visa etc.", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("61600", "Medical Co-Pay", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("62300", "Maintenance", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("62400", "Repairs", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("63500", "Whse Supply", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("64100", "Office Expenses", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("64200", "Misc Expense", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("65200", "Casual Labor", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("67500", "Meals & Entertain", atExpense)
        C = C + 1 : ReDim Preserve L(C) : L(C) = CreateGLAccountDef("69600", "Other Income", atOtherIncome)



        '  C = C + 1: ReDim Preserve L(C): L(C) = CreateGLAccountDef("5100", "Purchases", atCostOfGoodsSold)


        '  C = C + 1: ReDim Preserve L(C): L(C) = CreateGLAccountDef("")
        Count = C
        GLAccountList = L
    End Function

    Public Function QueryGLQBAccountMap(ByRef GLAccount As String) As String
        Dim L As Variant, X As Variant
        Dim M As String

        M = QBAccountMap
        For Each L In Split(M, vbCrLf)  ' list is not sorted
            X = Split(L, ":")
            If UBound(X) = 1 Then
                If Trim(X(0)) = GLAccount Then
                    QueryGLQBAccountMap = Trim(X(1))
                    Exit Function
                End If
            End If
        Next
    End Function

    Public Function GetQBSetupValue(ByVal Field As String, Optional ByVal Store As Long = 0) As String
        If Store = 0 Then Store = StoresSld
        Select Case LCase(Field)
            Case "useqb"
                GetQBSetupValue = TrueFalseString(StoreSettings(Store).bUseQB)
            Case "posttoloc1"
                GetQBSetupValue = TrueFalseString(StoreSettings(Store).bPostToLoc1)
            Case "file"
                GetQBSetupValue = GetConfigTableValue(IIf(QBPostAs(Store) = 1, "QB_FILE", "QB_FILE_" & Store))
                If Not IsServer() Then
                    GetQBSetupValue = Replace(GetQBSetupValue, LocalRoot, RemoteRoot, , , vbTextCompare)
                Else
                    GetQBSetupValue = Replace(GetQBSetupValue, RemoteRoot, LocalRoot, , , vbTextCompare)
                End If
            Case "mapfile"
                GetQBSetupValue = InventFolder() & "QBAccountMap" & IIf(QBPostAs(Store) = 1, "", Store) & ".txt"
            Case "xmlmajor" : GetQBSetupValue = GetConfigTableValue("QB_XML_MAJOR", "3") ' "2")
            Case "xmlminor" : GetQBSetupValue = GetConfigTableValue("QB_XML_MINOR", "0") ' "1")
            Case "qbfcver" : GetQBSetupValue = Val(GetConfigTableValue("QB_QBFC_VERSION", "5"))
            Case "qbrds" : GetQBSetupValue = GetConfigTableValue("QB_USE_RDS", "0")
            Case Else : Err.Raise -1, , "Invalid field: " & Field
    Exit Function
        End Select
    End Function

    Public Function SetQBSetupValue(ByVal Field As String, ByVal Value As String, Optional ByVal Store As Long = 0) As Boolean
        If Store = 0 Then Store = StoresSld
        Select Case LCase(Field)
            Case "useqb"
                Err.Raise -1, , "Cannot set useqb from here!"
    Case "posttoloc1"
                Err.Raise -1, , "Cannot set posttoloc1 from here!"
'      Get StoreInformation(Store).bPostToLoc1
'      frmSetup .chkPostToLoc1 = IIf(CBool(Value), 1, 0)
            Case "file"
                SetQBSetupValue = SetConfigTableValue(IIf(QBPostAs(Store) = 1, "QB_FILE", "QB_FILE_" & Store), Value)
            Case "xmlmajor" : SetQBSetupValue = SetConfigTableValue("QB_XML_MAJOR", Value)
            Case "xmlminor" : SetQBSetupValue = SetConfigTableValue("QB_XML_MINOR", Value)
            Case "qbfcver" : SetQBSetupValue = SetConfigTableValue("QB_QBFC_VERSION", Value)
            Case "qbrds" : SetQBSetupValue = SetConfigTableValue("QB_USE_RDS", Value)
            Case Else : Exit Function
        End Select
    End Function

End Module
