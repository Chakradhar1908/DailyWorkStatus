Imports QBFC5Lib
Imports QBFC10Lib
Module modQuickBooks_CommQBFC13
    Private mQBSM13 As QBFC13Lib.QBSessionManager
    Private mMsReq13 As QBFC13Lib.IMsgSetRequest
    Private mMsRsp13 As QBFC13Lib.IMsgSetResponse
    Private mQBActiveStore As Integer, mQBAS_JustSet As Boolean
    Public Function QB13_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC13Lib.IVendorRet
        Dim VQ As QBFC13Lib.IVendorQuery

        On Error Resume Next
        VQ = MSREQ13.AppendVendorQueryRq

        VQ.ORVendorListQuery.FullNameList.Add(Vendor)
        VQ = Nothing

        QB13_VendorQuery_Vendor = QB13_SendRequestsGetRet(RET, RetMsg)
        Exit Function
    End Function

    Public Function QB13_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As Object
        Dim Q As QBFC13Lib.IClassQuery
        QBSM13_Reset()

        On Error Resume Next
        Q = MSREQ13.AppendClassQueryRq
        Q.ORListQuery.FullNameList.Add(ClassName)
        Q = Nothing

        QB13_ClassQuery_Class = Nothing
        QB13_ClassQuery_Class = QB13_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB13_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC13Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC13Lib.ICustomerQuery
        QBSM13_Reset()
        Q = MSREQ13.AppendCustomerQueryRq
        Q.ORCustomerListQuery.FullNameList.Add(Name)
        '    Q.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameFilter.Name.SetValue Name
        Q = Nothing

        QB13_CustomerQuery_Name = QB13_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB13_AccountQuery_All(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As Object
        On Error Resume Next
        QBSM13_Reset()
        MSREQ13.AppendAccountQueryRq() ' nothing to set
        QB13_AccountQuery_All = QB13_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB13_VendorQuery_All(Optional ByRef RET As Integer = 0) As Object
        Dim ES As String
        On Error Resume Next
        QBSM13_Reset()
        MSREQ13.AppendVendorQueryRq() ' nothing to set
        QB13_VendorQuery_All = QB13_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB13ObjectsExist(Optional ByRef Msg As String = "") As Boolean
        Dim X As Object, Y As Object
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC13)"
        X = New QBFC13Lib.QBSessionManager
        '  Set X = CreateObject("QBFC13.QBSessionManager")
        Msg = "Could not create MsgSetRequest (QBFC13)"
        Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
        Msg = ""
        QB13ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
        Y = Nothing
        X = Nothing
    End Function

    Public Function QB13_SendRequests(Optional ByRef ErrString As String = "", Optional ByRef ErrNo As Integer = 0, Optional ByRef OnErr As QBFC10Lib.ENRqOnError = QBFC10Lib.ENRqOnError.roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB13Startup(ErrNo, ErrString) Then Exit Function
        MSREQ13.Attributes.OnError = OnErr
        MSRSP13 = QBSM13.DoRequests(MSREQ13)
        '  Debug.Print mmsrsp13.ToXMLString

        QB13_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Public Function QB13Shutdown() As Object
        ActiveLog("QB::QBShutdown", 1)
        If QBSessOpen Then QBSM13.EndSession() : QBSessOpen = False
        If QBConnOpen Then QBSM13.CloseConnection() : QBConnOpen = False
        mQBSM13 = Nothing
    End Function

    Public ReadOnly Property MSREQ13() As QBFC13Lib.IMsgSetRequest
        Get
            If mMsReq13 Is Nothing Then
                mMsReq13 = QBSM13.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
            End If
            MSREQ13 = mMsReq13
        End Get
    End Property

    Public Function QB13_SendRequestsGetRet(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As Object
        Dim RL As Object, En As Integer, ES As String
        If QB13_SendRequests(ES, En) Then
            RET = StatusCode
            RetMsg = StatusMsg
            If RET = 0 Then
                If Not Resp.Detail Is Nothing Then
                    On Error Resume Next
                    QB13_SendRequestsGetRet = Resp.Detail
                    RL = Resp.Detail
                    QB13_SendRequestsGetRet = RL.GetAt(0)
                    RL = Nothing
                End If
            End If
        Else
            RetMsg = ES
            RET = -1
            Select Case En
                Case -2147220458
                    RET = -2
                Case -2147220445
                    MessageBox.Show("(" & Hex(En) & ") " & ES)
                Case Else
                    MessageBox.Show("(" & Hex(En) & ") " & ES)
            End Select
        End If
        QBSM13_Reset()
    End Function

    Public Sub QBSM13_Reset()
        '  Set mQBSM = Nothing
        ActiveLog("QB::QBSM13_Reset", 3)
        If mQBAS_JustSet Or POMode("EDIT", "REC") Then
            mQBAS_JustSet = False
        Else
            QBActiveStore = 0
        End If

        MsReq13_Reset()
    End Sub

    Public Sub MsReq13_Reset()
        mMsReq13 = Nothing
        mMsRsp13 = Nothing
    End Sub

    Public Function QB13_SendRequestsGetRetList(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As Object
        Dim RL As Object, T() As Object, N As Integer, En As Integer, ES As String
        If QB_SendRequests(ES, En) Then
            RET = StatusCode
            RetMsg = StatusMsg
            If RET = 0 Then
                If (Not Resp.Detail Is Nothing) Then
                    RL = Resp.Detail
                    ReDim T(RL.Count - 1)
                    For N = 0 To RL.Count - 1
                        T(N) = RL.GetAt(N)
                    Next
                End If
            End If
        Else
            RetMsg = ES
            If En = -2147220458 Then
                RET = -2
            Else
                RET = -1
            End If
        End If
        QB13_SendRequestsGetRetList = T
        QBSM13_Reset()
    End Function

    Public Function QB13Startup(Optional ByRef RET As Integer = 0, Optional ByRef Msg As String = "") As Boolean
        Dim P As frmProgress
        Dim E As String
        On Error GoTo NoComm
        ActiveLog("QB::QBStartup", 1)
        If Not QBConnOpen Then
            P = New frmProgress
            P.Progress(0, , "Opening Connection to Quickbooks...", True, False)
            If QBUseRDS() Then
                QBSM13.OpenConnection2(QB_AppID, QB_AppNm, QBFC13Lib.ENConnectionType.ctRemoteQBD)
            Else
                QBSM13.OpenConnection2(QB_AppID, QB_AppNm, QBFC13Lib.ENConnectionType.ctLocalQBD)
            End If
            QBConnOpen = True
        End If
        If Not QBSessOpen Then
            If QBUseRDS() Then
                '      QBSM13.BeginSession "", omDontCare 'BFH20150325 - Added the file.. If you don't want it, leave it blank in the form
                QBSM13.BeginSession(QB_File, QBFC13Lib.ENOpenMode.omDontCare)
            Else
                QBSM13.BeginSession(QB_File, QBFC13Lib.ENOpenMode.omDontCare)
            End If
            QBSessOpen = True
        End If
        QB13Startup = True
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
        Exit Function
NoComm:
        RET = Err.Number
        Msg = Err.Description
        Err.Clear()
        QB13Shutdown()
        QB13Startup = False
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
    End Function

    Private Property MSRSP13() As QBFC13Lib.IMsgSetResponse
        Get
            MSRSP13 = mMsRsp13
        End Get
        Set(value As QBFC13Lib.IMsgSetResponse)
            mMsRsp13 = value
        End Set
    End Property

    Public ReadOnly Property QBSM13() As QBFC13Lib.QBSessionManager
        Get
            If mQBSM13 Is Nothing Then
                mQBSM13 = New QBFC13Lib.QBSessionManager
            End If
            QBSM13 = mQBSM13
        End Get
    End Property

    Private ReadOnly Property StatusCode() As Integer
        Get
            On Error Resume Next
            StatusCode = Resp.StatusCode
        End Get
    End Property

    Private ReadOnly Property StatusMsg(Optional LDesc As String = "") As String
        Get
            On Error Resume Next
            StatusMsg = Resp.StatusMessage
        End Get
    End Property

    Private ReadOnly Property Resp() As QBFC13Lib.IResponse
        Get
            On Error Resume Next
            Resp = Response(0)
        End Get
    End Property

    Private ReadOnly Property Response(Index As Integer) As QBFC13Lib.IResponse
        Get
            On Error Resume Next
            Response = ResponseList13.GetAt(Index)
        End Get
    End Property

    Public ReadOnly Property ResponseList13() As QBFC13Lib.IResponseList
        Get
            If Not MSRSP13 Is Nothing Then
                ResponseList13 = MSRSP13.ResponseList
            End If
        End Get
    End Property

    Public Function QB13CreateDeposit(
    Optional ByVal TxnDate As String = "",
    Optional ByVal AcctRef_ListID As String = "", Optional ByVal AcctRef_FullName As String = "",
    Optional ByVal FromAccount_ListID As String = "", Optional ByVal FromAccount_FullName As String = "",
    Optional ByVal Memo As String = "",
    Optional ByVal CheckNumber As String = "",
    Optional ByVal PayMethRef_ListID As String = "", Optional ByVal PayMethRef_FullName As String = "",
    Optional ByVal Amount As String = "",
    Optional ByVal EntityRef_ListID As String = "", Optional ByVal EntityRef_FullName As String = "",
    Optional ByVal ClassRef_ListID As String = "", Optional ByVal ClassRef_FullName As String = "") As Boolean
        Dim R As Integer, M As String
        QBSM13_Reset()
        QB13_AppendDespositAdd _
      (TxnDate, AcctRef_ListID, AcctRef_FullName, FromAccount_ListID, FromAccount_FullName,
      Memo, CheckNumber, PayMethRef_ListID, PayMethRef_FullName,
      Amount, EntityRef_ListID, EntityRef_FullName, ClassRef_ListID, ClassRef_FullName)

        QB13CreateDeposit = (QB13_SendRequestsSingle(R, M) <> 0) And R = 0
        QBSM13_Reset()
    End Function

    Public Function QB13_SendRequestsSingle(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "", Optional ByVal Notify As Boolean = True) As Integer
        Dim RR As QBFC13Lib.IResponse, II As Integer

        If QB13_SendRequests(RetMsg, RET, QBFC10Lib.ENRqOnError.roeContinue) Then
            For II = 0 To ResponseCount - 1
                RR = ResponseList13.GetAt(II)
                If RR.StatusCode <> 0 Then
                    RetMsg = RetMsg & IIf(Len(RetMsg) > 0, vbCrLf, "")
                    RetMsg = RetMsg & "Error processing " & RR.Type.GetAsString & ": " & RR.StatusMessage
                End If
            Next
            If Len(RetMsg) > 0 Then
                If Notify Then
                    If MsgBox("Error(s): " & vbCrLf & RetMsg, vbInformation, "Error(s)", , , 13) = vbCancel Then
                        QB13_SendRequestsSingle = -1
                        Exit Function
                    End If
                End If
            Else
                QB13_SendRequestsSingle = QB13_SendRequestsSingle + 1
            End If
        Else
            If Notify Then MessageBox.Show("Error communicating with QuickBooks (nothing done):" & vbCrLf & RetMsg)
        End If
    End Function

    Public Function QB13_AppendDespositAdd(
    Optional ByVal TxnDate As String = "",
    Optional ByVal AcctRef_ListID As String = "", Optional ByVal AcctRef_FullName As String = "",
    Optional ByVal FromAccount_ListID As String = "", Optional ByVal FromAccount_FullName As String = "",
    Optional ByVal Memo As String = "",
    Optional ByVal CheckNumber As String = "",
    Optional ByVal PayMethRef_ListID As String = "", Optional ByVal PayMethRef_FullName As String = "",
    Optional ByVal Amount As String = "",
    Optional ByVal EntityRef_ListID As String = "", Optional ByVal EntityRef_FullName As String = "",
    Optional ByVal ClassRef_ListID As String = "", Optional ByVal ClassRef_FullName As String = "") As Boolean
        On Error Resume Next
        Err.Clear()
        Dim Dep As QBFC13Lib.IDepositAdd
        Dep = MSREQ13.AppendDepositAddRq

        If AcctRef_ListID = "" And AcctRef_FullName = "" Then
            Err.Raise(-1, , "Must specify customer for invoice")
        End If
        If FromAccount_ListID = "" And FromAccount_FullName = "" Then
            Err.Raise(-1, , "Must specify customer for invoice")
        End If

        IfNNSetValue(Dep.TxnDate, TxnDate)
        IfNNSetValue(Dep.DepositToAccountRef.ListID, AcctRef_ListID)
        IfNNSetValue(Dep.DepositToAccountRef.FullName, AcctRef_FullName)
        IfNNSetValue(Dep.Memo, Memo)

        Dim D As Object
        D = Dep.DepositLineAddList.Append.ORDepositLineAdd.DepositInfo
        IfNNSetValue(D.EntityRef.ListID, EntityRef_ListID)
        IfNNSetValue(D.EntityRef.FullName, EntityRef_FullName)
        IfNNSetValue(D.AccountRef.ListID, FromAccount_ListID)
        IfNNSetValue(D.AccountRef.FullName, FromAccount_FullName)
        IfNNSetValue(D.Memo, Memo)
        IfNNSetValue(D.CheckNumber, CheckNumber)
        IfNNSetValue(D.PaymentMethodRef.ListID, PayMethRef_ListID)
        IfNNSetValue(D.PaymentMethodRef.FullName, PayMethRef_FullName)
        IfNNSetValue(D.ClassRef.ListID, ClassRef_ListID)
        IfNNSetValue(D.ClassRef.FullName, ClassRef_FullName)
        IfNNSetValue(D.Amount, Amount)
        D = Nothing
        QB13_AppendDespositAdd = True
    End Function

    Public Function QB13CreateJournalEntry(
    Optional ByVal TxnDate As String = "", Optional ByVal RefNumber As String = "",
    Optional ByVal Memo As String = "", Optional ByVal IsAdjustment As String = "",
    Optional ByVal DebitTxnLineID As String = "",
    Optional ByVal DebitAccountRef_ListID As String = "", Optional ByVal DebitAccountRef_FullName As Object = "",
    Optional ByVal DebitAmount As String = "", Optional ByVal DebitMemo As String = "",
    Optional ByVal DebitEntityRef_ListID As String = "", Optional ByVal DebitEntityRef_FullName As Object = "",
    Optional ByVal DebitClassRef_ListID As String = "", Optional ByVal DebitClassRef_FullName As Object = "",
    Optional ByVal CreditTxnLineID As String = "",
    Optional ByVal CreditAccountRef_ListID As String = "", Optional ByVal CreditAccountRef_FullName As Object = "",
    Optional ByVal CreditAmount As String = "", Optional ByVal CreditMemo As String = "",
    Optional ByVal CreditEntityRef_ListID As String = "", Optional ByVal CreditEntityRef_FullName As Object = "",
    Optional ByVal CreditClassRef_ListID As String = "", Optional ByVal CreditClassRef_FullName As Object = "") As Boolean

        Dim E As Integer, S As String
        QBSM13_Reset()

        QB13_AppendJournalEntryAdd _
      (TxnDate, RefNumber, Memo, IsAdjustment,
      DebitTxnLineID, DebitAccountRef_ListID, DebitAccountRef_FullName,
      DebitAmount, DebitMemo, DebitEntityRef_ListID, DebitEntityRef_FullName,
      DebitClassRef_ListID, DebitClassRef_FullName,
      CreditTxnLineID, CreditAccountRef_ListID, CreditAccountRef_FullName,
      CreditAmount, CreditMemo, CreditEntityRef_ListID, CreditEntityRef_FullName,
      CreditClassRef_ListID, CreditClassRef_FullName)

        QB13CreateJournalEntry = (QB13_SendRequestsSingle(E, S) <> 0) And E = 0
        QBSM13_Reset()
    End Function

    Public Function QB13_AppendJournalEntryAdd(
    Optional ByVal TxnDate As String = "", Optional ByVal RefNumber As String = "",
    Optional ByVal Memo As String = "", Optional ByVal IsAdjustment As String = "",
    Optional ByVal DebitTxnLineID As String = "",
    Optional ByVal DebitAccountRef_ListID As String = "", Optional ByVal DebitAccountRef_FullName As Object = "",
    Optional ByVal DebitAmount As String = "", Optional ByVal DebitMemo As String = "",
    Optional ByVal DebitEntityRef_ListID As String = "", Optional ByVal DebitEntityRef_FullName As Object = "",
    Optional ByVal DebitClassRef_ListID As String = "", Optional ByVal DebitClassRef_FullName As Object = "",
    Optional ByVal CreditTxnLineID As String = "",
    Optional ByVal CreditAccountRef_ListID As String = "", Optional ByVal CreditAccountRef_FullName As Object = "",
    Optional ByVal CreditAmount As String = "", Optional ByVal CreditMemo As String = "",
    Optional ByVal CreditEntityRef_ListID As String = "", Optional ByVal CreditEntityRef_FullName As Object = "",
    Optional ByVal CreditClassRef_ListID As String = "", Optional ByVal CreditClassRef_FullName As Object = "") As Boolean

        On Error Resume Next
        Err.Clear()
        Dim Jrn As QBFC13Lib.IJournalEntryAdd
        Jrn = MSREQ13.AppendJournalEntryAddRq
        IfNNSetValue(Jrn.TxnDate, TxnDate)
        IfNNSetValue(Jrn.RefNumber, RefNumber)
        IfNNSetValue(Jrn.Memo, Memo)
        IfNNSetValue(Jrn.IsAdjustment, IsAdjustment)
        Dim L As QBFC13Lib.IORJournalLine

        If DebitTxnLineID <> "" Or DebitAccountRef_ListID <> "" Or DebitAccountRef_FullName <> "" Or DebitAmount <> "" Or DebitMemo <> "" Or DebitEntityRef_ListID <> "" Or DebitEntityRef_FullName <> "" Or DebitClassRef_ListID <> "" Or DebitClassRef_FullName <> "" Then
            L = Jrn.ORJournalLineList.Append
            IfNNSetValue(L.JournalDebitLine.TxnLineID, DebitTxnLineID)
            IfNNSetValue(L.JournalDebitLine.AccountRef.ListID, DebitAccountRef_ListID)
            IfNNSetValue(L.JournalDebitLine.AccountRef.FullName, DebitAccountRef_FullName)
            IfNNSetValue(L.JournalDebitLine.Amount, DebitAmount)
            IfNNSetValue(L.JournalDebitLine.Memo, DebitMemo)
            IfNNSetValue(L.JournalDebitLine.EntityRef.ListID, DebitEntityRef_ListID)
            IfNNSetValue(L.JournalDebitLine.EntityRef.FullName, DebitEntityRef_FullName)
            IfNNSetValue(L.JournalDebitLine.ClassRef.ListID, DebitClassRef_ListID)
            IfNNSetValue(L.JournalDebitLine.ClassRef.FullName, DebitClassRef_FullName)
            L = Nothing
        End If
        If CreditTxnLineID <> "" Or CreditAccountRef_ListID <> "" Or CreditAccountRef_FullName <> "" Or CreditAmount <> "" Or CreditMemo <> "" Or CreditEntityRef_ListID <> "" Or CreditEntityRef_FullName <> "" Or CreditClassRef_ListID <> "" Or CreditClassRef_FullName <> "" Then
            L = Jrn.ORJournalLineList.Append
            IfNNSetValue(L.JournalCreditLine.TxnLineID, CreditTxnLineID)
            IfNNSetValue(L.JournalCreditLine.AccountRef.ListID, CreditAccountRef_ListID)
            IfNNSetValue(L.JournalCreditLine.AccountRef.FullName, CreditAccountRef_FullName)
            IfNNSetValue(L.JournalCreditLine.Amount, CreditAmount)
            IfNNSetValue(L.JournalCreditLine.Memo, CreditMemo)
            IfNNSetValue(L.JournalCreditLine.EntityRef.ListID, CreditEntityRef_ListID)
            IfNNSetValue(L.JournalCreditLine.EntityRef.FullName, CreditEntityRef_FullName)
            IfNNSetValue(L.JournalCreditLine.ClassRef.ListID, CreditClassRef_ListID)
            IfNNSetValue(L.JournalCreditLine.ClassRef.FullName, CreditClassRef_FullName)
        End If
        L = Nothing
        Jrn = Nothing
        QB13_AppendJournalEntryAdd = True
    End Function
End Module
