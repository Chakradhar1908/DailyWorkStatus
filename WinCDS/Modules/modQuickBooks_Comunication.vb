Imports QBFC5Lib
Module modQuickBooks_Comunication
    Private mMsReq As QBFC5Lib.IMsgSetRequest
    Private mQBSM As QBFC5Lib.QBSessionManager
    Private mMsRsp As QBFC5Lib.IMsgSetResponse
    Public Function QB5_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC5Lib.IVendorRet
        Dim VQ As QBFC5Lib.IVendorQuery
        '  qbsm5_reset ' BFH20091021 - REMOVED

        On Error Resume Next
        VQ = MsReq5.AppendVendorQueryRq
        VQ.ORVendorListQuery.FullNameList.Add(Vendor)
        VQ = Nothing

        QB5_VendorQuery_Vendor = QB5_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public ReadOnly Property MsReq5() As QBFC5Lib.IMsgSetRequest
        Get
            If mMsReq Is Nothing Then
                mMsReq = QBSM5.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
            End If
            MsReq5 = mMsReq
        End Get
    End Property

    Public ReadOnly Property QBSM5() As QBFC5Lib.QBSessionManager
        Get
            If mQBSM Is Nothing Then
                mQBSM = New QBFC5Lib.QBSessionManager
                '    Set mQBSM = CreateObject("QBFC5Lib.QBSessionManager")
            End If
            QBSM5 = mQBSM
        End Get
    End Property

    Public Function QB5_SendRequestsGetRet(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As Object
        Dim RL As Object, En As Integer, ES As String
        If QB_SendRequests(ES, En) Then
            RET = StatusCode
            RetMsg = StatusMsg
            If RET = 0 Then
                If Not Resp.Detail Is Nothing Then
                    On Error Resume Next
                    QB5_SendRequestsGetRet = Resp.Detail
                    RL = Resp.Detail
                    QB5_SendRequestsGetRet = RL.GetAt(0)
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
        QBSM5_Reset()
    End Function

    Public Function QB5_SendRequests(Optional ByRef ErrString As String = "", Optional ByRef ErrNo As Integer = 0, Optional ByRef OnErr As ENRqOnError = ENRqOnError.roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB5Startup(ErrNo, ErrString) Then Exit Function
        MsReq5.Attributes.OnError = OnErr
        MsRsp5 = QBSM5.DoRequests(MsReq5)
        '  Debug.Print mMsRsp.ToXMLString

        QB5_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

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

    Private ReadOnly Property Resp() As QBFC5Lib.IResponse
        Get
            On Error Resume Next
            Resp = Response(0)
        End Get
    End Property

    Public Sub QBSM5_Reset()
        '  Set mQBSM = Nothing
        ActiveLog("QB::qbsm5_reset", 3)
        If mQBAS_JustSet Or POMode("EDIT", "REC") Then
            mQBAS_JustSet = False
        Else
            QBActiveStore = 0
        End If

        MsReq5_Reset()
    End Sub

    Public Sub MsReq5_Reset()
        mMsReq = Nothing
        mMsRsp = Nothing
    End Sub

    Public Function QB5Startup(Optional ByRef RET As Integer = 0, Optional ByRef Msg As String = "") As Boolean
        Dim P As frmProgress
        Dim E As String
        On Error GoTo NoComm
        ActiveLog("QB::QBStartup", 1)
        If Not QBConnOpen Then
            P = New frmProgress
            P.Progress(0, , "Opening Connection to Quickbooks...", True, False)
            If QBUseRDS() Then
                QBSM5.OpenConnection2(QB_AppID, QB_AppNm, ENConnectionType.ctRemoteQBD)
            Else
                QBSM5.OpenConnection2(QB_AppID, QB_AppNm, ENConnectionType.ctLocalQBD)
            End If
            QBConnOpen = True
        End If
        If Not QBSessOpen Then
            If QBUseRDS() Then
                QBSM5.BeginSession(QB_File, ENOpenMode.omDontCare)
            Else
                QBSM5.BeginSession(QB_File, ENOpenMode.omDontCare)
            End If
            QBSessOpen = True
        End If
        QB5Startup = True
        If Not P Is Nothing Then P.ProgressClose() : P = Nothing
        Exit Function
NoComm:
        RET = Err.Number
        Msg = Err.Description
        Err.Clear()
        QB5Shutdown()
        QB5Startup = False
        If Not P Is Nothing Then P.ProgressClose()
        P = Nothing
    End Function

    Public Sub QB5Shutdown()
        ActiveLog("QB::QBShutdown", 1)
        If QBSessOpen Then QBSM5.EndSession()
        QBSessOpen = False
        If QBConnOpen Then QBSM5.CloseConnection()
        QBConnOpen = False
        mQBSM = Nothing
    End Sub

    Private Property MsRsp5() As QBFC5Lib.IMsgSetResponse
        Get
            MsRsp5 = mMsRsp
        End Get
        Set(value As QBFC5Lib.IMsgSetResponse)
            mMsRsp = value
        End Set
    End Property

    Private ReadOnly Property Response(Index As Integer) As QBFC5Lib.IResponse
        Get
            On Error Resume Next
            Response = responseList5.GetAt(Index)
        End Get
    End Property

    Public ReadOnly Property responseList5() As QBFC5Lib.IResponseList
        Get
            If Not MsRsp5 Is Nothing Then
                responseList5 = MsRsp5.ResponseList
            End If
        End Get
    End Property

    Public Function QB5_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As Object
        Dim Q As QBFC5Lib.IClassQuery
        QBSM5_Reset()

        On Error Resume Next
        Q = MsReq5.AppendClassQueryRq
        Q.ORListQuery.FullNameList.Add(ClassName)
        Q = Nothing

        QB5_ClassQuery_Class = Nothing
        QB5_ClassQuery_Class = QB5_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB5_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC5Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC5Lib.ICustomerQuery
        QBSM5_Reset()
        Q = MsReq5.AppendCustomerQueryRq
        Q.ORCustomerListQuery.FullNameList.Add(Name)
        Q = Nothing

        QB5_CustomerQuery_Name = QB5_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB5_AccountQuery_All(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As Object
        On Error Resume Next
        QBSM5_Reset()
        MsReq5.AppendAccountQueryRq() ' nothing to set
        QB5_AccountQuery_All = QB5_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB5_VendorQuery_All(Optional ByRef RET As Integer = 0) As Object
        Dim ES As String
        On Error Resume Next
        QBSM5_Reset()
        MsReq5.AppendVendorQueryRq() ' nothing to set
        QB5_VendorQuery_All = QB5_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB5_SendRequestsGetRetList(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As Object
        Dim RL As Object, T() As Object
        Dim N As Integer, En As Integer, ES As String
        If QB5_SendRequests(ES, En) Then
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
        QB5_SendRequestsGetRetList = T
        QBSM5_Reset()
    End Function

    Public Function QB5ObjectsExist(Optional ByRef Msg As String = "") As Boolean
        Dim X As Object, Y As Object
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC5)"
        X = New QBFC5Lib.QBSessionManager
        Msg = "Could not create MsgSetRequest (QBFC5)"
        Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
        Msg = ""
        QB5ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
        Y = Nothing
        X = Nothing
    End Function

    Public Function QB5CreateDeposit(
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
        QBSM5_Reset()
        QB5_AppendDespositAdd _
      (TxnDate, AcctRef_ListID, AcctRef_FullName, FromAccount_ListID, FromAccount_FullName,
      Memo, CheckNumber, PayMethRef_ListID, PayMethRef_FullName,
      Amount, EntityRef_ListID, EntityRef_FullName, ClassRef_ListID, ClassRef_FullName)

        QB5CreateDeposit = (QB5_SendRequestsSingle(R, M) <> 0) And R = 0
        QBSM5_Reset()
    End Function

    Public Function QB5_SendRequestsSingle(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "", Optional ByVal Notify As Boolean = True) As Integer
        Dim RR As QBFC5Lib.IResponse, II As Integer

        If QB_SendRequests(RetMsg, RET, QBFC10Lib.ENRqOnError.roeContinue) Then
            For II = 0 To ResponseCount - 1
                RR = responseList5.GetAt(II)
                If RR.StatusCode <> 0 Then
                    RetMsg = RetMsg & IIf(Len(RetMsg) > 0, vbCrLf, "")
                    RetMsg = RetMsg & "Error processing " & RR.Type.GetAsString & ": " & RR.StatusMessage
                End If
            Next
            If Len(RetMsg) > 0 Then
                If Notify Then
                    If MsgBox("Error(s): " & vbCrLf & RetMsg, vbInformation, "Error(s)", , , 10) = vbCancel Then
                        QB5_SendRequestsSingle = -1
                        Exit Function
                    End If
                End If
            Else
                QB5_SendRequestsSingle = QB5_SendRequestsSingle + 1
            End If
        Else
            If Notify Then MessageBox.Show("Error communicating with QuickBooks (nothing done):" & vbCrLf & RetMsg)
        End If
    End Function

    Public Function QB5_AppendDespositAdd(
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
        Dim Dep As QBFC5Lib.IDepositAdd
        Dep = MsReq5.AppendDepositAddRq

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

        Dim DepInfo As Object
        DepInfo = Dep.DepositLineAddList.Append.ORDepositLineAdd.DepositInfo
        IfNNSetValue(DepInfo.EntityRef.ListID, EntityRef_ListID)
        IfNNSetValue(DepInfo.EntityRef.FullName, EntityRef_FullName)
        IfNNSetValue(DepInfo.AccountRef.ListID, FromAccount_ListID)
        IfNNSetValue(DepInfo.AccountRef.FullName, FromAccount_FullName)
        IfNNSetValue(DepInfo.Memo, Memo)
        IfNNSetValue(DepInfo.CheckNumber, CheckNumber)
        IfNNSetValue(DepInfo.PaymentMethodRef.ListID, PayMethRef_ListID)
        IfNNSetValue(DepInfo.PaymentMethodRef.FullName, PayMethRef_FullName)
        IfNNSetValue(DepInfo.ClassRef.ListID, ClassRef_ListID)
        IfNNSetValue(DepInfo.ClassRef.FullName, ClassRef_FullName)
        IfNNSetValue(DepInfo.Amount, Amount)

        Dep = Nothing
        QB5_AppendDespositAdd = True
    End Function

    Public Function QB5CreateJournalEntry(
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
        QBSM5_Reset()

        QB5_AppendJournalEntryAdd _
      (TxnDate, RefNumber, Memo, IsAdjustment,
      DebitTxnLineID, DebitAccountRef_ListID, DebitAccountRef_FullName,
      DebitAmount, DebitMemo, DebitEntityRef_ListID, DebitEntityRef_FullName,
      DebitClassRef_ListID, DebitClassRef_FullName,
      CreditTxnLineID, CreditAccountRef_ListID, CreditAccountRef_FullName,
      CreditAmount, CreditMemo, CreditEntityRef_ListID, CreditEntityRef_FullName,
      CreditClassRef_ListID, CreditClassRef_FullName)

        QB5CreateJournalEntry = (QB5_SendRequestsSingle(E, S) <> 0) And E = 0
        QBSM5_Reset()
    End Function

    Public Function QB5_AppendJournalEntryAdd(
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
        Dim Jrn As QBFC5Lib.IJournalEntryAdd
        Jrn = MsReq5.AppendJournalEntryAddRq

        IfNNSetValue(Jrn.TxnDate, TxnDate)
        IfNNSetValue(Jrn.RefNumber, RefNumber)
        IfNNSetValue(Jrn.Memo, Memo)
        IfNNSetValue(Jrn.IsAdjustment, IsAdjustment)
        Dim L As QBFC5Lib.IORJournalLine

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
        QB5_AppendJournalEntryAdd = True
    End Function
End Module
