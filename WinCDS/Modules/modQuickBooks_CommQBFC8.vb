Imports QBFC10Lib

Module modQuickBooks_CommQBFC8
    Private mQBSM8 As QBFC8Lib.QBSessionManager
    Private mMsReq8 As QBFC8Lib.IMsgSetRequest
    Private mMsRsp8 As QBFC8Lib.IMsgSetResponse
    Private mQBActiveStore As Integer, mQBAS_JustSet As Boolean

    Public Function qb8_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC8Lib.IVendorRet
        Dim VQ As QBFC8Lib.IVendorQuery
        '  qbsm8_reset ' BFH2009821 - REMOVED

        On Error Resume Next
        VQ = MsReq8.AppendVendorQueryRq
        VQ.ORVendorListQuery.FullNameList.Add(Vendor)
        VQ = Nothing

        qb8_VendorQuery_Vendor = QB8_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB8_SendRequestsGetRet(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim RL As Object, En As Integer, ES As String
        If QB8_SendRequests(ES, En) Then
            RET = StatusCode
            RetMsg = StatusMsg
            If RET = 0 Then
                If Not Resp.Detail Is Nothing Then
                    On Error Resume Next
                    QB8_SendRequestsGetRet = Resp.Detail
                    RL = Resp.Detail
                    QB8_SendRequestsGetRet = RL.GetAt(0)
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
        QBSM8_Reset()
    End Function

    Public ReadOnly Property responseList8() As QBFC8Lib.IResponseList
        Get
            If Not MsRsp8 Is Nothing Then
                responseList8 = MsRsp8.ResponseList
            End If
        End Get
    End Property

    Private ReadOnly Property Response(Index As Integer) As QBFC8Lib.IResponse
        Get
            On Error Resume Next
            Response = responseList8.GetAt(Index)
        End Get
    End Property

    Private ReadOnly Property Resp() As QBFC8Lib.IResponse
        Get
            On Error Resume Next
            Resp = Response(0)
        End Get
    End Property

    Public Sub QBSM8_Reset()
        '  Set mQBSM = Nothing
        ActiveLog("QB::QBSM8_Reset", 3)
        If mQBAS_JustSet Or POMode("EDIT", "REC") Then
            mQBAS_JustSet = False
        Else
            QBActiveStore = 0
        End If

        MsReq8_Reset()
    End Sub

    Public Sub MsReq8_Reset()
        mMsReq8 = Nothing
        mMsRsp8 = Nothing
    End Sub

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
    Public ReadOnly Property MsReq8() As QBFC8Lib.IMsgSetRequest
        Get
            If mMsReq8 Is Nothing Then
                mMsReq8 = QBSM8.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
            End If
            MsReq8 = mMsReq8
        End Get
    End Property

    Public ReadOnly Property QBSM8() As QBFC8Lib.QBSessionManager
        Get
            If mQBSM8 Is Nothing Then
                mQBSM8 = New QBFC8Lib.QBSessionManager
            End If
            QBSM8 = mQBSM8
        End Get
    End Property

    Public Function QB8_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim Q As QBFC8Lib.IClassQuery
        QBSM8_Reset()
        On Error Resume Next
        Q = MsReq8.AppendClassQueryRq
        Q.ORListQuery.FullNameList.Add(ClassName)
        Q = Nothing

        QB8_ClassQuery_Class = Nothing
        QB8_ClassQuery_Class = QB8_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function qb8_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC8Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC8Lib.ICustomerQuery
        QBSM8_Reset()
        Q = MsReq8.AppendCustomerQueryRq
        With Q
            Q.ORCustomerListQuery.FullNameList.Add(Name)
        End With
        Q = Nothing

        qb8_CustomerQuery_Name = QB8_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB8_AccountQuery_All(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        On Error Resume Next
        QBSM8_Reset()
        MsReq8.AppendAccountQueryRq() ' nothing to set
        QB8_AccountQuery_All = QB8_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function qb8_VendorQuery_All(Optional ByRef RET As Integer = 0)
        Dim ES As String
        On Error Resume Next
        QBSM8_Reset()
        MsReq8.AppendVendorQueryRq() ' nothing to set
        qb8_VendorQuery_All = QB8_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB8_SendRequestsGetRetList(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
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
        QB8_SendRequestsGetRetList = T
        QBSM8_Reset()
    End Function

    Public Function QB8ObjectsExist(Optional ByRef Msg As String = "") As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC8)"
        X = New QBFC8Lib.QBSessionManager
        Msg = "Could not create MsgSetRequest (QBFC8)"
        Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
        Msg = ""
        QB8ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
        Y = Nothing
        X = Nothing
    End Function

    Public Function QB8_SendRequests(Optional ByRef ErrString As String = "", Optional ByRef ErrNo As Integer = 0, Optional ByRef OnErr As ENRqOnError = ENRqOnError.roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB8Startup(ErrNo, ErrString) Then Exit Function
        MsReq8.Attributes.OnError = OnErr
        MsRsp8 = QBSM8.DoRequests(MsReq8)
        '  Debug.Print mmsrsp8.ToXMLString

        QB8_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Private Property MsRsp8() As QBFC8Lib.IMsgSetResponse
        Get
            MsRsp8 = mMsRsp8
        End Get
        Set(value As QBFC8Lib.IMsgSetResponse)
            mMsRsp8 = value
        End Set
    End Property

    Public Function QB8Startup(Optional ByRef RET As Integer = 0, Optional ByRef Msg As String = "") As Boolean
        Dim P As frmProgress
        Dim E As String
        On Error GoTo NoComm
        ActiveLog("QB::QBStartup", 1)
        If Not QBConnOpen Then
            P = New frmProgress
            P.Progress(0, , "Opening Connection to Quickbooks...", True, False)
            If QBUseRDS() Then
                QBSM8.OpenConnection2(QB_AppID, QB_AppNm, QBFC8Lib.ENConnectionType.ctRemoteQBD)
            Else
                QBSM8.OpenConnection2(QB_AppID, QB_AppNm, QBFC8Lib.ENConnectionType.ctLocalQBD)
            End If
            QBConnOpen = True
        End If
        If Not QBSessOpen Then
            If QBUseRDS() Then
                '      QBSM8.BeginSession "", omDontCare 'BFH20150325 - Added the file.. If you don't want it, leave it blank in the form
                QBSM8.BeginSession(QB_File, QBFC8Lib.ENOpenMode.omDontCare)
            Else
                QBSM8.BeginSession(QB_File, QBFC8Lib.ENOpenMode.omDontCare)
            End If
            QBSessOpen = True
        End If
        QB8Startup = True
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
        Exit Function
NoComm:
        RET = Err.Number
        Msg = Err.Description
        Err.Clear()
        QB8Shutdown()
        QB8Startup = False
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
    End Function

    Public Function QB8Shutdown()
        ActiveLog("QB::QBShutdown", 1)
        On Error Resume Next
        If QBSessOpen Then
            QBSM8.EndSession()
            QBSessOpen = False
        End If
        If QBConnOpen Then QBSM8.CloseConnection() : QBConnOpen = False
        mQBSM8 = Nothing
    End Function

    Public Function qb8CreateDeposit(
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
        QBSM8_Reset()
        qb8_AppendDespositAdd _
      (TxnDate, AcctRef_ListID, AcctRef_FullName, FromAccount_ListID, FromAccount_FullName,
      Memo, CheckNumber, PayMethRef_ListID, PayMethRef_FullName,
      Amount, EntityRef_ListID, EntityRef_FullName, ClassRef_ListID, ClassRef_FullName)

        qb8CreateDeposit = (QB8_SendRequestsSingle(R, M) <> 0) And R = 0
        QBSM8_Reset()
    End Function

    Public Function QB8_SendRequestsSingle(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "", Optional ByVal Notify As Boolean = True) As Integer
        Dim RR As QBFC8Lib.IResponse, II As Integer

        If QB8_SendRequests(RetMsg, RET, ENRqOnError.roeContinue) Then
            For II = 0 To ResponseCount - 1
                RR = responseList8.GetAt(II)
                If RR.StatusCode <> 0 Then
                    RetMsg = RetMsg & IIf(Len(RetMsg) > 0, vbCrLf, "")
                    RetMsg = RetMsg & "Error processing " & RR.Type.GetAsString & ": " & RR.StatusMessage
                End If
            Next
            If Len(RetMsg) > 0 Then
                If Notify Then
                    If MsgBox("Error(s): " & vbCrLf & RetMsg, vbInformation, "Error(s)", , , 8) = vbCancel Then
                        QB8_SendRequestsSingle = -1
                        Exit Function
                    End If
                End If
            Else
                QB8_SendRequestsSingle = QB8_SendRequestsSingle + 1
            End If
        Else
            If Notify Then MessageBox.Show("Error communicating with QuickBooks (nothing done):" & vbCrLf & RetMsg)
        End If
    End Function

    Public Function qb8_AppendDespositAdd(
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
        Dim Dep As QBFC8Lib.IDepositAdd
        Dep = MsReq8.AppendDepositAddRq
        With Dep
            If AcctRef_ListID = "" And AcctRef_FullName = "" Then
                Err.Raise(-1, , "Must specify customer for invoice")
            End If
            If FromAccount_ListID = "" And FromAccount_FullName = "" Then
                Err.Raise(-1, , "Must specify customer for invoice")
            End If

            IfNNSetValue(.TxnDate, TxnDate)
            IfNNSetValue(.DepositToAccountRef.ListID, AcctRef_ListID)
            IfNNSetValue(.DepositToAccountRef.FullName, AcctRef_FullName)
            IfNNSetValue(.Memo, Memo)

            With .DepositLineAddList.Append.ORDepositLineAdd.DepositInfo
                IfNNSetValue(.EntityRef.ListID, EntityRef_ListID)
                IfNNSetValue(.EntityRef.FullName, EntityRef_FullName)
                IfNNSetValue(.AccountRef.ListID, FromAccount_ListID)
                IfNNSetValue(.AccountRef.FullName, FromAccount_FullName)
                IfNNSetValue(.Memo, Memo)
                IfNNSetValue(.CheckNumber, CheckNumber)
                IfNNSetValue(.PaymentMethodRef.ListID, PayMethRef_ListID)
                IfNNSetValue(.PaymentMethodRef.FullName, PayMethRef_FullName)
                IfNNSetValue(.ClassRef.ListID, ClassRef_ListID)
                IfNNSetValue(.ClassRef.FullName, ClassRef_FullName)
                IfNNSetValue(.Amount, Amount)
            End With
        End With
        Dep = Nothing
        qb8_AppendDespositAdd = True
    End Function

    Public Function qb8CreateJournalEntry(
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
        QBSM8_Reset()

        qb8_AppendJournalEntryAdd _
      (TxnDate, RefNumber, Memo, IsAdjustment,
      DebitTxnLineID, DebitAccountRef_ListID, DebitAccountRef_FullName,
      DebitAmount, DebitMemo, DebitEntityRef_ListID, DebitEntityRef_FullName,
      DebitClassRef_ListID, DebitClassRef_FullName,
      CreditTxnLineID, CreditAccountRef_ListID, CreditAccountRef_FullName,
      CreditAmount, CreditMemo, CreditEntityRef_ListID, CreditEntityRef_FullName,
      CreditClassRef_ListID, CreditClassRef_FullName)

        qb8CreateJournalEntry = (QB8_SendRequestsSingle(E, S) <> 0) And E = 0
        QBSM8_Reset()
    End Function

    Public Function qb8_AppendJournalEntryAdd(
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
        Dim Jrn As QBFC8Lib.IJournalEntryAdd
        Jrn = MsReq8.AppendJournalEntryAddRq
        With Jrn
            IfNNSetValue(.TxnDate, TxnDate)
            IfNNSetValue(.RefNumber, RefNumber)
            IfNNSetValue(.Memo, Memo)
            IfNNSetValue(.IsAdjustment, IsAdjustment)
            Dim L As QBFC8Lib.IORJournalLine

            If DebitTxnLineID <> "" Or DebitAccountRef_ListID <> "" Or DebitAccountRef_FullName <> "" Or DebitAmount <> "" Or DebitMemo <> "" Or DebitEntityRef_ListID <> "" Or DebitEntityRef_FullName <> "" Or DebitClassRef_ListID <> "" Or DebitClassRef_FullName <> "" Then
                L = .ORJournalLineList.Append
                With L.JournalDebitLine
                    IfNNSetValue(.TxnLineID, DebitTxnLineID)
                    IfNNSetValue(.AccountRef.ListID, DebitAccountRef_ListID)
                    IfNNSetValue(.AccountRef.FullName, DebitAccountRef_FullName)
                    IfNNSetValue(.Amount, DebitAmount)
                    IfNNSetValue(.Memo, DebitMemo)
                    IfNNSetValue(.EntityRef.ListID, DebitEntityRef_ListID)
                    IfNNSetValue(.EntityRef.FullName, DebitEntityRef_FullName)
                    IfNNSetValue(.ClassRef.ListID, DebitClassRef_ListID)
                    IfNNSetValue(.ClassRef.FullName, DebitClassRef_FullName)
                End With
                L = Nothing
            End If
            If CreditTxnLineID <> "" Or CreditAccountRef_ListID <> "" Or CreditAccountRef_FullName <> "" Or CreditAmount <> "" Or CreditMemo <> "" Or CreditEntityRef_ListID <> "" Or CreditEntityRef_FullName <> "" Or CreditClassRef_ListID <> "" Or CreditClassRef_FullName <> "" Then
                L = .ORJournalLineList.Append
                With L.JournalCreditLine
                    IfNNSetValue(.TxnLineID, CreditTxnLineID)
                    IfNNSetValue(.AccountRef.ListID, CreditAccountRef_ListID)
                    IfNNSetValue(.AccountRef.FullName, CreditAccountRef_FullName)
                    IfNNSetValue(.Amount, CreditAmount)
                    IfNNSetValue(.Memo, CreditMemo)
                    IfNNSetValue(.EntityRef.ListID, CreditEntityRef_ListID)
                    IfNNSetValue(.EntityRef.FullName, CreditEntityRef_FullName)
                    IfNNSetValue(.ClassRef.ListID, CreditClassRef_ListID)
                    IfNNSetValue(.ClassRef.FullName, CreditClassRef_FullName)
                End With
            End If
            L = Nothing
        End With
        Jrn = Nothing
        qb8_AppendJournalEntryAdd = True
    End Function

End Module

