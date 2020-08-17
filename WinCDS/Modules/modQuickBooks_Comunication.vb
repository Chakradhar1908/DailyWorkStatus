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

End Module
