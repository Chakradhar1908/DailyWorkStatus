Imports QBFC10Lib

Module modQuickBooks_CommQBFC10
    Private mQBSM10 As QBFC10Lib.QBSessionManager
    Private mMsReq10 As QBFC10Lib.IMsgSetRequest
    Private mMsRsp10 As QBFC10Lib.IMsgSetResponse
    Private mQBActiveStore As Integer, mQBAS_JustSet As Boolean

    Public Function QB10_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC10Lib.IVendorRet
        Dim VQ As QBFC10Lib.IVendorQuery
        '  qbsm10_reset ' BFH20091021 - REMOVED

        On Error Resume Next
        VQ = MsReq10.AppendVendorQueryRq
        VQ.ORVendorListQuery.FullNameList.Add(Vendor)
        VQ = Nothing

        QB10_VendorQuery_Vendor = QB10_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB10_SendRequestsGetRet(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim RL As Object, En As Integer, ES As String
        If QB10_SendRequests(ES, En) Then
            RET = StatusCode
            RetMsg = StatusMsg
            If RET = 0 Then
                If Not Resp.Detail Is Nothing Then
                    On Error Resume Next
                    QB10_SendRequestsGetRet = Resp.Detail
                    RL = Resp.Detail
                    QB10_SendRequestsGetRet = RL.GetAt(0)
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
        QBSM10_Reset()
    End Function

    Public ReadOnly Property responseList10() As QBFC10Lib.IResponseList
        Get
            If Not MsRsp10 Is Nothing Then
                responseList10 = MsRsp10.ResponseList
            End If
        End Get
    End Property

    Private ReadOnly Property Response(Index As Integer) As QBFC10Lib.IResponse
        Get
            On Error Resume Next
            Response = responseList10.GetAt(Index)
        End Get
    End Property

    Private ReadOnly Property Resp() As QBFC10Lib.IResponse
        Get
            On Error Resume Next
            Resp = Response(0)
        End Get
    End Property

    Public Sub QBSM10_Reset()
        '  Set mQBSM = Nothing
        ActiveLog("QB::QBSM10_Reset", 3)
        If mQBAS_JustSet Or POMode("EDIT", "REC") Then
            mQBAS_JustSet = False
        Else
            QBActiveStore = 0
        End If

        MsReq10_Reset()
    End Sub

    Public Sub MsReq10_Reset()
        mMsReq10 = Nothing
        mMsRsp10 = Nothing
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

    Public ReadOnly Property MsReq10() As QBFC10Lib.IMsgSetRequest
        Get
            If mMsReq10 Is Nothing Then
                mMsReq10 = QBSM10.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
            End If
            MsReq10 = mMsReq10
        End Get
    End Property

    Public ReadOnly Property QBSM10() As QBFC10Lib.QBSessionManager
        Get
            If mQBSM10 Is Nothing Then
                mQBSM10 = New QBFC10Lib.QBSessionManager
            End If
            QBSM10 = mQBSM10
        End Get
    End Property

    Public Function QB10_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim Q As QBFC10Lib.IClassQuery
        QBSM10_Reset()

        On Error Resume Next
        Q = MsReq10.AppendClassQueryRq
        Q.ORListQuery.FullNameList.Add(ClassName)
        Q = Nothing

        QB10_ClassQuery_Class = Nothing
        QB10_ClassQuery_Class = QB10_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB10_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC10Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC10Lib.ICustomerQuery
        QBSM10_Reset()
        Q = MsReq10.AppendCustomerQueryRq
        Q.ORCustomerListQuery.FullNameList.Add(Name)
        Q = Nothing

        QB10_CustomerQuery_Name = QB10_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB10_AccountQuery_All(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        On Error Resume Next
        QBSM10_Reset()
        MsReq10.AppendAccountQueryRq() ' nothing to set
        QB10_AccountQuery_All = QB10_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB10_SendRequestsGetRetList(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
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
        QB10_SendRequestsGetRetList = T
        QBSM10_Reset()
    End Function

    Public Function QB10_VendorQuery_All(Optional ByRef RET As Integer = 0)
        Dim ES As String
        On Error Resume Next
        QBSM10_Reset()
        MsReq10.AppendVendorQueryRq() ' nothing to set
        QB10_VendorQuery_All = QB10_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB10ObjectsExist(Optional ByRef Msg As String = "") As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC10)"
        X = New QBFC10Lib.QBSessionManager
        Msg = "Could not create MsgSetRequest (QBFC10)"
        Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
        Msg = ""
        QB10ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
        Y = Nothing
        X = Nothing
    End Function

    Public Function QB10_SendRequests(Optional ByRef ErrString As String = "", Optional ByRef ErrNo As Integer = 0, Optional ByRef OnErr As ENRqOnError = ENRqOnError.roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB10Startup(ErrNo, ErrString) Then Exit Function
        MsReq10.Attributes.OnError = OnErr
        MsRsp10 = QBSM10.DoRequests(MsReq10)
        '  Debug.Print mmsrsp10.ToXMLString

        QB10_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Private Property MsRsp10() As QBFC10Lib.IMsgSetResponse
        Get
            MsRsp10 = mMsRsp10
        End Get
        Set(value As QBFC10Lib.IMsgSetResponse)
            mMsRsp10 = value
        End Set
    End Property

    Public Function QB10Startup(Optional ByRef RET As Integer = 0, Optional ByRef Msg As String = "") As Boolean
        Dim P As frmProgress
        Dim E As String
        On Error GoTo NoComm
        ActiveLog("QB::QBStartup", 1)
        If Not QBConnOpen Then
            P = New frmProgress
            P.Progress(0, , "Opening Connection to Quickbooks...", True, False)
            If QBUseRDS() Then
                QBSM10.OpenConnection2(QB_AppID, QB_AppNm, ENConnectionType.ctRemoteQBD)
            Else
                QBSM10.OpenConnection2(QB_AppID, QB_AppNm, ENConnectionType.ctLocalQBD)
            End If
            QBConnOpen = True
        End If
        If Not QBSessOpen Then
            If QBUseRDS() Then
                '      QBSM10.BeginSession "", omDontCare 'BFH20150325 - Added the file.. If you don't want it, leave it blank in the form
                QBSM10.BeginSession(QB_File, ENOpenMode.omDontCare)
            Else
                QBSM10.BeginSession(QB_File, ENOpenMode.omDontCare)
            End If
            QBSessOpen = True
        End If
        QB10Startup = True
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
        Exit Function
NoComm:
        RET = Err.Number
        Msg = Err.Description
        Err.Clear()
        QB10Shutdown()
        QB10Startup = False
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
    End Function

    Public Function QB10Shutdown()
        ActiveLog("QB::QBShutdown", 1)
        If QBSessOpen Then QBSM10.EndSession() : QBSessOpen = False
        If QBConnOpen Then QBSM10.CloseConnection() : QBConnOpen = False
        mQBSM10 = Nothing
    End Function

End Module
