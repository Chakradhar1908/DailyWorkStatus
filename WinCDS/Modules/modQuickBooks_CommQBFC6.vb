Imports QBFC10Lib
Module modQuickBooks_CommQBFC6
    Private mQBSM6 As QBFC6Lib.QBSessionManager
    Private mMsReq6 As QBFC6Lib.IMsgSetRequest
    Private mMsRsp6 As QBFC6Lib.IMsgSetResponse
    Private mQBActiveStore As Integer, mQBAS_JustSet As Boolean

    Public Function QB6_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC6Lib.IVendorRet
        Dim VQ As QBFC6Lib.IVendorQuery

        On Error Resume Next
        VQ = MSREQ6.AppendVendorQueryRq
        With VQ
            .ORVendorListQuery.FullNameList.Add(Vendor)
            '    .ORVendorListQuery.VendorListFilter.ORNameFilter.NameFilter.Name.SetValue Vendor
        End With
        VQ = Nothing

        QB6_VendorQuery_Vendor = QB6_SendRequestsGetRet(RET, RetMsg)
        Exit Function
    End Function

    Public ReadOnly Property MSREQ6() As QBFC6Lib.IMsgSetRequest
        Get
            If mMsReq6 Is Nothing Then
                mMsReq6 = QBSM6.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
            End If
            MSREQ6 = mMsReq6
        End Get
    End Property

    Public ReadOnly Property QBSM6() As QBFC6Lib.QBSessionManager
        Get
            If mQBSM6 Is Nothing Then
                mQBSM6 = New QBFC6Lib.QBSessionManager
            End If
            QBSM6 = mQBSM6
        End Get
    End Property

    Public Function QB6_SendRequestsGetRet(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim RL As Object, En As Integer, ES As String
        If QB6_SendRequests(ES, En) Then
            RET = StatusCode
            RetMsg = StatusMsg
            If RET = 0 Then
                If Not Resp.Detail Is Nothing Then
                    On Error Resume Next
                    QB6_SendRequestsGetRet = Resp.Detail
                    RL = Resp.Detail
                    QB6_SendRequestsGetRet = RL.GetAt(0)
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
        QBSM6_Reset()
    End Function

    Public Sub QBSM6_Reset()
        '  Set mQBSM = Nothing
        ActiveLog("QB::QBSM6_Reset", 3)
        If mQBAS_JustSet Or POMode("EDIT", "REC") Then
            mQBAS_JustSet = False
        Else
            QBActiveStore = 0
        End If

        MsReq6_Reset()
    End Sub

    Public Sub MsReq6_Reset()
        mMsReq6 = Nothing
        mMsRsp6 = Nothing
    End Sub

    Private ReadOnly Property Resp() As QBFC6Lib.IResponse
        Get
            On Error Resume Next
            Resp = Response(0)
        End Get
    End Property

    Public ReadOnly Property ResponseList6() As QBFC6Lib.IResponseList
        Get
            If Not MSRSP6 Is Nothing Then
                ResponseList6 = MSRSP6.ResponseList
            End If
        End Get
    End Property

    Private ReadOnly Property Response(Index As Integer) As QBFC6Lib.IResponse
        Get
            On Error Resume Next
            Response = ResponseList6.GetAt(Index)
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

    Public Function QB6_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC6Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC6Lib.ICustomerQuery
        QBSM6_Reset()
        Q = MSREQ6.AppendCustomerQueryRq
        With Q
            Q.ORCustomerListQuery.FullNameList.Add(Name)
            '    Q.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameFilter.Name.SetValue Name
        End With
        Q = Nothing

        QB6_CustomerQuery_Name = QB6_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB6_AccountQuery_All(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        On Error Resume Next
        QBSM6_Reset()
        MSREQ6.AppendAccountQueryRq() ' nothing to set
        QB6_AccountQuery_All = QB6_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB6_SendRequestsGetRetList(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim RL As Object
        Dim T() As Object, N As Integer, En As Integer, ES As String
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
        QB6_SendRequestsGetRetList = T
        QBSM6_Reset()
    End Function

    Public Function QB6_VendorQuery_All(Optional ByRef RET As Integer = 0)
        Dim ES As String
        On Error Resume Next
        QBSM6_Reset()
        MSREQ6.AppendVendorQueryRq() ' nothing to set
        QB6_VendorQuery_All = QB6_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB6ObjectsExist(Optional ByRef Msg As String = "") As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC6)"
        X = New QBFC6Lib.QBSessionManager
        '  Set X = CreateObject("QBFC6.QBSessionManager")
        Msg = "Could not create MsgSetRequest (QBFC6)"
        Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
        Msg = ""
        QB6ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
        Y = Nothing
        X = Nothing
    End Function

    Public Function QB6_SendRequests(Optional ByRef ErrString As String = "", Optional ByRef ErrNo As Integer = 0, Optional ByRef OnErr As ENRqOnError = ENRqOnError.roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB6Startup(ErrNo, ErrString) Then Exit Function
        MSREQ6.Attributes.OnError = OnErr
        MSRSP6 = QBSM6.DoRequests(MSREQ6)
        '  Debug.Print mmsrsp6.ToXMLString

        QB6_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Private Property MSRSP6() As QBFC6Lib.IMsgSetResponse
        Get
            MSRSP6 = mMsRsp6
        End Get
        Set(value As QBFC6Lib.IMsgSetResponse)
            mMsRsp6 = value
        End Set
    End Property

    Public Function QB6Startup(Optional ByRef RET As Integer = 0, Optional ByRef Msg As String = "") As Boolean
        Dim P As frmProgress
        Dim E As String
        On Error GoTo NoComm
        ActiveLog("QB::QBStartup", 1)
        If Not QBConnOpen Then
            P = New frmProgress
            P.Progress(0, , "Opening Connection to Quickbooks...", True, False)
            If QBUseRDS() Then
                QBSM6.OpenConnection2(QB_AppID, QB_AppNm, QBFC6Lib.ENConnectionType.ctRemoteQBD)
            Else
                QBSM6.OpenConnection2(QB_AppID, QB_AppNm, QBFC6Lib.ENConnectionType.ctLocalQBD)
            End If
            QBConnOpen = True
        End If
        If Not QBSessOpen Then
            If QBUseRDS() Then
                '      QBSM6.BeginSession "", omDontCare 'BFH20150325 - Added the file.. If you don't want it, leave it blank in the form
                QBSM6.BeginSession(QB_File, QBFC6Lib.ENOpenMode.omDontCare)
            Else
                QBSM6.BeginSession(QB_File, QBFC6Lib.ENOpenMode.omDontCare)
            End If
            QBSessOpen = True
        End If
        QB6Startup = True
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
        Exit Function
NoComm:
        RET = Err.Number
        Msg = Err.Description
        Err.Clear()
        QB6Shutdown()
        QB6Startup = False
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
    End Function

    Public Function QB6Shutdown()
        ActiveLog("QB::QBShutdown", 1)
        If QBSessOpen Then QBSM6.EndSession() : QBSessOpen = False
        If QBConnOpen Then QBSM6.CloseConnection() : QBConnOpen = False
        mQBSM6 = Nothing
    End Function

End Module
