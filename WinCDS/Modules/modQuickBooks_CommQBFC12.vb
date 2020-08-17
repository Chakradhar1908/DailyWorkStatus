Imports QBFC10Lib
Module modQuickBooks_CommQBFC12
    Private mQBSM12 As QBFC12Lib.QBSessionManager
    Private mMsReq12 As QBFC12Lib.IMsgSetRequest
    Private mMsRsp12 As QBFC12Lib.IMsgSetResponse
    Private mQBActiveStore As Integer, mQBAS_JustSet As Boolean

    Public Function QB12_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC12Lib.IVendorRet
        Dim VQ As QBFC12Lib.IVendorQuery
        '  qbsm12_reset ' BFH20091221 - REMOVED

        On Error Resume Next
        VQ = msReq12.AppendVendorQueryRq
        With VQ
            .ORVendorListQuery.FullNameList.Add(Vendor)
        End With
        VQ = Nothing

        QB12_VendorQuery_Vendor = QB12_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB12_SendRequestsGetRet(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim RL As Object, En As Integer, ES As String
        If QB12_SendRequests(ES, En) Then
            RET = StatusCode
            RetMsg = StatusMsg
            If RET = 0 Then
                If Not Resp.Detail Is Nothing Then
                    On Error Resume Next
                    QB12_SendRequestsGetRet = Resp.Detail
                    RL = Resp.Detail
                    QB12_SendRequestsGetRet = RL.GetAt(0)
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
        QBSM12_Reset()
    End Function

    Public ReadOnly Property ResponseList12() As QBFC12Lib.IResponseList
        Get
            If Not MsRsp12 Is Nothing Then
                ResponseList12 = MsRsp12.ResponseList
            End If
        End Get
    End Property

    Private ReadOnly Property Response(Index As Integer) As QBFC12Lib.IResponse
        Get
            On Error Resume Next
            Response = ResponseList12.GetAt(Index)
        End Get
    End Property

    Private ReadOnly Property Resp() As QBFC12Lib.IResponse
        Get
            On Error Resume Next
            Resp = Response(0)
        End Get
    End Property

    Public Sub QBSM12_Reset()
        '  Set mQBSM = Nothing
        ActiveLog("QB::QBSM12_Reset", 3)
        If mQBAS_JustSet Or POMode("EDIT", "REC") Then
            mQBAS_JustSet = False
        Else
            QBActiveStore = 0
        End If

        MsReq12_Reset()
    End Sub

    Public Sub MsReq12_Reset()
        mMsReq12 = Nothing
        mMsRsp12 = Nothing
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

    Public ReadOnly Property msReq12() As QBFC12Lib.IMsgSetRequest
        Get
            If mMsReq12 Is Nothing Then
                mMsReq12 = QBSM12.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
            End If
            msReq12 = mMsReq12
        End Get
    End Property

    Public ReadOnly Property QBSM12() As QBFC12Lib.QBSessionManager
        Get
            If mQBSM12 Is Nothing Then
                mQBSM12 = New QBFC12Lib.QBSessionManager
            End If
            QBSM12 = mQBSM12
        End Get
    End Property

    Public Function QB12_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim Q As QBFC12Lib.IClassQuery
        QBSM12_Reset()

        On Error Resume Next
        Q = msReq12.AppendClassQueryRq
        With Q
            .ORListQuery.FullNameList.Add(ClassName)
        End With
        Q = Nothing

        QB12_ClassQuery_Class = Nothing
        QB12_ClassQuery_Class = QB12_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB12_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC12Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC12Lib.ICustomerQuery
        QBSM12_Reset()
        Q = msReq12.AppendCustomerQueryRq
        With Q
            Q.ORCustomerListQuery.FullNameList.Add(Name)
        End With
        Q = Nothing

        QB12_CustomerQuery_Name = QB12_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB12_AccountQuery_All(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        On Error Resume Next
        QBSM12_Reset()
        msReq12.AppendAccountQueryRq() ' nothing to set
        QB12_AccountQuery_All = QB12_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB12_SendRequestsGetRetList(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
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
        QB12_SendRequestsGetRetList = T
        QBSM12_Reset()
    End Function

    Public Function QB12_VendorQuery_All(Optional ByRef RET As Integer = 0)
        Dim ES As String
        On Error Resume Next
        QBSM12_Reset()
        msReq12.AppendVendorQueryRq() ' nothing to set
        QB12_VendorQuery_All = QB12_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB12ObjectsExist(Optional ByRef Msg As String = "") As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC12)"
        X = New QBFC12Lib.QBSessionManager
        Msg = "Could not create MsgSetRequest (QBFC12)"
        Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
        Msg = ""
        QB12ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
        Y = Nothing
        X = Nothing
    End Function

    Public Function QB12_SendRequests(Optional ByRef ErrString As String = "", Optional ByRef ErrNo As Integer = 0, Optional ByRef OnErr As ENRqOnError = ENRqOnError.roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB12Startup(ErrNo, ErrString) Then Exit Function
        msReq12.Attributes.OnError = OnErr
        MsRsp12 = QBSM12.DoRequests(msReq12)
        '  Debug.Print mmsrsp12.ToXMLString

        QB12_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Private Property MsRsp12() As QBFC12Lib.IMsgSetResponse
        Get
            MsRsp12 = mMsRsp12
        End Get
        Set(value As QBFC12Lib.IMsgSetResponse)
            mMsRsp12 = value
        End Set
    End Property

    Public Function QB12Startup(Optional ByRef RET As Integer = 0, Optional ByRef Msg As String = "") As Boolean
        Dim P As frmProgress
        Dim E As String
        On Error GoTo NoComm
        ActiveLog("QB::QBStartup", 1)
        If Not QBConnOpen Then
            P = New frmProgress
            P.Progress(0, , "Opening Connection to Quickbooks...", True, False)
            If QBUseRDS() Then
                QBSM12.OpenConnection2(QB_AppID, QB_AppNm, QBFC12Lib.ENConnectionType.ctRemoteQBD)
            Else
                QBSM12.OpenConnection2(QB_AppID, QB_AppNm, QBFC12Lib.ENConnectionType.ctLocalQBD)
            End If
            QBConnOpen = True
        End If
        If Not QBSessOpen Then
            If QBUseRDS() Then
                '      QBSM12.BeginSession "", omDontCare 'BFH20150325 - Added the file.. If you don't want it, leave it blank in the form
                QBSM12.BeginSession(QB_File, QBFC12Lib.ENOpenMode.omDontCare)
            Else
                QBSM12.BeginSession(QB_File, QBFC12Lib.ENOpenMode.omDontCare)
            End If
            QBSessOpen = True
        End If
        QB12Startup = True
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
        Exit Function
NoComm:
        RET = Err.Number
        Msg = Err.Description
        Err.Clear()
        QB12Shutdown()
        QB12Startup = False
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
    End Function

    Public Function QB12Shutdown()
        ActiveLog("QB::QBShutdown", 1)
        If QBSessOpen Then QBSM12.EndSession() : QBSessOpen = False
        If QBConnOpen Then QBSM12.CloseConnection() : QBConnOpen = False
        mQBSM12 = Nothing
    End Function

End Module
