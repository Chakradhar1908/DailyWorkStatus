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

End Module
