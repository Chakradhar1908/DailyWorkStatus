Imports QBFC5Lib
Imports QBFC10Lib
Module modQuickBooks_CommQBFC13
    Private mQBSM13 As QBFC13Lib.QBSessionManager
    Private mMsReq13 As QBFC13Lib.IMsgSetRequest
    Private mMsRsp13 As QBFC13Lib.IMsgSetResponse
    Private mQBActiveStore As Long, mQBAS_JustSet As Boolean
    Public Function QB13_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As QBFC13Lib.IVendorRet
        Dim VQ As QBFC13Lib.IVendorQuery

        On Error Resume Next
        VQ = MSREQ13.AppendVendorQueryRq

        VQ.ORVendorListQuery.FullNameList.Add(Vendor)
        VQ = Nothing

        QB13_VendorQuery_Vendor = QB13_SendRequestsGetRet(RET, RetMsg)
        Exit Function
    End Function

    Public Function QB13_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As Object
        Dim Q As QBFC13Lib.IClassQuery
        QBSM13_Reset()

        On Error Resume Next
        Q = MSREQ13.AppendClassQueryRq
        Q.ORListQuery.FullNameList.Add(ClassName)
        Q = Nothing

        QB13_ClassQuery_Class = Nothing
        QB13_ClassQuery_Class = QB13_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB13_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As QBFC13Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC13Lib.ICustomerQuery
        QBSM13_Reset()
        Q = MSREQ13.AppendCustomerQueryRq
        Q.ORCustomerListQuery.FullNameList.Add(Name)
        '    Q.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameFilter.Name.SetValue Name
        Q = Nothing

        QB13_CustomerQuery_Name = QB13_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB13_AccountQuery_All(Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As Object
        On Error Resume Next
        QBSM13_Reset()
        MSREQ13.AppendAccountQueryRq() ' nothing to set
        QB13_AccountQuery_All = QB13_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB13_VendorQuery_All(Optional ByRef RET As Long = 0) As Object
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

    Public Function QB13_SendRequests(Optional ByRef ErrString As String = "", Optional ByRef ErrNo As Long = 0, Optional ByRef OnErr As ENRqOnError = ENRqOnError.roeContinue) As Boolean
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

    Public Function QB13_SendRequestsGetRet(Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As Object
        Dim RL As Object, En As Long, ES As String
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

        MsReq13_Reset
    End Sub

    Public Sub MsReq13_Reset()
        mMsReq13 = Nothing
        mMsRsp13 = Nothing
    End Sub

    Public Function QB13_SendRequestsGetRetList(Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As Object
        Dim RL As Object, T() As Object, N As Long, En As Long, ES As String
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

    Public Function QB13Startup(Optional ByRef RET As Long = 0, Optional ByRef Msg As String = "") As Boolean
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

    Private ReadOnly Property StatusCode() As Long
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

End Module
