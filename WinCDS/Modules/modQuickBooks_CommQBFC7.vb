Imports QBFC10Lib
Module modQuickBooks_CommQBFC7
    Private mQBSM7 As QBFC7Lib.QBSessionManager
    Private mMsReq7 As QBFC7Lib.IMsgSetRequest
    Private mMsRsp7 As QBFC7Lib.IMsgSetResponse
    Private mQBActiveStore As Integer, mQBAS_JustSet As Boolean

    Public Function QB7_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC7Lib.IVendorRet
        Dim VQ As QBFC7Lib.IVendorQuery

        On Error Resume Next
        VQ = MSREQ7.AppendVendorQueryRq
        With VQ
            .ORVendorListQuery.FullNameList.Add(Vendor)
            '    .ORVendorListQuery.VendorListFilter.ORNameFilter.NameFilter.Name.SetValue Vendor
        End With
        VQ = Nothing

        QB7_VendorQuery_Vendor = QB7_SendRequestsGetRet(RET, RetMsg)
        Exit Function
    End Function

    Public Function QB7_SendRequestsGetRet(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim RL As Object, En As Integer, ES As String
        If QB7_SendRequests(ES, En) Then
            RET = StatusCode
            RetMsg = StatusMsg
            If RET = 0 Then
                If Not Resp.Detail Is Nothing Then
                    On Error Resume Next
                    QB7_SendRequestsGetRet = Resp.Detail
                    RL = Resp.Detail
                    QB7_SendRequestsGetRet = RL.GetAt(0)
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
        QBSM7_Reset()
    End Function

    Private ReadOnly Property Resp() As QBFC7Lib.IResponse
        Get
            On Error Resume Next
            Resp = Response(0)
        End Get
    End Property

    Public ReadOnly Property ResponseList7() As QBFC7Lib.IResponseList
        Get
            If Not MSRSP7 Is Nothing Then
                ResponseList7 = MSRSP7.ResponseList
            End If
        End Get

    End Property

    Private ReadOnly Property Response(Index As Integer) As QBFC7Lib.IResponse
        Get
            On Error Resume Next
            Response = ResponseList7.GetAt(Index)
        End Get
    End Property

    Public Sub QBSM7_Reset()
        '  Set mQBSM = Nothing
        ActiveLog("QB::QBSM7_Reset", 3)
        If mQBAS_JustSet Or POMode("EDIT", "REC") Then
            mQBAS_JustSet = False
        Else
            QBActiveStore = 0
        End If

        MsReq7_Reset()
    End Sub

    Public Sub MsReq7_Reset()
        mMsReq7 = Nothing
        mMsRsp7 = Nothing
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

    Public ReadOnly Property MSREQ7() As QBFC7Lib.IMsgSetRequest
        Get
            If mMsReq7 Is Nothing Then
                mMsReq7 = QBSM7.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
            End If
            MSREQ7 = mMsReq7
        End Get
    End Property

    Public ReadOnly Property QBSM7() As QBFC7Lib.QBSessionManager
        Get
            If mQBSM7 Is Nothing Then
                mQBSM7 = New QBFC7Lib.QBSessionManager
            End If
            QBSM7 = mQBSM7
        End Get
    End Property

    Public Function QB7_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim Q As QBFC7Lib.IClassQuery
        QBSM7_Reset()

        On Error Resume Next
        Q = MSREQ7.AppendClassQueryRq
        With Q
            .ORListQuery.FullNameList.Add(ClassName)
        End With
        Q = Nothing

        QB7_ClassQuery_Class = Nothing
        QB7_ClassQuery_Class = QB7_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB7_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As QBFC7Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC7Lib.ICustomerQuery
        QBSM7_Reset()
        Q = MSREQ7.AppendCustomerQueryRq
        With Q
            Q.ORCustomerListQuery.FullNameList.Add(Name)
            '    Q.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameFilter.Name.SetValue Name
        End With
        Q = Nothing

        QB7_CustomerQuery_Name = QB7_SendRequestsGetRet(RET, RetMsg)
    End Function

    Public Function QB7_AccountQuery_All(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        On Error Resume Next
        QBSM7_Reset()
        MSREQ7.AppendAccountQueryRq() ' nothing to set
        QB7_AccountQuery_All = QB7_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB7_SendRequestsGetRetList(Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "")
        Dim RL As Object, T() As Object
        Dim N As Integer, En As Integer, ES As String

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
        QB7_SendRequestsGetRetList = T
        QBSM7_Reset()
    End Function

    Public Function QB7_VendorQuery_All(Optional ByRef RET As Integer = 0)
        Dim ES As String
        On Error Resume Next
        QBSM7_Reset()
        MSREQ7.AppendVendorQueryRq() ' nothing to set
        QB7_VendorQuery_All = QB7_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB7ObjectsExist(Optional ByRef Msg As String = "") As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC7)"
        X = New QBFC7Lib.QBSessionManager
        '  Set X = CreateObject("QBFC7.QBSessionManager")
        Msg = "Could not create MsgSetRequest (QBFC7)"
        Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
        Msg = ""
        QB7ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
        Y = Nothing
        X = Nothing
    End Function

    Public Function QB7_SendRequests(Optional ByRef ErrString As String = "", Optional ByRef ErrNo As Integer = 0, Optional ByRef OnErr As ENRqOnError = ENRqOnError.roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB7Startup(ErrNo, ErrString) Then Exit Function
        MSREQ7.Attributes.OnError = OnErr
        MSRSP7 = QBSM7.DoRequests(MSREQ7)
        '  Debug.Print mmsrsp7.ToXMLString

        QB7_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Private Property MSRSP7() As QBFC7Lib.IMsgSetResponse
        Get
            MSRSP7 = mMsRsp7
        End Get
        Set(value As QBFC7Lib.IMsgSetResponse)
            mMsRsp7 = value
        End Set
    End Property

    Public Function QB7Startup(Optional ByRef RET As Integer = 0, Optional ByRef Msg As String = "") As Boolean
        Dim P As frmProgress
        Dim E As String
        On Error GoTo NoComm
        ActiveLog("QB::QBStartup", 1)
        If Not QBConnOpen Then
            P = New frmProgress
            P.Progress(0, , "Opening Connection to Quickbooks...", True, False)
            If QBUseRDS() Then
                QBSM7.OpenConnection2(QB_AppID, QB_AppNm, QBFC7Lib.ENConnectionType.ctRemoteQBD)
            Else
                QBSM7.OpenConnection2(QB_AppID, QB_AppNm, QBFC7Lib.ENConnectionType.ctLocalQBD)
            End If
            QBConnOpen = True
        End If
        If Not QBSessOpen Then
            If QBUseRDS() Then
                '      QBSM7.BeginSession "", omDontCare 'BFH20150325 - Added the file.. If you don't want it, leave it blank in the form
                QBSM7.BeginSession(QB_File, QBFC7Lib.ENOpenMode.omDontCare)
            Else
                QBSM7.BeginSession(QB_File, QBFC7Lib.ENOpenMode.omDontCare)
            End If
            QBSessOpen = True
        End If
        QB7Startup = True
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
        Exit Function
NoComm:
        RET = Err.Number
        Msg = Err.Description
        Err.Clear()
        QB7Shutdown()
        QB7Startup = False
        If Not P Is Nothing Then P.ProgressClose() : 
        P = Nothing
    End Function

    Public Function QB7Shutdown()
        ActiveLog("QB::QBShutdown", 1)
        If QBSessOpen Then QBSM7.EndSession() : QBSessOpen = False
        If QBConnOpen Then QBSM7.CloseConnection() : QBConnOpen = False
        mQBSM7 = Nothing
    End Function

End Module
