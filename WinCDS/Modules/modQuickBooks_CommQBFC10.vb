Module modQuickBooks_CommQBFC10
    Public Function QB10_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As qbfc10lib.IVendorRet
        Dim VQ As qbfc10lib.IVendorQuery
        '  qbsm10_reset ' BFH20091021 - REMOVED

        On Error Resume Next
  Set VQ = MsReq10.AppendVendorQueryRq
  VQ.ORVendorListQuery.FullNameList.Add Vendor
  Set VQ = Nothing
  
  Set QB10_VendorQuery_Vendor = QB10_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB10_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "")
        Dim Q As QBFC10Lib.IClassQuery
        QBSM10_Reset

        On Error Resume Next
  Set Q = MsReq10.AppendClassQueryRq
  Q.ORListQuery.FullNameList.Add ClassName
  Set Q = Nothing
  
  Set QB10_ClassQuery_Class = Nothing
  Set QB10_ClassQuery_Class = QB10_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB10_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String) As QBFC10Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC10Lib.ICustomerQuery
        QBSM10_Reset
  Set Q = MsReq10.AppendCustomerQueryRq
  Q.ORCustomerListQuery.FullNameList.Add Name
  Set Q = Nothing

  Set QB10_CustomerQuery_Name = QB10_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB10_AccountQuery_All(Optional ByRef RET As Long, Optional ByRef RetMsg As String)
        On Error Resume Next
        QBSM10_Reset
        MsReq10.AppendAccountQueryRq ' nothing to set
        QB10_AccountQuery_All = QB10_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB10_VendorQuery_All(Optional ByRef RET As Long)
        Dim ES As String
        On Error Resume Next
        QBSM10_Reset
        MsReq10.AppendVendorQueryRq ' nothing to set
        QB10_VendorQuery_All = QB10_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB10ObjectsExist(Optional ByRef Msg As String) As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC10)"
  Set X = New qbfc10lib.QBSessionManager
  Msg = "Could not create MsgSetRequest (QBFC10)"
  Set Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
  Msg = ""
        QB10ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
  Set Y = Nothing
  Set X = Nothing
End Function

    Public Function QB10_SendRequests(Optional ByRef ErrString As String, Optional ByRef ErrNo As Long, Optional ByRef OnErr As ENRqOnError = roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB10Startup(ErrNo, ErrString) Then Exit Function
        MsReq10.Attributes.OnError = OnErr
  Set MsRsp10 = QBSM10.DoRequests(MsReq10)
'  Debug.Print mmsrsp10.ToXMLString
  
  QB10_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Public Function QB10Shutdown()
        ActiveLog "QB::QBShutdown", 1
  If QBSessOpen Then QBSM10.EndSession : QBSessOpen = False
        If QBConnOpen Then QBSM10.CloseConnection : QBConnOpen = False
  Set mQBSM10 = Nothing
End Function

End Module
