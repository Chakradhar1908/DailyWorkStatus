Module modQuickBooks_CommQBFC12
    Public Function QB12_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As qbfc12lib.IVendorRet
        Dim VQ As qbfc12lib.IVendorQuery
        '  qbsm12_reset ' BFH20091221 - REMOVED

        On Error Resume Next
  Set VQ = msReq12.AppendVendorQueryRq
  With VQ
            .ORVendorListQuery.FullNameList.Add Vendor
  End With
  Set VQ = Nothing
  
  Set QB12_VendorQuery_Vendor = QB12_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB12_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "")
        Dim Q As QBFC12Lib.IClassQuery
        QBSM12_Reset

        On Error Resume Next
  Set Q = msReq12.AppendClassQueryRq
  With Q
            .ORListQuery.FullNameList.Add ClassName
  End With
  Set Q = Nothing
  
  Set QB12_ClassQuery_Class = Nothing
  Set QB12_ClassQuery_Class = QB12_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB12_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String) As QBFC12Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC12Lib.ICustomerQuery
        QBSM12_Reset
  Set Q = msReq12.AppendCustomerQueryRq
  With Q
            Q.ORCustomerListQuery.FullNameList.Add Name
  End With
  Set Q = Nothing

  Set QB12_CustomerQuery_Name = QB12_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB12_AccountQuery_All(Optional ByRef RET As Long, Optional ByRef RetMsg As String)
        On Error Resume Next
        QBSM12_Reset
        msReq12.AppendAccountQueryRq ' nothing to set
        QB12_AccountQuery_All = QB12_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB12_VendorQuery_All(Optional ByRef RET As Long)
        Dim ES As String
        On Error Resume Next
        QBSM12_Reset
        msReq12.AppendVendorQueryRq ' nothing to set
        QB12_VendorQuery_All = QB12_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB12ObjectsExist(Optional ByRef Msg As String) As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC12)"
  Set X = New qbfc12lib.QBSessionManager
  Msg = "Could not create MsgSetRequest (QBFC12)"
  Set Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
  Msg = ""
        QB12ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
  Set Y = Nothing
  Set X = Nothing
End Function

    Public Function QB12_SendRequests(Optional ByRef ErrString As String, Optional ByRef ErrNo As Long, Optional ByRef OnErr As ENRqOnError = roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB12Startup(ErrNo, ErrString) Then Exit Function
        msReq12.Attributes.OnError = OnErr
  Set MsRsp12 = QBSM12.DoRequests(msReq12)
'  Debug.Print mmsrsp12.ToXMLString
  
  QB12_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Public Function QB12Shutdown()
        ActiveLog "QB::QBShutdown", 1
  If QBSessOpen Then QBSM12.EndSession : QBSessOpen = False
        If QBConnOpen Then QBSM12.CloseConnection : QBConnOpen = False
  Set mQBSM12 = Nothing
End Function

End Module
