Module modQuickBooks_CommQBFC8
    Public Function qb8_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As qbfc8lib.IVendorRet
        Dim VQ As qbfc8lib.IVendorQuery
        '  qbsm8_reset ' BFH2009821 - REMOVED

        On Error Resume Next
  Set VQ = MsReq8.AppendVendorQueryRq
  VQ.ORVendorListQuery.FullNameList.Add Vendor
  Set VQ = Nothing
  
  Set qb8_VendorQuery_Vendor = QB8_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB8_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "")
        Dim Q As QBFC8Lib.IClassQuery
        QBSM8_Reset

        On Error Resume Next
  Set Q = MsReq8.AppendClassQueryRq
  Q.ORListQuery.FullNameList.Add ClassName
  Set Q = Nothing
  
  Set QB8_ClassQuery_Class = Nothing
  Set QB8_ClassQuery_Class = QB8_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function qb8_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String) As QBFC8Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC8Lib.ICustomerQuery
        QBSM8_Reset
  Set Q = MsReq8.AppendCustomerQueryRq
  With Q
            Q.ORCustomerListQuery.FullNameList.Add Name
  End With
  Set Q = Nothing

  Set qb8_CustomerQuery_Name = QB8_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB8_AccountQuery_All(Optional ByRef RET As Long, Optional ByRef RetMsg As String)
        On Error Resume Next
        QBSM8_Reset
        MsReq8.AppendAccountQueryRq ' nothing to set
        QB8_AccountQuery_All = QB8_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function qb8_VendorQuery_All(Optional ByRef RET As Long)
        Dim ES As String
        On Error Resume Next
        QBSM8_Reset
        MsReq8.AppendVendorQueryRq ' nothing to set
        qb8_VendorQuery_All = QB8_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB8ObjectsExist(Optional ByRef Msg As String) As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC8)"
  Set X = New qbfc8lib.QBSessionManager
  Msg = "Could not create MsgSetRequest (QBFC8)"
  Set Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
  Msg = ""
        QB8ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
  Set Y = Nothing
  Set X = Nothing
End Function

    Public Function QB8_SendRequests(Optional ByRef ErrString As String, Optional ByRef ErrNo As Long, Optional ByRef OnErr As ENRqOnError = roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB8Startup(ErrNo, ErrString) Then Exit Function
        MsReq8.Attributes.OnError = OnErr
  Set MsRsp8 = QBSM8.DoRequests(MsReq8)
'  Debug.Print mmsrsp8.ToXMLString
  
  QB8_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Public Function QB8Shutdown()
        ActiveLog "QB::QBShutdown", 1
On Error Resume Next
        If QBSessOpen Then
            QBSM8.EndSession
            QBSessOpen = False
        End If
        If QBConnOpen Then QBSM8.CloseConnection : QBConnOpen = False
  Set mQBSM8 = Nothing
End Function

End Module
