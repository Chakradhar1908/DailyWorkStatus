Imports Microsoft.VisualBasic.Interaction
Imports InetCtlsObjects.ProtocolConstants
Imports InetCtlsObjects.StateConstants
Imports InetCtlsObjects.DataTypeConstants
Public Class frmINet
    Private Declare Function GetTickCount Lib "kernel32" () As Integer  ' use for timerless timers
    Event inetStateChanged(ByVal State As InetCtlsObjects.StateConstants)
    Event inetConnected()
    Event inetRequestSent()
    Event inetResponseReceived()
    Event inetResponseCompleted()
    Event inetDisconnected()
    Event inetError()
    Event inetOtherState()

    Private Const DebugStateChange As Boolean = False

#Const DebugPrint = False

    Private mState As Integer, mEndTime As Date

    Public Function MakeRequest(ByVal URL As String, Optional ByVal Operation As String = "GET", Optional ByVal InputData As String = "", Optional ByVal InputHdrs As String = "", Optional ByVal DoSecure As TriState = vbUseDefault) As String
        Dim U As URLExtract
        U = ExtractUrl(URL)

        Protocol = Switch(DoSecure = vbTrue, icHTTPS, DoSecure = vbFalse, icHTTP, True, IIf(U.Scheme = "https", icHTTPS, icHTTP))
        Execute(URL, Operation, IIf(InputData = "" And Operation = "POST", U.Query, InputData), InputHdrs)
        Application.DoEvents()
        If INETTimeout = 0 Then INETTimeout = INETTimeout_Default
        EndTime(INETTimeout)
        Do While IsIn(State, icNone, icConnecting, icConnected, icReceivingResponse, icRequesting, icRequestSent)
            Application.DoEvents()
#If DebugPrint Then
Debug.Print State & " - " & INetState
#End If
            If EndTime() Then GoTo OutOfLoop
        Loop
OutOfLoop:
        'On Error Resume Next
        MakeRequest = GetResponse()
    End Function

    Public Function GetResponse() As String
        Dim X As Integer
        On Error Resume Next
        X = GetTickCount() + 10000
        Do While True
            Application.DoEvents()
            GetResponse = GetChunk(100000, icString)
            If GetResponse <> "" Then Exit Function
            If GetTickCount() > X Then Exit Function
            Application.DoEvents()
        Loop
        '  GetResponse = GetChunk(100000, icString)
    End Function

    Public Function GetChunk(ByVal Size As Integer, Optional ByVal DataType As InetCtlsObjects.DataTypeConstants = Nothing) As Object
        GetChunk = inet.GetChunk(Size, DataType)
    End Function

    Public Property Protocol() As InetCtlsObjects.ProtocolConstants
        Get
            Protocol = inet.Protocol
        End Get
        Set(value As InetCtlsObjects.ProtocolConstants)
            inet.Protocol = value
        End Set
    End Property

    Public Sub Execute(Optional ByVal URL As String = "", Optional ByVal Operation As String = "GET", Optional ByVal InputData As String = "", Optional ByVal InputHdrs As String = "")
        On Error Resume Next
        ' The .Execute is extremely broken...  you must not pass URL, it must be set in properties.. you must not pass the last 2 ars unless used..

        'If Not IsMissing(URL) Then inet.URL = URL
        If Not IsNothing(URL) Then inet.URL = URL

        'If (IsMissing(InputData) Or InputData = "") And (IsMissing(InputHdrs) Or InputHdrs = "") Then
        If (IsNothing(InputData) Or InputData = "") And (IsNothing(InputHdrs) Or InputHdrs = "") Then
            'inet.Execute(, Operation)
            inet.Execute("", Operation, "", "")
        ElseIf InputHdrs = "" Then
            'inet.Execute , Operation, InputData
            inet.Execute("", Operation, InputData, "")
        ElseIf InputData = "" Then
            'inet.Execute , Operation, , InputHdrs
            inet.Execute("", Operation, "", InputHdrs)
        Else
            'inet.Execute , Operation, InputData, InputHdrs
            inet.Execute("", Operation, InputData, InputHdrs)
        End If
        If Err.Description <> "" Then
#If DebugPrint Then
Debug.Print Err.Description
#End If
        End If
    End Sub

    Private Function EndTime(Optional ByVal Duration As Integer = 0) As Boolean
        If Duration = 0 Then
            EndTime = DateAfter2(Now, mEndTime, , "s") : Exit Function
        Else
            mEndTime = DateAdd("s", Duration, Now)
            EndTime = False
        End If
    End Function

    Public ReadOnly Property State() As InetCtlsObjects.StateConstants
        Get
            State = mState
        End Get
    End Property

    Public Function MakeRequestOnly(ByVal URL As String, Optional ByVal Operation As String = "GET", Optional ByVal InputData As String = "", Optional ByVal InputHdrs As String = "", Optional ByVal DoSecure As VBA.VbTriState = vbUseDefault) As Boolean
        Dim U As URLExtract
        U = ExtractUrl(URL)

        Protocol = Switch(DoSecure = vbTrue, icHTTPS, DoSecure = vbFalse, icHTTP, True, IIf(U.Scheme = "https", icHTTPS, icHTTP))
        Execute(URL, Operation, IIf(InputData = "" And Operation = "POST", U.Query, InputData), InputHdrs)
        Application.DoEvents()
        If INETTimeout = 0 Then INETTimeout = INETTimeout_Default
        EndTime(INETTimeout)
        Do While IsIn(State, icNone, icConnecting, icConnected, icRequesting, icRequestSent)
            Application.DoEvents()
#If DebugPrint Then
Debug.Print State & " - " & INetState
#End If
            If EndTime() Then GoTo OutOfLoop
        Loop
OutOfLoop:
        'On Error Resume Next
        MakeRequestOnly = True
    End Function

End Class