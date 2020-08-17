Public Class OrdConvertLAW
    Public Result As String
    Public Location as integer
    Public Cancelled As Boolean
    Public Poid as integer

    Private mStatus As String

    Public Sub SetupForm(ByVal nStyle As String, ByVal nStatus As String, ByVal nLocation As String, Optional ByVal Poid as integer = 0)
        Cancelled = False
        Style = nStyle
        Status = nStatus
        txtLoc.Text = nLocation
        ConvertTo = ""
    End Sub
    Public Property Style() As String
        Get
            Style = lblStyle.Text
        End Get
        Set(value As String)
            Dim M As New CInvRec
            lblStyle.Text = value
            If M.Load(value, "Style") Then
                txtAvailable.Text = CStr(M.Available)
                txtOnOrder.Text = CStr(M.QueryTotalOnOrder)
            Else
                txtAvailable.Text = ""
                txtOnOrder.Text = ""
            End If
            DisposeDA(M)
        End Set
    End Property

    Public Property Status() As String
        Get
            Status = mStatus
        End Get
        Set(value As String)
            mStatus = value
            Me.Text = "Convert " & value & " to..."
            Select Case UCase(Status)
                Case "LAW"
                    optST.Text = "Convert To S&tock"
                    optST.Visible = True
                    optSO.Visible = True
                    optPO.Visible = True
                    optTW.Visible = False
                Case "ST"
                    optST.Text = "Convert To LAW"
                    optST.Visible = True
                    optSO.Visible = True
                    optPO.Visible = True
                    optTW.Visible = True
                Case "PO"
                    optST.Text = "Convert To S&tock"
                    optST.Visible = True
                    optSO.Visible = False
                    optPO.Visible = False
                    optTW.Visible = False
                Case "POREC", "SOREC"
                    optNn.Visible = True
                    optPO.Visible = False
                    optSO.Visible = False
                    optST.Visible = False
                    optTW.Visible = False
                Case Else
                    MsgBox("Unknown status: " & value)
            End Select
        End Set
    End Property
    Public Property ConvertTo() As String
        Get
            If optNn.Checked Then ConvertTo = ""
            If optSO.Checked Then ConvertTo = "SO"
            If optPO.Checked Then ConvertTo = "PO"
            If optST.Checked Then ConvertTo = "ST"
            If optTW.Checked Then ConvertTo = "DELTW"
        End Get
        Set(value As String)
            Select Case UCase(value)
                Case "SO" : optSO.Checked = True
                Case "PO" : optPO.Checked = True
                Case "ST" : optST.Checked = True
                Case "DELTW" : optTW.Checked = True
                Case Else
                    optSO.Checked = False
                    optPO.Checked = False
                    optST.Checked = False
                    optTW.Checked = False
                    optNn.Checked = False
                    optNn.Checked = True
            End Select
            Result = ConvertTo
        End Set
    End Property
End Class