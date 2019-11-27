Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class CalendarInstr
    Private mDD As Date

    Public Property DeliveryDay() As Date
        Get
            DeliveryDay = mDD
        End Get
        Set(value As Date)
            mDD = value
            fraDate.Text = Format(mDD, "ddd, m/d/yyyy")
            LoadDeliveries()
        End Set
    End Property

    Private Sub LoadDeliveries()
        Dim RS As ADODB.Recordset, N As TreeNode
        Dim XIndex As Integer, SaleNm As String, SaleNo As String, SalePD As String
        Dim SaleEntry As String
        Dim SpRS As ADODB.Recordset, SpInstr As String

        RS = Calendar.GetDeliveryCalendarData(DeliveryDay, 1)
        tvwList.Nodes.Clear()

        Do Until RS.EOF

            XIndex = RS("Index").Value
            SaleNo = RS("SaleNo").Value
            SalePD = IfNullThenNilString(RS("PD").Value)
            SaleNm = RS("Name").Value
            Select Case SalePD
                Case "P", "S", "D" ' Do nothing
                Case Else
                    SalePD = "D"      ' Default to Delivery.
            End Select

            SaleEntry = Microsoft.VisualBasic.Left(SalePD & " " & SaleNm & New String(" ", 50), 20) & " " & SaleNo
            'N = tvwList.Nodes.Add(, , , SaleEntry)
            N = tvwList.Nodes.Add(SaleEntry)

            SpInstr = ""
            SpRS = GetRecordsetBySQL("SELECT Special FROM GrossMargin LEFT JOIN Mail ON Mail.Index = GrossMargin.MailIndex WHERE SaleNo='" & SaleNo & "'", , GetDatabaseAtLocation())
            If Not SpRS.EOF Then
                SpInstr = IfNullThenNilString(SpRS("Special").Value)
            End If
            SpRS = Nothing

            'N.Expanded = True
            N.Expand()
            If SpInstr <> "" Then
                Dim L As Object
                SpInstr = Replace(SpInstr, vbLf, vbCr)
                SpInstr = Replace(SpInstr, vbCr & vbCr, vbCr)
                For Each L In Split(SpInstr, vbCr)
                    'tvwList.Nodes.Add(N, tvwChild, , L)
                    'tvwList.Nodes(0).Nodes.Add(L)
                    N.Nodes.Add(L)
                Next
            End If

            RS.MoveNext
        Loop
        RS.Close
        RS = Nothing
    End Sub

    Private Sub CalendarInstr_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetButtonImage(cmdExit, "ok")
        SetButtonImage(cmdExit, 2)
    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub CalendarInstr_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        fraDate.Width = Me.Width - fraDate.Left - 20
        fraDate.Height = Me.Height - fraDate.Top - cmdExit.Height - 60
        tvwList.Width = fraDate.Width - tvwList.Left - 10
        tvwList.Height = fraDate.Height - 20
        cmdExit.Left = Me.Width / 2 - cmdExit.Width / 2
        cmdExit.Top = fraDate.Top + fraDate.Height + 10
    End Sub
End Class