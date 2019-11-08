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
            tvwList.Nodes.Add(SaleEntry)

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
                    tvwList.Nodes(0).Nodes.Add(L)
                Next
            End If

            RS.MoveNext
        Loop
        RS.Close
        RS = Nothing
    End Sub

End Class