Public Class InvReports
    Public Property Item() As Integer
        Get
            For Item = 1 To 20
                'If ItemCtrl(Item) Then Exit Property
                If ItemCtrl(Item).Checked = True Then Exit Property
            Next
            Item = -1
        End Get
        Set(value As Integer)
            If ItemCtrl(value) Is Nothing Then
                Item = 1
                'ItemCtrl(1) = False
                optList1.Checked = False
            Else
                'If ItemCtrl(value) Then ItemCtrl(value) = False
                If ItemCtrl(value).Checked = True Then ItemCtrl(value).Checked = False
                'ItemCtrl(value) = True
                ItemCtrl(value).Checked = True
            End If
        End Set
    End Property

    Private ReadOnly Property ItemCtrl(ByVal vData As Integer) As RadioButton
        Get
            If vData <= 0 Or vData > 20 Then Exit Property
            ItemCtrl = Nothing

            'ItemCtrl = optList(vData)
            Select Case vData
                Case 1
                    ItemCtrl = optList1
                Case 2
                    ItemCtrl = optList2
                Case 3
                    ItemCtrl = optList3
                Case 4
                    ItemCtrl = optList4
                Case 5
                    ItemCtrl = optList5
                Case 6
                    ItemCtrl = optList7
                Case 7
                    ItemCtrl = optList7
                Case 8
                    ItemCtrl = optList8
                Case 9
                    ItemCtrl = optList9
                Case 10
                    ItemCtrl = optList10
                Case 11
                    ItemCtrl = optList11
                Case 12
                    ItemCtrl = optList12
                Case 13
                    ItemCtrl = optList13
                Case 14
                    ItemCtrl = optList14
                Case 15
                    ItemCtrl = optList15
                Case 16
                    ItemCtrl = optList16
                Case 17
                    ItemCtrl = optList17
                Case 18
                    ItemCtrl = optList18
                Case 19
                    ItemCtrl = optList19
                Case 20
                    ItemCtrl = optList20
            End Select
        End Get
    End Property
End Class