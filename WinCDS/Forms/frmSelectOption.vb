Public Class frmSelectOption
    Private mOptionMode As ESelOpts
    Private mSelected As Object  '() as integer
    Private SkipCheck As Boolean  ' To avoid recursive checklist handling.

    Public Enum ESelOpts
        SelOpt_Default = 1
        SelOpt_List = 0
        SelOpt_OptionList = 1
        SelOpt_MultiList = 2
        SelOpt_ToItem = &H8
    End Enum

    Public Function SelectOptionArray(ByVal selTitle As String, ByVal selType As ESelOpts, ByRef selOptions() As Object, Optional ByVal SelectButtonCaption As String = "&Print", Optional ByVal PreSelChk As String = "x") 'as integer
        Dim Opt As Object
        Dim Preselected As Boolean, WithSelection As Boolean
        Dim ToItem As Boolean
        Dim FirstSelected As Integer
        Dim optSelectionArray() As Integer

        FirstSelected = -1
        ToItem = (selType And ESelOpts.SelOpt_ToItem)
        selType = selType And &H7

        mSelected = 0
        mOptionMode = selType
        Me.Text = selTitle
        cmdOk.Text = SelectButtonCaption

        ' Option mode 2 (checks) allows multiple selections, so we don't need the prompt entry.
        'If mOptionMode <> 2 Then lstSelection.AddItem "Select an Option"
        If mOptionMode <> 2 Then
            lstSelection.Items.Clear()
            lstSelection.Items.Add("Select an Option")
        End If

        For Each Opt In selOptions
            If Microsoft.VisualBasic.Left(Opt, Len(PreSelChk)) = PreSelChk Then
                Preselected = True
                Opt = Mid(Opt, Len(PreSelChk) + 1)
                WithSelection = True
            Else
                Preselected = False
            End If

            Select Case selType
                Case ESelOpts.SelOpt_List
                    'lstSelection.AddItem Opt
                    lstSelection.Items.Add(Opt)
                    If Preselected Then
                        SkipCheck = True
                        'lstSelection.Selected(lstSelection.NewIndex) = True
                        lstSelection.SetSelected(lstSelection.Items.Count - 1, True)
                        'If FirstSelected = -1 Then FirstSelected = lstSelection.NewIndex
                        If FirstSelected = -1 Then FirstSelected = lstSelection.Items.Count - 1
                        SkipCheck = False
                    End If
                Case ESelOpts.SelOpt_MultiList
                    'lstSelectionCheck.AddItem Opt
                    lstSelectionCheck.Items.Add(Opt)
                    If Preselected Then
                        SkipCheck = True
                        'lstSelectionCheck.Selected(lstSelectionCheck.NewIndex) = True
                        lstSelectionCheck.SetItemChecked(lstSelectionCheck.Items.Count - 1, True)
                        If FirstSelected = -1 Then FirstSelected = lstSelection.Items.Count - 1
                        SkipCheck = False
                    End If
                Case ESelOpts.SelOpt_OptionList
                    'Load optSelection(optSelection.UBound + 1)
                    'optSelection(optSelection.UBound).Caption = Opt
                    'If Preselected Then
                    '    SkipCheck = True
                    '    optSelection(optSelection.UBound) = True
                    '    If FirstSelected = -1 Then FirstSelected = lstSelection.NewIndex
                    '    SkipCheck = False
                    'End If

                    'NOTE: REPLACE THE ABOVVE COMMENTED BLOCK OF CODE WITH THE BELOW ONE.
                    Dim r As New RadioButton
                    Dim i As Integer

                    r.Text = Opt
                    'r.Name = "optSelection" & i
                    r.Name = "optSelection" & UBound(optSelectionArray) + 1
                    If Preselected Then
                        SkipCheck = True
                        r.Checked = True
                        If FirstSelected = -1 Then FirstSelected = lstSelection.Items.Count - 1
                        SkipCheck = False
                    End If
                    Me.Controls.Add(r)
                    ReDim optSelectionArray(UBound(optSelectionArray) + 1)
            End Select
        Next

        If Not WithSelection Then
            Select Case selType
                Case ESelOpts.SelOpt_List
                    'lstSelection.ListIndex = 1
                    lstSelection.SelectedIndex = 1
                Case ESelOpts.SelOpt_MultiList
                    'lstSelectionCheck.ListIndex = 0
                    lstSelectionCheck.SelectedIndex = 0
                Case ESelOpts.SelOpt_OptionList
                    'optSelection(1).Value = True

                    For Each c In Me.Controls
                        If c.Name = "optSelection1" Then
                            c.PerformClick()
                        End If
                    Next

            End Select
        Else
            Select Case selType
                Case ESelOpts.SelOpt_List
                    'lstSelection.ListIndex = 1
                    lstSelection.SelectedIndex = 1
                Case ESelOpts.SelOpt_MultiList
                    'lstSelectionCheck.ListIndex = 0
                    lstSelectionCheck.SelectedIndex = 0
                Case ESelOpts.SelOpt_OptionList
                    'optSelection(1).Value = True
                    For Each c In Me.Controls
                        If c.Name = "optSelection1" Then
                            c.PerformClick()
                        End If
                    Next
            End Select
        End If

        RearrangeControls()
        'Show vbModal
        ShowDialog()
        SelectOptionArray = mSelected

        If ToItem And Not selType = ESelOpts.SelOpt_MultiList Then
            On Error Resume Next
            If SelectOptionArray <= 0 Then
                SelectOptionArray = ""
            Else
                SelectOptionArray = selOptions(SelectOptionArray - 1 + LBound(selOptions))
            End If
        End If

        mSelected = 0
        mOptionMode = 0
    End Function

    'Private Function RearrangeControls()
    '    'Dim Opt As OptionButton
    '    Dim Opt As RadioButton
    '    Dim X As Integer, Y As Integer
    '    Dim TH As Integer

    '    Select Case mOptionMode
    '        Case ESelOpts.SelOpt_List                            ' Listbox
    '            'lstSelection.Move 60, 60, ScaleWidth - 120, ScaleHeight - 180 - cmdOk.Height
    '            lstSelection.Location = New Point(60, 60)
    '            lstSelection.Size = New Size(Me.ClientSize.Width - 120, Me.ClientSize.Height - 180 - cmdOk.Height)
    '            lstSelection.Visible = True
    '            Y = lstSelectionCheck.Top + lstSelectionCheck.Height + 60
    '        Case ESelOpts.SelOpt_MultiList                       ' Listbox with checks
    '            'lstSelectionCheck.Move 60, 60, ScaleWidth - 120, ScaleHeight - 180 - cmdOk.Height
    '            lstSelectionCheck.Location = New Point(60, 60)
    '            lstSelectionCheck.Size = New Size(Me.ClientSize.Width - 120, Me.ClientSize.Height - 180 - cmdOk.Height)
    '            lstSelectionCheck.Visible = True
    '            Y = lstSelectionCheck.Top + lstSelectionCheck.Height + 60
    '        Case ESelOpts.SelOpt_OptionList                      ' Option buttons
    '            'For Each Opt In optSelection
    '            '    If Opt.Index > 0 Then
    '            '        Opt.Move 60, 60 + Opt.Height * (Opt.Index - 1), ScaleWidth - 120
    '            '        Opt.Visible = True
    '            '    End If
    '            'Next

    '            'NOTE -->   THE ABOVE CODE BLOCK WILL BE REPLACED WITH THE BELOW ONE.  <--
    '            Dim rindex As Integer
    '            For Each Opt In Controls
    '                If Opt.Name Like "optSelection" Then
    '                    If Microsoft.VisualBasic.Right(Opt.Name, 1) > 0 Then
    '                        rindex = CInt(Microsoft.VisualBasic.Right(Opt.Name, 1)) - 1
    '                        Opt.Location = New Point(60, 60 + Opt.Height * rindex)
    '                        Opt.Size = New Size(ClientSize.Width - 120, ClientSize.Height)
    '                        Opt.Visible = True
    '                    End If

    '                End If
    '            Next

    '            'Y = optSelection(optSelection.UBound).Top + optSelection(0).Height + 60
    '    End Select
    '    'X = ScaleWidth / 2 - cmdOk.Width - 30
    '    'cmdOk.Move X, Y
    '    'cmdCancel.Move X + cmdOk.Width + 60, Y
    '    'TH = (Height - ScaleHeight) + cmdOk.Top + cmdOk.Height + 120
    '    If Not InRange(TH - 15, Height, TH + 15) Then Height = TH
    'End Function

    Public Function SelectOption(ByVal selTitle As String, ByVal selType As ESelOpts, ParamArray selOptions() As Object) As Object  ' As as integer
        Dim Arr() As Object
        Arr = selOptions
        SelectOption = SelectOptionArray(selTitle, selType, Arr)
    End Function

    Private Sub RearrangeControls()
        Dim Opt As RadioButton
        Dim X As Integer, Y As Integer
        Dim TH As Integer
        'Dim Ctrl As Control
        Dim OptIndex As Integer

        Select Case mOptionMode
            Case ESelOpts.SelOpt_List                            ' Listbox
                'lstSelection.Move 60, 60, ScaleWidth - 120, ScaleHeight - 180 - cmdOk.Height
                lstSelection.Location = New Point(6, 6)
                'lstSelection.Size = New Size(Width - 12, Height - 18 - cmdOk.Height)
                lstSelection.Size = New Size(Width - 33, Height - 54 - cmdOk.Height)
                lstSelection.Visible = True
                Y = lstSelectionCheck.Top + lstSelectionCheck.Height + 60
                'Y = lstSelectionCheck.Top + lstSelectionCheck.Height + 6
            Case ESelOpts.SelOpt_MultiList                       ' Listbox with checks
                'lstSelectionCheck.Move 60, 60, ScaleWidth - 120, ScaleHeight - 180 - cmdOk.Height
                lstSelectionCheck.Location = New Point(6, 6)
                'lstSelectionCheck.Size = New Size(Width - 12, Height - 18 - cmdOk.Height)
                lstSelectionCheck.Size = New Size(Width - 33, Height - 54 - cmdOk.Height)
                lstSelectionCheck.Visible = True
                Y = lstSelectionCheck.Top + lstSelectionCheck.Height + 60
                'Y = lstSelectionCheck.Top + lstSelectionCheck.Height + 6
            Case ESelOpts.SelOpt_OptionList                      ' Option buttons
                'For Each Opt In optSelection
                For Each Opt In Me.Controls
                    If Mid(Opt.Name, 1, 12) = "optSelection" Then
                        'If Opt.Index > 0 Then
                        If Len(Opt.Name) > 12 Then
                            OptIndex = Val(Mid(Opt.Name, 13))
                            If OptIndex > 0 Then
                                'Opt.Move 60, 60 + Opt.Height * (Opt.Index - 1), ScaleWidth - 120
                                Opt.Location = New Point(6, 6 + Opt.Height * OptIndex - 1)
                                Opt.Size = New Size(Width - 12, Me.Height)
                                'Opt.Visible = True
                                Opt.Visible = True
                            End If
                        End If
                    End If
                Next

                'Y = optSelection(optSelection.UBound).Top + optSelection(0).Height + 60
                For Each Opt In Me.Controls
                    If Len(Opt.Name) > 12 Then
                        OptIndex = Val(Mid(Opt.Name, 13))
                    End If
                Next
                Y = Opt.Top + optSelection.Height + 60
        End Select

        'X = ScaleWidth / 2 - cmdOk.Width - 30
        X = Width / 2 - cmdOk.Width - 3
        'cmdOk.Move X, Y
        cmdOk.Location = New Point(X, Y)
        'cmdCancel.Move X + cmdOk.Width + 60, Y
        cmdCancel.Location = New Point(X + cmdOk.Width + 6, Y)
        'TH = (Height - ScaleHeight) + cmdOk.Top + cmdOk.Height + 120
        TH = (Height - Me.ClientSize.Height) + cmdOk.Top + cmdOk.Height + 12
        If Not InRange(TH - 15, Height, TH + 15) Then Height = TH
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        mSelected = 0
        'Unload Me
        Me.Close()
    End Sub

    Private Sub cmdOk_Click(sender As Object, e As EventArgs) Handles cmdOk.Click
        Select Case mOptionMode
            Case ESelOpts.SelOpt_List
                'If lstSelection.ListIndex > 0 Then
                If lstSelection.SelectedIndex > 0 Then
                    'mSelected = lstSelection.ListIndex
                    mSelected = lstSelection.SelectedIndex
                Else
                    Exit Sub
                End If
            Case ESelOpts.SelOpt_MultiList
                Dim I As Integer, C As Integer
                'mSelected = Array()
                ReDim mSelected(C)
                'For I = 0 To lstSelectionCheck.ListCount - 1
                For I = 0 To lstSelectionCheck.Items.Count - 1
                    'If lstSelectionCheck.Selected(I) Then
                    If lstSelectionCheck.GetSelected(I) = True Then
                        ReDim Preserve mSelected(C)
                        'mSelected(C) = lstSelectionCheck.List(I)
                        mSelected(C) = lstSelectionCheck.SelectedItem
                        C = C + 1
                    End If
                Next
            Case ESelOpts.SelOpt_OptionList
                Dim Opt As RadioButton
                'For Each Opt In optSelection
                '    If Opt.Value = True Then mSelected = Opt.Index
                'Next

                For Each Opt In Me.Controls
                    If Mid(Opt.Name, 1, 12) = "optSelection" Then
                        If Len(Opt.Name) = 12 Then
                            If Opt.Checked = True Then
                                mSelected = 0
                            End If
                        ElseIf Len(Opt.Name) > 12 Then
                            If Opt.Checked = True Then
                                mSelected = Val(Mid(Opt.Name, 13))
                            End If
                        End If
                    End If
                Next
        End Select
        'Unload Me
        Me.Close()

    End Sub

    Private Sub frmSelectOption_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetButtonImage cmdOk
        'SetButtonImage cmdCancel
        'SetCustomFrame Me, ncBasicTool
        SetButtonImage(cmdOk, 2)
        SetButtonImage(cmdCancel, 3)
    End Sub

    Private Sub frmSelectOption_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        RearrangeControls()
    End Sub

    Private Function SpaceCount(ByVal strText As String) As Integer
        Dim I As Integer
        For I = 1 To Len(strText)
            If Mid(strText, I, 1) = " " Then
                SpaceCount = SpaceCount + 1
            Else
                Exit For
            End If
        Next
    End Function

    Private Sub lstSelection_DoubleClick(sender As Object, e As EventArgs) Handles lstSelection.DoubleClick
        'cmdOk.Value = True
        cmdOk_Click(cmdOk, New EventArgs)
    End Sub

    Private Sub lstSelectionCheck_MouseUp(sender As Object, e As MouseEventArgs) Handles lstSelectionCheck.MouseUp
        Dim cP As clsPopup, Res As Integer

        'If Button <> vbRightButton Then Exit Sub
        If e.Button <> e.Button.Right Then Exit Sub
        cP = New clsPopup
        cP.AddItem("  Select None")
        cP.AddItem("x Select All")
        'Res = cP.PopupMenu(hWnd)
        Res = cP.PopupMenu(Handle)
        DisposeDA(cP)

        On Error Resume Next    ' who knows what this might do...
        Dim I As Integer, J As Integer
        Select Case Res
            Case 1, 2
                'J = lstSelectionCheck.ListIndex
                J = lstSelectionCheck.SelectedIndex
                'LockWindowUpdate lstSelectionCheck.hWnd
                LockWindowUpdate(lstSelectionCheck.Handle)
                'For I = 0 To lstSelectionCheck.ListCount - 1
                For I = 0 To lstSelectionCheck.Items.Count - 1
                    'lstSelectionCheck.Selected(I) = Not (Res = 1)
                    lstSelectionCheck.SetSelected(I, Not (Res = 1))
                Next
                'lstSelectionCheck.ListIndex = J
                lstSelectionCheck.SelectedIndex = J
                LockWindowUpdate(0&)
        End Select
    End Sub

    Private Sub lstSelectionCheck_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles lstSelectionCheck.ItemCheck
        Dim Spaces As Integer, I As Integer

        If SkipCheck Then Exit Sub
        SkipCheck = True
        'If Item < 0 Then Exit Sub
        If e.Index < 0 Then Exit Sub

        'Spaces = SpaceCount(lstSelectionCheck.List(Item))
        Spaces = SpaceCount(lstSelectionCheck.SelectedItem)
        'For I = Item + 1 To lstSelectionCheck.ListCount - 1
        For I = e.Index + 1 To lstSelectionCheck.Items.Count - 1
            'If SpaceCount(lstSelectionCheck.List(I)) > Spaces Then
            If SpaceCount(lstSelectionCheck.SelectedItem) > Spaces Then
                'lstSelectionCheck.Selected(I) = lstSelectionCheck.Selected(Item)
                lstSelectionCheck.SetSelected(I, True)
            Else
                Exit For
            End If
        Next

        'For I = Item - 1 To 0 Step -1
        For I = e.Index - 1 To 0 Step -1
            'If SpaceCount(lstSelectionCheck.List(I)) < Spaces Then
            If SpaceCount(lstSelectionCheck.SelectedItem) < Spaces Then
                'Spaces = SpaceCount(lstSelectionCheck.List(I))
                Spaces = SpaceCount(lstSelectionCheck.SelectedItem)
                'lstSelectionCheck.Selected(I) = False
                lstSelectionCheck.SetSelected(I, False)
            End If
        Next
        'lstSelectionCheck.ListIndex = Item
        lstSelectionCheck.SelectedIndex = e.Index
        SkipCheck = False

    End Sub
End Class