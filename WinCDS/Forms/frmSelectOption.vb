Public Class frmSelectOption
    Private mOptionMode As ESelOpts
    Private mSelected '() as integer
    Private SkipCheck As Boolean  ' To avoid recursive checklist handling.

    Public Enum ESelOpts
        SelOpt_Default = 1
        SelOpt_List = 0
        SelOpt_OptionList = 1
        SelOpt_MultiList = 2
        SelOpt_ToItem = &H8
    End Enum
    Public Function SelectOptionArray(ByVal selTitle As String, ByVal selType As ESelOpts, ByRef selOptions() As Object, Optional ByVal SelectButtonCaption As String = "&Print", Optional ByVal PreSelChk As String = "x") 'as integer
        Dim Opt
        Dim Preselected As Boolean, WithSelection As Boolean
        Dim ToItem As Boolean
        Dim FirstSelected as integer



        FirstSelected = -1
        ToItem = (selType And ESelOpts.SelOpt_ToItem)
        selType = selType And &H7

        mSelected = 0
        mOptionMode = selType
        Me.Text = selTitle
        cmdOk.Text = SelectButtonCaption

        ' Option mode 2 (checks) allows multiple selections, so we don't need the prompt entry.
        'If mOptionMode <> 2 Then lstSelection.AddItem "Select an Option"
        If mOptionMode <> 2 Then lstSelection.Items.Add("Select an Option")

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
                    r.Name = "optSelection" & i
                    If Preselected Then
                        SkipCheck = True
                        r.Checked = True
                        If FirstSelected = -1 Then FirstSelected = lstSelection.Items.Count - 1
                        SkipCheck = False
                    End If
                    Me.Controls.Add(r)
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
    Private Function RearrangeControls()
        'Dim Opt As OptionButton
        Dim Opt As RadioButton
        Dim X as integer, Y as integer
        Dim TH as integer

        Select Case mOptionMode
            Case ESelOpts.SelOpt_List                            ' Listbox
                'lstSelection.Move 60, 60, ScaleWidth - 120, ScaleHeight - 180 - cmdOk.Height
                lstSelection.Location = New Point(60, 60)
                lstSelection.Size = New Size(Me.ClientSize.Width - 120, Me.ClientSize.Height - 180 - cmdOk.Height)
                lstSelection.Visible = True
                Y = lstSelectionCheck.Top + lstSelectionCheck.Height + 60
            Case ESelOpts.SelOpt_MultiList                       ' Listbox with checks
                'lstSelectionCheck.Move 60, 60, ScaleWidth - 120, ScaleHeight - 180 - cmdOk.Height
                lstSelectionCheck.Location = New Point(60, 60)
                lstSelectionCheck.Size = New Size(Me.ClientSize.Width - 120, Me.ClientSize.Height - 180 - cmdOk.Height)
                lstSelectionCheck.Visible = True
                Y = lstSelectionCheck.Top + lstSelectionCheck.Height + 60
            Case ESelOpts.SelOpt_OptionList                      ' Option buttons
                'For Each Opt In optSelection
                '    If Opt.Index > 0 Then
                '        Opt.Move 60, 60 + Opt.Height * (Opt.Index - 1), ScaleWidth - 120
                '        Opt.Visible = True
                '    End If
                'Next

                'NOTE -->   THE ABOVE CODE BLOCK WILL BE REPLACED WITH THE BELOW ONE.  <--
                Dim rindex As Integer
                For Each Opt In Controls
                    If Opt.name Like "optSelection" Then
                        If Microsoft.VisualBasic.Right(Opt.Name, 1) > 0 Then
                            rindex = CInt(Microsoft.VisualBasic.Right(Opt.Name, 1)) - 1
                            Opt.Location = New Point(60, 60 + Opt.Height * rindex)
                            Opt.Size = New Size(ClientSize.Width - 120, ClientSize.Height)
                            Opt.Visible = True
                        End If

                    End If
                Next

                'Y = optSelection(optSelection.UBound).Top + optSelection(0).Height + 60
        End Select
        'X = ScaleWidth / 2 - cmdOk.Width - 30
        'cmdOk.Move X, Y
        'cmdCancel.Move X + cmdOk.Width + 60, Y
        'TH = (Height - ScaleHeight) + cmdOk.Top + cmdOk.Height + 120
        If Not InRange(TH - 15, Height, TH + 15) Then Height = TH
    End Function

End Class