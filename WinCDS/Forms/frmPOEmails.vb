Public Class frmPOEmails
    Private mSType as integer

    Public Property sType() as integer
        Get
            sType = mSType
        End Get
        Set(value as integer)
            mSType = value
            fraSelect.Text = Microsoft.VisualBasic.Interaction.Switch(sType = 0, "Not Acknowledged:", True, "Overdue Orders:")
            dtpRunAsDate.Value = IIf(sType = 0, DateAdd("d", -10, Today), Today)
            cmdEditTemplate.Visible = mSType = 0 Or mSType = 1
            RefreshSelect
        End Set
    End Property

    Private Sub RefreshSelect()
        Dim S As String, RS As ADODB.Recordset
        Dim Ln As String, OK As Boolean

        lstSelect.Items.Clear()

        Select Case sType
            Case 0
                S = "SELECT DISTINCT PoNo, Vendor FROM [Po] WHERE PoDate>=#" & dtpRunAsDate.Value & "# AND IIF(IsNull(AckInv),'',AckInv)='' AND Posted<>'X' AND PrintPO<>'V' ORDER BY [PoNo]"
            Case 1
                S = "SELECT DISTINCT PoNo, Vendor FROM [Po] WHERE NOT IsNull(DueDate) AND DueDate < #" & dtpRunAsDate.Value & "# AND Posted<>'X' ORDER BY [PoNo]"
            Case Else
        End Select

        RS = GetRecordsetBySQL(S, , GetDatabaseInventory)

        Do While Not RS.EOF
            Dim N As String, Em As String
            Em = ""     ' make sure there's a valid vendor email address
            GetVendorName(IfNullThenNilString(RS("Vendor").Value), N, , , , , , , , Em)
            OK = (Em <> "")

            Ln = IIf(OK, "", "*") & AlignString(RS("PoNo").Value, 10, VBRUN.AlignConstants.vbAlignLeft) & " " & IfNullThenNilString(RS("Vendor").Value)


            'lstSelect.AddItem Ln
            lstSelect.Items.Add(Ln)
            'lstSelect.Selected(lstSelect.NewIndex) = OK
            lstSelect.SetSelected(lstSelect.Items.Count - 1, OK)
            RS.MoveNext
        Loop

        DisposeDA(RS)
    End Sub

End Class