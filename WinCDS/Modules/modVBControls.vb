Imports System.Windows
Module modVBControls
    Public Function SelectContents(ByRef txtBox As Object) As Boolean
        On Error Resume Next
        'txtBox.SelStart = 0
        'txtBox.SelLength = Len(txtBox.Text)

        txtBox.SelectionStart = 0
        txtBox.SelectionLength = Len(txtBox.Text)
        SelectContents = True
    End Function

    Public Function MoveControl(ByRef C As Object, Optional ByVal X As Integer = 0, Optional ByVal Y As Integer = -10001, Optional ByVal W As Integer = -10001, Optional ByVal H As Integer = -10001, Optional ByVal MakeVisible As Boolean = False, Optional ByVal ZOrderTop As Boolean = False) As Boolean
        Const NonVal As Integer = -10001
        On Error Resume Next
        If TypeName(C) = "Line" Then
            C.X1 = X
            C.Y1 = Y
            C.X2 = W
            C.Y2 = H
        Else
            If Y <> NonVal Then
                If W <> NonVal Then
                    If H <> NonVal Then
                        C.Move(X, Y, W, H)
                    Else
                        C.Move(X, Y, W)
                    End If
                Else
                    C.Move(X, Y)
                End If
            Else
                C.Move
            End If
        End If

        If MakeVisible Then C.Visible = True
        If ZOrderTop Then C.ZOrder(0)
    End Function
    Public Function ColorDatePicker(ByRef dtp As DateTimePicker)
        With dtp
            '.CalendarBackColor = vbWhite
            .CalendarForeColor = Color.Black
            .CalendarTitleBackColor = Color.Blue
            .CalendarTitleForeColor = Color.Black
            .CalendarTrailingForeColor = Color.Blue
        End With
    End Function

    Public Function SelectComboBoxItemData(ByRef Cbo As ComboBox, ByVal iData As Integer) As Boolean
        Dim I As Integer
        'For I = 0 To Cbo.Items.Count - 1
        '    If Cbo.itemData(I) = iData Then
        '        Cbo.ListIndex = I
        '        SelectComboBoxItemData = True
        '        Exit Function
        '    End If
        'Next

        For I = 0 To Cbo.Items.Count - 1
            If CType(Cbo.Items(I), ItemDataClass).ItemData = iData Then
                Cbo.SelectedIndex = I
                SelectComboBoxItemData = True
                Exit Function
            End If
        Next
    End Function

    Public Function AddItemToComboBox(ByVal Cbo As ComboBox, ByVal Item As String, Optional ByVal itemData As Integer = 0, Optional ByVal vIndex As Integer = -1) As Boolean
        AddItemToComboBox = ComboBoxAdd(Cbo, Item, itemData, vIndex) <> -1
    End Function

    Public Function ComboBoxAdd(ByVal Cbo As ComboBox, ByVal Item As String, Optional ByVal itemData As Integer = 0, Optional ByVal vIndex As Integer = -1) As Integer
        'Dim idc As ItemDataClass
        On Error GoTo Fail
        'idc = New ItemDataClass(Item, itemData)
        If vIndex = -1 Then
            'Cbo.AddItem Item
            If itemData <> 0 Then
                Cbo.Items.Add(New ItemDataClass(Item, itemData))
            Else
                Cbo.Items.Add(Item)
            End If
        Else
            'Cbo.AddItem(Item, vIndex)
            If itemData <> 0 Then
                Cbo.Items.Insert(vIndex, New ItemDataClass(Item, itemData))
            Else
                Cbo.Items.Insert(vIndex, Item)
            End If
        End If

        ComboBoxAdd = Cbo.Items.Count - 1
        Exit Function
Fail:
        ComboBoxAdd = -1
    End Function

    Public Function AddControlToForm(ByVal ClassName, Optional ByVal F = Nothing, Optional ByRef CtrlName = "") As Object
        On Error Resume Next
        If IsNothing(F) Then F = MainMenu
        If CtrlName = "" Then CtrlName = "ctrlAdded" & Second(Now) & "_" & Random(1000)
        Do While FormHasControl(F, CtrlName)
            CtrlName = CtrlName & "a"
        Loop
        'AddControlToForm = F.Controls.Add(ClassName, CtrlName)
        AddControlToForm = F.Controls.Add(ClassName)
    End Function

    Public Function FormHasControl(ByVal F As Form, ByVal cName As String) As Boolean
        FormHasControl = IsNotNothing(FormControl(F, cName))
    End Function

    Public Function FormControl(ByVal F As Form, ByVal cName As String) As Control
        Dim L
        For Each L In F.Controls
            If L.Name = cName Then FormControl = L : Exit Function
        Next
    End Function

    Public Function CenterForm(ByRef F) As Boolean
        Dim vF As Form
        On Error Resume Next
        vF = SelectForm(F)
        If IsNothing(vF) Then Exit Function
        'vF.Move(Screen.Width - vF.Width) / 2, (Screen.Height - vF.Height) / 2
        vF.Location = New Point((Screen.PrimaryScreen.Bounds.Width - vF.Width) / 2, (Screen.PrimaryScreen.Bounds.Height - vF.Height) / 2)
        CenterForm = True
    End Function

    Public Function SelectForm(ByRef F)
        Dim V As Integer
        On Error Resume Next
        If False Then
            'ElseIf IsObject(F) Then
        ElseIf F IsNot Nothing Then
            SelectForm = F
        ElseIf IsNumeric(F) Then
            'V = FitRange(0, Val(F), Forms.Count - 1)
            V = FitRange(0, Val(F), My.Application.OpenForms.Count - 1)
            'SelectForm = Forms(V)
            SelectForm = My.Application.OpenForms.Item(V)
        ElseIf TypeName(F) = "String" Then
            'For V = 0 To Forms.Count - 1
            For V = 0 To My.Application.OpenForms.Count - 1
                'If LCase(Forms(V).Name) = LCase(F) Then SelectForm = Forms(V) : Exit Function
                If LCase(My.Application.OpenForms.Item(V).Name) = LCase(F) Then SelectForm = My.Application.OpenForms.Item(V) : Exit Function
            Next
            'For V = 0 To Forms.Count - 1
            For V = 0 To My.Application.OpenForms.Count - 1
                If Left(LCase(My.Application.OpenForms.Item(V).Name), Len(F)) = LCase(F) Then SelectForm = My.Application.OpenForms.Item(V) : Exit Function
            Next
        End If
    End Function

    Public Sub EnableFrame(ByRef frm As Form, ByRef Fra As GroupBox, ByRef Enabled As Boolean, Optional ByRef ParentFrame As GroupBox = Nothing)
        Dim C As Control, OK As Boolean
        On Error Resume Next
        Fra.Enabled = Enabled
        For Each C In frm.Controls
            If ParentFrame Is Nothing Then
                'OK = (C.Container.Name = Fra.Name)
                OK = (C.Container Is Fra.Name)
            Else
                'OK = (C.Container = Fra) Or (C.Container = ParentFrame)  ' doesn't seem to work..
                OK = (C.Container Is Fra) Or (C.Container Is ParentFrame)  ' doesn't seem to work..
            End If
            If Err.Number <> 0 Then
                OK = False
                Err.Clear()
            End If
            If OK Then
                C.Enabled = Enabled
                'If TypeName(C) = "Frame" Then EnableFrame Frm, C, Enabled, fra
            End If
        Next
    End Sub

    Public Function FocusControl(ByRef C As Object) As Boolean
        On Error Resume Next    ' SetFocus can hard-fail..  This protects it without adding error handling in your sub
        C.Select
    End Function

    Public Function FocusSelect(ByRef txtBox As Object) As Boolean
        On Error Resume Next
        'txtBox.SetFocus
        txtBox.Select
        FocusSelect = SelectContents(txtBox)
    End Function

    Public Function MoveControlTo(ByRef C As Object, ByRef D As Object, Optional ByVal MakeVisible As Boolean = False, Optional ByVal ZOrderTop As Boolean = False) As Boolean
        MoveControl(C, D.Left, D.Top, D.Width, D.Height, MakeVisible, ZOrderTop)
    End Function

End Module
