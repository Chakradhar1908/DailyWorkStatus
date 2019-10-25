Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class UGridIO
#Const BackColors = False
#Const BackColorsInDevelopment = False

    Private ColoringColumns As Boolean

    Private mMaxCols As Integer           'Number of columns
    Private mMaxRows As Integer              'Number of rows

    Private mLoading As Integer              'Count of things moving the grid around without wanting events fired.
    Private Const J As Integer = 1
    Private mRefresh As Boolean
    Private mActivated As Boolean

    Private RowColChange_Active As Boolean

    Dim GridArray(,) As String      'Place to store the data
    Dim ColumnColors() As Integer

    Private mLastRow As Object
    Private mLastCol As Object
    Private mLostFocus As Boolean
    Private mCurrentCellModified As Boolean
    Private mKeyPressed As Boolean
    Public PrevProc As Integer

    Public Event AfterColEdit(ByVal ColIndex As Integer)
    Public Event AfterColUpdate(ByVal ColIndex As Integer)
    Public Event AfterDelete()
    Public Event AfterInsert()
    Public Event AfterUpdate()
    Public Event BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Public Event BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Object, Cancel As Integer)
    Public Event BeforeDelete(Cancel As Integer)
    Public Event BeforeInsert(Cancel As Integer)
    Public Event BeforeUpdate(Cancel As Integer)
    Public Event ButtonClick(ByVal ColIndex As Integer)
    Public Event ColEdit(ByVal ColIndex As Integer)
    'Public Event LostFocus()
    Public Event OnAddNew()
    Public Event RowColChange(LastRow As Object, LastCol As Object, newRow As Object, newCol As Object, ByRef Cancel As Boolean)
    Public Event RowDelete(LastRow As Object)
    Public Event Key_Press(KeyAscii As Integer)
    Public Event ClickEvent()
    Public Event Change()
    Public Event DblClick()
    Public Event SelChange(Cancel As Integer)
    Public Event TabKey()
    Public Event ValidateEvent(Cancel As Boolean)
    Public Event HeadClick(ByVal ColIndex As Integer)
    Public Event MouseMoveOverCell(ByVal Col As Integer, ByVal Row As Integer)
    Public Event Mouse_Move(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event MM(Cancel As Integer)
    Dim uhwnd As IntPtr

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        If MaxCols = 0 Then MaxCols = 2
        If MaxRows = 0 Then MaxRows = 10
        'RowColChange_Active = True
        ConnectData()
        'MessageBox.Show(AxDataGrid1.FirstRow & "-" & AxDataGrid1.Bookmark & "-" & AxDataGrid1.Text)
        AxDataGrid1.Row = 0
        AxDataGrid1.Col = 1
        AxDataGrid1.Text = ""
    End Sub

    Public Sub ConnectData()
        Dim c As New ADODB.Connection

        c.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\CDSData\Store1\NewOrder\cdsdata.mdb"
        c.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        c.Open()
        'MsgBox(c.State)
        'MessageBox.Show(c.State)
        Dim r As New ADODB.Recordset
        r.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        r.Open("select * from temptable", c, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
        AxDataGrid1.DataSource = r
        'c.Close()
    End Sub

    Public ReadOnly Property Hwnd() As Integer
        Get
            uhwnd = Me.Handle
            Return uhwnd
        End Get
    End Property

    Public Property MaxCols() As Integer
        Get
            Return mMaxCols
        End Get
        Set(value As Integer)
            mMaxCols = value
            ReDim Preserve ColumnColors(mMaxCols)
        End Set
    End Property

    Public Property MaxRows() As Integer
        Get
            Return mMaxRows
        End Get
        Set(value As Integer)
            mMaxRows = value
            On Error Resume Next
            'AxDataGrid1.ApproxCount=value

            ' This can error if GridArray hasn't been initialized yet, and that's okay.
            If value <> UBound(GridArray, 2) Then
                ReDim Preserve GridArray(UBound(GridArray, 1), value)
            End If
        End Set
    End Property

    '------> NOTE: HELPCONTEXTID PROPERTY IS NOT AVAILABLE IN .NET  <-----------
    'Public Property GridHelpContextID() As Integer
    'Get
    '    '        GridHelpContextID = DBGrid1.HelpContextID

    'End Get
    '    '    Set(value as integer)
    '    '        DBGrid1.HelpContextID = value
    '    '        'PropertyChanged("GridHelpContextID")
    '    '    End Set
    'End Property

    Public Property Text() As String
        Get
            Text = AxDataGrid1.Text
        End Get
        Set(value As String)
            AxDataGrid1.Text = value
        End Set
    End Property

    Public Sub AddColumn(ColOrder As Integer, Title As String, ColWidth As Integer, IsLocked As Boolean, AllowSizing As Boolean, Optional Align As Integer = 0, Optional IsVisible As Boolean = True)
        'Dim Col As MSDBGrid.Column
        Dim Col As MSDataGridLib.Column
        Col = GetColumn(ColOrder)
        With Col
            .Caption = Title
            .Width = ColWidth
            .AllowSizing = AllowSizing
            .Alignment = Align
            .Visible = IsVisible
            .Locked = IsLocked
        End With
    End Sub

    Public Function GetColumn(Index As Integer) As MSDataGridLib.Column
        On Error Resume Next

        With AxDataGrid1
            'Dim Cols As MSDBGrid.Columns
            'Set Cols = .Columns
            '   Debug.Print cols.Count
            If .Columns.Count = Index Then
                .Columns.Add(Index)
                .Columns(Index).Caption = "Column " & Index + 1
                .Columns(Index).Visible = True
            End If
            GetColumn = .Columns(Index)
            'Set Cols = Nothing
        End With
    End Function

    Public Sub Initialize()
        ReDim Preserve GridArray(0 To mMaxCols - 1, 0 To mMaxRows - 1)
        RowColChange_Active = True
        mRefresh = True
    End Sub

    Public Function GetDBGrid() As AxMSDataGridLib.AxDataGrid
        GetDBGrid = AxDataGrid1
        Return GetDBGrid
    End Function

    'Public Function GetDBGrid() As AxMSDBGrid.AxDBGrid
    '    GetDBGrid = AxDataGrid1
    'End Function
    Public Sub Clear()
        Dim C As Integer, R As Integer

        AxDataGrid1.ClearFields()

        On Error Resume Next
        For R = 0 To mMaxRows - 1
            For C = 0 To mMaxCols - 1
                SetValue(R, C, "")
            Next
        Next
        Refresh()
    End Sub

    Public Property Activated() As Boolean
        Get
            Return mActivated
        End Get
        Set(value As Boolean)
            mActivated = value
        End Set
    End Property

    Private Sub UGridIO_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        'AxDataGrid1.Left = 0
        'AxDataGrid1.Top = 0
        'AxDataGrid1.Width = Me.Width
        'AxDataGrid1.Height = Me.Height
        AxDataGrid1.Location = New Point(0, 0)
        AxDataGrid1.Size = New Size(Width, Height)
    End Sub

    Private Sub UGridIO_Enter(sender As Object, e As EventArgs) Handles MyBase.Enter
        On Error Resume Next
        AxDataGrid1.Select()
    End Sub

    'NOTE::: This visiblechanged event is for both usercontrol_hide and usercontrol_show events of vb 6.0.
    Private Sub UGridIO_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        'This visiblechanged event is for both usercontrol_hide and usercontrol_show events of vb 6.0.
        If Me.Visible = True Then
#If BackColors Then
  #If Not BackColorsInDevelopment Then
  If Not IsDevelopment Then
  #End If
    HookForm
    ColorColumns
  #If Not BackColorsInDevelopment Then
  End If
  #End If
#End If
        Else
#If BackColors Then
  #If Not BackColorsInDevelopment Then
  If Not IsDevelopment Then
  #End If
    UnHookForm
  #If Not BackColorsInDevelopment Then
  End If
  #End If
#End If

        End If
    End Sub

    Private Sub UGridIO_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        ' Debug.Print "UserControl_KeyPress " & KeyAscii
    End Sub

    Public Sub Refresh(Optional PreserveRow As Boolean = False)
        Dim Col As Integer
        Dim Row As Integer
        Dim Fro As Object

        On Error Resume Next
        'With AxDataGrid1
        With AxDataGrid1
            Col = .Col
            Row = .Row
            Fro = .FirstRow - 1

            RowColChange_Active = False
            ' mRefresh = True
            ' Update the current field in the array, if needed.
            If CurrentCellModified() Then
                ' This is good, but not good enough.
                ' We need to update the row, if not the whole grid.
                SetValue(Val(Fro) + Row, Col, .Text)
            End If
            .Refresh()
            'DBGrid1UnboundReadData
            '       mRefresh = False
            '     .col = col
            If (PreserveRow) Then
                If Row <> -1 And .Row <> (Val(Fro) + Row) Then
                    .FirstRow = Fro + 1
                    .Row = Row
                    'Debug.Print "Moving grid to row " & Fro & "+" & Row & "."
                End If
            End If
            '     If (row <> -1) Then .row = row
            RowColChange_Active = True
        End With
    End Sub

    Public Function CurrentCellModified() As Boolean
        CurrentCellModified = mCurrentCellModified
    End Function

    Public Sub SetValue(ByVal Row As Integer, ByVal Col As Integer, ByVal Value As String)
        ' Removed redim call MJK20030521 because it's unnecessary and slow.
        '    If UBound(GridArray, 1) <= Col - 1 Or UBound(GridArray, 2) <= Row - 1 Then
        '      ReDim Preserve GridArray(0 To mMaxCols - 1, 0 To mMaxRows - 1) As String
        '    End If
        ' col = 7
        'On Error Resume Next
        GridArray(Col, Row) = Value
        If Col = Me.Col And Row = Me.Row Then mCurrentCellModified = False
    End Sub

    Public Property Col() As Integer
        Get
            Return AxDataGrid1.Col
        End Get
        Set(value As Integer)
            AxDataGrid1.Col = value
        End Set
    End Property

    Public Property Row() As Integer
        Get
            On Error Resume Next
            'Return Row = AxDataGrid1.Row + IIf(AxDataGrid1.FirstRow = "", 0, AxDataGrid1.FirstRow)
            'Return Row = AxDataGrid1.Row + IIf(AxDataGrid1.FirstRow = "", 0, AxDataGrid1.FirstRow)
            Row = AxDataGrid1.Row
            Return Row
        End Get
        Set(value As Integer)
            On Error Resume Next
            'Dim i As Integer
            'AxDataGrid1.ApproxCount
            MakeRowVisible(value)  ' Debug this before distributing..
            'AxDataGrid1.Row = value - AxDataGrid1.FirstRow
            AxDataGrid1.Row = value - (AxDataGrid1.FirstRow - 1)
        End Set
    End Property

    Public Sub MakeRowVisible(ByRef RowNum As Integer)
        Dim FirstRow As Integer
        FirstRow = AxDataGrid1.FirstRow
        If FirstRow > 0 Then FirstRow = 0
        With AxDataGrid1
            If FirstRow > RowNum Then
                .FirstRow = RowNum - .VisibleRows + 1 ' Move up
                'ElseIf .FirstRow + .VisibleRows <= RowNum Then
            ElseIf FirstRow + .VisibleRows <= RowNum Then
                .FirstRow = RowNum   ' Move down
            End If
        End With
    End Sub

    Public Function LastRowUsed(Optional BlankLinesAllowed As Integer = -1) As Integer
        Dim I As Integer, J As Integer, HasData As Boolean, BlankCount As Integer

        For I = 0 To MaxRows - 1
            HasData = False
            For J = 0 To MaxCols - 1
                If GetValue(I, J) <> "" Then HasData = True : Exit For ' mark and move to the next row
            Next
            If HasData = True Then
                LastRowUsed = I
            Else
                If BlankLinesAllowed > 0 Then
                    BlankCount = BlankCount + 1
                    If BlankCount > BlankLinesAllowed Then Exit Function
                End If
            End If
        Next
    End Function

    Public Function GetValue(ByVal Row As Integer, ByVal Col As Integer) As String
        On Error Resume Next
        GetValue = GridArray(Col, Row)
    End Function

    Public Function GetValueDisplay(Row As Integer, Col As Integer) As String
        On Error Resume Next
        '   With DBGrid1
        '   End With
        GetValueDisplay = GridArray(Col, Row)
    End Function

    Public Sub SetValueDisplay(Row As Integer, Col As Integer, Value As String)
        'Debug.Print "SetValueDisplay: " & " " & row & " " & col & " " & value & " ";

        SetValue(Row, Col, Value)

        'With AxDataGrid1
        With AxDataGrid1
            Dim BM As Object, OldCol As Integer
            On Error Resume Next
            BM = .Bookmark
            OldCol = .Col

            'Dim TopRow As Object, BottomRow As Object
            Dim TopRow As Integer, BottomRow As Integer
            TopRow = .FirstRow - 1
            'If TopRow = 1 Then TopRow = 0
            If (.VisibleRows = 0) Then
                BottomRow = 0
            Else
                BottomRow = .RowBookmark(.VisibleRows - 1)
                'If (.VisibleRows <> 0) Then BottomRow = .RowBookmark(.VisibleRows - 1)
            End If
            '        Debug.Print "A: " & DBGrid1.Bookmark & ", " & bm & ", " & TopRow & ", " & BottomRow
            '      StoreUserData row, col, value
            .Col = Col
            '        If (row < items_per_page) Then
            If (TopRow <= Row) And (Row <= BottomRow) Then
                '       .row = row
                On Error GoTo softError
                On Error Resume Next
                '.Bookmark = Str(Row)
                '.Bookmark = .Bookmark + 1
                .Row = Row

                On Error GoTo softError
                On Error Resume Next
                '.Row = 1
                If Value = ".00" Then Value = ""
                .Text = Value
                .FirstRow = TopRow + 1
            End If
            On Error GoTo AnError
            'bm = .Bookmark
            On Error Resume Next
            '.Bookmark = BM
            .Col = OldCol
        End With
        '      Debug.Print "B: " & DBGrid1.Bookmark & ", " & bm & ", " & TopRow & ", " & BottomRow
        Exit Sub
softError:
        '    Debug.Print "C: " & DBGrid1.Bookmark & ", " & bm & ", " & TopRow & ", " & BottomRow
        MessageBox.Show("SetValueDisplay Bookmark error: Row=" & Row)
        ' Resume Next
        Exit Sub
AnError:
        '   Debug.Print "B: " & DBGrid1.Bookmark & ", " & bm & ", " & TopRow & ", " & BottomRow
        MessageBox.Show("SetValueDisplay Bookmark error: Row=" & Row)
        ' Resume Next
        Exit Sub
    End Sub

    Public Sub SetValueDisplayNew(Row As Integer, Col As Integer, Value As String)
        SetValue(Row, Col, Value)
        Dim vRow As Object
        vRow = Format(Row, "#")

        With AxDataGrid1
            Dim TopRow As Object, BottomRow As Object
            TopRow = .FirstRow
            BottomRow = .RowBookmark(.VisibleRows - 1)
            Dim S As Object
            Dim T As Object

            T = .Bookmark
            S = .FirstRow
            StoreUserData(Row, Col, Value)
            If (Val(TopRow) <= Val(vRow)) And (Val(vRow) <= Val(BottomRow)) Then
                .Bookmark = vRow
                .Bookmark = T
                '    .FirstRow = s:
            End If
            '  If (Val(TopRow) <= Val(s)) And (Val(s) <= Val(BottomRow)) Then
            '.FirstRow = s:
        End With
    End Sub

    Private Function StoreUserData(bookm As Object, colm As Integer, userval As Object) As Boolean
        Dim Index As Integer

        Index = IndexFromBookmark(bookm, False)
        If Index < 0 Or Index >= mMaxRows Or colm < 0 Or colm >= mMaxCols Then
            StoreUserData = False
        Else
            StoreUserData = True
            GridArray(colm, Index) = userval
            If bookm = Row And colm = Col Then mCurrentCellModified = False
        End If
    End Function

    Private Function IndexFromBookmark(bookm As Object, ReadPriorRows As Boolean) As Integer
        'If IsNull(bookm) Then
        If bookm Is Nothing Then
            If ReadPriorRows Then
                IndexFromBookmark = mMaxRows
            Else
                IndexFromBookmark = -1
            End If
        Else
            Dim Index As Integer
            Index = Val(bookm)
            If Index < 0 Or Index >= mMaxRows Then Index = -2000
            IndexFromBookmark = Index
        End If
    End Function

    Private Function GetRelativeBookmark(bookm As Object, relpos As Integer) As Object
        Dim Index As Integer

        Index = IndexFromBookmark(bookm, False) + relpos
        If Index < 0 Or Index >= mMaxRows Then
            GetRelativeBookmark = Nothing
        Else
            GetRelativeBookmark = MakeBookmark(Index)
        End If
    End Function

    Private Function MakeBookmark(Index As Integer) As Object
        MakeBookmark = Str(Index)
    End Function

    Private Function GetNewBookmark() As Object
        ReDim Preserve GridArray(0 To mMaxCols - 1, 0 To mMaxRows)
        GetNewBookmark = MakeBookmark(mMaxRows)
        mMaxRows = mMaxRows + 1
    End Function

    Public Function GetUserData(bookm As Object, colm As Integer) As Object
        Dim Index As Integer

        Index = IndexFromBookmark(bookm, False)
        If Index < 0 Or Index >= mMaxRows Or colm < 0 Or colm >= mMaxCols Then
            GetUserData = Nothing
        Else
            GetUserData = GridArray(colm, Index)
        End If
    End Function

    Public Property Loading() As Boolean
        Get
            Loading = mLoading > 0
        End Get
        Set(value As Boolean)
            'mLoading = mLoading + IIf(value, 1, -1)  ' Increment or decrement the counter.
            If value = True Then
                'mLoading = mLoading + J
                mLoading = 1
            Else
                'mLoading = mLoading - 1
                mLoading = 0
            End If
        End Set
    End Property

    Public ReadOnly Property LostFocusFlag() As Boolean
        Get
            Return LostFocusFlag = mLostFocus
        End Get
    End Property

    Public Sub MoveRowDown(Optional Value As Integer = 1)
        'AxDataGrid1.Scroll(0, Value)
        AxDataGrid1.Scroll(0, Value)
    End Sub

    Public Sub MoveRowUp(Optional Value As Integer = 1)
        'AxDataGrid1.Scroll(0, -Value)
        AxDataGrid1.Scroll(0, -Value)
    End Sub

    'Private Sub UserControl_Initialize()
    ' Intialize event not there in VB.net. Use sub new.
    '    If (MaxCols = 0) Then MaxCols = 2
    '    If (MaxRows = 0) Then MaxRows = 10
    'End Sub

    'NOTE -----------------------AfterColEdit is not required. It is not used in BillOSale form in vb6.0----------------------
    Private Sub AxDataGrid1_AfterColEdit(sender As Object, e As AxMSDataGridLib.DDataGridEvents_AfterColEditEvent) Handles AxDataGrid1.AfterColEdit
        RaiseEvent AfterColEdit(e.colIndex)
    End Sub

    'NOTE -----------------------AfterColUpdate is not required. It is not used in BillOSale form in vb6.0----------------------
    Private Sub AxDataGrid1_AfterColUpdate(sender As Object, e As AxMSDataGridLib.DDataGridEvents_AfterColUpdateEvent) Handles AxDataGrid1.AfterColUpdate
        RaiseEvent AfterColUpdate(e.colIndex)
    End Sub

    Private Sub AxDataGrid1_AfterDelete(sender As Object, e As EventArgs) Handles AxDataGrid1.AfterDelete
        RaiseEvent AfterDelete()
    End Sub

    Private Sub AxDataGrid1_AfterInsert(sender As Object, e As EventArgs) Handles AxDataGrid1.AfterInsert
        RaiseEvent AfterInsert()
    End Sub

    Private Sub AxDataGrid1_AfterUpdate(sender As Object, e As EventArgs) Handles AxDataGrid1.AfterUpdate
        RaiseEvent AfterUpdate()
    End Sub

    Private Sub AxDataGrid1_BeforeColEdit(sender As Object, e As AxMSDataGridLib.DDataGridEvents_BeforeColEditEvent) Handles AxDataGrid1.BeforeColEdit
        RaiseEvent BeforeColEdit(e.colIndex, e.keyAscii, e.cancel)
    End Sub

    Private Sub AxDataGrid1_BeforeColUpdate(sender As Object, e As AxMSDataGridLib.DDataGridEvents_BeforeColUpdateEvent) Handles AxDataGrid1.BeforeColUpdate
        RaiseEvent BeforeColUpdate(e.colIndex, e.oldValue, e.cancel)
        If Not e.cancel Then mCurrentCellModified = True
    End Sub

    Private Sub AxDataGrid1_BeforeDelete(sender As Object, e As AxMSDataGridLib.DDataGridEvents_BeforeDeleteEvent) Handles AxDataGrid1.BeforeDelete
        RaiseEvent BeforeDelete(e.cancel)
    End Sub

    Private Sub AxDataGrid1_BeforeInsert(sender As Object, e As AxMSDataGridLib.DDataGridEvents_BeforeInsertEvent) Handles AxDataGrid1.BeforeInsert
        RaiseEvent BeforeInsert(e.cancel)
    End Sub

    Private Sub AxDataGrid1_BeforeUpdate(sender As Object, e As AxMSDataGridLib.DDataGridEvents_BeforeUpdateEvent) Handles AxDataGrid1.BeforeUpdate
        RaiseEvent BeforeUpdate(e.cancel)
    End Sub

    Private Sub AxDataGrid1_ButtonClick(sender As Object, e As AxMSDataGridLib.DDataGridEvents_ButtonClickEvent) Handles AxDataGrid1.ButtonClick
        RaiseEvent ButtonClick(e.colIndex)
    End Sub

    Private Sub AxDataGrid1_Change(sender As Object, e As EventArgs) Handles AxDataGrid1.Change
        RaiseEvent Change()
    End Sub

    'Private Sub AxDataGrid1_ClickEvent(sender As Object, e As EventArgs) Handles AxDataGrid1.ClickEvent
    '    mKeyPressed = False
    '    RaiseEvent ClickEvent()
    'End Sub

    Private Sub AxDataGrid1_ColEdit(sender As Object, e As AxMSDataGridLib.DDataGridEvents_ColEditEvent) Handles AxDataGrid1.ColEdit
        RaiseEvent ColEdit(e.colIndex)
    End Sub

    'Private Sub AxDataGrid1_HeadClick(sender As Object, e As AxMSDataGridLib.DDataGridEvents_HeadClickEvent) Handles AxDataGrid1.HeadClick
    '    RaiseEvent HeadClick(e.colIndex)
    'End Sub

    Private Sub AxDataGrid1_KeyDownEvent(sender As Object, e As AxMSDataGridLib.DDataGridEvents_KeyDownEvent) Handles AxDataGrid1.KeyDownEvent
        mKeyPressed = (e.keyCode = 9)
    End Sub

    Private Sub AxDataGrid1_KeyUpEvent(sender As Object, e As AxMSDataGridLib.DDataGridEvents_KeyUpEvent) Handles AxDataGrid1.KeyUpEvent
        If e.keyCode = 9 And mKeyPressed Then
            ProcessTabKeys()  ' Replaces timer object, MJK 20030414
            'AxDataGrid1_RowColChange(AxDataGrid1, New AxMSDataGridLib.DDataGridEvents_RowColChangeEvent(AxDataGrid1.Row, AxDataGrid1.Col))
        End If
    End Sub

    Private Sub AxDataGrid1_KeyPressEvent(sender As Object, e As AxMSDataGridLib.DDataGridEvents_KeyPressEvent) Handles AxDataGrid1.KeyPressEvent
        RaiseEvent Key_Press(e.keyAscii)
        mKeyPressed = True
    End Sub

    Private Sub ProcessTabKeys()
        '-----------> NOTE: ProcessTabKey name has been changed to ProcessTabKeys. Because ProcessTabKey is a keyword in vb.net <------------

        ' This is fired from the KeyUp event, and only if
        ' the tab key was properly hit.
        'Dim c As MSDBGrid.Columns = DBGrid1.get_Columns

        RaiseEvent TabKey()

        If CurrentCellModified() Then
            ' Move to next cell.
            mCurrentCellModified = False
            Dim I As Integer, ColChanged As Boolean
            For I = Col + 1 To MaxCols - 1
                'If DBGrid1.Columns(I).Width > 0 And DBGrid1.Columns(I).Visible Then Col = Col + 1 : ColChanged = True : Exit For
                'If c.Columns(I).Width > 0 And c.Columns(I).Visible Then Col = Col + 1 : ColChanged = True : Exit For
                If AxDataGrid1.Columns(I).Width > 0 And AxDataGrid1.Columns(I).Visible Then Col = Col + 1 : ColChanged = True : Exit For
            Next
            If Not ColChanged Then
                Row = Row + 1
                Col = 0
            End If
        End If
    End Sub

    'Private Sub AxDataGrid1_MouseMoveEvent(sender As Object, e As AxMSDataGridLib.DDataGridEvents_MouseMoveEvent) Handles AxDataGrid1.MouseMoveEvent
    '    RaiseEvent Mouse_Move(e.button, e.shift, e.x, e.y)
    '    RaiseEvent MouseMoveOverCell(AxDataGrid1.ColContaining(e.x), AxDataGrid1.RowContaining(e.y))
    'End Sub

    Private Sub AxDataGrid1_OnAddNew(sender As Object, e As EventArgs) Handles AxDataGrid1.OnAddNew
        RaiseEvent OnAddNew()
    End Sub

    Private Sub AxDataGrid1_SelChange(sender As Object, e As AxMSDataGridLib.DDataGridEvents_SelChangeEvent) Handles AxDataGrid1.SelChange
        RaiseEvent SelChange(e.cancel)
    End Sub

    Private Sub AxDataGrid1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles AxDataGrid1.Validating
        RaiseEvent ValidateEvent(e.Cancel)
    End Sub



    '----> NOTE: UNBOUNDREADDATA EVENT IS NOT AVAILABLE IN VB.NET FOR AXDATAGRID CONTROL.   <-------
    'Private Sub DBGrid1_UnboundReadData(sender As Object, e As AxMSDBGrid.DBGridEvents_UnboundReadDataEvent)
    '    'Debug.Print "DBGrid1_UnboundReadData: ", RowBuf.RowCount, StartLocation, ReadPriorRows
    '    If Not mRefresh Then Exit Sub

    '    Dim bookm As Object
    '    bookm = e.startLocation

    '    Dim relpos As Integer
    '    If e.readPriorRows Then relpos = -1 Else relpos = 1

    '    Dim rowsFetched As Integer
    '    rowsFetched = 0
    '    Dim I As Integer, J As Integer

    '    For I = 0 To e.rowBuf.RowCount - 1
    '        ' Get the bookmark of the next available row
    '        bookm = GetRelativeBookmark(bookm, relpos)

    '        ' If the next is BOF or EOF, then done
    '        'If IsNull(bookm) Then Exit For
    '        If bookm Is Nothing Then Exit For

    '        For J = 0 To e.rowBuf.ColumnCount - 1
    '            e.rowBuf.Value(I, J) = GetUserData(bookm, J)
    '        Next

    '        ' Set the bookmark for the row
    '        e.rowBuf.Bookmark(I) = bookm

    '        ' Increment the count of fetched rows
    '        rowsFetched = rowsFetched + 1
    '    Next

    '    ' tell the grid how many rows were fetched
    '    e.rowBuf.RowCount = rowsFetched
    '    'DBGrid1.ApproxCount = Me.MaxRows + 1
    'End Sub

    '---> NOTE: UNBOUNDWRITEDATA IS NOT AVAILABLE IN VB.NET FOR AXDATAGRID CONTROL.  <------
    'Private Sub DBGrid1_UnboundWriteData(sender As Object, e As AxMSDBGrid.DBGridEvents_UnboundWriteDataEvent) Handles DBGrid1.UnboundWriteData
    '    '        'Debug.Print "DBGrid1_UnboundWriteData: ", RowBuf.RowCount, WriteLocation
    '    '        ' Assume that a Visual Basic for Applications function
    '    '        ' StoreUserData(bookm, col, value)
    '    '        ' takes a row bookmark, a column index, and a variant with the
    '    '        ' appropriate data to be stored in an array or database. The
    '    '        ' returns True if the data is acceptable and can be stored,
    '    '        ' False otherwise.

    '    '        ' Loop over all the columns of the row, storing non-Null values
    '    Dim I As Integer
    '    For I = 0 To e.rowBuf.ColumnCount - 1
    '        'If Not IsNull(RowBuf.Value(0, I)) Then
    '        If Not IsNothing(e.rowBuf.Value(0, 1)) Then
    '            If Not StoreUserData(e.writeLocation, I, e.rowBuf.Value(0, I)) Then

    '                ' storage of the data has failed. Fail the update
    '                e.rowBuf.RowCount = 0 ' tell the grid the update                        failed
    '                Exit Sub            ' it failed, so exit the event
    '            End If
    '        End If
    '    Next
    'End Sub

    '---> NOTE: UNBOUNDADDDATA EVENT IS NOT AVAILABLE IN VB.NET FOR AXDATAGRID CONTROL    <-----
    'Private Sub DBGrid1_UnboundAddData(sender As Object, e As AxMSDBGrid.DBGridEvents_UnboundAddDataEvent) Handles DBGrid1.UnboundAddData
    '    Dim c As MSDBGrid.Columns = DBGrid1.get_Columns
    '    ' Assume that a Visual Basic for Applications function
    '    ' StoreUserData(bookm, col, value) takes a row bookmark,
    '    ' a column index, and a variant with the appropriate data to be
    '    ' stored in an array or database. The StoreUserData()function
    '    ' returns True if the data is acceptable and can be stored ,
    '    ' False otherwise.

    '    ' First, get a bookmark for the new row. Do this with a VB
    '    ' for Applications function GetNewBookmark(), which allocates
    '    ' a new row of data in the storage media (array or database),
    '    ' and returns a variant containing a bookmark for that added row.
    '    e.newRowBookmark = GetNewBookmark()

    '    ' Loop over all the columns of the row, storing non-Null
    '    ' values
    '    Dim NewVal As Object
    '    Dim I As Integer
    '    For I = 0 To e.rowBuf.ColumnCount - 1
    '        NewVal = e.rowBuf.Value(0, I)
    '        'If IsNull(NewVal) Then
    '        If NewVal Is Nothing Then
    '            ' the RowBuf does not contain a value for this column.
    '            ' A default value should be set. A convenient value
    '            ' is the default value for the column.
    '            'NewVal = DBGrid1.Columns(I).DefaultValue
    '            NewVal = c(I).DefaultValue
    '        End If

    '        ' Now store the new values.
    '        If Not StoreUserData(e.newRowBookmark, I, NewVal) Then
    '            ' storage of the data has failed. Delete the added
    '            ' row using a Visual Basic for Applications function
    '            'DeleteRow, which takes a bookmark as an argument.
    '            ' Also, fail the update by clearing the RowCount.

    '            DeleteRow(e.newRowBookmark)
    '            e.rowBuf.RowCount = 0 ' tell the grid the update failed
    '            Exit Sub            ' it failed, so exit the event
    '        End If
    '    Next
    'End Sub

    Public Function DeleteRow(bookm As Object) As Boolean
        Dim Index As Integer

        Index = IndexFromBookmark(bookm, False)
        If Index < 0 Or Index >= mMaxRows Then
            DeleteRow = False
            Exit Function
        End If

        mMaxRows = mMaxRows - 1

        'Shift the data in the array
        Dim I As Integer
        For I = Index To mMaxRows - 1
            Dim J As Integer
            For J = 0 To mMaxCols - 1
                GridArray(J, I) = GridArray(J, I + 1)
                GridArray(J, I + 1) = ""
            Next
        Next

        ReDim Preserve GridArray(0 To mMaxCols - 1, 0 To mMaxRows - 1)

        DeleteRow = True
    End Function

    '---> NOTE: UNBOUNDDELETEROW EVENT IS NOT AVAILABLE IN VB.NET FOR AXDATAGRID CONTROL <----
    'Private Sub DBGrid1_UnboundDeleteRow(sender As Object, e As AxMSDBGrid.DBGridEvents_UnboundDeleteRowEvent) Handles DBGrid1.UnboundDeleteRow
    '    'Shift the data in the array
    '    ' This works nicely, but something is firing multiple
    '    ' UnboundReadData events which subsequently clear (or hide)
    '    ' the grid data.
    '    '  Debug.Print Bookmark, CInt(Bookmark)
    '    Dim Index As Integer
    '    Index = Val(e.bookmark)

    '    Dim I As Integer
    '    For I = Index To mMaxRows - 2
    '        Dim J As Integer
    '        For J = 0 To mMaxCols - 1
    '            GridArray(J, I) = GridArray(J, I + 1)
    '        Next
    '    Next
    '    For J = 0 To mMaxCols - 1
    '        GridArray(J, mMaxRows - 1) = ""
    '    Next
    '    mRefresh = False
    '    RaiseEvent RowDelete(e.bookmark)
    '    DBGrid1.Scroll(0, -1)
    '    mRefresh = True
    'End Sub

    Public Property firstrow() As Integer
        Get
            On Error Resume Next
            'firstrow = Val(axdatagrid1.firstrow)
            firstrow = Val(AxDataGrid1.FirstRow)
            'return firstrow
        End Get
        Set(value As Integer)
            On Error Resume Next
            'axdatagrid1.firstrow = value
            AxDataGrid1.FirstRow = value
        End Set
    End Property

    '---------NOTE: FONT, FONTNAME AND FONTSIZE PROPERTIES ARE NOT USED IN IMPLEMENTATION AREA(BILLOSALE FORM). SO COMMENTED.
    'Public Property Font() As Font
    '    Get
    '        'On Error Resume Next
    '        Return AxDataGrid1.Font
    '    End Get
    '    Set(value As Font)
    '        'On Error Resume Next
    '        AxDataGrid1.Font = value
    '        'PropertyChanged "Font"
    '        'OnFontChanged("Font")
    '    End Set
    'End Property

    'Public Property FontName() As String
    '    Get
    '        Return AxDataGrid1.Font.Name
    '    End Get
    '    Set(value As String)
    '        'DBGrid1.Font.Name = value
    '        'DBGrid1.Font = New Font(f.FontFamily, value)
    '        AxDataGrid1.Font = New Font(AxDataGrid1.Font.Name, AxDataGrid1.Font.Size, AxDataGrid1.Font.Style)
    '        'PropertyChanged "FontSize"
    '        'OnFontNameChanged("FontSize")
    '    End Set
    'End Property

    'Public Property FontSize() As String
    '    Get
    '        Return AxDataGrid1.Font.Size
    '    End Get
    '    Set(value As String)
    '        'DBGrid1.Font.Size = vData
    '        'DBGrid1.Font = New Font(f.FontFamily, value)
    '        AxDataGrid1.Font = New Font(AxDataGrid1.Font.Name, AxDataGrid1.Font.Size, AxDataGrid1.Font.Style)
    '        'propertychanged will fire readproperties and writeproperties events. 
    '        'Both readpropeties and writeproperties events are not avilable in vb.net.
    '        'Read and write properties functionality will be automatically done in vb.net.

    '        'PropertyChanged "Fontsize"
    '        'OnFontSizeChanged("fontSize")
    '    End Set
    'End Property

    Public Function FirstEmptyRow() As Integer
        FirstEmptyRow = LastRowUsed() + 1
    End Function

    Public Function ForceRowSave()
        Dim R As Integer

        ForceRowSave = Nothing
        'With AxDataGrid1
        With AxDataGrid1
            R = .Row
            If .Row = 0 Then .Row = 1 Else .Row = 0
            .Row = R
        End With
    End Function

    Public Function ColContaining(ByVal X As Single) As Integer
        ColContaining = AxDataGrid1.ColContaining(X)
    End Function

    Public Function RowContaining(ByVal Y As Single) As Integer
        RowContaining = AxDataGrid1.RowContaining(Y)
    End Function

    Public Function ColLeft(ByVal Col As Integer) As Single
        Dim I As Integer

        ColLeft = -1
        'For I = 1 To AxDataGrid1.Width
        '    If AxDataGrid1.ColContaining(I) = Col Then ColLeft = I : Exit Function
        'Next
        For I = 1 To AxDataGrid1.Width
            If AxDataGrid1.ColContaining(I) = Col Then ColLeft = I : Exit Function
        Next
    End Function

    Public Function RowTop(ByVal RowNum As Integer) As Single
        Dim I As Integer

        RowTop = -1
        'For I = 1 To AxDataGrid1.Height
        '    If AxDataGrid1.RowContaining(I) = RowNum Then RowTop = I : Exit Function
        'Next
        For I = 1 To AxDataGrid1.Height
            If AxDataGrid1.RowContaining(I) = RowNum Then RowTop = I : Exit Function
        Next
        '  RowTop = DBGrid1.RowTop(RowNum)
    End Function

    Public Sub AdjustControlToCell(Ctrl As Control, ByVal RowNum As Integer, ByVal ColNum As Integer, Left As Integer, Top As Integer)
        Dim L As Single, T As Single, W As Single, H As Single
        'Dim C As MSDBGrid.Column
        Dim C As MSDataGridLib.Column

        C = AxDataGrid1.Columns(ColNum)
        L = C.Left + Left
        T = AxDataGrid1.RowTop(RowNum) + Top
        W = C.Width
        H = AxDataGrid1.RowHeight
        C = Nothing

        On Error Resume Next
        'Ctrl.Move(L, T)
        Ctrl.Location = New Point(L, T)
        Ctrl.Width = W
        Ctrl.Height = H  ' An error I got using .Move L,T,W,H.. --> 'Height' property is read-only
    End Sub

    Public Sub ColorColumns(Optional ByVal WP As Integer = 0, Optional ByVal LP As Integer = 0)
        'Dim C As MSDBGrid.Column
        Dim C As MSDataGridLib.Column
        Dim L As Integer, T As Integer, W As Integer, H As Integer, I As Integer
        Dim R As Boolean

        If ColoringColumns Then Exit Sub
        ColoringColumns = True

        On Error Resume Next
        For I = 1 To MaxCols
            If ColumnBackColor(I) <> vbWhite Then
                C = AxDataGrid1.Columns(I)
                If C Is Nothing Then GoTo NoColumn

                'L = C.Left / VB6.TwipsPerPixelX
                'T = CInt(DBGrid1.RowTop(0) / Screen.TwipsPerPixelY)
                'W = C.Width / Screen.TwipsPerPixelX - 1
                'H = DBGrid1.Height / Screen.TwipsPerPixelY - 5
                C = Nothing

                R = DrawRectangle(AxDataGrid1.hWnd, L, T, W, H, ColumnBackColor(I), 25)
                'R = DrawRectangle(hwnd, L, T, W, H, ColumnBackColor(I), 25)
                Debug.Print(IIf(R, "SUCCESS: ", "FAILED:  ") & "i=" & I & ", Col=" & DescribeColor(ColumnBackColor(I)) & ", hwnd = " & Hwnd & ", (" & L & "x" & T & ")... " & W & "," & H)
            End If
NoColumn:
        Next
        ColoringColumns = False
    End Sub

    Public Property ColumnBackColor(ByVal N As Integer, Optional ByVal newCol As Integer = 0) As Integer
        Get
            On Error Resume Next
            ColumnBackColor = vbWhite
            ColumnBackColor = ColumnColors(N)
            If ColumnBackColor = 0 Then ColumnBackColor = vbWhite
            Return ColumnBackColor
        End Get
        Set(value As Integer)
            On Error Resume Next
            ColumnColors(value) = newCol
        End Set
    End Property

    Public Sub HookForm()
        Dim ug As UGridIO_PaintDelegate

        ug = AddressOf UGridIO_Paint
        UGridIO_AddHook(AxDataGrid1.hWnd, Me)
        PrevProc = SetWindowLong(AxDataGrid1.hWnd, GWL_WNDPROC, ug.ToString)
    End Sub

    Public Sub UnHookForm()
        SetWindowLong(AxDataGrid1.hWnd, GWL_WNDPROC, PrevProc)
        UGridIO_AddHook(AxDataGrid1.hWnd, Nothing)
    End Sub

    Public Sub SampleData()
        '   ReDim Preserve GridArray(0 To mMaxCols - 1, 0 To mMaxRows - 1) As String
        Dim I, J
        For I = 0 To mMaxCols - 1
            For J = 0 To mMaxRows - 1
                GridArray(I, J) = "Row" + Str(J) + ", Col" + Str(I)
            Next
        Next
    End Sub

    Private Sub AxDataGrid1_DblClick(sender As Object, e As EventArgs) Handles AxDataGrid1.DblClick
        RaiseEvent DblClick()
    End Sub

    Private Sub AxDataGrid1_Leave(sender As Object, e As EventArgs) Handles AxDataGrid1.Leave
        ' This should automatically fire the usercontrol's LostFocus event.
        mLostFocus = True
        '  Dim Cancel As Boolean
        '  RaiseEvent RowColChange(mLastRow, mLastCol, mLastRow, mLastCol, Cancel)
        '  RaiseEvent UserControl.LostFocus
    End Sub

    Private Sub AxDataGrid1_ScrollEvent(sender As Object, e As AxMSDataGridLib.DDataGridEvents_ScrollEvent) Handles AxDataGrid1.ScrollEvent
        Dim BM As Object

        On Error Resume Next
        If e.cancel = False Then
            mRefresh = False
            BM = AxDataGrid1.FirstRow
            mRefresh = True
            AxDataGrid1.Refresh()
            AxDataGrid1.FirstRow = BM
        End If
    End Sub

    Public Sub AxDataGrid1_RowColChange(sender As Object, e As AxMSDataGridLib.DDataGridEvents_RowColChangeEvent) Handles AxDataGrid1.RowColChange
        '---------> NOTE: RowColChange event will execute automatically as per WinCDS new sale requirement, 
        '---------> which Is happening in existing WinCDS of vb 6.0. But in vb.net, it will executing for user action only.
        '---------> So, created a new sub function and pulled this code in to it to execute automatically by calling it from cmdApplyBillOSale_Click event of BillOSale form.
        Dim Cancel As Boolean
        Dim OldCol As Object
        'Debug.Print "DBGrid1_RowColChange: ", LastRow, LastCol, DBGrid1.row, DBGrid1.FirstRow

        If Loading Then Exit Sub
        mLastRow = e.lastRow - 1
        mLastCol = e.lastCol
        mLostFocus = False
        If AxDataGrid1.Row = -1 Then Exit Sub

        OldCol = AxDataGrid1.Col  '+NEW 2003-01-20

        If RowColChange_Active Then
            'RaiseEvent RowColChange(mLastRow, mLastCol, Val(AxDataGrid1.Row) + Val(AxDataGrid1.FirstRow), AxDataGrid1.Col, Cancel)
            RaiseEvent RowColChange(mLastRow, mLastCol, AxDataGrid1.Row + 0, AxDataGrid1.Col, Cancel)
            Loading = True
            If Cancel Then
                ' Change to lastrow+col  ' This causes horrible looping, and I don't have time to debug it now, so un-movement code goes in the lostfocus events.
                '      row = LastRow
                '      col = LastCol
                On Error Resume Next
                AxDataGrid1.Col = e.lastCol   '+NEW 2003-01-20
                AxDataGrid1.Row = e.lastRow
            Else
                mCurrentCellModified = False
            End If
            Loading = False
        End If
    End Sub

    Public Sub Update()
        Refresh(True)
    End Sub

    Private Sub DBGrid1UnboundReadData()
        'Dim RowBuf As MSDataGridLib.
    End Sub

    Public Sub Axdatagrid1RowColumnChange(ByVal lastrow As Object, ByVal lastcol As Integer)
        'Debug.Print "DBGrid1_RowColChange: ", LastRow, LastCol, DBGrid1.row, DBGrid1.FirstRow
        Dim Cancel As Boolean
        Dim OldCol As Object

        If Loading Then Exit Sub

        mLastRow = lastrow
        mLastCol = lastcol
        mLostFocus = False

        'If AxDataGrid1.Row = -1 Then Exit Sub
        'If AxDataGrid1.Row = -1 Then AxDataGrid1.Row = 0

        OldCol = AxDataGrid1.Col  '+NEW 2003-01-20

        If RowColChange_Active Then
            'RaiseEvent RowColChange(mLastRow, mLastCol, Val(AxDataGrid1.Row) + Val(AxDataGrid1.FirstRow), AxDataGrid1.Col, Cancel)
            'BillOSale.UGridIO1RowColChange(mLastRow, mLastCol, AxDataGrid1.Row + 0, AxDataGrid1.Col, Cancel)
            RaiseEvent RowColChange(mLastRow, mLastCol, AxDataGrid1.Row + 0, AxDataGrid1.Col, Cancel)
            Loading = True
            If Cancel Then
                ' Change to lastrow+col  ' This causes horrible looping, and I don't have time to debug it now, so un-movement code goes in the lostfocus events.
                '      row = LastRow
                '      col = LastCol
                On Error Resume Next
                AxDataGrid1.Col = lastcol   '+NEW 2003-01-20
                AxDataGrid1.Row = lastrow
            Else
                mCurrentCellModified = False
            End If
            Loading = False
        End If

    End Sub

    'Public Sub Axdatagrid1RowColChange(sender As Object, Optional e As AxMSDataGridLib.DDataGridEvents_RowColChangeEvent = New AxMSDataGridLib.DDataGridEvents_RowColChangeEvent(lastRow:=LastRowUsed, 1) Handles AxDataGrid1.RowColChange
    '    AxDataGrid1_RowColChange(AxDataGrid1, e)
    'End Sub
    Public Sub MoveRow(I As Integer)
        If I < 0 Then I = 0
        With AxDataGrid1
            .Bookmark = .RowBookmark(I)
        End With
    End Sub
End Class
