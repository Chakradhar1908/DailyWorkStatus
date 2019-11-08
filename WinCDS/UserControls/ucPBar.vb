Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports System.Drawing.Printing
Public Class ucPBar
    Private mMin As Integer
    Private mMax As Integer
    Private mValue As Integer
    Private mStyle As Integer
    Private mShowDuration As Boolean
    Private mShowRemaining As Boolean

    Private StartTime As Date

    Private mBorderStyle As Integer
    Private mHasCaption As Boolean

    Private Const d_Max As Integer = 100
    Private Const d_Min As Integer = 1
    Private Const d_Value As Integer = 1
    Private Const d_Style As Integer = 1
    Private Const d_ShowDuration As Boolean = False
    Private Const d_ShowRemaining As Boolean = True

    Private Const d_BorderStyle As Integer = 0
    Private Const d_HasCaption As Boolean = True

    'Private Const d_ForeColor As Integer = vbWhite
    Private d_ForeColor As Color = Color.White
    Private Const d_FillColor As Integer = vbBlue
    'Private Const d_BackColor As Integer = vbButtonFace
    Private d_BackColor As Color = SystemColors.ButtonFace
    Private Const d_FontName As String = "Arial"

    Public Event Clickk()
    Public Event DblClick()

    Private Sub ValidateRange()
        If Min >= Max Then Max = Min + 1
        If Value < Min Then Value = Min
        If Value > Max Then Value = Max
    End Sub

    Public Property Max() As Integer
        Get
            Max = mMax
        End Get
        Set(value As Integer)
            mMax = value
            ValidateRange()
            Redraw(True)
        End Set
    End Property

    Public Property Min() As Integer
        Get
            Min = mMin
        End Get
        Set(value As Integer)
            mMin = value
            ValidateRange()
            Redraw(True)
        End Set
    End Property

    Public Property Value() As Integer
        Get
            Value = mValue
            If Value = 0 Then StartTime = Now
        End Get
        Set(value As Integer)
            mValue = value
            ValidateRange()
            If ShowRemaining And value = 0 Then TimeRemaining()
            Redraw(True)
        End Set
    End Property

    Public Property Style() As Integer
        Get
            Style = mStyle
        End Get
        Set(value As Integer)
            mStyle = value
            Redraw(True)
        End Set
    End Property

    Public ReadOnly Property Duration() As String
        Get
            Dim S As Integer
            S = DateDiff("s", StartTime, Now)
            Duration = "" & Format(S \ 60, "00") & ":" & Format(S Mod 60, "00")
        End Get
    End Property

    Public Property BorderStyle() As Integer
        Get
            BorderStyle = mBorderStyle
        End Get
        Set(value As Integer)
            mBorderStyle = value
            Redraw(True)
        End Set
    End Property

    Public Property HasCaption() As Boolean
        Get
            HasCaption = mHasCaption
        End Get
        Set(value As Boolean)
            mHasCaption = value
            Redraw(True)
        End Set
    End Property

    Public Property ForeColorNew() As Color
        Get
            'ForeColorNew = UserControl.ForeColor
            ForeColorNew = Me.ForeColor
        End Get
        Set(value As Color)
            Me.ForeColor = value
            Redraw(True)
        End Set
    End Property

    'NOTE: THE BELOW LINES ARE COMMENTED, BECAUSE IN VB.NET FillColor PROPERTY IS NOT AVAILABLE.
    'Public Property Get FillColor() As OLE_COLOR
    '    FillColor = UserControl.FillColor
    'End Property
    'Public Property Let FillColor(ByVal vData As OLE_COLOR)
    '    UserControl.FillColor = vData
    '    Redraw True
    'End Property

    Public Property BackColorNew() As Color
        Get
            BackColorNew = Me.BackColor
        End Get
        Set(value As Color)
            Me.BackColor = value
            Redraw(True)
        End Set
    End Property

    Public Property FontName() As String
        Get
            'FontName = UserControl.FontName
            FontName = Me.Font.Name
        End Get
        Set(value As String)
            'UserControl.FontName = vData
            Me.Font = New Font(Me.Font, value)
            Redraw(True)
        End Set
    End Property

    Public Property ShowDuration() As Boolean
        Get
            ShowDuration = mShowDuration
        End Get
        Set(value As Boolean)
            mShowDuration = value
        End Set
    End Property

    Public Property ShowRemaining() As Boolean
        Get
            ShowRemaining = mShowRemaining
        End Get
        Set(value As Boolean)
            mShowRemaining = value
        End Set
    End Property

    Private Sub ucPBar_Click(sender As Object, e As EventArgs) Handles MyBase.Click
        RaiseEvent Clickk()
    End Sub

    Private Sub ucPBar_DoubleClick(sender As Object, e As EventArgs) Handles MyBase.DoubleClick
        RaiseEvent DblClick()
    End Sub

    Private Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Min = d_Min
        Max = d_Max
        Value = d_Value
        Style = d_Style

        StartTime = Now

        BorderStyle = d_BorderStyle
        HasCaption = d_HasCaption

        ForeColorNew = d_ForeColor
        'FillColor = d_FillColor  -> Commeneted, because in vb.net fillcolor property is not there.
        BackColorNew = d_BackColor
        FontName = d_FontName
    End Sub

    Private Sub ucPBar_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
        Redraw()
    End Sub

    Private Sub Redraw(Optional ByVal WithRefresh As Boolean = False)
        Dim Rng As Integer, X As Integer, Tx As String, RW As Double, RH As Double
        Dim A As Color, B As Integer
        A = ForeColorNew
        'B = FillColor

        Dim TxFont As Font = New Font("Arial", Me.Font.Size, FontStyle.Regular)

        Rng = Max - Min
        If Rng <= 0 Then Exit Sub ' prevent divbyzero, jic
        X = Me.ClientSize.Width * (Value - Min) / Rng
        'UserControl.ForeColor = B

        Dim g As Graphics = Me.CreateGraphics
        Dim BlackPen As Pen
        If Value <> Min Then
            'UserControl.Line(0, 0)-(X, ScaleHeight), , BF
            g.DrawLine(Pens.Black, New Point(0, 0), New PointF(X, Me.ClientSize.Height))

        End If
        If Value <> Max Then
            'UserControl.ForeColor = UserControl.BackColor
            Me.ForeColorNew = Me.BackColorNew
            'UserControl.Line(X + 1, 0)-(ScaleWidth, ScaleHeight), , BF
            g.DrawLine(Pens.Black, New Point(X + 1, 0), New Point(Me.ClientSize.Width, Me.ClientSize.Height))
        End If

        'UserControl.FontSize = 22
        Me.Font = New Font(Me.Font, 22)
        'UserControl.FontBold = True
        Me.Font = New Font(Me.Font, FontStyle.Bold)

        If BorderStyle <> 0 Then
            'UserControl.ForeColor = vbBlack
            Me.ForeColorNew = Color.Black
            If BorderStyle > 2 Then
                'UserControl.DrawWidth = 5
                BlackPen = New Pen(Color.Black, 5)
                'UserControl.Line(0, 0)-(ScaleWidth - 10, ScaleHeight - 10), , B
                g.DrawLine(BlackPen, New Point(0, 0), New Point(Me.ClientSize.Width - 10, Me.ClientSize.Height - 10))
            Else
                'UserControl.DrawWidth = 2
                BlackPen = New Pen(Color.Black, 2)
                'UserControl.Line(0, 0)-(ScaleWidth, ScaleHeight), , B
                g.DrawLine(BlackPen, New Point(0, 0), New PointF(Me.ClientSize.Width, Me.ClientSize.Height))
            End If
            'UserControl.DrawWidth = 1

        End If

        Tx = "" & Format((Value - Min) / Rng * 100, "0") & "%"
        If HasCaption And Value <> Min Then

            'RH = UserControl.TextHeight(Tx)

            TxFont = New Font(Font, Tx)
            Dim StringSize As SizeF = g.MeasureString(Tx, TxFont)
            RH = StringSize.Height
            Do While RH > Me.ClientSize.Height
                'UserControl.FontSize = UserControl.FontSize - 2
                Me.Font = New Font(Me.Font.Name, Me.Font.Size - 2)
                'RH = UserControl.TextHeight(Tx)
                TxFont = New Font(Font, Tx)
                StringSize = g.MeasureString(Tx, TxFont)
                RH = StringSize.Height
            Loop
            'RW = UserControl.TextWidth(Tx)
            RW = StringSize.Width
            'UserControl.CurrentX = ScaleWidth / 2 - RW / 2
            'UserControl.CurrentY = ScaleHeight / 2 - RH / 2
            g.DrawString(Tx, TxFont, New SolidBrush(Color.Black), Me.ClientSize.Width / 2 - RW / 2, Me.ClientSize.Height / 2 - RH / 2)
            'UserControl.ForeColor = A
            Me.ForeColorNew = A
            'UserControl.DrawMode = vbXorPen
            'UserControl.Print Tx
            Dim Puc As PrintUserControlForucPBar = New PrintUserControlForucPBar   '-> These four lines(PrintUsercontrolForucPBar class) are replacement for the above line UserControl.Print Tx
            Puc.PrintText = Tx
            Puc.PrintTextFont = TxFont
            Puc.Print()

            '    UserControl.DrawMode = vbCopyPen
        End If

        'UserControl.FillColor = B  '-> FillColor property does not exist for user control in vb.net
        'UserControl.ForeColor = A
        Me.ForeColorNew = A

        Dim StringSize2 As SizeF
        If IsDevelopment() Or ShowDuration Then
            Dim D As String
            D = Duration
            'UserControl.ForeColor = vbWhite
            Me.ForeColorNew = Color.White
            'UserControl.FontName = "Arial"
            'UserControl.FontSize = 6
            Me.Font = New Font("Arial", 6)
            TxFont = New Font("Arial", 6)
            StringSize2 = g.MeasureString(D, TxFont)
            'UserControl.CurrentX = 30 ' Width - TextWidth(D)
            'UserControl.CurrentY = Height - TextHeight(D)
            'UserControl.Print D
            g.DrawString(D, TxFont, New SolidBrush(Color.Black), 30, Height - StringSize2.Height)
        End If

        If IsDevelopment() Or ShowRemaining Then
            Dim R As String
            R = TimeRemaining(Max - Value, FitRange(25, Max / 10, 250))
            If Microsoft.VisualBasic.Left(R, 3) = "00:" Then
                R = Mid(R, 4)
                If Microsoft.VisualBasic.Left(R, 3) = "00:" Then R = Mid(R, 4) & "s"
                If Microsoft.VisualBasic.Left(R, 1) = "0" Then R = Mid(R, 2)
            End If
            'UserControl.ForeColor = vbBlack
            Me.ForeColorNew = Color.Black
            'UserControl.FontName = "Arial"
            'UserControl.FontSize = 7.5
            Me.Font = New Font("Arial", 7.5)
            TxFont = New Font("Arial", 7.5)
            StringSize2 = g.MeasureString(R, TxFont)
            'UserControl.CurrentX = Width - TextWidth(R) - 30
            'UserControl.CurrentY = Height - TextHeight(R)
            'UserControl.Print R
            g.DrawString(R, TxFont, New SolidBrush(Color.Black), Width - StringSize2.Width - 30, Height - StringSize2.Height)
        End If

        If WithRefresh Then Refresh()
    End Sub

End Class
