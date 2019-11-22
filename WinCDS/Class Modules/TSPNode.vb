Public Class TSPNode
    Private Const Radius As Integer = 2
    Private Const DIAMETER As Integer = 2 * Radius + 1
    Private Const DEFAULT_DELAY_INCREMENT As Integer = 15

    Public X As Integer, Y As Integer

    Public Address As String
    Public City As String, State As String, Zip As String

    Public StopTime As Integer
    Private mWindowFrom As Date, mWindowTo As Date

    Public Visited As Boolean
    Public Name As String
    Public IsDepot As Boolean

    Public Property WindowFrom() As Date
        Get
            WindowFrom = mWindowFrom
        End Get
        Set(value As Date)
            mWindowFrom = TimeValue(value)
        End Set
    End Property

    Public Property WindowTo() As Date
        Get
            WindowTo = mWindowTo
        End Get
        Set(value As Date)
            mWindowTo = TimeValue(value)
        End Set
    End Property

    Public Sub Setup(Optional ByVal new_Name As String = "", Optional ByVal vX As Integer = 0, Optional ByVal vY As Integer = 0, Optional ByVal nStopTime As Integer = 0, Optional ByVal nWindowFrom As Date = #12:00:00 AM#, Optional ByVal nWindowTo As Date = #11:59:59 PM#, Optional ByVal nAddress As String = "", Optional ByVal nCity As String = "", Optional ByVal nSt As String = "", Optional ByVal nZip As String = "")
        Name = new_Name
        X = vX
        Y = vY
        StopTime = nStopTime
        WindowFrom = nWindowFrom
        WindowTo = nWindowTo
        Address = nAddress
        City = nCity
        State = nSt
        Zip = nZip
    End Sub

    Public Sub DrawNode(ByVal vPic As PictureBox, Optional ByVal XOff As Integer = 0, Optional ByVal YOff As Integer = 0, Optional ByVal XScale As Single = 1.0#, Optional ByVal YScale As Single = 1.0#)
        Dim Sz As Integer
        'The below variable declaration is to draw line and also to print text in picturebox. Because vb.net does not have picturebox.line method.
        Dim Bmp As Bitmap
        Dim G As Graphics
        Dim P As Pen
        Dim PictureInPicturebox As Boolean

        If Not vPic.Image Is Nothing Then
            Bmp = New Bitmap(vPic.Image)
            G = Graphics.FromImage(Bmp)
            P = New Pen(Color.Black)
            PictureInPicturebox = True
        End If
        Sz = 20
        vPic.ForeColor = Color.Yellow
        'vPic.Line((X - XOff) * XScale - Sz, (Y - YOff) * YScale - Sz)-((X - XOff) * XScale + Sz, (Y - YOff) * YScale + Sz), , BF

        If PictureInPicturebox = True Then
            G.DrawLine(P, (X - XOff) * XScale - Sz, (Y - YOff) * YScale - Sz, (X - XOff) * XScale + Sz, (Y - YOff) * YScale + Sz)
        End If
        'vPic.CurrentX = (X - XOff) * XScale + Sz
        'vPic.CurrentY = (Y - YOff) * YScale + Sz
        'vPic.ForeColor = vbRed
        'vPic.FontName = "Arial"
        'vPic.FontSize = 8
        'vPic.Print Name
        G = vPic.CreateGraphics
        G.DrawString(Name, New Font("Arial", 8, FontStyle.Regular), Brushes.Yellow, (X - XOff) * XScale + Sz, (Y - YOff) * YScale + Sz)
    End Sub

    Public Function HasWindow() As Boolean
        HasWindow = DateDiff("n", #12:00:00 AM#, WindowFrom) <> 0 Or DateDiff("n", #11:59:59 PM#, WindowTo)
    End Function

    Public Function IsBeforeWindow(ByVal T As Date) As Boolean
        IsBeforeWindow = DateDiff("n", TimeValue(T), WindowFrom) > 0
    End Function

    Public Function TimeToWindow(ByVal T As Date) As Integer
        If IsInWindow(T) Then
            TimeToWindow = 0
        ElseIf IsBeforeWindow(T) Then
            TimeToWindow = DateDiff("n", TimeValue(T), WindowFrom)
        Else
            TimeToWindow = DateDiff("n", TimeValue(T), WindowTo)
        End If
    End Function

    Public Function IsInWindow(ByVal T As Date) As Boolean
        IsInWindow = DateDiff("n", T, WindowFrom) <= 0 And DateDiff("n", T, WindowTo) >= 0
    End Function

    Public Function TimeRemainingInWindow(ByVal T As Date) As Integer
        If IsAfterWindow(T) Then
            TimeRemainingInWindow = 0
        Else
            TimeRemainingInWindow = DateDiff("n", TimeValue(T), WindowTo)
        End If
    End Function

    Public Function WindowMissPenalty(ByVal T As Date) As Single
        If DateDiff("n", T, WindowFrom) > 0 Or DateDiff("n", T, WindowTo) < 0 Then
            WindowMissPenalty = MISSED_WINDOW_PENALTY
        Else
            WindowMissPenalty = 0
        End If
    End Function

    Public Function IsAfterWindow(ByVal T As Date) As Boolean
        IsAfterWindow = DateDiff("n", WindowTo, TimeValue(T)) > 0
    End Function

End Class
