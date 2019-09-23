Imports System.Drawing.Printing
Public Class PrintUserControlForucPBar
    Inherits PrintDocument
    Public PrintText As String
    Public PrintTextFont As Font = New Font("Arial", 10)

    Protected Overrides Sub OnPrintPage(ByVal e As PrintPageEventArgs)
        MyBase.OnPrintPage(e)
        e.Graphics.DrawString(PrintText, PrintTextFont, Brushes.Black, 0, 0)
        e.HasMorePages = False
    End Sub
End Class
