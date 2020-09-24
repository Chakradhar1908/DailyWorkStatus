Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class SelectPrinter
    Public SmallTags As Boolean, TagSize As String
    Dim printer As New Printer

    Public Sub PrintTags(ByVal nStyle As String, ByVal nDesc As String, ByVal nLanded As String,
    ByVal nList As String, ByVal nOnSale As String, ByVal nDeptNo As String, ByVal nCode As String,
    ByVal nVendor As String, ByVal nStock As String, ByVal nComments As String,
    Optional ByVal ParentForm As Form = Nothing, Optional ByVal KitMode As Boolean = False,
    Optional ByVal nPictureFile As String = "", Optional ByVal nPrintPicture As Integer = 0,
    Optional ByVal AutoPrintTags As Boolean = False, Optional ByVal DefaultTagSize As String = "",
    Optional ByVal DefaultTicketPath As String = "", Optional ByVal nHidePricing As Boolean = False)
        '      ' This function gets the data from outside and shows the form..

        '      Dim X
        '      X = Printer.DeviceName

        '      If PrSel.GetSelectedPrinter Is Nothing Then
        '          PrintingAllowed = False
        '      Else
        '          PrintingAllowed = True
        '      End If

        '      Load Me
        'LoadTagInfo nStyle, nDesc, nLanded, nList, nOnSale, nDeptNo, nCode, nVendor, nStock, nComments
        'HidePricing = nHidePricing


        '      AllowRecLabelPrinting = True
        '      KitTag = KitMode  ' boolean, means this is a kit.
        '      PictureFile = nPictureFile
        '      PrintPicture = nPrintPicture

        '      If AutoPrintTags Then
        '          AutoPrint DefaultTagSize, DefaultTicketPath, ParentForm
        'ElseIf ParentForm Is Nothing Then
        '          Show vbModal
        'Else
        '          Show vbModal, ParentForm
        '  If UCase(ParentForm.Name) = "EDITPO" Then
        '              ' Awful hack, but oh well..
        '              EditPO.SaveTagPrintingOptions TagSize, TicketPath
        '  End If
        '      End If
        '      If Not SmallTags Then
        '          SetPrinter X
        'End If
        '      KitTag = False
    End Sub

    Public Function PrintSoldTags(ByVal Style As String, Optional ByVal LastName As String = "", Optional ByVal SaleNo As String = "", Optional ByVal Q As Integer = 1) As Integer
        'print dymo labels
        Dim Counter As Byte, OriginalPrint As String, InvData As New CInvRec
        Dim P As Object, SQL As String

        Dim Tx As Integer
        'If Q <= 0 Then Q = Quantity

        On Error Resume Next
        If Not InvData.Load(Style, "Style") Then
            DisposeDA(InvData)
            Exit Function
        End If

        OriginalPrint = printer.DeviceName

        For Counter = 1 To Q
            If Not SetDymoPrinter() Then  ' Yes, it's inside the loop
                MessageBox.Show("Dymo Printer Required!", "WinCDS")
                Exit Function
            End If

            printer.FontSize = 14
            printer.CurrentX = 0
            printer.CurrentY = 0

            printer.Orientation = vbPRORLandscape

            printer.FontSize = 32
            printer.FontBold = True
            printer.Print("SOLD") 'PO.AckInv
            Tx = printer.CurrentX * 1.1
            printer.FontBold = False

            printer.CurrentY = 0
            printer.FontSize = 14
            If LastName <> "" Then printer.CurrentX = Tx : printer.Print("Cust: " & LastName)
            If SaleNo <> "" Then printer.CurrentX = Tx : printer.Print("Sale: " & SaleNo)
            printer.CurrentX = Tx : printer.Print("Date: " & DateFormat(Now))

            printer.EndDoc()

            printer.Orientation = vbPRORPortrait
            If OriginalPrint <> "" Then
                If Not SetPrinter(OriginalPrint) Then
                    MessageBox.Show("Could not restore the original printer!", "Original Printer")
                End If
            End If
        Next

        DisposeDA(InvData)

        PrintSoldTags = Q
    End Function
End Class