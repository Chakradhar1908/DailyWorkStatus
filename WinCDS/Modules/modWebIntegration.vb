Module modWebIntegration
    Private Count As Integer
    Private dL As Object

    Public Sub Generate2DataCSV(Optional ByRef FN As String = "", Optional ByRef WebSite As String = "")
        Dim SQL As String, RS As ADODB.Recordset, FNum As Integer, Line As String

        On Error Resume Next
        Kill(FN)
        On Error GoTo 0

        Count = 0
        dL = GetDepartmentList()

        FNum = FreeFile()
        'Open FN For Output As #FNum
        FileOpen(FNum, FN, OpenMode.Output)
        Print(FNum, GenerateCSVHeader())

        SQL = "SELECT [2Data].*, iif(isnull(Search.RN),0,1) AS Orderable FROM [2DATA] LEFT JOIN [Search] ON [2Data].Style = [Search].Style"
        '  SQL = SQL & " WHERE Trim([2Data].Dept)='1'"
        RS = GetRecordsetBySQL(SQL, , GetDatabaseInventory)
        Do Until RS.EOF
            Line = GenerateCSVLine(RS, WebSite)
            If Len(Trim(Line)) > 0 Then Print(FNum, Line)
            RS.MoveNext()
        Loop

        'Close #FNum
        FileClose(FNum)
    End Sub

    Private Function GenerateCSVHeader() As String
        Dim T As String
        T = "id,name,code,price,sale-price,ship-weight,orderable,taxable" ' mandatory fields

        T = T & ",caption,abstract"
        T = T & ",label"
        T = T & ",manufacturer"

        T = T & ",product-url"
        T = T & ",department,department-url"

        GenerateCSVHeader = T
    End Function

    Private Function GenerateCSVLine(ByRef RS As ADODB.Recordset, Optional ByRef URLStart As String = "") As String
        If Right(URLStart, 1) = "/" Then URLStart = Left(URLStart, Len(URLStart) - 1)
        If IfNullThenNilString(RS("Vendor").Value) = "" Then Exit Function
        If IfNullThenNilString(RS("VendorNo").Value) = "" Then Exit Function
        If IfNullThenZero(RS("Orderable").Value) = 0 Then Exit Function

        '  If Count > 20 Then Exit Function
        Count = Count + 1

        GenerateCSVLine = ""

        AddCSVElement(GenerateCSVLine, GetYahooCVSID(RS("Style").Value))            ' id
        AddCSVElement(GenerateCSVLine, RS("Style").Value)                           ' name
        AddCSVElement(GenerateCSVLine, RS("Style").Value)                           ' code
        AddCSVElement(GenerateCSVLine, RS("List").Value)                            ' price
        AddCSVElement(GenerateCSVLine, RS("OnSale").Value)                          ' sale-price
        AddCSVElement(GenerateCSVLine, 0)                                     ' ship-weight
        AddCSVElement(GenerateCSVLine, YesNo(RS("Orderable").Value = 1))            ' orderable
        AddCSVElement(GenerateCSVLine, "Yes")                                 ' taxable

        AddCSVElement(GenerateCSVLine, Trim(IfNullThenNilString(RS("Desc").Value)))       ' caption
        AddCSVElement(GenerateCSVLine, Trim(IfNullThenNilString(RS("Comments").Value)))   ' abstract
        AddCSVElement(GenerateCSVLine, IfNullThenZero(RS("RN").Value))              ' label

        AddCSVElement(GenerateCSVLine, IfNullThenNilString(RS("Vendor").Value))     ' manufacturer

        AddCSVElement(GenerateCSVLine, URLStart & "/product/" & GetYahooCVSID(RS("Style").Value) & ".html") ' product-url

        Dim T As Integer
        T = RS("Dept").Value
        If T < LBound(dL) Or T > UBound(dL) Then T = LBound(dL)
        AddCSVElement(GenerateCSVLine, dL(T))  ' department name
        AddCSVElement(GenerateCSVLine, frmAutoWeb.SiteDepartmentURL(dL(T)))  ' department url
    End Function

    Private Sub AddCSVElement(ByRef Line As String, ByVal Item As String)
        Dim NItem As String
        If Len(Line) > 0 Then Line = Line & ","
        If InStr(Item, """") > 0 Or InStr(Item, ",") > 0 Then
            Item = Replace(Item, """", """""")
            Item = """" & Item & """"
        End If
        Line = Line & Item
    End Sub

    Public Function GetYahooCVSID(ByVal Style As String) As String
        Dim T As String
        T = Style
        T = Replace(T, "\", "--bslsh--")
        T = Replace(T, "/", "--slash--")
        T = Replace(T, "*", "--astar--")
        T = Replace(T, " ", "--space--")
        T = Replace(T, "(", "--lparn--")
        T = Replace(T, ")", "--rparn--")
        T = Replace(T, "[", "--lbrak--")
        T = Replace(T, "]", "--rbrak--")
        T = Replace(T, "_", "--uscor--")
        GetYahooCVSID = LCase(T)
    End Function
End Module
