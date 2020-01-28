Public Class cStringBuilder
    Private m_sString As String
    Private m_iChunkSize As Integer
    Private m_iPos As Integer
    Private m_iLen As Integer
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
      (pDst As Object, pSrc As Object, ByVal ByteLen As Integer)

    Public Sub Append(ByRef sThis As String)
        Dim lLen As Integer
        Dim lLenPlusPos As Integer

        ' Append an item to the string:
        'lLen = LenB(sThis)
        lLen = Len(sThis)
        lLenPlusPos = lLen + m_iPos
        If lLenPlusPos > m_iLen Then
            Dim lTemp As Integer

            lTemp = m_iLen
            Do While lTemp < lLenPlusPos
                lTemp = lTemp + m_iChunkSize
            Loop

            m_sString = m_sString & Space((lTemp - m_iLen) \ 2)
            m_iLen = lTemp
        End If

        'CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
        CopyMemory(UnsignedAdd(StrPtr(m_sString), m_iPos), StrPtr(sThis), lLen)
        m_iPos = m_iPos + lLen
    End Sub

    Private Function UnsignedAdd(ByRef Start As Integer, ByRef Incr As Integer) As Integer
        ' This function is useful when doing pointer arithmetic,
        ' but note it only works for positive values of Incr

        If Start And &H80000000 Then 'Start < 0
            UnsignedAdd = Start + Incr
        ElseIf (Start Or &H80000000) < -Incr Then
            UnsignedAdd = Start + Incr
        Else
            UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
        End If
    End Function
End Class
