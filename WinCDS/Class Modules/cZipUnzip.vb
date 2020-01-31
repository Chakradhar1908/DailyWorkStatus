Public Class cZipUnzip
    Private m_sZipFile As String
    Private m_iCount As Integer
    Private m_tZipContents() As tZipContents
    Private m_sUnzipFolder As String
    Private m_tDCL As DCLIST
    Private Structure tZipContents
        Dim sName As String
        Dim sFolder As String
        Dim lSize As Integer
        Dim lPackedSize As Integer
        Dim lFactor As Integer
        Dim sMethod As String
        Dim DDate As Date
        Dim lCrc As Integer
        Dim fEncryped As Boolean
        Dim fSelected As Boolean
    End Structure
    Public Enum EUZOverWriteResponse
        euzDoNotOverwrite = 100
        euzOverwriteThisFile = 102
        euzOverwriteAllFiles = 103
        euzOverwriteNone = 104
    End Enum
    Public Event OverwritePrompt(ByVal sFIle As String, ByRef eResponse As EUZOverWriteResponse)
    Public Event Progress(ByVal lCount As Integer, ByVal sMsg As String)
    Public Event PasswordRequest2(ByRef sPassword As String, ByRef bCancel As Boolean)
    Public Event Cancel(ByVal sMsg As String, ByRef bCancel As Boolean)

    Public Property ZipFile() As String
        Get
            ZipFile = m_sZipFile
        End Get
        Set(value As String)
            m_sZipFile = value
            m_iCount = 0
            Erase m_tZipContents
        End Set
    End Property

    Public Property UnzipFolder() As String
        Get
            UnzipFolder = m_sUnzipFolder
            m_tDCL.lpszExtractDir = m_sUnzipFolder
        End Get
        Set(value As String)
            m_sUnzipFolder = value
        End Set
    End Property

    Public Property ExtractOnlyNewer() As Boolean
        Get
            ExtractOnlyNewer = (m_tDCL.ExtractOnlyNewer <> 0)      ' 1=extract only newer
        End Get
        Set(value As Boolean)
            m_tDCL.ExtractOnlyNewer = Math.Abs(CInt(value))      ' 1=extract only newer
        End Set
    End Property

    Public Property OverwriteExisting() As Boolean
        Get
            OverwriteExisting = (m_tDCL.noflag <> 0)
        End Get
        Set(value As Boolean)
            m_tDCL.noflag = Math.Abs(CInt(value))
        End Set
    End Property

    Public Function Directory() As Integer
        Dim S(0 To 0) As String
        m_tDCL.lpszZipFN = m_sZipFile
        m_tDCL.lpszExtractDir = vbNullChar
        m_tDCL.nvflag = 1
        modZipUnzip.VBUnzip(Me, m_tDCL, 0, S, 0, S)
    End Function

    Public ReadOnly Property FileCount() As Integer
        Get
            FileCount = m_iCount
        End Get
    End Property

    Public ReadOnly Property FileDirectory(ByVal lIndex As Integer) As String
        Get
            FileDirectory = m_tZipContents(lIndex).sFolder
        End Get
    End Property

    Public Property UseFolderNames() As Boolean
        Get
            UseFolderNames = (m_tDCL.ndflag <> 0)
        End Get
        Set(value As Boolean)
            m_tDCL.ndflag = Math.Abs(CInt(value))
        End Set
    End Property

    Public Function Unzip() As Boolean
        Dim sInc() As String
        Dim iIncCount As Integer
        Dim S() As String
        Dim I As Integer
        If (m_sZipFile <> "") Then
            If (m_iCount > 0) Then
                For I = 1 To m_iCount
                    If (m_tZipContents(I).fSelected) Then
                        iIncCount = iIncCount + 1
                        ReDim Preserve sInc(0 To iIncCount - 1)
                        sInc(iIncCount) = ReverseSlashes(m_tZipContents(I).sFolder, m_tZipContents(I).sName)
                    End If
                Next
                If (iIncCount = m_iCount) Then
                    iIncCount = 0
                    ReDim sInc(0 To 0)
                End If
            End If
            m_tDCL.lpszZipFN = m_sZipFile
            m_tDCL.nvflag = 0
            m_tDCL.lpszExtractDir = m_sUnzipFolder
            Unzip = (modZipUnzip.VBUnzip(Me, m_tDCL, iIncCount, sInc, 0, S) <> 0)
        End If
    End Function

    Private Function ReverseSlashes(ByVal sFolder As String, ByVal sFIle As String) As String
        Dim sOut As String
        Dim iPos As Integer, iLastPos As Integer

        If Len(sFolder) > 0 And sFolder <> vbNullChar Then
            sOut = sFolder & "/" & sFIle
            iLastPos = 1
            Do
                iPos = InStr(iLastPos, sOut, "\")
                If (iPos <> 0) Then
                    Mid(sOut, iPos, 1) = "/"
                    iLastPos = iPos + 1
                End If
            Loop While iPos <> 0
            ReverseSlashes = sOut
        Else
            ReverseSlashes = sFIle
        End If
    End Function

    Friend Sub OverwriteRequest(ByVal sFIle As String, ByRef eResponse As EUZOverWriteResponse)
        RaiseEvent OverwritePrompt(sFIle, eResponse)
    End Sub

    Friend Sub ProgressReport(ByVal sMsg As String)
        RaiseEvent Progress(1, sMsg)
    End Sub

    Friend Sub PasswordRequest(ByRef sPassword As String, ByRef bCancel As Boolean)
        RaiseEvent PasswordRequest2(sPassword, bCancel)
    End Sub

    Friend Sub DirectoryListAddFile(
      ByVal sFileName As String,
      ByVal sFolder As String,
      ByVal DDate As Date,
      ByVal lSize As Integer,
      ByVal lCrc As Integer,
      ByVal fEncrypted As Boolean,
      ByVal lFactor As Integer,
      ByVal sMethod As String
   )
        If (sFileName <> vbNullChar) And Len(sFileName) > 0 Then
            m_iCount = m_iCount + 1
            'ReDim Preserve m_tZipContents(1 To m_iCount) As tZipContents
            ReDim Preserve m_tZipContents(0 To m_iCount - 1)
            m_tZipContents(m_iCount - 1).sName = sFileName
            m_tZipContents(m_iCount - 1).sFolder = sFolder
            m_tZipContents(m_iCount - 1).DDate = DDate
            m_tZipContents(m_iCount - 1).lSize = lSize
            m_tZipContents(m_iCount - 1).lCrc = lCrc
            m_tZipContents(m_iCount - 1).lFactor = lFactor
            m_tZipContents(m_iCount - 1).sMethod = sMethod
            m_tZipContents(m_iCount - 1).fEncryped = fEncrypted
            ' Default to selected:
            m_tZipContents(m_iCount - 1).fSelected = True
        End If
    End Sub

    Friend Sub Service(ByVal sMsg As String, ByRef bCancel As Boolean)
        RaiseEvent Cancel(sMsg, bCancel)
    End Sub

End Class
