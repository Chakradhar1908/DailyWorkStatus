Module dbaseloModule
    Public Function dbClose() As Boolean
        '::::dbClose
        ':::SUMMARY
        ': Close Accounting DB
        ':::DESCRIPTION
        ': This function is used to close the Database.
        ':
        ': Whichever connection was previously opened by the dbOpen is closed by this function.
        ':::PARAMETERS
        ':::SEE ALSO
        ': dbOpen
        ':::RETURN
        ': Returns true
        On Error Resume Next
        dbGen.dbClose
  Set dbGen = Nothing
  dbClose = True
    End Function

End Module
