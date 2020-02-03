Module modINIInstall
    Public Function VerifyProgramFeatures() As Boolean
        TryInstallINIFeature "amazon", True ' this is the only one being lost regularly right now...  Add as needed
    End Function

End Module
