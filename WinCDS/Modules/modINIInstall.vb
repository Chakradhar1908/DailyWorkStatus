Module modINIInstall
    Public Function VerifyProgramFeatures() As Boolean
        TryInstallINIFeature("amazon", True) ' this is the only one being lost regularly right now...  Add as needed
    End Function

    Public Function TryInstallINIFeature(ByVal FeatureName As String, Optional ByVal AsStartup As Boolean = False) As Boolean
        Dim Replaced As Boolean

        If Not FileExists(StoreFolder(1) & FeatureName & ".dat") Then Exit Function

        If IsFormLoaded("frmSetupAmazon") Then frmSetupAmazon.Cancelled = True : frmSetupAmazon.Hide
        If IsFormLoaded("frmSetupAshley") Then frmSetupAshley.Cancelled = True : frmSetupAshley.Hide
        If IsFormLoaded("frmSetupDispatchTrack") Then frmSetupDispatchTrack.Cancelled = True : frmSetupDispatchTrack.Hide
        If IsFormLoaded("frmSetupEquifax") Then frmSetupEquifax.Cancelled = True : frmSetupEquifax.Hide
        If IsFormLoaded("frmSetupPersonalization") Then frmSetupPersonalization.Cancelled = True : frmSetupPersonalization.Hide
        If IsFormLoaded("frmSetupRevolving") Then frmSetupRevolving.Cancelled = True : frmSetupRevolving.Hide
        If IsFormLoaded("frmSetupCCXCharge") Then frmSetupCCXCharge.Cancelled = True : frmSetupCCXCharge.Hide
        If IsFormLoaded("frmSetupCCTransactionCentral") Then frmSetupCCTransactionCentral.Cancelled = True : frmSetupCCTransactionCentral.Hide
        If IsFormLoaded("frmSetupCredomatic") Then frmSetupCCCredomatic.Cancelled = True : frmSetupCCCredomatic.Hide
        If IsFormLoaded("frmSetupTRAX") Then frmSetupTRAX.Cancelled = True : frmSetupTRAX.Hide

        If AsStartup Then
            Select Case LCase(FeatureName)
                Case "amazon"
                    If StoreSettings.AmazonKeyID = "" Then Replaced = True Else Exit Function
                Case Else
            End Select
        End If

        LogFile("InstallINI", "TryInstallINIFeature - " & FeatureName, False)

        InstallINIToStoreSettings(StoreFolder(1) & FeatureName & ".dat")
        ResetStoreSettings()
        If IsFormLoaded("frmSetup") Then LoadFrmSetupFromStoreInformation(StoreSettings)

        If Replaced Then LogFile("InstallINI", "TryInstallINIFeature - " & FeatureName & "  - Data was replaced.", False)

        TryInstallINIFeature = True
    End Function

End Module
