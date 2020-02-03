Module modDiagnostics
    Public Sub LogFolders()
        ActiveLog "Init::InventFolder: " & InventFolder(), 1
  ActiveLog "Init::PXFolder: " & PXFolder(), 1
  ActiveLog "Init::FXFolder: " & FXFolder(), 1
  ActiveLog "Init::AppFolder: " & AppFolder(), 1
  ActiveLog "Init::UpdateFolder: " & UpdateFolder(), 1
  If IsDevelopment() Then ActiveLog "Init::DEV MODE", 1

  LogInformationFiles
    End Sub

End Module
