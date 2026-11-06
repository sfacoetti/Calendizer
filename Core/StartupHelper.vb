''' <summary>Avvio con Windows tramite la chiave Run dell'utente (nessun permesso di amministratore). L'app parte nascosta con --tray.</summary>
Public Module StartupHelper
    Private Const RunKey As String = "Software\Microsoft\Windows\CurrentVersion\Run"
    Private Const ValueName As String = "Calendizer"

    Public Function IsEnabled() As Boolean
        Using k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunKey)
            Return k IsNot Nothing AndAlso k.GetValue(ValueName) IsNot Nothing
        End Using
    End Function

    Public Sub SetEnabled(enabled As Boolean)
        Using k = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RunKey)
            If enabled Then
                k.SetValue(ValueName, $"""{Environment.ProcessPath}"" --tray")
            Else
                k.DeleteValue(ValueName, False)
            End If
        End Using
    End Sub
End Module
