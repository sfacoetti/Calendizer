Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.AspNetCore.Components.WebView
Imports Microsoft.AspNetCore.Components.WebView.WindowsForms
Imports Microsoft.Extensions.DependencyInjection
Imports OrgaLendar.Core

Public Module Program
    <STAThread>
    Public Sub Main()
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2)
        Application.EnableVisualStyles()
        Application.Run(New MainForm())
    End Sub
End Module

Public Class MainForm
    Inherits Form

    Private ReadOnly _state As New CalendarState(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OrgaLendar", "data.json"))
    Private ReadOnly _web As New BlazorWebView()

    Public Sub New()
        Text = "OrgaLendar"
        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath)
        Width = 1300
        Height = 820

        Dim sc As New ServiceCollection()
        sc.AddWindowsFormsBlazorWebView()
        sc.AddSingleton(_state)
        _web.HostPage = "wwwroot\index.html"
        _web.Services = sc.BuildServiceProvider()
        _web.Dock = DockStyle.Fill
        _web.RootComponents.Add(Of OrgaLendar.UI.App)("#app", Nothing)
        Controls.Add(_web)
        _web.BlazorWebViewInitialized = AddressOf OnWebInit ' nel pacchetto è un campo delegato, non un Event
    End Sub

    Private Sub OnWebInit(sender As Object, e As BlazorWebViewInitializedEventArgs)
        AddHandler _state.ExternalDrop, AddressOf OnExternalDrop
    End Sub


    ' Il drop esterno (Outlook) lo riceve la pagina, che calcola lo slot e avvisa qui.
    Private Sub OnExternalDrop(day As Date, minutes As Integer)
        ' Outlook è bloccato finché il drop non termina: la lettura COM va fatta dopo, non qui dentro.
        BeginInvoke(Async Sub()
                        Await Task.Delay(300)
                        Try
                            Dim mails = OutlookHelper.GetSelectedMails()
                            If mails.Count = 0 Then
                                MessageBox.Show("Nessuna mail selezionata in Outlook.", "OrgaLendar")
                                Return
                            End If
                            For Each m In mails
                                _state.AddMail(m.Subject, m.SenderName, m.SenderEmail, m.Body, day, minutes)
                            Next
                        Catch ex As Exception
                            MessageBox.Show("Impossibile leggere la mail da Outlook: " & ex.Message, "OrgaLendar", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End Try
                    End Sub)
    End Sub
End Class
