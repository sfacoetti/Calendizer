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
        ' Una sola istanza: la seconda avvisa la prima di mostrarsi ed esce
        Dim createdNew As Boolean
        Using instanceLock As New Threading.Mutex(True, "OrgaLendar.SingleInstance", createdNew)
            Dim showEvt As New Threading.EventWaitHandle(False, Threading.EventResetMode.AutoReset, "OrgaLendar.Show")
            If Not createdNew Then
                showEvt.Set()
                Return
            End If
            Application.Run(New MainForm(Environment.GetCommandLineArgs().Contains("--tray"), showEvt))
        End Using
    End Sub
End Module

Public Class MainForm
    Inherits Form

    Private ReadOnly _state As New CalendarState(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OrgaLendar", "data.json"))
    Private ReadOnly _web As New BlazorWebView()

    Private ReadOnly _tray As New NotifyIcon()
    Private ReadOnly _timer As New Timer() With {.Interval = 30000}
    Private ReadOnly _autostart As New ToolStripMenuItem("Avvia con Windows")
    Private _startHidden As Boolean
    Private _exiting As Boolean

    ' All'avvio con Windows la finestra resta nascosta (solo icona accanto all'orologio)
    Protected Overrides Sub SetVisibleCore(value As Boolean)
        If _startHidden Then
            _startHidden = False
            value = False
        End If
        MyBase.SetVisibleCore(value)
    End Sub

    Private Sub ShowWindow()
        Show()
        WindowState = FormWindowState.Normal
        Activate()
        _tray.Icon = Icon ' torna l'icona normale: il promemoria è stato visto
    End Sub

    ' Icona normale con un pallino rosso: segnala un promemoria non ancora visto
    Private _alertIcon As Icon
    Private Function AlertIcon() As Icon
        If _alertIcon Is Nothing Then
            Using bmp As New Bitmap(Icon.ToBitmap(), 32, 32)
                Using g = Graphics.FromImage(bmp)
                    g.FillEllipse(Brushes.White, 16, 0, 16, 16)
                    g.FillEllipse(Brushes.Red, 18, 2, 12, 12)
                End Using
                _alertIcon = Icon.FromHandle(bmp.GetHicon()) ' shortcut: handle mai liberato, ne basta uno per sessione
            End Using
        End If
        Return _alertIcon
    End Function

    Private Sub ToggleAutostart(sender As Object, e As EventArgs)
        StartupHelper.SetEnabled(Not StartupHelper.IsEnabled())
        _autostart.Checked = StartupHelper.IsEnabled()
    End Sub

    ' Notch sul bordo destro: si crea/distrugge in base all'opzione nelle impostazioni
    Private _notch As NotchForm
    Private Sub ApplyNotch()
        If _state.Notch AndAlso _notch Is Nothing Then
            _notch = New NotchForm(Icon)
            AddHandler _notch.Opened, AddressOf ShowWindow
            _notch.ShowNotch()
        End If
        If _notch IsNot Nothing Then _notch.DwellMs = CInt(_state.NotchDelay * 1000)
        If Not _state.Notch AndAlso _notch IsNot Nothing Then
            _notch.Close()
            _notch.Dispose()
            _notch = Nothing
        End If
    End Sub

    Private Sub CheckReminders(sender As Object, e As EventArgs)
        Dim due = _state.TakeDueReminders(DateTime.Now)
        If due.Count = 0 Then Return
        Dim lines = due.Take(5).Select(Function(a) $"{a.Start:HH:mm} {a.Title}")
        If Not Visible Then _tray.Icon = AlertIcon() ' finestra nascosta: il pallino resta finché non la apri
        _tray.ShowBalloonTip(10000, "OrgaLendar", String.Join(vbLf, lines) & If(due.Count > 5, vbLf & $"…e altre {due.Count - 5}", ""), ToolTipIcon.Info)
    End Sub

    Public Sub New(startHidden As Boolean, showEvt As Threading.EventWaitHandle)
        _startHidden = startHidden
        CreateHandle() ' serve già ora (anche con avvio nascosto) per ricevere la richiesta della seconda istanza
        Threading.ThreadPool.RegisterWaitForSingleObject(showEvt, Sub(s, t) BeginInvoke(AddressOf ShowWindow), Nothing, -1, False)
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

        ' Area di notifica: chiudere la finestra la nasconde, si esce dal menu dell'icona
        _tray.Icon = Icon
        _tray.Text = "OrgaLendar"
        AddHandler _autostart.Click, AddressOf ToggleAutostart
        Dim menu As New ContextMenuStrip()
        AddHandler menu.Opening, Sub(s, e) _autostart.Checked = StartupHelper.IsEnabled() ' può essere cambiato dalle impostazioni
        menu.Items.Add("Apri", Nothing, Sub(s, e) ShowWindow())
        menu.Items.Add(_autostart)
        menu.Items.Add("Esci", Nothing, Sub(s, e)
                                             _exiting = True
                                             Close()
                                         End Sub)
        _tray.ContextMenuStrip = menu
        AddHandler _tray.DoubleClick, Sub(s, e) ShowWindow()
        AddHandler _tray.BalloonTipClicked, Sub(s, e) ShowWindow()
        _tray.Visible = True
        AddHandler _timer.Tick, AddressOf CheckReminders
        _timer.Start()
        CheckReminders(Nothing, EventArgs.Empty) ' recupera subito i promemoria scattati ad app chiusa
        AddHandler _state.Changed, Sub() BeginInvoke(AddressOf ApplyNotch)
        ApplyNotch()
        AddHandler FormClosing, Sub(s, e)
                                    If Not _exiting AndAlso e.CloseReason = CloseReason.UserClosing Then
                                        e.Cancel = True
                                        Hide()
                                    Else
                                        _tray.Visible = False
                                        _notch?.Close()
                                    End If
                                End Sub
    End Sub

    Private Sub OnWebInit(sender As Object, e As BlazorWebViewInitializedEventArgs)
        AddHandler _state.ExternalDrop, AddressOf OnExternalDrop
        AddHandler _state.ExportRequested, AddressOf OnExport
    End Sub

    Private Sub OnExport(fromDay As Date, toDay As Date)
        BeginInvoke(Sub()
                        Dim items = _state.Activities.Where(Function(a) a.Start.Date >= fromDay.Date AndAlso a.Start.Date <= toDay.Date).OrderBy(Function(a) a.Start).ToList()
                        If items.Count = 0 Then
                            MessageBox.Show("Nessuna attività nell'intervallo scelto.", "OrgaLendar")
                            Return
                        End If
                        Using dlg As New SaveFileDialog() With {.Filter = "Calendario (*.ics)|*.ics", .FileName = "OrgaLendar.ics", .DefaultExt = "ics"}
                            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
                            Try
                                IcsExporter.Save(items, dlg.FileName)
                                MessageBox.Show($"Esportate {items.Count} attività in {dlg.FileName}", "OrgaLendar")
                            Catch ex As Exception
                                MessageBox.Show("Impossibile salvare il file: " & ex.Message, "OrgaLendar", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            End Try
                        End Using
                    End Sub)
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
                                _state.AddMail(m, day, minutes)
                            Next
                        Catch ex As Exception
                            MessageBox.Show("Impossibile leggere la mail da Outlook: " & ex.Message, "OrgaLendar", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End Try
                    End Sub)
    End Sub
End Class
