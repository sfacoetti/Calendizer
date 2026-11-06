Imports System.IO
Imports System.Text.Json

Public Enum ViewMode
    Day
    WorkWeek
    Week
    Month
End Enum

Public Class Activity
    Public Property Id As String = Guid.NewGuid().ToString("N")
    Public Property Title As String = ""
    Public Property Sender As String = ""
    Public Property SenderEmail As String = ""
    Public Property Body As String = ""
    ' Collegamento alla mail Outlook; vuoti per attività libere e per quelle create prima di questa funzione
    Public Property Done As Boolean = False
    ''' <summary>EntryID dell'evento nel calendario Outlook "Calendizer" (sincronizzazione a una via); vuoto = mai sincronizzata.</summary>
    Public Property OutlookEventId As String = ""
    ''' <summary>Preavviso in minuti per questa attività; Nothing = usa quello predefinito, 0 = nessuno.</summary>
    Public Property Reminder As Integer?
    Public Property Notified As Boolean = False
    Public Property MailEntryId As String = ""
    Public Property MailStoreId As String = ""
    Public Property MailReceived As DateTime?
    Public Property MailMessageId As String = ""
    Private _start As DateTime
    ' Kind sempre Unspecified: orario "a muro", senza fuso nel JSON
    Public Property Start As DateTime
        Get
            Return _start
        End Get
        Set(v As DateTime)
            _start = DateTime.SpecifyKind(v, DateTimeKind.Unspecified)
        End Set
    End Property
    Public Property Minutes As Integer = 30
    ' <summary>Colore scelto a mano; vuoto = automatico per mittente.</summary>
    Public Property Color As String = ""
    Public Function EndTime() As DateTime
        Return Start.AddMinutes(Minutes)
    End Function
End Class

Public Class ColorRule
    ''' <summary>Dominio (@hexpol.com) o indirizzo completo.</summary>
    Public Property Key As String = ""
    Public Property Color As String = "#9e9e9e"
End Class

Public Class LayoutItem
    Public Property Act As Activity
    Public Property Col As Integer
    Public Property Cols As Integer = 1
End Class

Public Class StoreData
    Public Property Activities As New List(Of Activity)
    Public Property Rules As New List(Of ColorRule)
    ''' <summary>Preavviso predefinito in minuti; 0 = disattivato.</summary>
    Public Property Reminder As Integer = 0
    ''' <summary>"auto" (segue Windows), "light" o "dark".</summary>
    Public Property Theme As String = "auto"
    Public Property HideDone As Boolean = False
    Public Property Notch As Boolean = False
    Public Property NotchDelay As Double = 1.5
    Public Property LastArchive As String = ""
End Class

Public Class CalendarState
    Public Const FirstMinute As Integer = 8 * 60
    Public Const LastMinute As Integer = 19 * 60
    Public Const SlotMinutes As Integer = 30

    Private ReadOnly _path As String
    Private _data As New StoreData

    Public Property View As ViewMode = ViewMode.WorkWeek
    Public Property CurrentDate As Date = Date.Today
    Public Property SidebarOpen As Boolean = True
    Public Event Changed()
    Public Event ExternalDrop(day As Date, minutes As Integer)
    ''' <summary>La UI chiede l'esportazione; il file lo sceglie e lo scrive l'Host (finestra "Salva con nome").</summary>
    Public Event ExportRequested(fromDay As Date, toDay As Date)

    Public Sub RequestExport(fromDay As Date, toDay As Date)
        RaiseEvent ExportRequested(fromDay, toDay)
    End Sub

    Public ReadOnly Property Activities As List(Of Activity)
        Get
            Return _data.Activities
        End Get
    End Property

    Public ReadOnly Property Rules As List(Of ColorRule)
        Get
            Return _data.Rules
        End Get
    End Property

    Public Sub New(path As String)
        _path = path
        BackupDaily()
        Try
            If File.Exists(_path) Then _data = JsonSerializer.Deserialize(Of StoreData)(File.ReadAllText(_path)) ' errore di lettura: si parte vuoti ma il file resta intatto finché non si salva
        Catch ex As JsonException
            File.Copy(_path, _path & ".bak", True) ' backup prima che un salvataggio sovrascriva il file illeggibile
            _data = New StoreData
        End Try
        If _data.Rules.Count = 0 Then
            _data.Rules.Add(New ColorRule With {.Key = "@hexpol.com", .Color = "#1e90ff"})
            _data.Rules.Add(New ColorRule With {.Key = "@in4tek.it", .Color = "#e53935"})
            _data.Rules.Add(New ColorRule With {.Key = "@eurgomma.it", .Color = "#43a047"})
        End If
    End Sub

    ''' <summary>Una copia al giorno in backup\data-AAAAMMGG.json, tenute le ultime 7. Un errore qui non deve impedire l'avvio.</summary>
    Private Sub BackupDaily()
        Try
            If Not File.Exists(_path) Then Return
            Dim dir = Path.Combine(Path.GetDirectoryName(_path), "backup")
            Directory.CreateDirectory(dir)
            Dim dest = Path.Combine(dir, $"data-{Date.Today:yyyyMMdd}.json")
            If Not File.Exists(dest) Then File.Copy(_path, dest)
            For Each old In Directory.GetFiles(dir, "data-*.json").OrderByDescending(Function(f) f).Skip(7)
                File.Delete(old)
            Next
        Catch
        End Try
    End Sub

    ''' <summary>Salva e notifica la UI.</summary>
    Public Sub Touch()
        Directory.CreateDirectory(Path.GetDirectoryName(_path))
        File.WriteAllText(_path, JsonSerializer.Serialize(_data))
        RaiseEvent Changed()
    End Sub

    ''' <summary>Attività libera (non da mail): il colore è sempre esplicito.</summary>
    Public Sub AddNote(title As String, body As String, color As String, day As Date, minutes As Integer, duration As Integer)
        Activities.Add(New Activity With {.Title = title, .Body = body, .Color = color, .Start = day.Date.AddMinutes(minutes),
                                          .Minutes = Math.Max(SlotMinutes, Math.Min(duration, LastMinute - minutes))})
        Touch()
    End Sub

    Public Sub RaiseExternalDrop(day As Date, minutes As Integer)
        RaiseEvent ExternalDrop(day, minutes)
    End Sub

    Public Sub Navigate()
        RaiseEvent Changed()
    End Sub

    ''' <param name="minutes">Minuti dalla mezzanotte; negativo = 9:00.</param>
    Public Sub AddMail(m As MailInfo, day As Date, minutes As Integer)
        If minutes < 0 Then minutes = 9 * 60
        Activities.Add(New Activity With {.Title = m.Subject, .Sender = m.SenderName, .SenderEmail = m.SenderEmail, .Body = m.Body,
                                          .MailEntryId = m.EntryId, .MailStoreId = m.StoreId, .MailMessageId = m.MessageId,
                                          .MailReceived = If(m.Received = DateTime.MinValue, CType(Nothing, DateTime?), m.Received),
                                          .Start = day.Date.AddMinutes(minutes)})
        Touch()
    End Sub

    ''' <param name="minutes">Minuti dalla mezzanotte; negativo = mantiene l'orario (vista mese).</param>
    Public Sub Move(id As String, day As Date, minutes As Integer)
        Dim a = Activities.FirstOrDefault(Function(x) x.Id = id)
        If a Is Nothing Then Return
        a.Notified = False ' spostata: il promemoria vale per il nuovo orario
        If minutes < 0 Then
            a.Start = day.Date + a.Start.TimeOfDay
        Else
            minutes = Math.Max(FirstMinute, Math.Min(minutes, LastMinute - a.Minutes))
            a.Start = day.Date.AddMinutes(minutes)
        End If
        Touch()
    End Sub

    Public Sub Resize(id As String, minutes As Integer)
        Dim a = Activities.FirstOrDefault(Function(x) x.Id = id)
        If a Is Nothing Then Return
        Dim startMin = CInt(a.Start.TimeOfDay.TotalMinutes)
        a.Minutes = Math.Max(SlotMinutes, Math.Min(minutes, LastMinute - startMin))
        Touch()
    End Sub

    Public Sub Remove(id As String)
        Activities.RemoveAll(Function(x) x.Id = id)
        Touch()
    End Sub

    ''' <summary>Sincronizzazione a una via (app → Outlook). Lenta: chiamare fuori dal thread della UI. Restituisce il riepilogo.</summary>
    Public Function SyncOutlook() As String
        Dim items = Activities.ToList()
        Dim summary = OutlookHelper.SyncCalendar(items)
        Touch()
        Return summary
    End Function

    Public Function ColorOf(a As Activity) As String
        Return If(String.IsNullOrEmpty(a.Color), ColorFor(a.SenderEmail), a.Color)
    End Function

    Public Property Theme As String
        Get
            Return _data.Theme
        End Get
        Set(v As String)
            _data.Theme = v
            Touch()
        End Set
    End Property

    ''' <summary>Notch sul bordo destro dello schermo (la disegna l'Host).</summary>
    Public Property Notch As Boolean
        Get
            Return _data.Notch
        End Get
        Set(v As Boolean)
            _data.Notch = v
            Touch()
        End Set
    End Property

    ''' <summary>Secondi di permanenza del mouse sulla striscia prima che compaia la notch.</summary>
    Public Property NotchDelay As Double
        Get
            Return _data.NotchDelay
        End Get
        Set(v As Double)
            _data.NotchDelay = v
            Touch()
        End Set
    End Property

    Public Property HideDone As Boolean
        Get
            Return _data.HideDone
        End Get
        Set(v As Boolean)
            _data.HideDone = v
            Touch()
        End Set
    End Property

    ''' <summary>Ultima cartella Outlook scelta per l'archiviazione (EntryID), salvata tra le sessioni.</summary>
    Public Property LastArchive As String
        Get
            Return _data.LastArchive
        End Get
        Set(v As String)
            _data.LastArchive = v
            Touch()
        End Set
    End Property

    Public Property DefaultReminder As Integer
        Get
            Return _data.Reminder
        End Get
        Set(v As Integer)
            _data.Reminder = v
            For Each a In Activities : a.Notified = False : Next
            Touch()
        End Set
    End Property

    Public Function ReminderOf(a As Activity) As Integer
        Return If(a.Reminder, _data.Reminder)
    End Function

    Public Sub SetReminder(id As String, minutes As Integer?)
        Dim a = Activities.FirstOrDefault(Function(x) x.Id = id)
        If a Is Nothing Then Return
        a.Reminder = minutes
        a.Notified = False
        Touch()
    End Sub

    ''' <summary>Attività il cui preavviso è scattato e non ancora notificate (anche saltate a app chiusa); le segna come notificate.</summary>
    Public Function TakeDueReminders(now As DateTime) As List(Of Activity)
        ' shortcut: recupera solo le ultime 24 h, per non notificare attività vecchie alla prima attivazione
        Dim due = Activities.Where(Function(a) Not a.Done AndAlso Not a.Notified AndAlso ReminderOf(a) > 0 AndAlso
                                       now >= a.Start.AddMinutes(-ReminderOf(a)) AndAlso a.Start > now.AddHours(-24)).OrderBy(Function(a) a.Start).ToList()
        If due.Count > 0 Then
            due.ForEach(Sub(a) a.Notified = True)
            Touch()
        End If
        Return due
    End Function

    Public Sub ToggleDone(id As String)
        Dim a = Activities.FirstOrDefault(Function(x) x.Id = id)
        If a Is Nothing Then Return
        a.Done = Not a.Done
        Touch()
    End Sub

    Public Sub SetColor(id As String, color As String)
        Dim a = Activities.FirstOrDefault(Function(x) x.Id = id)
        If a Is Nothing Then Return
        a.Color = color
        Touch()
    End Sub

    ''' <summary>Regola sull'indirizzo completo, poi sul dominio.</summary>
    Public Function ColorFor(email As String) As String
        email = If(email, "").Trim()
        Dim domain = If(email.Contains("@"), email.Substring(email.LastIndexOf("@"c)), "")
        Dim hit = Rules.FirstOrDefault(Function(x) String.Equals(x.Key.Trim(), email, StringComparison.OrdinalIgnoreCase))
        If hit Is Nothing Then hit = Rules.FirstOrDefault(Function(x) String.Equals(x.Key.Trim().TrimStart("@"c), domain.TrimStart("@"c), StringComparison.OrdinalIgnoreCase) AndAlso domain <> "")
        Return If(hit IsNot Nothing, hit.Color, "#9e9e9e")
    End Function

    ''' <summary>Affianca le attività sovrapposte di un giorno (stile Outlook).</summary>
    Public Function LayoutDay(day As Date) As List(Of LayoutItem)
        Dim items = Activities.Where(Function(a) a.Start.Date = day.Date AndAlso Not (HideDone AndAlso a.Done)).OrderBy(Function(a) a.Start).ThenByDescending(Function(a) a.Minutes)
        Dim result As New List(Of LayoutItem)
        Dim cluster As New List(Of LayoutItem)
        Dim colEnds As New List(Of DateTime)
        Dim clusterEnd = DateTime.MinValue
        For Each a In items
            If cluster.Count > 0 AndAlso a.Start >= clusterEnd Then
                Flush(cluster, colEnds.Count, result)
                colEnds.Clear()
                clusterEnd = DateTime.MinValue
            End If
            Dim c = colEnds.FindIndex(Function(e) e <= a.Start)
            If c < 0 Then
                colEnds.Add(a.EndTime())
                c = colEnds.Count - 1
            Else
                colEnds(c) = a.EndTime()
            End If
            cluster.Add(New LayoutItem With {.Act = a, .Col = c})
            If a.EndTime() > clusterEnd Then clusterEnd = a.EndTime()
        Next
        Flush(cluster, colEnds.Count, result)
        Return result
    End Function

    Private Shared Sub Flush(cluster As List(Of LayoutItem), cols As Integer, result As List(Of LayoutItem))
        For Each i In cluster
            i.Cols = Math.Max(1, cols)
        Next
        result.AddRange(cluster)
        cluster.Clear()
    End Sub
End Class
