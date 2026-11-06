Option Strict Off

Public Class MailInfo
    Public Property Subject As String
    Public Property SenderName As String
    Public Property SenderEmail As String
    Public Property Body As String
    ' Collegamento alla mail: EntryID cambia se la mail viene spostata, StoreID e MessageId servono a ritrovarla
    Public Property EntryId As String = ""
    Public Property StoreId As String = ""
    Public Property Received As DateTime
    Public Property MessageId As String = ""
End Class

Public Class FolderInfo
    Public Property Path As String
    Public Property EntryId As String
    Public Property StoreId As String
End Class

''' <summary>Legge le mail selezionate in Outlook classic via COM (late binding, nessun interop).</summary>
Public Module OutlookHelper
    Private Const MsgIdProp As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
    Private Const NotFound As String = "Mail non trovata in Outlook: è stata eliminata o non è più raggiungibile."

    ''' <summary>Ritrova la mail: prima per EntryID+StoreID, poi (se spostata a mano) cercando il Message-ID. Aggiorna il collegamento dell'attività.</summary>
    Private Function ResolveMail(app As Object, a As Activity) As Object
        If String.IsNullOrEmpty(a.MailEntryId) Then Return Nothing
        Dim item As Object = Nothing
        Try : item = app.Session.GetItemFromID(a.MailEntryId, a.MailStoreId) : Catch : End Try
        If item Is Nothing AndAlso Not String.IsNullOrEmpty(a.MailMessageId) Then
            Dim filter As String = "@SQL=""" & MsgIdProp & """ = '" & a.MailMessageId.Replace("'", "''") & "'"
            ' shortcut: ricerca cartella per cartella in tutti gli archivi, lenta su caselle molto grandi
            For i = 1 To CInt(app.Session.Stores.Count)
                Try : item = FindIn(app.Session.Stores.Item(i).GetRootFolder(), filter) : Catch : End Try
                If item IsNot Nothing Then Exit For
            Next
        End If
        If item IsNot Nothing Then LinkTo(a, item)
        Return item
    End Function

    Private Function FindIn(folder As Object, filter As String) As Object
        If CInt(folder.DefaultItemType) = 0 Then
            Try
                Dim items As Object = folder.Items.Restrict(filter)
                If CInt(items.Count) > 0 Then Return items.Item(1)
            Catch
            End Try
        End If
        For i = 1 To CInt(folder.Folders.Count)
            Dim r As Object = Nothing
            Try : r = FindIn(folder.Folders.Item(i), filter) : Catch : End Try
            If r IsNot Nothing Then Return r
        Next
        Return Nothing
    End Function

    Private Sub LinkTo(a As Activity, item As Object)
        a.MailEntryId = CStr(item.EntryID)
        a.MailStoreId = CStr(item.Parent.StoreID)
    End Sub

    ''' <summary>Apre la mail in Outlook. Restituisce un messaggio d'errore, o Nothing se è andata a buon fine. Può aggiornare il collegamento di a.</summary>
    Public Function OpenMail(a As Activity) As String
        If String.IsNullOrEmpty(a.MailEntryId) Then Return "Questa attività non ha un collegamento a una mail."
        Try
            Dim item As Object = ResolveMail(CreateObject("Outlook.Application"), a)
            If item Is Nothing Then Return NotFound
            item.Display()
            Return Nothing
        Catch ex As Exception
            Return NotFound
        End Try
    End Function

    ''' <summary>Archivi (caselle) di Outlook: Path = nome, StoreId = identificativo. Veloce.</summary>
    Public Function ListStores() As List(Of FolderInfo)
        Dim result As New List(Of FolderInfo)
        Dim app As Object = CreateObject("Outlook.Application")
        For i = 1 To CInt(app.Session.Stores.Count)
            Try
                Dim s As Object = app.Session.Stores.Item(i)
                result.Add(New FolderInfo With {.Path = CStr(s.DisplayName), .StoreId = CStr(s.StoreID)})
            Catch
            End Try
        Next
        Return result
    End Function

    ''' <summary>Cartelle di posta di un solo archivio. Lenta (decine di secondi su caselle Exchange online): chiamare fuori dal thread della UI.</summary>
    Public Function ListFolders(storeId As String) As List(Of FolderInfo)
        Dim result As New List(Of FolderInfo)
        Dim app As Object = CreateObject("Outlook.Application")
        For i = 1 To CInt(app.Session.Stores.Count)
            Dim s As Object = app.Session.Stores.Item(i)
            If CStr(s.StoreID) <> storeId Then Continue For
            Dim root As Object = s.GetRootFolder()
            AddFolders(root, CStr(root.Name), result)
            Exit For
        Next
        Return result
    End Function

    Private Sub AddFolders(parent As Object, path As String, result As List(Of FolderInfo))
        For i = 1 To CInt(parent.Folders.Count)
            Try
                Dim f As Object = parent.Folders.Item(i)
                Dim p As String = path & " \ " & CStr(f.Name)
                If CInt(f.DefaultItemType) = 0 Then result.Add(New FolderInfo With {.Path = p, .EntryId = CStr(f.EntryID), .StoreId = CStr(f.StoreID)})
                AddFolders(f, p, result)
            Catch
            End Try
        Next
    End Sub

    ''' <summary>Sposta la mail dell'attività nella cartella indicata e aggiorna il collegamento. Restituisce un errore, o Nothing se ok.</summary>
    Public Function MoveMail(a As Activity, folderEntryId As String, folderStoreId As String) As String
        Try
            Dim app As Object = CreateObject("Outlook.Application")
            Dim item As Object = ResolveMail(app, a)
            If item Is Nothing Then Return NotFound
            Dim moved As Object = item.Move(app.Session.GetFolderFromID(folderEntryId, folderStoreId))
            LinkTo(a, moved)
            Return Nothing
        Catch ex As Exception
            Return "Impossibile spostare la mail: " & ex.Message
        End Try
    End Function

    Private Const CalName As String = "OrgaLendar"

    ' Il calendario "OrgaLendar" è una sottocartella del calendario predefinito: mai si tocca il calendario vero
    Private Function GetSyncFolder(app As Object) As Object
        Dim cal As Object = app.Session.GetDefaultFolder(9) ' olFolderCalendar
        For i = 1 To CInt(cal.Folders.Count)
            If CStr(cal.Folders.Item(i).Name) = CalName Then Return cal.Folders.Item(i)
        Next
        Return cal.Folders.Add(CalName, 9)
    End Function

    ''' <summary>Evento già collegato all'attività, solo se sta davvero nel calendario OrgaLendar.</summary>
    Private Function FindEvent(app As Object, folder As Object, id As String) As Object
        If String.IsNullOrEmpty(id) Then Return Nothing
        Try
            Dim item As Object = app.Session.GetItemFromID(id, folder.StoreID)
            If CStr(item.Parent.EntryID) = CStr(folder.EntryID) Then Return item
        Catch
        End Try
        Return Nothing
    End Function

    Private Function EventBody(a As Activity) As String
        Dim parts As New List(Of String)
        If Not String.IsNullOrEmpty(a.Sender) Then parts.Add("Da: " & a.Sender & " <" & a.SenderEmail & ">")
        If Not String.IsNullOrEmpty(a.Body) Then parts.Add(a.Body)
        Return String.Join(vbCrLf, parts)
    End Function

    Private Function OrgaId(ev As Object) As String
        Try
            Dim p As Object = ev.UserProperties.Find("OrgaId")
            Return If(p Is Nothing, "", CStr(p.Value))
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Rende il calendario OrgaLendar uguale all'app: un evento per attività, riconosciuto dalla proprietà nascosta OrgaId.
    ''' Quello che nella cartella non corrisponde a nessuna attività (eliminate, duplicati, eventi senza OrgaId) viene eliminato
    ''' (finisce in Elementi eliminati). La cartella è solo di OrgaLendar: non metterci eventi propri.
    ''' </summary>
    Public Function SyncCalendar(items As List(Of Activity)) As String
        Dim app As Object = CreateObject("Outlook.Application")
        Dim folder As Object = GetSyncFolder(app)
        Dim created = 0, updated = 0, deleted = 0, failed = 0
        ' stato attuale della cartella: OrgaId -> evento (il primo vince, gli altri sono duplicati da eliminare)
        Dim byId As New Dictionary(Of String, Object)
        Dim unclaimed As New List(Of Object)
        For i = 1 To CInt(folder.Items.Count)
            Dim ev As Object = folder.Items.Item(i)
            Dim oid As String = OrgaId(ev)
            If oid <> "" AndAlso Not byId.ContainsKey(oid) Then byId(oid) = ev Else unclaimed.Add(ev)
        Next
        Dim claimed As New HashSet(Of String)
        For Each a In items
            Try
                Dim ev As Object = Nothing
                If Not byId.TryGetValue(a.Id, ev) Then ev = FindEvent(app, folder, a.OutlookEventId) ' eventi creati prima dell'OrgaId
                If ev IsNot Nothing Then
                    If claimed.Contains(CStr(ev.EntryID)) Then ev = Nothing ' già assegnato a un'altra attività: se ne crea uno nuovo
                End If
                Dim isNew As Boolean = (ev Is Nothing)
                Dim prevKey As String = If(isNew, "", CStr(ev.EntryID)) ' l'ID com'era nella scansione, prima di salvare
                If isNew Then ev = folder.Items.Add(1) ' olAppointmentItem
                Dim subj As String = If(a.Done, "✓ ", "") & a.Title
                Dim body As String = EventBody(a)
                If isNew OrElse OrgaId(ev) <> a.Id OrElse CStr(ev.Subject) <> subj OrElse CDate(ev.Start) <> a.Start OrElse CDate(ev.End) <> a.EndTime() OrElse CStr(ev.Body) <> body Then
                    Dim p As Object = ev.UserProperties.Find("OrgaId")
                    If p Is Nothing Then p = ev.UserProperties.Add("OrgaId", 1) ' olText
                    p.Value = a.Id
                    ev.Subject = subj
                    ev.Start = a.Start
                    ev.End = a.EndTime()
                    ev.Body = body
                    ev.ReminderSet = False ' i promemoria li gestisce OrgaLendar
                    ev.Save()
                    If isNew Then created += 1 Else updated += 1
                End If
                a.OutlookEventId = CStr(ev.EntryID)
                claimed.Add(a.OutlookEventId)
                If prevKey <> "" Then claimed.Add(prevKey)
                byId.Remove(a.Id)
            Catch
                failed += 1
            End Try
        Next
        ' tutto ciò che resta (eventi di attività eliminate, duplicati, senza OrgaId) va via
        ' (gli eventi di attività esistenti ma andate in errore restano: si riprova alla prossima)
        Dim live As New HashSet(Of String)(items.Select(Function(x) x.Id))
        unclaimed.AddRange(byId.Where(Function(kv) Not live.Contains(kv.Key)).Select(Function(kv) kv.Value))
        For Each ev In unclaimed
            Try
                If Not claimed.Contains(CStr(ev.EntryID)) Then ev.Delete() : deleted += 1
            Catch
                failed += 1
            End Try
        Next
        Return $"Calendario Outlook ""{CalName}"": {created} create, {updated} aggiornate, {deleted} eliminate" & If(failed > 0, $", {failed} con errore", "") & "."
    End Function

    Public Function GetSelectedMails() As List(Of MailInfo)
        Dim result As New List(Of MailInfo)
        Dim app As Object = CreateObject("Outlook.Application")
        Dim sel As Object = app.ActiveExplorer().Selection
        For i = 1 To CInt(sel.Count)
            Dim m As Object = sel.Item(i)
            Try
                Dim email As String = CStr(m.SenderEmailAddress)
                If CStr(m.SenderEmailType) = "EX" Then ' Exchange: l'indirizzo è un DN, serve l'SMTP
                    Try
                        email = CStr(m.Sender.GetExchangeUser().PrimarySmtpAddress)
                    Catch
                    End Try
                End If
                Dim info As New MailInfo With {.Subject = CStr(m.Subject), .SenderName = CStr(m.SenderName), .SenderEmail = email, .Body = CStr(m.Body)}
                ' i dati di collegamento sono facoltativi: se uno manca la mail si importa comunque
                Try : info.EntryId = CStr(m.EntryID) : Catch : End Try
                Try : info.StoreId = CStr(m.Parent.StoreID) : Catch : End Try
                Try : info.Received = CDate(m.ReceivedTime) : Catch : End Try
                Try : info.MessageId = CStr(m.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")) : Catch : End Try
                result.Add(info)
            Catch ex As Exception
                ' elemento non-mail (appuntamento, riunione...): ignorato
            End Try
        Next
        Return result
    End Function
End Module
