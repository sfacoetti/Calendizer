Imports System.IO
Imports System.Text

''' <summary>Esporta attività in iCalendar (.ics), importabile in Outlook e altri calendari.</summary>
Public Module IcsExporter
    ' Orari "a muro" senza fuso (floating): l'evento resta alla stessa ora locale su qualsiasi calendario
    Private Const Fmt As String = "yyyyMMdd'T'HHmmss"

    Private Function Esc(s As String) As String
        Return If(s, "").Replace("\", "\\").Replace(";", "\;").Replace(",", "\,").Replace(vbCrLf, "\n").Replace(vbLf, "\n").Replace(vbCr, "\n")
    End Function

    ''' <summary>Righe oltre 75 byte vanno spezzate (RFC 5545); non si taglia mai in mezzo a un carattere.</summary>
    Private Sub AddLine(sb As StringBuilder, line As String)
        Dim bytes = 0
        Dim first = True
        For Each ch In line.EnumerateRunes()
            Dim n = ch.Utf8SequenceLength
            If bytes + n > If(first, 75, 74) Then
                sb.Append(vbCrLf).Append(" ")
                bytes = 0
                first = False
            End If
            sb.Append(ch.ToString())
            bytes += n
        Next
        sb.Append(vbCrLf)
    End Sub

    Public Function Build(items As IEnumerable(Of Activity)) As String
        Dim sb As New StringBuilder
        Dim stamp = DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'")
        AddLine(sb, "BEGIN:VCALENDAR")
        AddLine(sb, "VERSION:2.0")
        AddLine(sb, "PRODID:-//OrgaLendar//IT")
        For Each a In items
            AddLine(sb, "BEGIN:VEVENT")
            AddLine(sb, "UID:" & a.Id & "@orgalendar")
            AddLine(sb, "DTSTAMP:" & stamp)
            AddLine(sb, "DTSTART:" & a.Start.ToString(Fmt))
            AddLine(sb, "DTEND:" & a.EndTime().ToString(Fmt))
            AddLine(sb, "SUMMARY:" & Esc(If(a.Done, "✓ ", "") & a.Title))
            Dim desc As New List(Of String)
            If Not String.IsNullOrEmpty(a.Sender) Then desc.Add("Da: " & a.Sender & " <" & a.SenderEmail & ">")
            If Not String.IsNullOrEmpty(a.MailMessageId) Then desc.Add("Message-ID: " & a.MailMessageId)
            If Not String.IsNullOrEmpty(a.Body) Then desc.Add(a.Body)
            If desc.Count > 0 Then AddLine(sb, "DESCRIPTION:" & Esc(String.Join(vbLf, desc)))
            If a.Done Then AddLine(sb, "X-ORGALENDAR-DONE:TRUE")
            AddLine(sb, "END:VEVENT")
        Next
        AddLine(sb, "END:VCALENDAR")
        Return sb.ToString()
    End Function

    Public Sub Save(items As IEnumerable(Of Activity), path As String)
        File.WriteAllText(path, Build(items), New UTF8Encoding(False))
    End Sub
End Module
