Option Strict Off

Public Class MailInfo
    Public Property Subject As String
    Public Property SenderName As String
    Public Property SenderEmail As String
    Public Property Body As String
End Class

''' <summary>Legge le mail selezionate in Outlook classic via COM (late binding, nessun interop).</summary>
Public Module OutlookHelper
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
                result.Add(New MailInfo With {.Subject = CStr(m.Subject), .SenderName = CStr(m.SenderName), .SenderEmail = email, .Body = CStr(m.Body)})
            Catch ex As Exception
                ' elemento non-mail (appuntamento, riunione...): ignorato
            End Try
        Next
        Return result
    End Function
End Module
