Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

''' <summary>
''' Popup dei promemoria in basso a destra: campanella che squilla in alto a sinistra, X in alto a destra, una riga per attività
''' con il colore dell'attività (come nel calendario) e bordo che lampeggia nel colore della prima. Clic: apre l'app; Esc o X: chiude.
''' </summary>
Public Class ReminderPopup
    Inherits Form

    Public Event Opened()

    Private Const W As Integer = 400, HeadH As Integer = 34, Pad As Integer = 10, Gap As Integer = 6, Border As Integer = 3
    Private ReadOnly _items As New List(Of (Text As String, Color As Color, Rect As Rectangle))
    Private ReadOnly _more As String
    Private ReadOnly _moreRect As Rectangle
    Private ReadOnly _accent As Color
    Private ReadOnly _timer As New Timer() With {.Interval = 40}
    Private ReadOnly _textFont As New Font("Segoe UI", 10)
    Private ReadOnly _headFont As New Font("Segoe UI", 9.5F, FontStyle.Bold)
    Private ReadOnly _bellFont As New Font("Segoe UI Emoji", 12)
    Private _tick As Integer
    Private _k As Single = 1

    Public Sub New(items As IEnumerable(Of (Text As String, Color As String)), more As Integer)
        FormBorderStyle = FormBorderStyle.None
        ShowInTaskbar = False
        TopMost = True
        StartPosition = FormStartPosition.Manual
        BackColor = Color.White
        Text = "Calendizer – promemoria"
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
        _k = DeviceDpi / 96.0F
        Dim y = S(HeadH)
        For Each it In items
            Dim c = ParseColor(it.Color)
            Dim h = TextRenderer.MeasureText(it.Text, _textFont, New Size(S(W - 2 * Pad - 18), 0), TextFormatFlags.WordBreak).Height + S(12)
            _items.Add((it.Text, c, New Rectangle(S(Pad), y, S(W - 2 * Pad), h)))
            y += h + S(Gap)
        Next
        If more > 0 Then
            _more = $"…e altre {more}"
            _moreRect = New Rectangle(S(Pad), y, S(W - 2 * Pad), S(20))
            y += S(20) + S(Gap)
        End If
        _accent = If(_items.Count > 0, _items(0).Color, Color.SteelBlue)
        ClientSize = New Size(S(W), y + S(Pad - Gap))
        AddHandler _timer.Tick, Sub(s, e)
                                    _tick += 1
                                    Invalidate()
                                End Sub
        _timer.Start()
    End Sub

    Private Function S(v As Integer) As Integer
        Return CInt(v * _k)
    End Function

    Private Shared Function ParseColor(html As String) As Color
        Try
            Return ColorTranslator.FromHtml(html)
        Catch
            Return Color.Gray ' colore non valido nei dati: il popup deve comparire comunque
        End Try
    End Function

    Private ReadOnly Property CloseRect As Rectangle
        Get
            Return New Rectangle(ClientSize.Width - S(32), S(5), S(24), S(24))
        End Get
    End Property

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit

        ' Campanella: per ~1 s ogni 3 s oscilla attorno al gancio, poi sta ferma
        Dim phase = _tick Mod 75
        Dim angle = If(phase < 25, CSng(22 * Math.Sin(phase * 0.9) * (1 - phase / 25.0)), 0F)
        Dim pivot As New PointF(S(20), S(8))
        Dim st = g.Save()
        g.TranslateTransform(pivot.X, pivot.Y)
        g.RotateTransform(angle)
        Using b As New SolidBrush(_accent)
            g.DrawString(ChrW(&HD83D) & ChrW(&HDD14), _bellFont, b, -S(11), -S(3))
        End Using
        g.Restore(st)
        TextRenderer.DrawText(g, "Promemoria", _headFont, New Point(S(36), S(9)), Color.FromArgb(34, 34, 34))
        TextRenderer.DrawText(g, "✕", _headFont, CloseRect, Color.FromArgb(102, 102, 102), TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)

        ' Righe: sfondo tenue + barra a sinistra nel colore dell'attività, come nel calendario
        For Each it In _items
            Using bg As New SolidBrush(Color.FromArgb(64, it.Color)), bar As New SolidBrush(it.Color)
                g.FillRectangle(bg, it.Rect)
                g.FillRectangle(bar, it.Rect.X, it.Rect.Y, S(4), it.Rect.Height)
            End Using
            Dim tr = New Rectangle(it.Rect.X + S(12), it.Rect.Y + S(6), it.Rect.Width - S(18), it.Rect.Height - S(6))
            TextRenderer.DrawText(g, it.Text, _textFont, tr, Color.FromArgb(34, 34, 34), TextFormatFlags.WordBreak)
        Next
        If _more IsNot Nothing Then TextRenderer.DrawText(g, _more, _textFont, _moreRect, Color.FromArgb(102, 102, 102), TextFormatFlags.Left)

        ' Bordo lampeggiante (~2 volte al secondo) nel colore della prima attività
        Dim blink = (_tick \ 12) Mod 2 = 0
        Using p As New Pen(If(blink, _accent, Color.FromArgb(221, 221, 221)), S(Border)) With {.Alignment = PenAlignment.Inset}
            g.SmoothingMode = SmoothingMode.None
            g.DrawRectangle(p, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1)
        End Using
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        Cursor = Cursors.Hand
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        If Not CloseRect.Contains(e.Location) Then RaiseEvent Opened()
        Close()
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If keyData = Keys.Escape Then
            Close()
            Return True
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            _timer.Dispose()
            _textFont.Dispose()
            _headFont.Dispose()
            _bellFont.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
