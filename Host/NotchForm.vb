Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

''' <summary>
''' Notch sul bordo destro dello schermo principale, centrata in verticale. A riposo è una striscia sottile; se il mouse ci resta
''' sopra 3 secondi si allarga in una pillola di vetro (acrilico di Windows 11) con l'icona dell'app. Un clic apre l'app.
''' </summary>
Public Class NotchForm
    Inherits Form

    Public Event Opened()

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As Integer, ByRef value As Integer, size As Integer) As Integer
    End Function

    Private Const DwellMs As Integer = 3000 ' permanenza sulla striscia prima che compaia: evita attivazioni per caso
    Private Const LeaveMs As Integer = 600  ' tempo fuori dalla pillola prima che si ritiri
    Private Const StripW As Single = 4
    Private Const StripH As Single = 90
    Private Const PillW As Single = 76
    Private Const PillH As Single = 124

    Private ReadOnly _icon As Icon
    Private ReadOnly _timer As New Timer() With {.Interval = 30}
    Private _scale As Single = 1
    Private _w As Single = StripW
    Private _expanded As Boolean
    Private _dwellStart As DateTime = DateTime.MinValue
    Private _leaveStart As DateTime = DateTime.MinValue
    Private _needLeave As Boolean ' dopo un clic: il mouse deve prima uscire dalla zona
    Private _glass As Boolean

    Public Sub New(appIcon As Icon)
        _icon = appIcon
        FormBorderStyle = FormBorderStyle.None
        ShowInTaskbar = False
        TopMost = True
        StartPosition = FormStartPosition.Manual
        BackColor = Color.Black ' sul vetro di Windows il nero diventa trasparente
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
        AddHandler _timer.Tick, AddressOf OnTick
    End Sub

    ' Non ruba il focus e non compare in Alt+Tab
    Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean = True
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H80 Or &H8000000 ' TOOLWINDOW | NOACTIVATE
            Return cp
        End Get
    End Property

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        _scale = DeviceDpi / 96.0F
        Dim noRound = 1 ' gli angoli li fa la Region
        DwmSetWindowAttribute(Handle, 33, noRound, 4)
        ' Sfocatura acrilica dietro la finestra (API non documentata ma stabile da Windows 10): a differenza dello
        ' sfondo DWM di Windows 11 rispetta la Region, quindi gli angoli restano arrotondati. Tinta bianca leggera = vetro chiaro.
        Dim accent As New AccentPolicy With {.AccentState = 4, .GradientColor = &H40FFFFFF}
        Dim size = Marshal.SizeOf(accent)
        Dim ptr = Marshal.AllocHGlobal(size)
        Try
            Marshal.StructureToPtr(accent, ptr, False)
            Dim data As New WindowCompositionAttributeData With {.Attribute = 19, .Data = ptr, .SizeOfData = size}
            _glass = SetWindowCompositionAttribute(Handle, data) <> 0
        Finally
            Marshal.FreeHGlobal(ptr)
        End Try
        If Not _glass Then ' ripiego: grigio scuro semitrasparente
            BackColor = Color.FromArgb(38, 42, 50)
            Opacity = 0.92
        End If
    End Sub

    <StructLayout(LayoutKind.Sequential)>
    Private Structure AccentPolicy
        Public AccentState As Integer, AccentFlags As Integer, GradientColor As Integer, AnimationId As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure WindowCompositionAttributeData
        Public Attribute As Integer
        Public Data As IntPtr
        Public SizeOfData As Integer
    End Structure

    <DllImport("user32.dll")>
    Private Shared Function SetWindowCompositionAttribute(hwnd As IntPtr, ByRef data As WindowCompositionAttributeData) As Integer
    End Function

    Public Sub ShowNotch()
        Show()
        ApplyGeometry() ' dopo Show: ora il DPI reale dello schermo è noto
        _timer.Start()
    End Sub

    Private Function Scr() As Rectangle
        Return Screen.PrimaryScreen.Bounds ' shortcut: solo lo schermo principale
    End Function

    Private Function Progress() As Single
        Return Math.Max(0, Math.Min(1, (_w - StripW) / (PillW - StripW)))
    End Function

    ' Dimensione, posizione (bordo destro, al centro in verticale) e sagoma arrotondata
    Private Sub ApplyGeometry()
        Dim p = Progress()
        Dim w = CInt(_w * _scale)
        Dim h = CInt((StripH + (PillH - StripH) * p) * _scale)
        Dim s = Scr()
        SetBounds(s.Right - w, s.Top + (s.Height - h) \ 2, w, h)
        Using path = Shape(w, h)
            Region = New Region(path)
        End Using
        Invalidate()
    End Sub

    ' Rettangolo con gli angoli sinistri arrotondati; la parte destra esce dallo schermo e viene tagliata
    Private Shared Function Shape(w As Integer, h As Integer) As GraphicsPath
        Dim r = Math.Max(1, Math.Min(w \ 2, h \ 2))
        Dim d = r * 2
        Dim fw = w + r
        Dim path As New GraphicsPath()
        path.AddArc(0, 0, d, d, 180, 90)
        path.AddArc(fw - d, 0, d, d, 270, 90)
        path.AddArc(fw - d, h - d, d, d, 0, 90)
        path.AddArc(0, h - d, d, d, 90, 90)
        path.CloseFigure()
        Return path
    End Function

    Private Sub OnTick(sender As Object, e As EventArgs)
        Dim pos = Cursor.Position
        Dim s = Scr()
        Dim now = DateTime.Now
        Dim zone As New Rectangle(s.Right - 6, s.Top + (s.Height - CInt(StripH * _scale)) \ 2, 6, CInt(StripH * _scale))
        If Not _expanded Then
            If Not zone.Contains(pos) Then
                _needLeave = False
                _dwellStart = DateTime.MinValue
            ElseIf Not _needLeave Then
                If _dwellStart = DateTime.MinValue Then _dwellStart = now
                If (now - _dwellStart).TotalMilliseconds >= DwellMs Then
                    _expanded = True
                    _leaveStart = DateTime.MinValue
                End If
            End If
        Else
            If Bounds.Contains(pos) Then
                _leaveStart = DateTime.MinValue
            ElseIf _leaveStart = DateTime.MinValue Then
                _leaveStart = now
            ElseIf (now - _leaveStart).TotalMilliseconds >= LeaveMs Then
                Collapse()
            End If
        End If
        Dim target = If(_expanded, PillW, StripW)
        If Math.Abs(target - _w) > 0.4F Then
            _w += (target - _w) * 0.3F
            ApplyGeometry()
        ElseIf _w <> target Then
            _w = target
            ApplyGeometry()
        End If
        Cursor = If(_expanded, Cursors.Hand, Cursors.Default)
    End Sub

    Private Sub Collapse()
        _expanded = False
        _dwellStart = DateTime.MinValue
        _needLeave = True
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        If Not _expanded Then Return
        Collapse()
        RaiseEvent Opened()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim w = Width, h = Height
        ' bordo chiaro + riflesso nella metà alta: l'effetto "vetro"
        Using path = Shape(w, h)
            ' a riposo la striscia deve essere "appena visibile": riempimento chiaro che sfuma mentre la pillola si apre
            Using fill As New SolidBrush(Color.FromArgb(CInt(170 * (1 - Progress())), 255, 255, 255))
                g.FillPath(fill, path)
            End Using
            Using edge As New Pen(Color.FromArgb(90, 40, 50, 70), 1.2F * _scale) ' bordo scuro: stacca il vetro dagli sfondi chiari
                g.DrawPath(edge, path)
            End Using
            Using pen As New Pen(Color.FromArgb(170, 255, 255, 255), 1.2F * _scale)
                g.TranslateTransform(1.2F * _scale, 1.2F * _scale) ' bordo chiaro appena dentro: effetto spessore
                g.DrawPath(pen, path)
                g.ResetTransform()
            End Using
            g.SetClip(path)
            Using br As New LinearGradientBrush(New Rectangle(0, 0, Math.Max(1, w), Math.Max(1, h \ 2)), Color.FromArgb(70, 255, 255, 255), Color.FromArgb(0, 255, 255, 255), 90.0F)
                g.FillRectangle(br, 0, 0, w, h \ 2)
            End Using
            g.ResetClip()
        End Using
        If Progress() > 0.5F Then
            Dim sz = CInt(36 * _scale)
            g.DrawIcon(_icon, New Rectangle((w - sz) \ 2 - CInt(4 * _scale), (h - sz) \ 2, sz, sz))
        End If
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then _timer.Dispose()
        MyBase.Dispose(disposing)
    End Sub
End Class
