Imports System.IO
Imports System.IO.Compression
Imports System.Net.Http
Imports System.Text.Json

''' <summary>Aggiornamento da GitHub Releases (repo pubblico, nessun token). La release deve avere un asset .zip con exe + wwwroot + WebView2Loader.dll alla radice.</summary>
Public Module UpdateHelper
    Private Const Repo As String = "sfacoetti/Calendizer"

    Public Function CurrentVersion() As Version
        Return Norm(Reflection.Assembly.GetEntryAssembly().GetName().Version)
    End Function

    ' 1.0.0 e 1.0.0.0 devono risultare uguali
    Private Function Norm(v As Version) As Version
        Return New Version(v.Major, Math.Max(v.Minor, 0), Math.Max(v.Build, 0))
    End Function

    ''' <summary>Restituisce (versione, url dello zip) se su GitHub c'è una release più nuova, altrimenti Nothing. Gli errori di rete risalgono al chiamante.</summary>
    Public Async Function CheckAsync() As Task(Of (Version As Version, ZipUrl As String)?)
        Using http = NewClient()
            Dim json = Await http.GetStringAsync($"https://api.github.com/repos/{Repo}/releases/latest")
            Using doc = JsonDocument.Parse(json)
                Dim root = doc.RootElement
                Dim v As Version = Nothing
                If Not Version.TryParse(root.GetProperty("tag_name").GetString().TrimStart("v"c), v) Then Return Nothing
                v = Norm(v)
                If v <= CurrentVersion() Then Return Nothing
                For Each a In root.GetProperty("assets").EnumerateArray()
                    If a.GetProperty("name").GetString().EndsWith(".zip", StringComparison.OrdinalIgnoreCase) Then
                        Return (v, a.GetProperty("browser_download_url").GetString())
                    End If
                Next
            End Using
        End Using
        Return Nothing
    End Function

    ''' <summary>Scarica e decomprime lo zip, poi lancia uno script che attende la chiusura dell'app, fa un backup della cartella, copia i file e riavvia. Poi chiude l'app.</summary>
    Public Async Function ApplyAsync(zipUrl As String) As Task
        Dim work = Path.Combine(Path.GetTempPath(), "Calendizer-update")
        If Directory.Exists(work) Then Directory.Delete(work, True)
        Dim newDir = Path.Combine(work, "new")
        Directory.CreateDirectory(work)
        Dim zipPath = Path.Combine(work, "u.zip")
        Using http = NewClient()
            Using s = Await http.GetStreamAsync(zipUrl)
                Using f = File.Create(zipPath)
                    Await s.CopyToAsync(f)
                End Using
            End Using
        End Using
        ZipFile.ExtractToDirectory(zipPath, newDir)
        Dim exe = Environment.ProcessPath
        Dim inst = AppContext.BaseDirectory.TrimEnd("\"c)
        Dim bak = Path.Combine(work, "backup")
        ' data.json sta in %APPDATA%, non nella cartella dell'app: l'update non lo tocca
        File.WriteAllText(Path.Combine(work, "update.cmd"),
            "@echo off" & vbCrLf &
            "ping -n 4 127.0.0.1 >nul" & vbCrLf &
            $"xcopy ""{inst}"" ""{bak}\"" /E /I /Y /Q >nul" & vbCrLf &
            $"xcopy ""{newDir}"" ""{inst}\"" /E /Y /Q >nul" & vbCrLf &
            $"start """" ""{exe}""" & vbCrLf)
        Process.Start(New ProcessStartInfo("cmd.exe", $"/c ""{Path.Combine(work, "update.cmd")}""") With {.CreateNoWindow = True, .UseShellExecute = False})
        Environment.Exit(0)
    End Function

    Private Function NewClient() As HttpClient
        Dim c As New HttpClient() With {.Timeout = TimeSpan.FromMinutes(2)}
        c.DefaultRequestHeaders.UserAgent.ParseAdd("Calendizer-updater") ' GitHub rifiuta richieste senza User-Agent
        Return c
    End Function
End Module
