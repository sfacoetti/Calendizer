# Calendizer

Calendario desktop per Windows per trasformare le email di Outlook (classic) in attività. VB.NET WinForms + Blazor WebView.

## Funzioni

- **Viste:** giorno, settimana lavorativa (lun-ven), settimana, mese. Slot da 30 minuti, 8:00-19:00. Colonne e righe si adattano alla finestra.
- **Mail da Outlook:** trascina una o più mail selezionate in Outlook sul calendario; diventano attività (30 min) nello slot di rilascio. Titolo e mittente sono visibili; l'icona busta in basso a destra apre il testo completo.
- **Anteprima di drop:** durante il trascinamento un rettangolo evidenzia lo slot di destinazione (anche per le attività già presenti).
- **Spostamento e durata:** trascina un'attività per cambiare giorno/ora; trascina il bordo inferiore per cambiare la durata.
- **Sovrapposizioni:** attività nello stesso slot si dividono lo spazio, come in Outlook.
- **Colori:** sfondo semitrasparente in base al mittente (Impostazioni: dominio `@esempio.it` o indirizzo completo → colore; l'indirizzo ha la precedenza). Colore modificabile per singola attività.
- **Menu tasto destro:** colore, "Automatico", cestino, "Sposta a" (data e ora).
- **Attività libere:** doppio clic su una casella libera, oppure clic/trascinamento su più caselle e Invio → popup con titolo, testo, colore, durata.
- **Selezione e Canc:** clic su un'attività la seleziona; Canc la elimina (nessuna conferma).
- **Barra laterale sinistra** a scomparsa: viste e impostazioni.

## Struttura

| Progetto | Linguaggio | Contenuto |
|---|---|---|
| `Core` | VB.NET | modello, salvataggio JSON, colori, layout sovrapposizioni, lettura Outlook (COM, late binding) |
| `UI` | C# (Razor) | `App.razor`, `orga.js` (drag/resize/drop), `orga.css`, icone SVG. Razor non è supportato da VB, quindi la UI è in una libreria C# minima |
| `Host` | VB.NET WinForms | finestra, BlazorWebView, icona app, lettura mail dopo il drop |

Il drop esterno (Outlook) è ricevuto dalla pagina, che calcola lo slot e avvisa l'host; l'host aspetta 300 ms e legge la selezione di Outlook via COM (durante il drag Outlook è bloccato).

## Requisiti

- Windows, .NET 8 SDK (build) / runtime (uso), WebView2 Runtime, Outlook classic per il drop.

## Build ed esecuzione

```
dotnet build
Host\bin\Debug\net8.0-windows\Calendizer.exe
```

## Dati

`%AppData%\Calendizer\data.json` (attività e regole colore). Se il file è illeggibile viene copiato in `data.json.bak` prima di ripartire vuoti.

## Icone

SVG da Iconify: Streamline Plump Color e Streamline Flex Color (`UI/wwwroot/icons/`). Icona app: `memes-comment-reply` (Plump).

## Sviluppo

Vedi `PLAN.md` (piano, stato, verifiche da fare a mano, limiti noti).

## Licenze

Le icone del cestino (`trash.svg`) e di chiusura (`close.svg`) vengono dal set Streamline Sharp Color di Streamline (https://github.com/webalys-hq/streamline-vectors), licenza CC BY 4.0 (https://creativecommons.org/licenses/by/4.0/), tramite Iconify.
