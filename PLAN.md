# PLAN Calendizer
Obiettivo: app desktop VB.NET WinForms + Blazor WebView; calendario 4 viste (giorno, sett. lavorativa, settimana, mese), slot 30', 8-19; drop mail da Outlook classic -> attività; drag/resize; overlap affiancato; colori per dominio/indirizzo; sidebar a scomparsa + settings.
Struttura: Core (VB, stato/dati/Outlook COM) | UI (C# RCL, solo .razor/js/css, Razor non supporta VB) | Host (VB WinForms)

## Versione 1 (fatta)
1. Scheletro 3 progetti + sln — prova: `dotnet build` ok
2. Core: modello, JSON, colori, layout overlap — prova: test console layout/colori
3. UI: viste, sidebar, modali, settings — prova: build + avvio app
4. JS: drag/resize/drop slot — prova: test nel WebView
5. Host: drop da Outlook (COM letto DOPO il drop) — prova: manuale con Outlook (drop reale provato: funziona)
6. Evoluzioni: anteprima drop, menu tasto destro, attività libere, selezione caselle, Canc, icone

## Evoluzioni (tutte fatte nel codice, build pulita; "provato" = confermato dall'utente)
1. Collegamento mail (EntryID, StoreID, data, Message-ID in `Activity`) — provato con drop reale
2. Attività completate (`Done`, menu, stile sbiadito/barrato/spunta) — da provare
3. Apri mail in Outlook (menu, doppio clic sulla busta, popup di lettura) — provato dal popup; doppio clic corretto (clic singolo ritardato di 250 ms)
4. Icona nell'area di notifica + promemoria (predefinito nelle impostazioni, per attività dal menu, recupero ultime 24 h, avvio con Windows `--tray`) — da provare
5. Archivia in cartella Outlook (elenco caselle → cartelle, caricamento in background, ricerca per Message-ID se spostata a mano) — da provare
6. Tema Chiaro / Scuro / Segui Windows (variabili CSS, `data-theme`) — provato a vista
7. Esporta `.ics` (intervallo da/a, "Salva con nome") — formato controllato, importazione in Outlook da provare
8. Robustezza: backup giornaliero di `data.json` (ultime 7) e pubblicazione `dotnet publish Host\Calendizer.vbproj -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o <cartella>` (exe + `wwwroot` + `WebView2Loader.dll`, serve .NET 8 Desktop Runtime) — backup e avvio dell'exe pubblicato provati
9. Sincronizzazione a una via con Outlook: barra in alto, sottocartella "Calendizer" del calendario predefinito, evento riconosciuto da proprietà `OrgaId`; ciò che non corrisponde ad attività viene eliminato (Elementi eliminati) — corretta dopo duplicati e orfani, da riprovare
10. Impostazioni: interruttore "Avvia con Windows" (`StartupHelper`) e "Notch sul bordo destro" (`Host/NotchForm.vb`: striscia al centro del bordo destro dello schermo principale, dopo 3 s di mouse sopra diventa una pillola di vetro sfocato con l'icona; clic = apre l'app; icona nell'area di notifica mantenuta) — aspetto verificato con uno screenshot di prova, integrazione nell'app da provare
11. Altro: una sola istanza (la seconda riporta in primo piano la prima), Esc chiude ogni popup + X in alto a destra, pulsanti moderni con icone, "Nascondi completate", ultima cartella di archivio salvata, conferma con Canc, pallino rosso sull'icona per promemoria non visti

## Da verificare a mano
- Drag reale da Outlook classic su più monitor / DPI diversi
- Vista mese: drag, menu, selezione (provata poco)
- Mittenti Exchange (SMTP via `GetExchangeUser`)
- Archivio/spostamento mail su mail di test (cache Exchange, cartelle condivise); lettura in background di Outlook (cartelle, sincronizzazione)
- Promemoria: balloon e avvio nascosto `--tray`; una sola istanza; tema scuro; esportazione in Outlook; sincronizzazione dopo la correzione (da 10 a 4 eventi)

## Limiti noti
- Attività fuori dalle 8:00-19:00 non sono visibili
- L'anteprima di drop si affianca alle attività presenti senza spostarle (al rilascio il layout si ridistribuisce)
- La selezione a trascinamento resta nella colonna di partenza
- Il cestino nel popup e nel menu elimina senza conferma (solo Canc la chiede)
- La sincronizzazione Outlook è a una via e manuale; la cartella Calendizer è solo dell'app

Decisioni prese: archiviazione solo su richiesta; promemoria predefinito + valore per attività; tema a 3 voci; export solo `.ics`; sincronizzazione a una via; niente trascinamento al contrario verso le cartelle di Outlook.
