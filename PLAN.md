# PLAN OrgaLendar
Obiettivo: app desktop VB.NET WinForms + Blazor WebView; calendario 4 viste (giorno, sett. lavorativa, settimana, mese), slot 30', 8-19; drop mail da Outlook classic -> attività; drag/resize; overlap affiancato; colori per dominio/indirizzo; sidebar a scomparsa + settings.
Struttura: Core (VB, stato/dati/Outlook COM) | UI (C# RCL, solo .razor/js/css, Razor non supporta VB) | Host (VB WinForms)

1. Scheletro 3 progetti + sln — prova: `dotnet build` ok
2. Core: modello, JSON, colori, layout overlap — prova: test console layout/colori
3. UI: viste, sidebar, modali, settings — prova: build + avvio app
4. JS: drag/resize/drop slot — prova: test nel WebView
5. Host: drop da Outlook (COM letto DOPO il drop) — prova: manuale con Outlook (non automatizzabile qui)
6. Verifica finale, riepilogo — FATTO tranne drop reale da Outlook (da provare a mano)

Stato: versione 1 completata + evoluzioni (anteprima drop, menu tasto destro, attività libere, selezione caselle, Canc, icone). Prossimi passi in `todo.md`.
