# TODO OrgaLendar

Ordine consigliato. Ogni passo si chiude con build pulita + prova nell'app.

## Prossimi passi

1. **Collegamento alla mail** (piccolo) — alla creazione da drop, salvare EntryID e StoreID di Outlook, data di ricezione e Message-ID (`PR_INTERNET_MESSAGE_ID`). L'EntryID cambia se la mail viene spostata: con gli altri due dati la si ritrova. Le attività già create restano senza collegamento. Base per i passi 3 e 5.
2. **Attività completate** (piccolo) — campo `Done`; voce nel menu tasto destro; stile: sbiadita, titolo barrato, spunta (`icons/confirm.svg`). Opzionale: filtro per nascondere le completate.
3. **Apri mail in Outlook** (piccolo) — voce nel menu e doppio clic sull'icona busta. Se la mail non si trova più, messaggio chiaro.
4. **Icona nell'area di notifica + promemoria** (medio) — le notifiche richiedono l'app in esecuzione: icona accanto all'orologio, chiusura = minimizza, avvio automatico con Windows. Preavviso per attività (nessuno / 5 / 15 / 30 min). Recupero delle notifiche saltate alla riapertura. Notifiche già incluse in Windows, nessuna libreria aggiuntiva.
5. **Archivia / sposta in cartella Outlook** (medio) — pulsante "Archivia" e scelta della cartella dall'albero di Outlook (COM). Dopo lo spostamento aggiornare il collegamento (passo 1).
6. **Robustezza** (piccolo) — copia di sicurezza automatica di `data.json`; pubblicazione come eseguibile singolo.

## Decisioni da prendere

- Completando un'attività, la mail va archiviata in automatico o solo su richiesta?
- Archivio: cartella Archivio di Outlook o cartella scelta per ogni mail?
- Promemoria: preavviso unico per tutte le attività o scelto per ciascuna?

## Da verificare a mano

- Drag reale da Outlook classic: l'evento di drop arriva alla pagina (provato solo con eventi simulati).
- Vista mese: drag, menu, selezione (provata poco).
- Più monitor / DPI diversi.
- Mittenti Exchange (indirizzo SMTP via `GetExchangeUser`).
- Archivio/spostamento mail: provare su mail di test (cache Exchange, cartelle condivise).

## Limiti noti

- Attività fuori dalle 8:00-19:00 non sono visibili.
- L'anteprima di drop si affianca alle attività presenti senza spostarle (al rilascio il layout si ridistribuisce).
- La selezione a trascinamento resta nella colonna di partenza.
- Canc elimina senza conferma.
