# AGENTS.md - Documentazione e Guida del Progetto

Questo documento serve come guida di riferimento completa per qualsiasi agente o sviluppatore che lavora su questo repository. Descrive l'architettura, le modifiche implementate, i file del progetto e i flussi operativi automatizzati.

---

## ⚠️ REGOLA AUREA: PROTOCOLLO DI DEPLOY OBBLIGATORIO PER LA MINI APP (ZERO DIMENTICANZE)

Esistono **DUE repository Git distinti** su questa macchina:
1. `C:\!VR\App tariffa` (Repo principale/backend: Google Apps Script bot, test suite, file sorgente `index.html`, `sw.js`).
2. `C:\!VR\CalcoloTariffaVR-app` (Repo GitHub Pages collegato a `https://sxs46.github.io/CalcoloTariffaVR-app/`).

> [!CAUTION]
> **PROTOCOLLO TASSATIVO PER OGNI MODIFICA A `index.html` o `sw.js` (DA ESEGUIRE SEMPRE SENZA ECCEZIONI):**
> 1. **Copia File**: Copiare immediatamente `index.html` e `sw.js` da `C:\!VR\App tariffa` a `C:\!VR\CalcoloTariffaVR-app`.
> 2. **Cache Busting Service Worker**: In entrambi i repo, assicurarsi di incrementare `CACHE_NAME = 'vr-app-shell-vX'` in `sw.js` e la query di registrazione in `index.html` (`./sw.js?v=X`).
> 3. **Git Push GitHub Pages**: Entrare in `C:\!VR\CalcoloTariffaVR-app`, fare `git add index.html sw.js`, `git commit -m "..."` e `git push origin main`.
> 4. **Bust Cache URL Telegram nel Bot GAS**: In `google_apps_script_bot.js`, bumpare la costante `MINI_APP_URL` (es. da `v=3.20.0` a `v=3.21.0`). Se questo parametro non viene aggiornato, Telegram apre la webview con la cache vecchia e l'utente NON vedrà le modifiche!
> 5. **Deploy Backend GAS**: Eseguire `npx @google/clasp push --force` e `npx @google/clasp deploy -i AKfycbxZZMRCEy3OmmuQht3ABI3_o4YOFfrITiWRnQxQz4hL9rU5LG2zT6HXLmgmVOwQMw -d "..."`.
> 6. **Test Suite**: Eseguire `node test_suite.js` (tutti i test devono passare a 0 errori).
> 7. **Verifica Live**: Controllare tramite curl / web request che `https://sxs46.github.io/CalcoloTariffaVR-app/index.html` risponda con il nuovo contenuto prima di notificare l'utente.

---

## 1. Panoramica del Sistema

Il sistema automatizza il calcolo delle tariffe lavorative e la compilazione dei consuntivi orari tra quattro componenti:
1. **Bot Telegram (`google_apps_script_bot.js`)**: Interfaccia conversazionale utilizzata dall'utente per registrare, modificare e archiviare le giornate lavorative (Casa e Trasferta), gestendo tariffe base, straordinari, extra ed eventi consecutivi.
2. **Mini App Telegram (`index.html`, `sw.js`)**: Interfaccia grafica WebApp moderna per Telegram (ospitata su `https://sxs46.github.io/CalcoloTariffaVR-app/`), sincronizzata in tempo reale con il backend Google Apps Script.
3. **Google Sheet**: Database centrale strutturato in due sezioni principali:
   - **Casa**: Righe 3–26 (Col A: Data, Col B: Inizio, Col C: Fine, Col E: Compenso, Col F: Extra, Col G: Consecutivo/Totale Evento).
   - **Trasferta**: Righe 31–54 (stessa struttura).
   - **Tariffe**: Colonna J (J3: Casa Short, J5: Casa Base, J6: Casa Straordinario, J31: Trasferta Short, J33: Trasferta Base, J34: Trasferta Straordinario).
4. **Userscript Tampermonkey (`videorent_sync.user.js`)**: Script del browser che gira su `https://g.videorent.it/`. Legge i dati dal Google Sheet tramite endpoint Web App GET e compila automaticamente gli slot orari e i compensi delle card degli eventi nel planner di Videorent.

---

## 2. File del Repository

| File | Scopo |
| :--- | :--- |
| `index.html` | Frontend SPA della Mini App Telegram (PWA). |
| `sw.js` | Service Worker per il caching e offline support della Mini App Telegram. |
| `test_suite.js` | Test suite automatica end-to-end con mock sandbox GAS (>860 test). |
| `videorent_sync.user.js` | Userscript principale v2.2 per `g.videorent.it` con cache locale, sync sequenziale e bulk sync. |
| `tampermonkey_loader.user.js` | Script ponte per Tampermonkey: carica ed esegue `videorent_sync.user.js` dal disco locale via `@require`. |
| `google_apps_script_bot.js` | Backend completo: Webhook Bot Telegram (`doPost`) + Web App JSON (`doGet`) per Videorent. |
| `aggiorna_bot.bat` | Script batch Windows: esegue `clasp push` e `clasp deploy` in produzione con 1 clic. |
| `.clasp.json` | Configurazione Clasp collegata allo Script ID del progetto Google Apps Script. |
| `.claspignore` | Filtro per Clasp: carica su Google solo `google_apps_script_bot.js` e `appsscript.json`. |
| `appsscript.json` | Manifest del progetto Apps Script (Timezone Europe/Rome, V8 runtime, WebApp permissions). |
| `Calcolo Tariffa.xlsx` | Modello di foglio di calcolo Excel originale con le formule tariffarie. |
| `.gitignore` | Esclusioni per Git (node_modules, file temporanei, OS metadata). |

---

## 3. Tutte le Modifiche e Migliorie Implementate

### A. Ottimizzazione delle Prestazioni (Eliminazione Collo di Bottiglia ~10s ➔ ~0.4s)
- **Problema**: Le funzioni leggevano le celle di Google Sheets una ad una con centinaia di chiamate sincrone `sheet.getRange(r, c).getValue()`.
- **Soluzione Batch Reading**: Riscritte tutte le funzioni di lettura per prelevare intere matrici con `getValues()` e `getDisplayValues()` in 1 o 2 chiamate aggregate:
  - `readSectionForSync`: Endpoint GET per Videorent (ridotto da ~300 chiamate a 4 chiamate batch).
  - `getRecentEntries`: Menu "Modifica Giornate" del bot Telegram (ridotto da ~250 chiamate a 2 chiamate batch).
  - `existsEntryForDate`: Controllo giorni adiacenti (lettura colonna A in blocco).
  - `insertChronologicalEntry`: Ricerca posizione cronologica e verifica spazio libero.
  - `getEntryByRow`: Dettaglio singola riga.
  - `getRatesFromSheet`: Lettura tariffe (singola lettura intervallo `J3:J34`).

### B. Cache Locale Persistente (*Stale-While-Revalidate*)
- Utilizzo di `GM_setValue` / `GM_getValue` (con fallback su `localStorage`).
- **Apertura Istantanea a 0 ms**: Quando l'utente apre `my-planner-list.php`, le card vengono decorate immediatamente con i dati in memoria locale.
- **Revalidation Silenziosa**: Se la cache è più vecchia di 60 secondi, parte una verifica in background; se i dati sono cambiati, aggiorna la cache e ricarica a caldo le card senza ricaricare la pagina.
- **Anti-Flicker**: Aggiunto attributo `data-vr-sync-sig` su ogni card: evita il ridisegno dell'HTML se lo stato e i dati dell'evento non sono cambiati.

### C. Precaricamento Prima del Login
- Esteso `@match` a tutto il dominio `https://g.videorent.it/*`.
- Quando l'utente atterra sulla schermata di login (`/login.php` o `/`), se la cache locale ha più di 3 minuti, lo script **avvia il download dei dati in background**.
- Al completamento del login, quando si atterra sul planner, i dati sono già pronti in locale.

### D. Sincronizzazione Sequenziale Robusta
- Sostituito `Promise.all` in `executeSync` con un ciclo `async/await` sequenziale con pausa di respiro di 120ms tra uno slot e l'altro.
- Previene il blocco di sessione PHP (*session lock*) o errori di database su Videorent.
- Feedback visivo in tempo reale nel badge:  
  `⏳ Giorno 1/3...` ➔ `⏳ Giorno 2/3...` ➔ `⏳ Salvataggio totale...` ➔ `✅ Compilato con successo`.

### E. Bulk Sync e Widget Informativo Intelligente
- Funzione `getPendingReadyEvents`: rileva tutte le card che hanno copertura 100% nel foglio e importo non ancora registrato su Videorent.
- Se ci sono eventi da compilare, compare il tasto verde: `⚡ Compila tutti (N)`.
- Se tutti gli eventi coperti sono già compilati, il widget mostra esplicitamente lo stato:  
  `⚡ Foglio: 28 gg (in cache) • X già compilati ✅`, evitando dubbi sull'assenza del tasto bulk.

### F. Riepilogo Contabile all'Archiviazione del Mese
- In `archiveAndClearMonth()`, prima di cancellare le righe attive, il bot calcola:
  - Giornate Casa
  - Giornate Trasferta
  - Extra totali maturati
  - Fatturato complessivo in Euro (€)
- Invia un report Telegram completo di chiusura mese indicando anche il nome della scheda di backup creata.

### G. Gestione Sicura Token con Fallback
- Introdotta la funzione `getScriptConfig()` che interroga `PropertiesService.getScriptProperties()` per `TELEGRAM_TOKEN`, `ALLOWED_USER_ID`, `SYNC_SECRET_KEY`, `SPREADSHEET_ID`, `TELEGRAM_WEBHOOK_SECRET`, mantenendo i valori costanti come fallback trasparente.

### H. Input Orari Flessibile e Suggerimenti Asimmetrici nel Bot Telegram
- **Digitazione Diretta**: In qualsiasi momento in cui il bot chiede l'orario di inizio o fine (apertura giornata, inserimento retroattivo, modifica giornate), l'utente può digitare l'orario direttamente nel messaggio di chat (es. `8`, `8.30`, `19`, `19.30`) senza dover prima premere pulsanti aggiuntivi; il bot lo interpreta automaticamente senza uscire dal flusso.
- **Rimozione Pulsante "Altro orario..."**: La tastiera inline è stata snellita eliminando il pulsante ridondante.
- **Tastiere Orarie Asimmetriche**:
  - **Inizio Giornata**: 3 pulsanti con orari precedenti (`-90m`, `-60m`, `-30m`) e 1 successivo (`+30m`) rispetto all'orario arrotondato attuale, facilitando la registrazione di turni già iniziati.
  - **Fine Giornata**: 1 pulsante con orario precedente (`-30m`) e 3 successivi (`+30m`, `+60m`, `+90m`) rispetto all'orario arrotondato attuale, velocizzando la chiusura di turni che si protraggono.

### I. Conferma Obbligatoria per Annullamento ed Eliminazione Giornate
- **Annulla Giornata Attiva**: Premendo il pulsante `🗑️ Annulla Giornata` dalla schermata Home con turno aperto, il bot non cancella più immediatamente la giornata. Mostra ora una card di avviso con i dati del turno (tipologia, orario di inizio, data) e richiede conferma esplicita (`🗑️ Sì, Annulla Definitivamente` oppure `🔙 No, Mantieni Giornata`).
- **Elimina Giornata Registrata**: Migliorata la finestra di conferma prima di eliminare una riga salvata dal foglio Google, evidenziando data, tipologia, orario, compenso maturato e numero di riga.

### J. Riepilogo Mese in Tempo Reale nel Bot Telegram
- **Pulsante e Comandi Dedicati**: Aggiunto il pulsante `📊 Riepilogo Mese` nella Home del bot e i comandi chat `/stats`, `/mese`, `mese`.
- **Statistiche Istantanee**: Calcola in tempo reale dal foglio:
  - Giornate Casa e Trasferta lavorate.
  - Totale ore lavorate e compenso complessivo provvisorio.
  - Extra accumulati.
  - Evidenzia se un turno odierno è attualmente in corso di svolgimento.

### K. Vista Settimanale Rapida nel Bot Telegram
- **Pulsante e Comandi Dedicati**: Aggiunto il pulsante `📅 Settimana` nella Home del bot e i comandi chat `/settimana`, `settimana`.
- **Mini-Calendario Settimanale**: Mostra tutti i 7 giorni della settimana corrente (da Lunedì a Domenica) indicando:
  - Eventi lavorati con tipologia (🏠/🚗), orari e compenso.
  - Eventuale turno attualmente aperto (`🟢 In corso (dalle HH.mm)`).
  - Giorni di riposo (`—`).
  - Totale giornate lavorate e fatturato maturato nella settimana.

### L. Rilevamento Discrepanze Importi e Protezione Bulk Sync su Videorent
- **Allerta Visivo Differenza**: Se una card sul planner di Videorent ha già un importo registrato che differisce dal totale calcolato dal foglio Google, viene visualizzato un badge ambra:  
  `⚠️ Differenza: VR X € vs Foglio Y €`.
- **Pulsante di Correzione Mirata**: Compare il pulsante `🔄 Correggi a Y €` protetto da una finestra di dialogo di conferma per evitare sovrascritture accidentali.
### M. Correzione Sicurezza Webhook Telegram (Risoluzione Webhook Bypass)
- **Problema**: La verifica del token di sicurezza in `doPost(e)` scattava solo se l'header veniva inviato. Una chiamata HTTP diretta priva di header poteva bypassare il controllo.
- **Risoluzione**: Reso il controllo obbligatorio e categorico: se `TELEGRAM_WEBHOOK_SECRET` è configurato, qualsiasi richiesta priva dell'header esatto `X-Telegram-Bot-Api-Secret-Token` (o parametro secret) viene respinta all'istante.

### N. Promemoria Automatico Chiusura Turno dopo 14 Ore
- **Trigger Temporizzato Intelligente**: All'apertura del turno (`OPEN_DAY`), il bot programma un trigger temporizzato tramite `ScriptApp.newTrigger("send14HourReminder")` a 14 ore di distanza.
- **Pulizia Automatica**: Alla chiusura regolare o all'annullamento della giornata, il trigger pendente viene eliminato per non occupare risorse.
- **Notifica con Azioni Dirette**: Se il turno è ancora aperto dopo 14 ore, invia un alert Telegram con pulsanti inline per chiudere il turno all'istante o annullarlo.

### O. Rilevamento Anomalie di Durata e Sovrapposizioni Orarie
- **Controllo Turno Eccessivo**: Se un turno inserito o modificato supera le 16 ore consecutive, il bot richiede conferma esplicita all'utente evidenziando il totale ore calcolato.
- **Rilevamento Sovrapposizione Oraria**: Scansiona la tabella (Casa e Trasferta) per verificare se esiste già una giornata registrata con orari che si intersecano con quelli del nuovo turno. Se presente, mostra una card di allerta con orari e tipologia, consentendo di confermare o correggere.

### P. Gestione e Correzione Turni Notturni (a Cavallo della Mezzanotte)
- **Giorno di Registrazione**: Confermato che il bot assegna sempre la data di inizio del turno in colonna A.
- **Formula Foglio Google**: Se un turno termina il giorno successivo (`endDecimal < startDecimal`), viene aggiunto `+1.0` (24 ore) al valore decimale scritto in colonna C. In questo modo la formula `=C-B` del Foglio calcola la durata positiva corretta (es. dalle 20:00 alle 02:00 = 6 ore) e abilita il corretto calcolo di straordinari e compenso, mentre il formato cella `"hh.mm"` continua a mostrare visivamente l'orario `02.00`.

### Q. Sistema Multi-Utente, Onboarding con Approvazione di Simone e Setup Rapido
- **Anagrafica Multi-Utente (`USER_DIRECTORY`)**: Gestita dinamicamente tramite `PropertiesService`. Simone (`ADMIN_ID: 12025766`) è l'amministratore supremo.
- **Flusso di Registrazione e Approvazione**:
  - Quando un collega avvia il bot con `/start`, viene registrato come `PENDING_APPROVAL`.
  - Simone riceve una notifica interattiva su Telegram con i bottoni `[✅ Approva ed Assegna Foglio]` e `[❌ Rifiuta]`.
- **Clonazione Automatica del Foglio di Calcolo**:
  - All'approvazione, Google Apps Script clona il modello `TEMPLATE_SPREADSHEET_ID` su Google Drive creando un foglio dedicato (es. `Calcolo Tariffa - Nome Collega`).
  - La funzione `cleanSpreadsheetForNewUser` azzera automaticamente tutte le righe di lavoro esistenti (Casa 3–26 e Trasferta 31–54) e le celle delle tariffe `J2:J6` e `J30:J34`, lasciando intatte tutte le formule del foglio.
  - Viene generata una `syncKey` univoca e crittograficamente sicura per la sincronizzazione browser.
- **Wizard Interattivo per le Tariffe Personali**:
  - **Campi vuoti nel wizard tariffe**: Al collega non viene precompilata alcuna tariffa; inserisce i propri valori da zero.a zero.
  - **Flessibilità Mezza Giornata**: Il bot chiede esplicitamente se il collega ha un accordo per la mezza giornata oppure se per lui è sempre giornata intera (sia per Casa che per Trasferta).
    - Se *Sempre giornata intera*: imposta la soglia ore a 0 (`J2=0` o `J30=0`) e il compenso mezza giornata uguale alla giornata base (`J3=J5` o `J31=J33`).
    - Se *Mezza giornata presente*: chiede la soglia ore (es. 4 o 5) e il compenso pattuito.
  - Configura inoltre lo straordinario orario per Casa e Trasferta.
- **Installazione 1-Clic e Aggiornamenti Automatici Nativi (@updateURL / @downloadURL)**:
  - Passo 1: Link diretti per l'installazione dell'estensione Tampermonkey dallo store ufficiale (Chrome Web Store / Edge Add-ons).
  - Passo 2: Pulsante interattivo *⚡ Installa con 1 Clic*: sfrutta l'intercettazione nativa di Tampermonkey per i file `.user.js` (`?action=install_script&key=<syncKey>&format=raw&file=videorent_sync.user.js`), aprendo direttamente il prompt di installazione/reinstallazione nel browser.
  - **Aggiornamento Automatico a Zero Fatica**: Lo script generato contiene le direttive `@updateURL` e `@downloadURL` collegate alla Web App Google Apps Script. Quando Simone pubblica aggiornamenti o miglioramenti, Tampermonkey verifica periodicamente l'URL e **aggiorna automaticamente lo script del collega in background**, senza che il collega debba ricaricare o copiare codice manualmente.
  - Disponibili comunque i pulsanti di fallback per scaricare il file `.user.js` o copiare il codice negli appunti.
- **Pannello Amministrazione per Simone (`👥 Gestione Colleghi`)**:
  - Presente nel menu Home solo per Simone (`ADMIN_ID`).
  - Consente di consultare l'elenco dei colleghi, aprire il loro foglio Drive personale con un clic, aprire direttamente la pagina di installazione script personale del collega, sospendere/riattivare l'accesso, rigenerare la chiave Sync o riavviare il wizard tariffe.

### R. Risoluzione Webhook Redirect 302 e Navigazione In-Place nel Bot
- **Eliminazione errore 302 Webhook**: In `doPost(e)`, sostituito `ContentService.createTextOutput` con `HtmlService.createHtmlOutput("OK")` per eliminare il redirect 302 a `googleusercontent.com` che causava la ripetizione automatica dei messaggi da parte di Telegram.
- **Navigazione In-Place fluida**: Nel pulsante `CMD_CANCEL` e nelle conferme, richiamato `renderHome(chatId, null, messageId)` per aggiornare il messaggio sul posto evitando l'invio di menù duplicati.

### S. Configurazione Completa Tariffe e Regola Weekend/Festivi
- **Menu Impostazioni Dettagliato (`⚙️ Impostazioni`)**:
  - Mostra il quadro completo dei parametri attuali sia per Casa sia per Trasferta: Base giornata, Straordinario orario e stato Mezza Giornata (attiva con soglia ore e compenso, oppure disattivata).
  - Sottomenu dedicati `🏠 Configura Casa` e `🚗 Configura Trasferta`: consentono di modificare individualmente Base (`K5`/`K33`), Straordinario (`K6`/`K34`) o gestire la Mezza Giornata (`K2:K3` o `K30:K31`).
  - Gestione Mezza Giornata flessibile: possibilità di attivarla impostando ore e compenso, oppure disattivarla (impostando ore a 0 e compenso pari alla base).
  - Isolamento multi-utente: ogni modifica viene scritta sul foglio personale del singolo utente senza intaccare gli altri.
- **Regola Festivi e Weekend Flessibile (Applicazione Tariffa Trasferta su Giornate Casa)**:
  - Rilevamento automatico del calendario italiano: calcola tutte le festività fisse, le festività mobili (Pasqua e Lunedì dell'Angelo/Pasquetta con algoritmo Meeus/Jones/Butcher) e i giorni di weekend (Sabato e Domenica).
  - **Tre modalità configurabili individualmente per ciascun collaboratore**:
    1. `ALL`: **Sabato, Domenica e Festivi** (tutti i festivi nazionali e sia sabato sia domenica a tariffa Trasferta anche da Casa).
    2. `SUNDAY_AND_HOLIDAYS`: **Solo Domenica e Festivi** (solo domenica e festivi nazionali a tariffa Trasferta; al **Sabato** viene applicata la normale tariffa base Casa).
    3. `NONE`: **Disattivata** (nessuna maggiorazione per Casa né a weekend né a festivi).
  - Configurabile liberamente dall'utente in `⚙️ Impostazioni` (toggle a 3 stati), da Simone nel pannello colleghi `👥 Gestione Colleghi`, e nel wizard interattivo iniziale di onboarding (`🎈 3. REGOLA WEEKEND E FESTIVI`).
  - Quando attiva su un giorno idoneo, il bot applica la maggiorazione differenziale di Trasferta (es. 30 € se Casa è 220 € e Trasferta è 250 €) in colonna G (Extra), preservando intatte le formule native del foglio.
### T. Risoluzione Criticità Multi-Utente, Espansione Dinamica, PDF Telegram e Troupe Videorent
- **Promemoria 14h Multi-Utente Robusto (`ensureHourlyMaintenanceTrigger`)**:
  - Eliminati i trigger temporizzati dinamici per singolo turno (che si cancellavano a vicenda quando un collega chiudeva un turno e rischiavano di superare il limite di 20 trigger Google).
  - Introdotto 1 singolo trigger orario ricorrente di manutenzione che scansiona in loop gli utenti attivi e notifica solo chi ha superato le 14 ore con turno aperto.
  - Controllo automatico esclusivo su turni aperti > 14 ore (il promemoria mattutino delle ore 08:00 è stato rimosso per garantire la massima discrezione del bot).
- **Rilevamento Sovrapposizioni Turni Notturni di Ieri**:
  - `findOverlappingEntry` ora verifica anche la giornata precedente (`targetDate - 1`): se il turno di ieri notte è terminato dopo la mezzanotte e invade l'orario di inizio odierno, scatta l'allerta sovrapposizione evidenziando `(Notturno di ieri)`.
- **Espansione Dinamica Tabelle e Batch Shifting Istantaneo**:
  - `ensureAvailableRows`: se le 24 righe fisse di Casa o Trasferta si saturano, la funzione inserisce automaticamente le righe necessarie prima della riga di totale, copiando formule relative (colonne D, E, G) e formati, consentendo alle formule di somma (`=SUM`) di espandersi senza mai bloccarsi.
  - Spostamento righe retroattivo convertito da 160 chiamate a cella singola a una singola operazione matriciale `range.copyTo()`, rendendo l'inserimento retroattivo istantaneo (< 0.2s).
- **Esportazione PDF Contabile su Telegram (`/pdf`, `/fattura`)**:
  - Generazione di un documento consuntivo A4 verticale professionale ed elegante (`generateMonthPdfHtml`), privo di formule grezze e colonne di servizio tecniche (H-U). Include intestazione con nome collaboratore, mese, card statistiche (Giornate Casa/Trasferta, Compenso base, Extra e Totale spettante) e tabella cronologica pulita con subtotali e totale complessivo, inviato direttamente in chat Telegram (`sendDocument`).
- **Rilevamento Troupe/Colleghi su Videorent**:
  - Nello Userscript Tampermonkey, `extractEventCrew` rileva i colleghi assegnati alla stessa commessa e aggiunge un badge visuale elegante: `👥 Troupe: Nome Collega`.
  - Cache locale univoca isolata per utente: `vr_sheet_sync_cache_<prefix>`, azzerando interferenze tra profili diversi.

### U. Ottimizzazione Reattività Selezione Orari (Eliminazione Collo di Bottiglia ~4.5s ➔ ~0.3s)
- **Problema di Latenza Riscontrato**: Quando l'utente premeva un pulsante per scegliere l'orario (inizio o fine turno), il bot impiegava diversi secondi a prendere l'input.
- **Risoluzione Colli di Bottiglia**:
  - **Cache Stato Trigger di Manutenzione**: In `ensureHourlyMaintenanceTrigger`, la chiamata di sistema `ScriptApp.getProjectTriggers()` (che impiegava ~1.2s ad ogni webhook `doPost` e veniva persino duplicata all'apertura turno) ora viene saltata istantaneamente verificando la proprietà persistente `HOURLY_TRIGGER_INSTALLED` in `PropertiesService`, azzerando l'overhead a 0 ms.
  - **Validazione Batch Unificata in Memoria (`inspectShiftValidation`)**: Quando si seleziona l'orario di fine, anziché eseguire 3 aperture separate del foglio Google e 6 letture isolate (`findOverlappingEntry` + 2x `existsEntryForDate`), il bot apre il foglio una sola volta ed esegue una singola lettura aggregata delle 3 colonne A:C. Tutte le verifiche (durata >16h, sovrapposizioni odierne, turni notturni di ieri e presenza di turni ieri/domani per eventi consecutivi) vengono calcolate in memoria istantaneamente (~1ms).
  - **Zero Query nel Dialogo Consecutivi**: Il risultato relativo alla giornata successiva (`hasNext`) viene memorizzato nello stato utente, rendendo istantanea anche la risposta a "Fa parte dello stesso evento?".

### V. Ottimizzazioni di Velocità, Roundtrip Telegram e Sync Videorent
- **Batching Istantaneo in `getSectionBoundaries`**:
  - Sostituito il ciclo con chiamate remote cella per cella (`getValue()` e `getFormula()` ripetute 35–50 volte) con un'unica operazione batch `sheet.getRange(1, 1, lastRow, 5).getValues()` / `getFormulas()`. Risolto il collo di bottiglia durante l'inserimento/espansione tabelle (< 0.05s anziché ~1.2s).
- **Chiusura Turno a Singola Chiamata Telegram**:
  - In `finalizeEntry`, eliminata la sequenza ridondante `editMessage/sendMessage` + `renderHome(chatId)` che causava 2 roundtrip consecutivi alle API di Telegram. Estratta la funzione `getHomeKeyboard(chatId, state)` e allegata la tastiera principale direttamente al messaggio di riepilogo registrazione con un'unica chiamata.
- **Cache Tariffe Utente con Invalidation (`CacheService`)**:
  - In `getRatesFromSheet`, implementata la cache ad accesso ultra-rapido (`CacheService.getScriptCache()`, TTL 6 ore). All'aggiornamento tariffe tramite menu Impostazioni (`updateSheetRate`), la cache del foglio specifico viene invalidata istantaneamente.
- **Velocizzazione Sync Slot Multi-Giorno su Videorent**:
  - Nello Userscript `videorent_sync.user.js` (e nel generatore backend `generateUserScriptContent`), ridotto l'intervallo di sicurezza tra slot consecutivi dello stesso evento da 120ms a 60ms, dimezzando i tempi di compilazione su commesse lunghe senza rischi di session lock.

### W. Risoluzione Separa/Unisci Evento di Ieri e Batch Shifting Istantaneo
- **Correzione TypeError su `getDifferenceInDays`**:
  - La serializzazione JSON in `PropertiesService` trasformava `item.date` in stringa ISO, causando `TypeError: d2.getFullYear is not a function` in `getDifferenceInDays`.
  - Normalizzati difensivamente tutti i parametri data in `getDifferenceInDays` e risolta la data corrente direttamente dalla cella Google Sheet in `getPreviousDayStatus`.
- **Riscrittura `toggleEventConnection` con Batch `copyTo`**:
  - Eliminato il ciclo cella per cella `copyCellData` (~160 chiamate API, ~15s).
  - Sostituito con operazione nativa batch `range.copyTo(targetRange)` (< 0.05s).
  - **Separazione**: integra `ensureAvailableRows`, sposta `item.row .. lastUsedRow` in basso di 1 riga e pulisce `item.row` (colonne A, B, C, F), preservando le formule native di colonna D, E, G.
  - **Unione**: sposta in alto verso `item.row - 1` e pulisce l'ultima riga in coda.
  - Aggiornamento in-place del menu e inversione dinamica del pulsante (Separa ⇄ Unisci).

### X. Test Suite Automatica End-to-End e Protezione Pre-Deploy (`test_suite.js`)
- **Simulatore Gas Sandbox**: Emulazione fedele a 0 ms di `SpreadsheetApp`, `PropertiesService` (con serializzazione reale delle stringhe), `UrlFetchApp`, `Session`, `Utilities` e `CacheService`.
- **43 Scenari di Collaudo**:
  - Navigazione chat, apertura turno, chiusura turno e formule compenso.
  - **Test di regressione bug** su serializzazione data in "Separa da Evento Ieri" e "Unisci".
  - Correzione automatica arrotondamento `totalHours` contro imprecisioni floating point IEEE-754.
  - Calcolo tariffe festive/weekend con maggiorazione trasferta su casa.
  - Generazione HTML consuntivo PDF e test endpoint `doGet` per Videorent.
  - Flusso apertura turno anticipato con domanda consecutivi e chiusura senza domande.
  - Sincronizzazione Google Calendar (nomenclatura titoli, toggle, variazioni orari e cancellazioni).
- **Gate Pre-Deploy in `aggiorna_bot.bat`**: Esegue `node test_suite.js` prima di `clasp push` e `deploy`, bloccando il rilascio in produzione se anche un solo test fallisce.

### Y. Apertura Turno Semplificata e Sync Automatico Google Calendar ("Videorent")
- **Spostamento Domanda "Stesso Evento di Ieri" all'Apertura Turno**:
  - In `OPEN_DAY`, dopo la scelta dell'orario di inizio (`SET_START_...` o input testo), il bot controlla se ieri (`targetDate - 1`) è registrato un turno.
  - Se presente, pone subito la domanda (`OPEN_EVENT_SAME` / `OPEN_EVENT_DIFF`) e memorizza `state.activeDay.skipBefore`.
  - Alla chiusura serale (`MENU_END_DAY` / `SET_END_...`), la domanda su ieri viene saltata del tutto: l'utente passa direttamente alla selezione degli extra, azzerando l'attrito cognitivo dopo turni lunghi.
- **Sincronizzazione Automatica Google Calendar (`Videorent`)**:
  - Calendario secondario dedicato denominato `"Videorent"` gestito via `CalendarApp`.
  - **Nomenclatura Intelligente dei Titoli**:
    - Bologna Confermato: `Videorent`
    - Bologna In attesa di conferma: `Videorent?`
    - Fuori Bologna Confermato: `Videorent [Città]` (es. `Videorent Roma`)
    - Fuori Bologna In attesa di conferma: `Videorent [Città]?` (es. `Videorent Roma?`)
  - **Dati Completi Evento**: Luogo/indirizzo reale valorizzato (per indicazioni stradali su Google Maps), date/orari precisi, descrizione dettagliata (commessa, cliente, referente, troupe, ID).
  - **Notifiche Telegram Proattive a Simone**: alert immediati su variazioni di orari/date di impegni esistenti, passaggio da provvisorio a confermato, o cancellazione/rimozione di eventi futuri da Videorent (con eliminazione automatica dal calendario).
  - **Controllo Privacy & Toggle**: Sincronizzazione disabilitata di default per massima riservatezza; abilitabile e disabilitabile in qualsiasi momento dal menu `⚙️ Impostazioni` (`SETTING_TOGGLE_CALENDAR_SYNC`).

### Z. Risoluzione Mancata Sincronizzazione Eventi su Google Calendar
- **Problema Riscontrato**: Il calendario secondario "Videorent" era stato creato correttamente e il pulsante `📅 Calendar` compariva sul planner di Videorent, ma gli eventi non comparivano sul Google Calendar.
- **Causa Radice**: In `doPost(e)`, la condizione `if (!syncKeyReceived && e.postData && e.postData.contents)` non veniva eseguita perché `syncKeyReceived` era già valorizzato dalla query string dell'URL (`?key=...`). Di conseguenza, `parsedBody` rimaneva nullo, `actionReceived` rimaneva vuoto, il controllo `actionReceived === "calendar_sync"` falliva e lo script ricadeva sul controllo del secret token Telegram, restituendo `Unauthorized`.
- **Risoluzione Implementata**:
  - In `google_apps_script_bot.js` (`doPost`), `parsedBody` viene parsato prioritariamente da `postData.contents`; `syncKeyReceived` e `actionReceived` vengono estratti sia dai parametri URL che dal body JSON con fallback trasparente.
  - Se `parsedBody.events` è un array, la richiesta viene instradata automaticamente alla sincronizzazione del calendario anche in assenza esplicita di `action`.
  - In `parseCalendarDateTime`, aggiunto supporto nativo robusto sia per formati `DD/MM/YYYY` sia per `YYYY-MM-DD` (con o senza ore), garantendo che `ev.start` e `ev.end` estratti da `data-event` o dal testo della card vengano convertiti senza disallineamenti di fuso orario UTC.
  - Nello Userscript `videorent_sync.user.js`, l'endpoint della richiesta include esplicitamente `action=calendar_sync` sia nell'URL che nel body JSON, e al clic manuale su `📅 Calendar` viene mostrato un report visivo immediato con il conteggio degli eventi creati, aggiornati o rimossi.
  - Aggiunto il test 10.5 in `test_suite.js` (44/44 test superati al 100%).

### AA. Risoluzione Falsa Cancellazione Eventi in Paginazione e Cambio Scheda
- **Problema Riscontrato**: Navigando tra le pagine di "Prossimi Eventi" (es. passando da Pagina 1 a Pagina 2 o 3) o aprendo la scheda "Storico Eventi", gli eventi della pagina precedente venivano erroneamente considerati eliminati da Videorent, cancellati da Google Calendar e notificati a cascata con falsi allarmi su Telegram.
- **Causa Radice**: Il backend `syncVideorentEventsToCalendar` assumeva che ogni richiesta contenesse il catalogo completo di tutti gli eventi futuri, eliminando qualsiasi evento presente in `syncMap` ma assente dal payload. Navigando su Pagina 2, gli eventi di Pagina 1 non erano nel DOM della pagina attiva, innescando la cancellazione spuria.
- **Risoluzione Implementata**:
  - **Filtro Scheda Storico**: In `videorent_sync.user.js`, la funzione `isStoricoTab` rileva se l'utente si trova in `tab=storico` o sulla scheda attiva "Storico Eventi", bloccando l'esecuzione della sincronizzazione ed evitando qualsiasi impatto sul calendario futuro.
  - **Scansione Multi-Pagina Completa (`collectAllFutureEvents`)**: Nello Userscript, quando ci si trova su "Prossimi Eventi", lo script estrae le card visibili e scansiona in background (via `fetch` asincrono con credenziali di sessione) tutte le pagine successive/precedenti di paginazione (`.pagination a`, `.page-link`, pulsanti numerici), unificando in memoria l'intero catalogo di commesse future senza duplicati.
  - **Protezione Server-Side (`isFullCatalog`)**: Il backend accetta il flag `isFullCatalog`. Se il sync è parziale o non garantito (`isFullCatalog: false`), il backend esegue **esclusivamente l'upsert** (creazione/aggiornamento) e **preserva intatti al 100% tutti gli eventi già salvati in `syncMap` senza cancellarli**.
  - La rimozione automatica da Google Calendar e la notifica Telegram di cancellazione scattano **esclusivamente** se il client ha verificato con certezza l'intero catalogo (`isFullCatalog: true`) e una commessa futura risulta effettivamente scomparsa da tutte le pagine del gestionale.
  - Test suite estesa con i test 10.4a e 10.4b (**46/46 test superati al 100%**).

### AB. Risoluzione Definitiva Duplicazione Eventi e Auto-Purge Copie Esistenti
- **Problema Riscontrato**: Gli eventi sul Google Calendar venivano creati doppi; passando da Pagina 1 a Pagina 2 ne rimaneva uno solo, e tornando a Pagina 1 tornava a duplicare.
- **Causa Radice**:
  1. *Race Condition da Chiamate Concorrenti*: Nello Userscript `videorent_sync.user.js`, `syncAllEventsToCalendar()` veniva invocato all'interno del loop periodico `processCards()` ogni 1.5s senza lock in-flight immediato e senza debounce, inviando richieste HTTP concorrenti prima della conclusione della sincronizzazione precedente.
  2. *Mancata Ricerca Calendar-First nel Backend*: In `google_apps_script_bot.js`, il backend si fidava unicamente della chiave in memoria `syncMap`. Se la commessa non era presente in `syncMap` (ad es. per disallineamento cache tra pagine), chiamava ciecamente `cal.createEvent` creando una seconda copia sul calendario anche se l'evento era già presente.
  3. *Fluttuazione Paginazione*: La combinazione di cancellazioni in paginazione e ricreazione senza deduplicazione sul calendario causava l'effetto alternato "doppio ➔ singolo ➔ doppio".
- **Risoluzione Implementata**:
  - **Ricerca Calendar-First Diretta**: Il backend prima di ogni operazione cerca su Google Calendar (`cal.getEvents`) gli eventi esistenti nel range temporale matching l'ID commessa (`ID Videorent: #${evIdStr}` o titolo `... #${evIdStr}`). Se l'evento esiste già sul calendario, lo riutilizza ed aggiorna titolo/orari/luogo senza mai creare una nuova voce.
  - **Auto-Purge Automatico dei Duplicati**: Se nel calendario Google vengono trovate due o più copie dello stesso ID commessa (create in precedenza), il backend conserva la copia primaria ed **elimina immediatamente all'istante le copie in eccesso (`deleteEvent()`)**, bonificando lo storico del calendario al primo passaggio.
  - **Debounce e Lock Robusto nel Frontend**: Nello Userscript, la sincronizzazione del calendario è stata disaccoppiata dal loop delle card con un timer debounced (`scheduleCalendarSync(1500)`) e flag di mutua esclusione atomico `isSyncingCalendar = true` impostato prima di qualsiasi richiesta asincrona.
  - **Riconoscimento Paginazione Numerica Universale**: Allargato il rilevamento dei pulsanti di paginazione numerici a qualsiasi elemento button/a con valore numerico <= 50, garantendo che se ci sono pagine multiple non raggiungibili via URL il backend riceva sempre `isFullCatalog = false`, proteggendo da qualsiasi cancellazione errata.
  - **Test Suite Estesa a 48 Scenari**: Aggiunto il test 10.4c che simula la presenza di copie duplicate orfane sul calendario e ne verifica l'avvenuta auto-eliminazione e convergenza ad 1 solo evento (**48/48 test superati al 100%**).

### AC. Catalogo Multi-Pagina Persistente, Sync Parziale Sicuro e Protezione Mass-Drop
- **Problema Riscontrato**: Navigando da Pagina 1 a Pagina 2 del planner di Videorent, gli eventi di Pagina 1 venivano cancellati da Google Calendar, scatenando una valanga di allerte Telegram di rimozione. Tornando a Pagina 1, accadeva l'opposto per Pagina 2.
- **Causa Radice**:
  1. *Fragilità del Background Fetching*: Lo scraping in background via `fetch()` delle altre pagine non funzionava (i pulsanti `1` e `2` di Videorent utilizzano logica AJAX interna senza parametri URL standard); il fetch restituiva 0 card ma con status HTTP 200, inducendo lo script a credere erroneamente che le card visibili fossero l'intero catalogo (`isFullCatalog: true`).
  2. *Cancellazione Server-Side su Catalogo Incompleto*: Il backend Apps Script, ricevendo `isFullCatalog: true` con solo le card di Pagina 2, deduceva che tutte le commesse di Pagina 1 fossero state annullate da Videorent e le cancellava in blocco inviando una notifica Telegram per ciascuna.
- **Risoluzione Implementata**:
  - **Catalogo Multi-Pagina Persistente Client-Side (`vr_planner_catalog_future_events`)**:
    - Nello Userscript `videorent_sync.user.js`, memorizza persistentemente (`GM_setValue` / `localStorage`) l'unione di tutte le commesse future visualizzate.
    - Quando l'utente visualizza Pagina 1, le card di Pagina 1 vengono registrate; quando si sposta su Pagina 2, le card di Pagina 2 vengono unite al catalogo esistente senza perdere Pagina 1.
    - Il payload inviato a Google Calendar include costantemente **l'intero insieme di tutti gli eventi conosciuti**, azzerando il rischio che un evento risulti mancante solo perché risiede su un'altra pagina.
    - Pulizia automatica locale per eventi conclusi con data inizio > 24 ore nel passato.
  - **Enforcement `isFullCatalog = false` per Sync Automatico**:
    - Durante la navigazione ordinaria tra pagine, lo Userscript invia categoricamente `isFullCatalog: false`. In questo modo il backend opera esclusivamente in modalità upsert (creazione/aggiornamento) e non elimina mai commesse remote né genera allerte spurie.
  - **Protezione Anti-Mass-Drop nel Backend (`google_apps_script_bot.js`)**:
    - Aggiunto un guard difensivo: se anche una richiesta dovesse riportare `isFullCatalog: true` ma presenta la scomparsa simultanea di 4 o più eventi futuri, il server la degrada automaticamente a sync parziale (`safeFullCatalog = false`), bloccando qualsiasi cancellazione di massa e azzerando le notifiche a raffica su Telegram.
  - **Test Suite a 49 Scenari (`test_suite.js`)**:
    - Aggiunto il test `10.4d` che verifica l'intervento del guard anti-mass-drop a fronte di 5 eventi mancanti simulati (**49/49 test superati al 100%**).

### AD. Ottimizzazioni Prestazionali ad Alto Impatto (Batch Calendar, Cache doGet e MutationObserver)
- **Problema/Opportunità**:
  1. *Latenza Chiamate CalendarApp*: Per ogni commessa nel catalogo (es. 20 eventi), il backend invocava singolarmente `cal.getEvents(...)`, spendendo ~250-300ms a chiamata (fino a 6 secondi totali per sincronizzazione).
  2. *Latenza Apertura Foglio in doGet*: Ogni richiesta GET dello Userscript per ottenere i consuntivi del foglio Google invocava `SpreadsheetApp.openById()` e la scansione batch delle righe (~800-1200ms).
  3. *Polling Aggressivo nel Browser*: Lo Userscript eseguiva `processCards` su `setInterval` fisso a 1.5 secondi ri-armando continuamente `scheduleCalendarSync(1500)` anche a riposo e introducendo fino a 1500ms di ritardo nel decorare le card dopo un cambio pagina o filtro.
- **Risoluzione Implementata**:
  - **Batch Calendar Lookup e Indicizzazione in Memoria (`syncVideorentEventsToCalendar`)**:
    - Pre-calcola l'intervallo temporale globale (da min `startDate - 48h` a max `endDate + 48h`) tra tutte le commesse in arrivo.
    - Esegue **1 singola chiamata aggregata** a `CalendarApp.getEvents(start, end)`.
    - Indicizza gli eventi restituiti in una `Map` JavaScript in memoria filtrando per `#ID` commessa.
    - L'accesso a ciascun evento durante il ciclo passa da ~300ms a **0.001ms** (tempo totale di interrogazione calendario ridotto da ~6s a ~0.3s, **~20x più veloce**).
  - **Cache ad Accesso Ultra-Rapido per i Dati Foglio (`CacheService`)**:
    - In `getAllEntriesForSync`, implementata la memorizzazione nella cache Apps Script (`CacheService.getScriptCache()`, TTL 30 minuti).
    - Risposta dell'endpoint GET per Videorent ridotta da ~1000ms a **~15ms** (**~60x più veloce**).
    - Invalidation automatica trasparente: ogni volta che l'utente inserisce una giornata (`finalizeEntry`), modifica orari (`updateEntryInSheet`), cancella una riga (`clearRowInSheet`), unisce/separa eventi (`toggleEventConnection`) o archivia il mese (`archiveAndClearMonth`), la cache viene invalidata istantaneamente.
  - **Reattività Istantanea a 0ms con `MutationObserver` nel Frontend**:
    - Nel planner di Videorent, introdotto `MutationObserver` sul contenitore delle card: non appena l'utente clicca su Pagina 2, cambia filtro o scheda, le nuove card vengono intercettate e decorate istantaneamente in **80ms** (anziché attendere fino a 1.5 secondi).
    - Fallback periodico esteso da 1.5s a 4s per azzerare l'utilizzo di CPU e il consumo di batteria a riposo.

### AE. Filtro Presenza Tecnico e Notifiche Telegram Differenziate (Già Confermato vs Da Confermare)
- **Contesto e Requisiti di Business**:
  - Nel planner di Videorent, i lavori assegnati al tecnico si distinguono per lo stato di accettazione/presenza di Simone:
    1. *In attesa di accettazione*: pulsanti `[ Conferma Presenza ]` e `[ Rifiuta lavoro ]`, badge `[ ⚠️ Conferma la tua presenza! ]`.
    2. *Presenza confermata*: pulsante `[ ✔️ Hai confermato ]` e badge `[ 👤 Presenza confermata ]`.
  - Requisito: **I lavori non ancora confermati dal tecnico non devono assolutamente apparire su Google Calendar**, ma Simone deve ricevere una notifica Telegram con lo stato opportuno per ricordarsi di confermare o rifiutare la commessa su Videorent.
- **Risoluzione Implementata**:
  - **Rilevamento Nello Userscript (`extractCardCalendarData`)**:
    - Rileva in modo insensibile a maiuscole/minuscole i testi dei pulsanti e badge: presenza di `conferma presenza`, `rifiuta lavoro`, `conferma la tua presenza` imposta `isTechnicianConfirmed = false`; presenza di `hai confermato`, `presenza confermata` imposta `isTechnicianConfirmed = true`.
  - **Filtro Categorico Calendario nel Backend (`syncVideorentEventsToCalendar`)**:
    - Se `isTechnicianConfirmed === false`:
      - Qualsiasi evento esistente su Google Calendar per quel lavoro viene rimosso immediatamente (`deleteEvent()`), impedendo la comparsa di eventi non ancora approvati.
      - Al primo avvistamento su Videorent (evento futuro), invia notifica Telegram dedicata: `⚠️ Videorent - Nuovo Lavoro da Confermare!` con dettagli e invito ad agire su Videorent.
      - Memorizza lo stato in `syncMap` con `calendarEventId: null` e `isTechnicianConfirmed: false`.
    - Se `isTechnicianConfirmed === true`:
      - Se è un lavoro nuovo che appare già confermato: crea l'evento su Google Calendar e notifica `🆕 Videorent - Nuovo Lavoro (Già Confermato)`.
      - Se è un lavoro precedentemente in attesa (`prevSync.isTechnicianConfirmed === false`) che Simone ha appena approvato: crea/aggiorna l'evento su Calendar e notifica `✅ Videorent - Presenza Confermata!`.
  - **Protezione Anti-Spam Bulk Catalog**:
    - `isInitialCatalogSync`: se `syncMap` è vuoto e riceve più di 2 eventi (primo sync o riconnessione), popola il catalogo senza generare raffiche di notifiche push per i vecchi lavori.
  - **Test Suite a 58 Scenari (`test_suite.js`)**:
    - Aggiunti test unitari `10.6`, `10.7` e `10.8` per verificare l'esclusione dal Calendar di eventi non confermati, l'invio degli alert di conferma e la rimozione immediata degli eventi annullati/cancellati (**58/58 test superati al 100%**).

### AF. Risoluzione Rilevazione Eventi Annullati, Partizionamento Pagine e Rimozione da Calendar
- **Problema Riscontrato**:
  - Quando un evento (es. 16-17 novembre) veniva cancellato da Videorent, il bot non lo rimuoveva dal calendario e non inviava l'alert di annullamento.
  - *Causa radice frontend*: `vr_planner_catalog_future_events` accumulava gli eventi in un dizionario piatto senza mai cancellare quelli futuri rimossi dal planner. Inoltre, `isFullCatalog` era forzato a `false`, impedendo al backend di scansionare le assenze nel catalogo.
- **Risoluzione Implementata**:
  - **Rilevamento `isCancelled` nello Userscript (`extractCardCalendarData`)**:
    - Scansiona il testo della card e l'oggetto dati cercando occorrenze di `"annullat"`, `"cancellat"`, `"rifiutat"`. Se presenti, imposta `isCancelled = true`.
  - **Catalogo Partizionato per Pagina (`vr_planner_catalog_pages_v3`)**:
    - Il catalogo memorizza gli eventi per pagina (`catalogData.pages[String(curPage)] = pageEvents`). Quando la pagina 1 viene visitata, i vecchi eventi di pagina 1 vengono sovrascritti, eliminando istantaneamente quelli cancellati.
    - La vecchia cache piatta `vr_planner_catalog_future_events` viene bonificata all'avvio.
  - **Scansione Pagine in Background e `isFullCatalog` Affidabile**:
    - Se ci sono più pagine (`totalPages > 1`), `fetchOtherPlannerPages` scarica le pagine non ancora visitate in background e aggrega tutti gli eventi futuri impostando `isFullCatalog = true`.
  - **Gestione Backend e Rimozione Difensiva (`google_apps_script_bot.js`)**:
    - Gestione immediata di `ev.isCancelled: true`: elimina l'evento da Google Calendar, invia l'alert Telegram `🚨 Videorent - Lavoro Annullato o Rimosso!` ed esclude la commessa dal nuovo stato.
    - Rimozione difensiva per ID commessa (`calEventsById.get(String(oldId))`) per eliminare l'evento anche se `calendarEventId` risultava nullo o non aggiornato.

### AG. Scansione Multi-Pagina Automatica in Background e Semplificazione Copy Lavori da Confermare
- **Problemi Riscontrati**:
  - *Scorrimento Pagine (Slide verso Storico)*: Man mano che i lavori passano allo storico, le commesse scivolano dalla pagina 2 alla pagina 1. Se l'utente non apriva manualmente la pagina 2, eventi cancellati o modificati su altre pagine rischiavano di non essere aggiornati subito.
  - *Copy Rumoroso / Non Gradito*: Gli eventi da riconfermare inviavano il messaggio `🚨 Videorent - Presenza da Riconfermare!` con la nota `🗑️ Rimosso temporaneamente dal calendario Google in attesa di conferma`, creando confusione e copy prolisso.
- **Risoluzioni Implementate**:
  - **Copy Semplificato e Pulito (`google_apps_script_bot.js`)**:
    - Eliminato qualsiasi riferimento a "Presenza da Riconfermare" o "Rimosso temporaneamente dal calendario".
    - Tutti i lavori non ancora confermati dal tecnico vengono notificati con il testo chiaro `⚠️ *Videorent - Nuovo Lavoro da Confermare!*` (con flag `notifiedUnconfirmed: true` per azzerare lo spam).
    - La rimozione dal calendario per eventi non confermati avviene in modo silenzioso e trasparente.
  - **Scansione Multi-Pagina Automatica in Background (`videorent_sync.user.js`)**:
    - Implementato `fetchPageViaIframe(pageNumber)`: carica le pagine 2, 3, ecc. in un iframe nascosto ed estrae le card senza interferire con la UI dell'utente.
    - Implementato `scanPlannerViaBackgroundIframe`: se l'utente naviga su qualsiasi pagina di Videorent (`index.php`, magazzino, ecc.), lo script avvia in background una scansione silenziosa del planner per sincronizzare l'intero catalogo eventi a Google Calendar senza bisogno di aprire manualmente `my-planner-list.php`.
### AH. Eliminazione Cicli Inutili e Stato di Riposo "In Attesa di Modifiche"
- **Problema Riscontrato**:
  - Al ricaricamento della pagina o durante la navigazione, sembrava che lo script ripartisse ogni volta da zero (aprendo iframe, ricalcolando il catalogo e rieseguendo sync continui).
  - *Causa radice*:
    1. `lastCalendarPayloadHash` era una variabile solo in memoria RAM: ad ogni ricaricamento pagina tornava `null`, inducendo lo script a considerare tutto non sincronizzato e inviare una chiamata POST a Google Apps Script.
    2. In `processCards`, veniva chiamata `scheduleCalendarSync(2500)` incondizionatamente. Poiché un `setInterval(..., 4000)` richiamava `processCards`, si creava un ciclo infinito in background ogni 4 secondi.
    3. `fetchOtherPlannerPages` interrogava nuovamente le pagine secondarie ad ogni ciclo senza verificare la freschezza della cache locale.
- **Risoluzione Implementata**:
  - **Persistenza Hash in `localStorage`**: `lastCalendarPayloadHash` viene salvato in `localStorage`. Al reload, lo script confronta immediatamente l'hash in ~2ms: se invariato, non invia alcuna chiamata di rete e non apre alcun iframe.
  - **Stato di Riposo Silenzioso ("In attesa di modifiche")**:
    - Una volta verificato che il catalogo è aggiornato, lo script rimane a riposo (0% CPU, 0 richieste di rete).
    - Il badge mostra subito `📅 Calendar ✅`.
  - **Disaccoppiamento `processCards` da `scheduleCalendarSync`**:
    - `scheduleCalendarSync` viene eseguito una sola volta all'avvio e poi reattivamente solo quando `MutationObserver` rileva reali nuovi nodi nel DOM (es. cambio pagina/scheda dell'utente).
### AI. Scansione Unificata (Storico 2 Mesi + Prossimi) e Widget Intelligente (Mese Attivo + Ultimo Giorno)
- **Problema Riscontrato**:
  - Il testo `"X già compilati"` occupava spazio senza utilità operativa.
  - L'indicazione `"Foglio: 28 gg"` contava genericamente tutte le righe del foglio senza isolare il mese in corso e senza indicare l'ultima data registrata.
  - La compilazione bulk scansionava solo le card della scheda attiva nel DOM: le commesse concluse finite nello **"Storico Eventi"** non venivano rilevate se l'utente si trovava sui "Prossimi Eventi", richiedendo cambi manuali di scheda.
- **Risoluzione Implementata**:
  - **Widget Informativo Intelligente (Opzione A)**:
    - Mostra il mese di riferimento delle registrazioni attive e il conteggio delle sole giornate del mese con l'ultima data lavorata:  
      Es: `⚡ Settembre: 14 gg (ultimo: 18/09)`.
    - Determina dinamicamente il mese dalle date reali scritte nel foglio: se è il 1° del mese ma il foglio contiene le registrazioni del mese precedente (non ancora archiviato), mostra correttamente il mese in chiusura senza mostrare "0 gg".
    - Se ci sono lavori da compilare: evidenzia `• ⚠️ N da compilare` con il pulsante verde `[ ⚡ Compila tutti (N) ]`.
    - Se tutti i lavori registrati sono compilati: conferma `• Tutto compilato ✅`.
  - **Scansione Unificata Storico (Ultimi 2 Mesi) + Prossimi**:
    - Funzione `fetchStoricoEvents`: scarica in background la scheda `tab=storico` estraendo le commesse degli ultimi 65 giorni (~2 mesi, allineato al ciclo di fatturazione).
    - Cache locale persistente `vr_storico_events_cache` con TTL di 10 minuti.
    - `getUnifiedPendingReadyEvents`: aggrega le commesse di entrambe le schede (Prossimi + Storico), verificando la copertura del foglio e l'assenza di importo su Videorent.
  - **Compilazione Bulk a 1 Clic Globale (`executeBulkSync`)**:
    - Il pulsante `⚡ Compila tutti (N)` carica orari e totale fattura su Videorent per tutti gli eventi pronti, sia dello Storico sia dei Prossimi.
    - `executeSync` gestisce in modo null-safe sia le card presenti nel DOM visibile (con aggiornamento badge in tempo reale) sia quelle residenti nell'altra scheda (tramite chiamate API dirette).

### AJ. Schede Mensili Automatiche e Rotazione Dinamica a 6 Mesi (Zero Archiviazione Manuale)
- **Problema Riscontrato**:
  - In precedenza l'archiviazione mensile era un'operazione manuale bloccante: se un utente lavorava l'ultimo giorno del mese (es. 30 Settembre) e apriva un nuovo turno il giorno successivo (1° Ottobre) prima di archiviare il mese, il vecchio turno rischiava di essere sovrascritto o cancellato dall'archiviazione tardiva.
  - La sincronizzazione con Videorent leggeva solo la prima scheda `ss.getSheets()[0]`.
- **Risoluzione Implementata**:
  - **Routing Automatico della Scheda per Data del Turno (`getOrCreateMonthSheet`)**:
    - Ogni mese ha una scheda dedicata intitolata `Mese Anno` (es. `Settembre 2026`, `Ottobre 2026`).
    - All'inserimento di una giornata (live o retroattiva), il bot rileva la data del turno e individua la scheda mensile corrispondente.
    - Se la scheda non esiste ancora (es. primo turno del nuovo mese), viene generata all'istante clonando il template, preservando al 100% tutte le formule matematiche (colonne D, E, G) e le tariffe personali (colonna J), mentre le sole celle di lavoro (A..C, F) vengono pulite.
    - Se esiste un solo foglio preesistente (es. "Foglio 1"), viene automaticamente rinominato al mese corrente senza toccare alcun dato esistente.
  - **Multi-Month Sync per Videorent (`getAllEntriesForSync` & `doGet`)**:
    - L'endpoint `doGet` scansiona tutte le schede mensili attive (ultimi 2-4 mesi) aggregando gli eventi in un unico flusso JSON.
    - In questo modo le commesse a cavallo di mese o non ancora compilate nello storico vengono coperte al 100%.
  - **Rotazione Dinamica e Pulizia Automatica a 6 Mesi (`pruneOldMonthSheets`)**:
    - In background, il bot verifica periodicamente le date delle schede mensili.
    - Qualsiasi scheda più vecchia di 6 mesi (>= 6 mesi trascorsi) viene eliminata automaticamente per mantenere il foglio leggero, veloce e pulito, preservando sempre almeno una scheda attiva.
  - **Supporto Confini Mese in Validazione e Settimana**:
    - `inspectShiftValidation`: se il turno è il 1° del mese, controlla anche il foglio del mese precedente per prevenire sovrapposizioni con turni notturni di ieri e verificare eventi consecutivi.
    - `renderCurrentWeekView`: per le settimane a cavallo tra due mesi (es. 28 Settembre - 4 Ottobre), interroga entrambe le schede mensili assemblando perfettamente il consuntivo settimanale.
    - `getRecentEntries`: elenca le giornate degli ultimi mesi indicando la scheda di appartenenza, consentendo modifica e cancellazione mirata con `SELECT_ROW_<row>_<idx>`.
  - **Test Suite Automatica Estesa a 79 Scenari (`test_suite.js`)**:
    - Implementato `MockSheet` e `MockSpreadsheet` con supporto multi-scheda.
    - Aggiunto `Gruppo 11` con 17 nuovi test che certificano la creazione schede, isolamento dati, sync multi-mese, consuntivo settimanale e pruning a 6 mesi (**79/79 test superati al 100%**).

### AK. Barretta Rapida Salva-Spazio e Ripristino Tasto Indietro Smartphone
- **Problema Riscontrato**:
  - La Reply Keyboard persistente fissa in basso (`is_persistent: true`) occupava lo slot di input/tastiera dello smartphone. Su Android e nelle gesture di navigazione, il sistema intercettava il tasto/gesture "Indietro" a livello del pannello tastiera anziché a livello di chat, impedendo all'utente di uscire rapidamente alla lista conversazioni di Telegram con 1 tocco (obbligando a premere la freccia interna dell'app in alto a sinistra). Inoltre, occupava permanentemente spazio verticale sul display.
- **Risoluzione Implementata**:
  - **Sblocco Totale Tasto Indietro (`remove_keyboard: true`)**:
    - `getPersistentKeyboard` invia ora categoricamente `remove_keyboard: true`, ordinando al client Telegram di chiudere ed eliminare qualsiasi pannello tastiera custom precedentemente memorizzato.
    - Il tasto/gesture di sistema "Indietro" dello smartphone torna a funzionare istantaneamente al 100% per uscire dalla chat con un solo tocco.
  - **Barretta Rapida Salva-Spazio Dinamica (`sendQuickActionBar`)**:
    - Quando arrivano notifiche esterne da Videorent (lavori da confermare, variazioni, cancellazioni):
      1. Tutte le notifiche arrivano pulite e consecutive in blocco (zero "Home in mezzo").
      2. Al termine del batch, il bot invia in fondo **un'unica barretta compatta inline** (2 sole righe salva-spazio):
         - *In Standby*: `⚡ Azioni rapide:` con `[ 🟢 Inizia Giornata ] [ 🏠 Menu Completo ]`.
         - *In Servizio*: `🟢 Turno in corso: Tipologia dalle HH:mm` con `[ 🔴 Termina Giornata ] [ 🗑️ Annulla Giornata ]` e `[ 🏠 Menu Completo ]`.
  - **Deduplicazione Automatica Intelligente (`deleteMessage`)**:
    - Il bot memorizza `state.lastQuickBarMessageId`. Se in seguito arrivano nuove notifiche, cancella silenziosamente la barretta precedente prima di inviare quella nuova in fondo, evitando qualsiasi messaggio duplicato o spazzatura nella cronologia chat.
  - **Espansione In-Place e Azione Immediata**:
    - Toccando `🏠 Menu Completo`, la barretta si trasforma sul posto nella Home completa con statistiche del mese e settimana.
    - Toccando `🟢 Inizia Giornata` o `🔴 Termina Giornata`, la selezione orario e le relative domande di fine/extra avvengono direttamente lì in fondo sul posto.
  - **Test Suite Automatica Aggiornata (`test_suite.js`)**:
    - Gruppo 12 aggiornato per verificare il rilascio di `remove_keyboard: true`, la matrice pulsanti compatta in standby e turno attivo, l'invio e la cancellazione deduplicata via `deleteMessage`, e la compatibilità con comandi testuali rapidi (`/inizia`, `/termina`, `/fine`, `/stop`, `🟢 Inizia Giornata`).

### AM. Architettura Relazionale a ID Commessa (Colonna A), Bot Zero-Dialog e Matching Istantaneo
- **Motivazione**: Sostituita la precedente logica basata sulla contiguità fisica delle righe e sulle righe vuote di separazione con un modello relazionale fondato sull'**ID Commessa** (`PLAN_ID_COLUMN.md`):
  - **Colonna A**: `ID Commessa` (es. `#42576`, oppure vuota per lavori estemporanei/extra).
  - **Colonna B**: `Data`
  - **Colonna C**: `Inizio`
  - **Colonna D**: `Fine`
  - **Colonna E**: `Totale Ore` (`=IF(OR(ISBLANK(C3);ISBLANK(D3)); ""; IF(D3<C3; (D3+1)-C3; D3-C3))`)
  - **Colonna F**: `Compenso`
  - **Colonna G**: `Extra`
  - **Colonna H**: `Totale Evento` (formula aggregata che somma tutti i turni con lo stesso ID tramite `SOMMA.SE` e visualizza il totale sull'ultima riga del blocco commessa, oppure singolo compenso per i turni senza ID: `=IF(B3=""; ""; IF(A3=""; F3+G3; IF(A3<>A4; SOMMA.SE(A$3:A$26; A3; F$3:F$26) + SOMMA.SE(A$3:A$26; A3; G$3:G$26); "")))`).
  - **Colonna K**: `Tariffe` (`K2:K6` Casa, `K30:K34` Trasferta, traslate a destra di 1 colonna).
- **Bot Zero-Dialog**:
  - Eliminata la domanda all'apertura del turno (`AWAITING_OPEN_PREV_EVENT`) e alla chiusura (`AWAITING_PREV_EVENT`/`AWAITING_NEXT_EVENT`): il turno si apre e si chiude all'istante senza domande di raccordo.
  - Al salvataggio, il bot cerca la data del turno nella mappa catalogo di Videorent (`PLANNER_CATALOG_MAP_<userId>` in `PropertiesService`): se la commessa è presente, scrive in automatico l'ID (es. `#42576`) in Colonna A; altrimenti lascia la cella vuota.
  - Scrittura compatta (`rowsNeeded = 1`) senza inserimento di righe vuote separatorie.
- **Matching Userscript O(1)**:
  - In `videorent_sync.user.js` (e in `generateUserScriptContent`), la funzione `analyzeEventCoverage` confronta prima l'ID commessa esatto della card con l'ID presente nel foglio e la data, con fallback trasparente su data/orario se assente.
- **Suite di Test Estesa a 88/88**:
  - Mock del foglio aggiornato all'architettura a 8 colonne e tariffe in Colonna K.
  - Test Gruppo 9 riscritto per validare il flusso zero-dialog e l'assegnazione automatica dell'ID Commessa in Colonna A (**88/88 test superati con successo**).

### AN. Calcolo Corretto Eventi Multi-Mese, Zero-Fabrication Orari, Date Sicure e Ordinamento Inverso Schede
- **Calcolo Corretto Eventi a Cavallo di Mese (`analyzeEventCoverage`)**:
  - Risolto errore di calcolo per eventi estesi tra due mesi (es. `#41923 ESPID 2026`, 31 Maggio - 5 Giugno: Maggio 320 € + Giugno 1605 € = 1925 €).
  - La formula `lastEntry.eventTotal` nel foglio di Giugno considerava solo le righe di Giugno (1605 €).
  - Lo script ora calcola `calculatedInvoice = matchedEntries.reduce((acc, it) => acc + (it.dayPay || 0), 0)`, sommando i compensi reali di tutte le singole giornate dell'evento tra tutti i mesi coinvolti, azzerando le false discrepanze.
- **Divieto Assoluto Fabbricazione Orari (`extractCompiledEventData`)**:
  - In precedenza, se un evento multi-giorno (es. #42707, 20-21 settembre) non aveva input di orari dettagliati nel DOM, lo script creava orari fittizi accorpando l'inizio del primo giorno alla fine dell'ultimo giorno (`16:00 - 16:30`), generando il falso avviso `⚠️ Differenza Orari`.
  - Implementata la regola ferrea di Zero-Fabrication: `hasRealSlots = false` se mancano campi orario dettagliati nel DOM per eventi multi-giorno; il controllo discrepanza orari viene soppresso evitando avvisi errati.
- **Parsing Sicuro Date Italiane e ISO (`parseDateTimeSafe`)**:
  - Risolto bug in `collectAllFutureEvents` dove `new Date("20/09/2026 16:00")` restituiva `Invalid Date` / `NaN` scartando eventi futuri dalla sincronizzazione con Google Calendar.
  - `parseDateTimeSafe` gestisce in modo affidabile sia il formato italiano `DD/MM/YYYY [HH:mm]` che i formati ISO `YYYY-MM-DD`.
- **Ordinamento Decrescente delle Schede Mensili (`sortMonthSheetsDescending`)**:
  - Il mese più recente è sempre la prima scheda a sinistra nel Google Sheet (es. Ottobre > Settembre > Giugno > Maggio ...).
  - Implementato tramite `Spreadsheet.setActiveSheet(sheet)` e `ss.moveActiveSheet(1)` su tutte le schede mese in ordine cronologico decrescente.
  - All'avvio di un nuovo mese, `getOrCreateMonthSheet` riordina istantaneamente le schede ponendo il nuovo mese al primo posto e scalando i precedenti.
- **Suite di Test Estesa a 158 Scenari (`test_suite.js`)**:
  - Aggiunto `Gruppo 17` con collaudo automatico completo di `parseDateTimeSafe`, somma compensi multi-mese, zero slots fittizi e ordinamento inverso schede (**158/158 test superati con successo**).

### AP. Scansione Veloce Storico, Fedeltà Orari, Mese di Febbraio e Push Immediato Presenza Keepalive
- **Scansione Completa e Veloce dello Storico via Fetch Diretta (`scanStoricoForMissingJobs`)**:
  - Eliminato l'iframe invisibile con navigazione click (che dopo pagina 1 falliva e si bloccava senza mai raggiungere Febbraio o Gennaio).
  - Sostituito con scansione sequenziale diretta tramite `fetch` e parsing con `DOMParser`: interroga tutte le pagine dello Storico dell'anno in corso a 60ms di intervallo, raccogliendo tutte le commesse (incluso Febbraio) in meno di 1 secondo e creando automaticamente le schede mensili mancanti nel Google Sheet.
- **Estrazione Fedele degli Orari ed Eliminazione Orari Fittizi (`extractCompiledEventData`)**:
  - Estrazione orari reali da `cardText` (`INIZIO: ...`, `FINE: ...` e pattern `ORARIO: HH:mm - HH:mm`), evitando categoricamente il fallback arbitrario `08:00 - 17:00`.
  - Gestione turni notturni unici a cavallo della mezzanotte: se un turno termina la notte seguente (es. 21:00 - 03:00, durata <= 18h), viene trattato come singola giornata lavorativa registrata sulla data di inizio con orari effettivi.
  - In `autoImportVideorentJobs` e `extractCompiledEventData`: eliminato categoricamente qualsiasi fallback o stima oraria inventata (niente `h+9`, niente `09.00 - 18.00`). Se l'orario di inizio o fine non è presente autenticamente su Videorent, la commessa NON viene importata e viene visualizzato il badge di avviso esplicito sulla card: `⚠️ Orari non specificati nel gestionale`.
- **Intercettazione Immediata della Conferma Presenza con Keepalive (0 ms)**:
  - Registrato listener `click` in capture phase (`document.addEventListener('click', ..., true)`): intercetta istantaneamente il click sui pulsanti "Conferma presenza", "Conferma la tua presenza", "Rifiuta lavoro", prima ancora del refresh o della chiusura della pagina.
  - Genera e invia il payload `calendar_sync` a Google Apps Script tramite `fetch` con `keepalive: true` e `GM_xmlhttpRequest`. Anche se l'utente chiude la scheda un millisecondo dopo aver cliccato, la richiesta HTTP viene completata dal browser in background: la notifica Telegram e l'aggiornamento Google Calendar partono all'istante senza bisogno di ricaricare la pagina o riaprire il gestionale.
- **Suite di Test Estesa a 180 Scenari (`test_suite.js`)**:
  - Aggiunto `Gruppo 18` con 18 nuovi test automatici per orari reali, turni notturni, creazione scheda Febbraio, payload keepalive e divieto assoluto di orari fittizi (**180/180 test superati al 100%**).

### AQ. Priorità Assoluta Dati Gestionale, Sincronizzazione Storico, Zero-Fabrication e Ordinamento Decrescente Schede
- **Principio Fondamentale: Priorità del Gestionale**:
  - Qualsiasi dato presente a gestionale Videorent ha la priorità rispetto al Google Sheet e deve essere scritto nel foglio di calcolo.
  - Se il mese corrispondente (es. `Febbraio 2026`) non esiste nel foglio, Apps Script clona il template, crea la scheda, azzera le righe e richiama `sortMonthSheetsDescending(ss)` per collocarla nel corretto ordine cronologico decrescente (es. Ottobre > Settembre > ... > Febbraio > Gennaio).
- **Zero-Fabrication Orari & Preservazione Giornata e Compenso**:
  - Se nel gestionale una commessa è consuntivata/compilata ma mancano gli orari specifici di inizio e/o fine, il sistema NON inventa orari fittizi (niente 09.00 - 18.00 né +9h).
  - La riga viene inserita nella scheda mensile con la Data e l'ID commessa, le celle orario C e D lasciate vuote `""` (segnalando in modo chiaro l'assenza del dato orario), e l'intero importo pattuito/consuntivato dal gestionale (`invoice_amount` o tariffa base) assegnato in Colonna G (Extra / Totale).
  - In questo modo il consuntivo economico, il fatturato e le giornate tornano al centesimo, preservando l'autenticità dei dati.
- **Sincronizzazione Storico Tempestiva al Ricaricamento Pagina e Widget**:
  - Nello Userscript `videorent_sync.user.js` v2.6, allentato il debounce bloccante per permettere la scansione automatica tempestiva all'apertura/ricaricamento (`F5`) dello Storico (`isStoricoTab()`).
  - Rimosso il requisito `compiled.hasRealSlots` sia da `readCardsFromDoc` che da `checkAndAutoImportMissingJobs`: qualsiasi commessa compilata viene importata nel foglio.
  - Aggiunto il pulsante dedicato `📥 Storico` nel widget floating per forzare la scansione completa di tutte le pagine dello Storico in 1 clic.
- **Suite di Test Estesa a 184 Scenari (`test_suite.js`)**:
  - Validati con successo i test 18.5, 18.6 e 18.7 per la priorità del gestionale, celle orario vuote, compenso a consuntivo e ordine decrescente delle schede mensili (**184/184 test superati al 100%**).

### AR. Risoluzione Approvazione Colleghi, Assegnazione Nome Reale e Duplicazione Foglio Resiliente
- **Problema Riscontrato**:
  - Alla ricezione della notifica di richiesta accesso di un collega, premendo *"Approva ed assegna foglio"* Telegram mostrava un caricamento indefinito e sembrava non succedere nulla.
  - Causa 1: `DriveApp.getFileById()` richiedeva un'autorizzazione OAuth interattiva a Google Drive (`https://www.googleapis.com/auth/drive`), bloccando l'esecuzione dell'handler Webhook `doPost` che gira in background come servizio.
  - Causa 2: `answerCallback` veniva chiamato solo dopo la duplicazione e pulizia del foglio, lasciando la rotellina di caricamento su Telegram per diversi secondi.
  - Causa 3: Il nome del collega veniva desunto forzatamente dall'account Telegram (spesso nickname come `area00`), senza possibilità per Simone di impostare o modificare il vero nome e cognome reale del collaboratore.
- **Risoluzione Implementata**:
  - **Duplicazione Foglio Senza Dipendenza DriveApp (`duplicateTemplateSpreadsheet`)**:
    - Metodo primario: `SpreadsheetApp.openById(TEMPLATE_SPREADSHEET_ID).copy(copyName)`: utilizza unicamente lo scope nativo Google Sheets già autorizzato da sempre, azzerando qualsiasi problema o blocco di autorizzazione.
    - Fallback resiliente: se `copy()` dovesse fallire, interviene `DriveApp.makeCopy` con gestione errori puntuale.
  - **Assegnazione Nome Reale al Collega (`ADMIN_ASK_NAME_` & `ADMIN_WAITING_USER_NAME`)**:
    - Nella notifica di nuova richiesta, Simone ha due opzioni:
      - `[ ✏️ Inserisci Nome e Approva ]`: il bot chiede a Simone di digitare il nome reale in chat (es. *Fabrizio Verdi*). Alla digitazione, il bot salva il nome, clona il foglio con titolo `"Calcolo Tariffa - Fabrizio Verdi"`, attiva il wizard al collega e notifica Simone con link a Drive.
      - `[ ⚡ Approva Rapido ("Nome Telegram") ]`: approva istantaneamente usando il nome Telegram.
  - **Rinomina Collega in Qualsiasi Momento (`ADMIN_ASK_RENAME_` & `executeUserRename`)**:
    - Nella scheda del collega (`👥 Gestione Colleghi`), aggiunto il pulsante sempre disponibile `[ ✏️ Rinomina Collega ]`.
    - Digitando il nuovo nome, il bot aggiorna l'anagrafica e rinomina automaticamente il foglio Google Drive (`Spreadsheet.rename("Calcolo Tariffa - Nuovo Nome")`).
  - **Feedback Istantaneo e Fallback Anti-Errore Telegram**:
    - `answerCallback(cb.id, "Creazione foglio in corso...", false)` scatta all'istante al tocco del pulsante, rimuovendo il loader.
    - `editMessage` e `sendMessageWithKeyboard` implementano fallback automatico senza `parse_mode` se Telegram restituisce errori di formattazione Markdown, garantendo la consegna dei messaggi al 100%.
  - **Suite di Test Estesa a 240 Scenari (`test_suite.js`)**:
    - Validata la richiesta di inserimento nome, il salvataggio personalizzato, l'approvazione del collega, la duplicazione foglio nativa e la rinomina dinamica (**240/240 test superati con successo**).

### AS. Pulizia Funzioni Superflue Backend e Home Sempre Ultimo Messaggio
- **Rimozione Chiamata Ridondante `pruneOldMonthSheets`**:
  - Eliminata la scansione di rotazione schede (`pruneOldMonthSheets`) dalla finalizzazione della registrazione giornaliera (`finalizeEntry`). La rotazione a 6 mesi rimane attiva esclusivamente durante l'archiviazione mensile (`archiveAndClearMonth`), azzerando operazioni superflue ad ogni chiusura turno.
- **Eliminazione Funzioni Dummy Residue**:
  - Rimosse le funzioni vuote `schedule14HourReminder` e `cancel14HourReminder` e relative chiamate: il monitoraggio a 14 ore è affidato interamente al trigger orario centralizzato `handleHourlyMaintenance`, evitando chiamate e funzioni fittizie.
- **Eliminazione Funzioni e Azioni di Debug Orfane**:
  - Rimosse `findOverlappingEntry` e `getTimeSuggestions`, sostituite e rese obsolete dai componenti batch in-memory.
  - Rimossi i test endpoint `inspect_maps`, `delete_mock_sheets`, `dump_sheet_raw` e `populate_user_shifts` da `doGet`.
  - Rimossi `populatePreviousMonths` e `deleteMockMonths` dal codice di produzione (trasferiti direttamente nella test suite locale `test_suite.js`).
- **Home Sempre Come Ultimo Messaggio all'Apertura Turno**:
  - Alla selezione dell'orario di inizio turno (o alla scelta su stesso evento di ieri), il bot elimina il messaggio interattivo precedente con `deleteMessage` e invia un nuovo messaggio `renderHome(chatId)` in fondo alla chat.
  - In questo modo la card Home aggiornata con stato "In Servizio" e la tastiera completa (`🔴 Termina Giornata`, `🗑️ Annulla Giornata`, `📊 Riepilogo Mese`, ecc.) è categoricamente l'ultimo messaggio in chat, senza messaggi duplicati o dispersi.

### AT. Integrazione Google Calendar Multi-Utente, Condivisione Email e Onboarding Semplificato
- **Google Calendar Dedicato Multi-Utente (`Videorent - Nome Collega`)**:
  - Eliminato il vincolo hardcoded `Number(user.id) !== ADMIN_ID` che riservava Calendar a Simone: tutti gli utenti attivi possono ora sincronizzare i propri turni.
  - `getOrCreateVideorentCalendar(user)`:
    - Per Simone (`ADMIN_ID`): crea e gestisce il calendario `Videorent`.
    - Per ciascun collega: crea e gestisce un calendario secondario dedicato `Videorent - [Nome Collega]`.
    - Salva l'ID del calendario nell'anagrafica utente (`calendarId`) per lookups diretti a 0 ms senza ricerche per nome ripetute.
- **Condivisione Nativa Google Calendar via Email (`shareVideorentCalendar`)**:
  - Il collega inserisce la propria email Gmail/Google (da chat o menu impostazioni): Apps Script aggiunge automaticamente l'indirizzo come editor/viewer del proprio calendario turni (`cal.addEditor(email)`).
  - Il calendario compare direttamente dentro l'app **Google Calendar sullo smartphone** del collega (Android o iPhone), sincronizzato in tempo reale.
  - Anche Simone, dal pannello `👥 Gestione Colleghi`, può visualizzare e impostare l'email Google del collega con 1 clic (`ADMIN_ASK_CAL_EMAIL_`).
- **Menu e Comandi Dedicati (`renderCalendarMenu`, `/calendario`, `/calendar`)**:
  - Accessibile dal pulsante `📅 Google Calendar` in `⚙️ Impostazioni` o digitando `/calendario` / `/calendar`.
  - Mostra lo stato del calendario (nome, email collegata, sync attiva/disattivata), il pulsante per collegare/modificare l'email e il toggle per attivare o disattivare la sincronizzazione.
- **Onboarding e Messaggi Chiariti (Zero Confusione)**:
  - **Questionario Tariffe Preservato al 100%**: il wizard a passi consecutivi è mantenuto per garantire la riservatezza e la personalizzazione tariffaria di ciascun professionista.
  - Al termine del wizard, il messaggio chiarisce che **il bot Telegram è operativo al 100% da smartphone** per registrare giornate, calcolare compensi ed esportare PDF, offrendo subito il pulsante per collegare Google Calendar.
  - La guida Tampermonkey chiarisce che l'estensione browser è **opzionale per PC/Mac** (solo per chi compila i consuntivi da computer) e include gli aggiornamenti automatici nativi.
- **Aggiornamento Automatico Nativo Tampermonkey per i Colleghi (v3.9.4)**:
  - Allineato il generatore backend `generateUserScriptContent`: include ora la sincronizzazione automatica del calendario (`sendCalendarSyncPayload`, `extractCardCalendarData`, `scheduleCalendarSync`), intercettazione presenza immediata keepalive e versione bumped a `3.9.4`.
  - Con le direttive `@updateURL` e `@downloadURL`, Tampermonkey si aggiorna automaticamente in background sul browser dei colleghi senza interventi manuali.
- **Suite di Test a 259 Scenari (`test_suite.js`)**:
  - Aggiunto Gruppo 21 con 19 nuovi test automatici a copertura di: calendari dedicati, condivisione email, sync eventi per colleghi, notifiche Telegram, comandi chat, userscript 3.9.4 e guida (**259/259 test superati con successo**).

### AU. Risoluzione Sovrascrittura Stato Turno Aperto di Simone durante Abilitazione/Modifica Colleghi
- **Problema Riscontrato**:
  - Se Simone aveva un turno di lavoro aperto (in corso di svolgimento) e abilitava un nuovo collega (tramite `[ ✏️ Inserisci Nome e Approva ]` o `[ ✏️ Rinomina Collega ]` o collegava l'email Google Calendar), il suo turno aperto spariva e veniva cancellato.
  - *Causa Radice*: Nei callback admin (`ADMIN_ASK_NAME_`, `ADMIN_ASK_RENAME_`, `ADMIN_ASK_CAL_EMAIL_`), il codice eseguiva:
    `setUserState(adminId, { action: "ADMIN_WAITING_USER_NAME", targetUserId: targetId, promptMessageId: messageId });`
    Questo creava un nuovo oggetto letterale sovrascrivendo l'intero stato utente di Simone (`PropertiesService.getUserProperties()`), distruggendo all'istante l'oggetto `state.activeDay` in cui era memorizzato il turno in corso.
- **Risoluzione Implementata**:
  - **Preservazione Integrale dello Stato Preesistente di Simone**:
    - Tutte le funzioni amministrative ora recuperano lo stato esistente con `const adminState = getUserState(adminId);`, impostano esclusivamente i campi di attesa azione (`adminState.action = ...; adminState.targetUserId = ...; adminState.promptMessageId = ...;`) e salvano con `setUserState(adminId, adminState);`.
    - L'oggetto `adminState.activeDay` (con data, tipologia, orario di inizio e dettagli) rimane al 100% preservato e inviolato.
  - **Pulizia Pulita e Sicura in `clearCurrentAction`**:
    - `clearCurrentAction(chatId)` ora elimina anche `targetUserId` e `promptMessageId`, preservando sempre `activeDay`.
    - In `ADMIN_APPROVE_` diretto (pulsante `[ ⚡ Approva come ... ]`), viene invocato preventivamente `clearCurrentAction(adminId)` per resettare eventuali azioni rimaste in sospeso senza intaccare il turno.
  - **Navigazione Diretta alla Home**:
    - Aggiunto il pulsante `[ 🏠 Torna al Menu Principale ]` (`CMD_CANCEL`) nelle card di conferma approvazione (`executeUserApproval`), rinomina (`executeUserRename`) e nella scheda utente (`renderAdminUserDetail`), consentendo a Simone di tornare alla Home del proprio turno con un singolo tap.
  - **Suite di Test a 272 Scenari (`test_suite.js`)**:
    - Aggiunto Gruppo 22 con 13 test dedicati: verifica che l'apertura del turno di Simone, l'arrivo della richiesta di un collega, la richiesta del nome reale, l'invio del nome in chat, la rinomina, l'approvazione diretta e il ritorno alla Home mantengano categoricamente il turno `state.activeDay` aperto, integro e chiudibile senza errori (**272/272 test superati al 100%**).

### AV. Risoluzione Raffica Messaggi Telegram all'Apertura Gestionale (Anti-Flapping Presenza & Protezione Paginazione)
- **Problema Riscontrato**:
  - All'apertura del gestionale Videorent arrivava a raffica una sequenza confusa di messaggi Telegram:
    1. *Falsa Cancellazione e Ricreazione Eventi*: Commesse future residenti in pagina 2 (es. `#42430`, `#42521`, `#42743`) venivano notificate come `🚨 Lavoro Annullato o Rimosso!` e subito dopo come `🆕 Nuovo Lavoro (Già Confermato)`.
    2. *Flapping Stato Presenza*: Commesse confermate (es. `#42657`, `#42743`) oscillavano avanti e indietro nello stesso minuto tra `⚠️ Nuovo Lavoro da Confermare!` e `✅ Presenza Confermata!`.
- **Cause Radice Individuate**:
  - *Cancellazione Prematura Cache Pagine*: All'avvio della pagina, prima che il componente di paginazione HTML completasse il rendering, `detectTotalPages()` restituiva temporaneamente `1`. La funzione `collectAllFutureEvents` eseguiva un ciclo distruttivo `delete catalogData.pages[p]` cancellando la pagina 2 da `localStorage`, e impostava erroneamente `isFullCatalog = true`.
  - *Soglia Anti-Mass-Drop Troppo Alta*: In Google Apps Script la soglia per bloccare cancellazioni accidentali era `candidateRemovals.length >= 4`. Mancando i 3 eventi di pagina 2 (3 < 4), il backend li considerava rimossi da Videorent, li eliminava da Google Calendar e inviava gli alert Telegram. 2 secondi dopo, con la pagina 2 caricata in background, li reinseriva da zero.
  - *Risoluzione Presenza Invertita*: In `extractCardCalendarData`, `hasPendingConfirm` controllava la presenza di `"rifiuta lavoro"` ed era valutato PRIMA di `hasConfirmedPresence`. Poiché su Videorent una card confermata presenta sia il badge `[ ✔️ Hai confermato ]` sia il bottone per rifiutare l'incarico `[ Rifiuta lavoro ]`, lo script declassava la card a non confermata (`isTechnicianConfirmed = false`). Il push del click utente inviava `true`, ma il background sync successivo rimetteva `false`, scatenando un ping-pong continuo.
  - *Doppia Richiesta Concorrente*: `sendImmediatePresenceSync` lanciava simultaneamente sia `fetch(..., { keepalive: true })` sia `GM_xmlhttpRequest`, generando due esecuzioni parallele in conflitto in Apps Script.
- **Risoluzioni Implementate**:
  - **Priorità Assoluta allo Stato Confermato (`extractCardCalendarData`)**:
    - `hasConfirmedPresence` (badge `[ Hai confermato ]`, `[ Presenza confermata ]`) ha precedenza assoluta e categorica.
    - `hasPendingConfirm` verifica esclusivamente bottoni di accettazione pendente (`"conferma presenza"`, `"conferma la tua presenza"`, `"accetta presenza"`), rimuovendo categoricamente `"rifiuta lavoro"` dalla verifica di pending.
    - Se una presenza è confermata, non può mai essere declassata per la semplice presenza dell'opzione di rifiuto.
  - **Protezione Paginazione e Cache Multi-Pagina (`collectAllFutureEvents`)**:
    - Eliminato il ciclo di cancellazione cieca delle pagine `catalogData.pages`.
    - Introdotto `effectiveTotalPages = Math.max(catalogData.totalPages || 1, totalPages)` per proteggere la memoria delle pagine durante il caricamento iniziale.
    - `isFullCatalog` viene impostato su `true` SOLO quando tutte le pagine fino a `effectiveTotalPages` sono state effettivamente scansionate.
  - **Singolo Canale di Rete e Debounce Anti-Race**:
    - `sendImmediatePresenceSync` trasmette su singolo canale (priorità a `GM_xmlhttpRequest` per Tampermonkey, fallback su `fetch keepalive`).
    - Integrato cooldown temporale di 3.5s per evitare che `MutationObserver` scateni sincroni sovrapposti durante il cambio di stato.
  - **Soglia di Protezione Backend Rafforzata (`candidateRemovals >= 2`)**:
    - Nel backend Google Apps Script, se 2 o più eventi futuri mancano contemporaneamente da un sync, la richiesta viene degradata a sync parziale (`safeFullCatalog = false`), azzerando le cancellazioni massive accidentali e lo spam Telegram.
  - **Suite di Test Estesa a 295 Scenari (`test_suite.js`)**:
    - Aggiunto Test 10.10: verifica il blocco preventivo di cancellazioni multiple (2+ eventi futuri assenti) e la tenuta dello stato di presenza (**295/295 test superati al 100%**).
    - Aggiornate le date di collaudo a date dinamiche future per garantire l'immutabilità dei test nel tempo.


### AW. Risoluzione Doppio Prompt Orario di Fine e Menu Sempre a Portata di Mano in Fondo
- **Problemi Riscontrati**:
  1. *Menu Disperso in Alto*: Durante la modifica di giornate passate o l'inserimento di orari/date via chat, il bot editava messaggi vecchi o inviava notifiche intermedie, lasciando il menu e i tasti di navigazione molto in alto nella cronologia Telegram. L'utente doveva continuamente scorrere in alto per continuare.
  2. *Orario di Fine Chiesto Due Volte*:
     - *Causa Radice 1*: In `handleShiftOpenStart`, la chiamata `clearCurrentAction(chatId)` veniva immediatamente sovrascritta da `setUserState(chatId, state)` con la variabile locale `state` non ripulita, lasciando l'utente permanentemente nello stato `CHOOSE_START_TIME` con `mode: OPEN_DAY`. Quando l'utente digitava l'orario di fine (es. `18` o `18.30`), il bot lo interpretava erroneamente come un nuovo inizio turno! Quando poi l'utente premeva "Termina Giornata", il bot chiedeva l'orario di fine una seconda volta.
     - *Causa Radice 2*: `parseTime` non riconosceva la virgola italiana (`18,30`, `18,00`), spazi (`18 30`) né suffissi (`18h`, `19 ore`), scartando l'input con errore e richiedendo l'orario di fine.
     - *Causa Radice 3*: `getEndTimeSuggestions` calcolava suggerimenti legati all'ora attuale del giorno anziché all'orario di inizio del turno inserito, offrendo bottoni incongruenti.
- **Migliorie e Risoluzioni Implementate**:
  - **Chiusura Immediata Turno Aperto con Digitazione Diretta (Zero Prompt Duplicati)**:
    - Se l'utente ha un turno attivo aperto e scrive direttamente un orario valido in chat (es. `18`, `18.30`, `18,30`, `19`), il bot lo acquisisce all'istante come orario di chiusura turno e avanza direttamente alla scelta Extra, senza resettare a Home né richiedere conferme duplicate.
  - **Robustezza di `parseTime`**:
    - Supporto universale per virgola decimale (`18,30`), spazio (`18 30`), suffissi (`18h`, `19 ore`) e singoli decimali (`8.3` -> `08.30`).
  - **Suggerimenti Fine Turno Contestuali all'Inizio (`getEndTimeSuggestions(startTime)`)**:
    - Se è noto l'orario di inizio (es. 08.00), i bottoni mostrano automaticamente la durata lavorativa reale: +8h (16.00), +8.5h (16.30), +9h (17.00 - giornata standard), +9.5h (17.30), +10h (18.00).
  - **Tasti di Navigazione Rapida Sempre in Fondo alla Chat**:
    - Nella lista `renderRecentEntriesList`: aggiunti bottoni `[ ➕ Inserisci Passata ]` e `[ 🏠 Home ]` direttamente a fondo lista.
    - Nel dettaglio giornata `renderEditItemMenu`: aggiunti bottoni `[ 🔙 Lista Giornate ]` e `[ 🏠 Home ]`.
    - In `EDIT_CHANGE_START`, `EDIT_CHANGE_END`, `EDIT_CHANGE_EXTRA`: aggiunto pulsante `[ 🔙 Annulla e torna alla giornata ]` (`CANCEL_EDIT_ITEM`). All'invio del messaggio chat, il prompt precedente viene rimosso (`deleteMessage`) e il bot re-invia una singola card pulita aggiornata in fondo alla chat con banner di successo.
    - In `finalizeEntry`: eliminato il messaggio muto finale (`🏁 Turno concluso`), sostituito da una card completa in fondo alla chat con il pulsante `[ 📝 Inserisci un'altra passata ]`, `[ ✏️ Modifica Giornate ]`, `[ 📊 Riepilogo Mese ]` e `[ 🏠 Menu Principale ]`.
  - **Suite di Test Estesa a 330 Scenari (`test_suite.js`)**:
    - Aggiunto Gruppo 26 con 21 nuovi test automatici a copertura di: parseTime con virgola/spazi/suffissi, suggerimenti orari contestuali, chiusura diretta senza doppio prompt, pulsanti rapidi in lista e finalizzazione consuntivo (**330/330 test superati al 100%**).

### AX. Accesso Diretto Foglio Google Personale per Colleghi
- **Pulsante Dedicato in Telegram**:
  - Aggiunto inline button `[ 📊 Apri Foglio Google ]` in `⚙️ Impostazioni` (`renderSettings`), `📊 Riepilogo Mese` (`renderMonthStats`), e in `📋 Modifica Giornate` (`renderRecentEntriesList`).
  - Apre nativamente l'URL del foglio di calcolo personale del singolo collega (`https://docs.google.com/spreadsheets/d/${userSheetId}/edit`) sia nel browser da PC sia nell'app Google Fogli da smartphone con 1 tocco a latenza zero.
- **Comandi Chat Dedicati**:
  - Aggiunti comandi rapidi `/foglio`, `foglio`, `/sheet`, `sheet`: il bot invia una card di riepilogo con link markdown diretto e pulsante inline.
- **Isolamento Multi-Utente**:
  - Ciascun collaboratore riceve rigorosamente il link al proprio foglio Google personale creato al momento dell'approvazione, mantenendo la riservatezza tra colleghi e Simone.

### AY. Risoluzione Bug Modifica Giornate Vecchie e Date Picker a Scorrimento Indietro
- **Problemi Riscontrati**:
  1. *Bug Grave Modifica Fine Turno Vecchio*: Se un collega o Simone tentava di modificare l'orario di chiusura di una giornata passata da "Modifica Giornate" mentre aveva un turno aperto (o dopo un turno), il bot intercettava l'orario digitato all'inizio di `handleMessage` come chiusura del turno aperto odierno, salvando il nuovo orario sull'ultima giornata registrata invece che sulla giornata selezionata, mentre la giornata vecchia veniva lasciata con orario `00` (o `0.0`).
  2. *Impossibilità di Scorrere Indietro per Giornate Non Inserite*: In "Inserisci Passata" (`promptDateSelection`) erano presenti solo "Oggi" e "Ieri" o l'inserimento manuale da chat (`SET_DATE_CUSTOM`). Non c'era modo di selezionare con pulsanti rapidi i giorni precedenti (es. -2gg, -3gg, -4gg, -5gg) o scorrere indietro nel tempo per recuperare turni arretrati. Inoltre, in "Modifica Giornate" la lista mostrava solo le prime 8 giornate senza possibilità di sfogliare all'indietro le registrazioni più vecchie.
- **Migliorie e Risoluzioni Implementate**:
  - **Priorità Assoluta Edit Turni (`EDIT_ENTRY_*`)**:
    - Spostati i gestori di modifica giornata (`EDIT_ENTRY_START`, `EDIT_ENTRY_END`, `EDIT_ENTRY_EXTRA`) in cima a `handleMessage`, prima di qualsiasi intercettazione del turno aperto odierno.
    - Se l'utente sta modificando una registrazione selezionata, il bot aggiorna direttamente e categoricamente la riga di quella registrazione (`updateRowInSheet`), preservando al 100% l'integrità del turno odierno in corso.
  - **Protezione Anti-Overwrite 00 in `updateRowInSheet`**:
    - Riscritto `updateRowInSheet` per aggiornare la cella di fine turno (Col D) esclusivamente quando un orario di fine valido è specificato, impedendo categoricamente a valori vuoti o non definiti di sovrascrivere l'orario con `0.0` (`00.00`).
  - **Paginazione Completa in Modifica Giornate (`renderRecentEntriesList`)**:
    - La funzione ora preleva fino a 60 registrazioni e implementa una paginazione dinamica a blocchi di 8 con indicatore (`Pagina 1/N`) e pulsanti rapidi `[ ⬅️ Precedenti ]` e `[ ➡️ Più Recenti ]` (`EDIT_LIST_PAGE_`).
  - **Date Picker a Scorrimento Indietro (`promptDateSelection`)**:
    - Integrati pulsanti rapidi per i giorni precedenti con giorno della settimana e data (`Gio 22/10`, `Mer 21/10`, `Mar 20/10`, `Lun 19/10`).
    - Aggiunto pulsante `[ ⬅️ Giorni Precedenti ]` con paginazione a scorrimento (`SET_DATE_OFFSET_`), che consente all'utente di retrocedere nel tempo a blocchi di 6 giorni e selezionare qualsiasi data passata non ancora registrata con un singolo tocco.
  - **Suite di Test Estesa a 362 Scenari (`test_suite.js`)**:
    - Aggiunti Gruppo 27 (16 test apertura foglio personale e comandi) e Gruppo 28 (16 test risoluzione bug modifica orari, paginazione lista e date picker a scorrimento), portando la suite a **362/362 test superati al 100%**.

### AZ. Protezione Fogli Google Colleghi da Modifiche Accidentali (Sola Visualizzazione /preview & Sheet Protection)
- **Problema di Sicurezza e Integrità Dati**:
  - Dando accesso diretto al foglio di calcolo personale, i colleghi rischiavano di toccare inavvertitamente celle contenenti formule complesse (es. calcolo automatico ore colonna D, compenso orario colonna E, maggiorazioni colonna G o totali mensili) da browser o da smartphone, corrompendo la struttura del foglio.
- **Risoluzione e Architettura di Protezione**:
  - **Differenziazione URL Dinamica (`getUserSheetUrl`)**:
    - Centralizzata la risoluzione degli URL in `getUserSheetUrl(chatId, targetSheetId)`.
    - **Per i Colleghi**: i link generati dal bot (pulsante in `⚙️ Impostazioni`, `📊 Riepilogo Mese`, `✏️ Modifica Giornate` e comandi chat `/foglio`, `foglio`) rimandano a `https://docs.google.com/spreadsheets/d/<sheetId>/preview`. La modalità `/preview` di Google Sheets apre il foglio in una visualizzazione pulita, completamente in sola lettura, priva di toolbar di modifica e senza cursore attivo sulle celle, impedendo cancellazioni accidentali e fat-finger da touch-screen.
    - **Per Simone (Admin `12025766`)**: mantiene l'accesso completo in modalità di modifica (`/edit`) sia per il proprio foglio sia per i fogli dei colleghi dal pannello di amministrazione.
  - **Pulsante Chat Intelligente e Avviso di Sicurezza**:
    - I pulsanti inline per i colleghi recitano chiaramente `👁️ Visualizza Foglio Google` (anziché `📊 Apri Foglio Google`).
    - Il comando `/foglio` invia un messaggio rassicurante e chiaro: *"Il foglio si apre in sola visualizzazione protetta da smartphone o browser per evitare modifiche o cancellazioni involontarie di formule e dati. Per qualsiasi modifica (orari, extra, date), usa comodamente i pulsanti di questo bot!"*.
  - **Protezione Warning Nativa sul Foglio (`mainSheet.protect().setWarningOnly(true)`)**:
    - In fase di clonazione/inizializzazione foglio (`cleanSpreadsheetForNewUser`), viene applicata la protezione nativa a livello foglio con avviso (`setWarningOnly(true)`). Anche se un utente dovesse aprire il foglio dal proprio account Google in modalità normale, qualsiasi tentativo di modifica accidentale genera un alert di conferma esplicito a schermo intero da parte di Google Sheets.
  - **Estensione Suite di Test a 378 Scenari (`test_suite.js`)**:
    - Test Gruppo 27 estesi con verifiche automatiche su `/preview` per colleghi, `/edit` per admin, differenziazione tastiere e helper `getUserSheetUrl`, portando la suite a **378/378 test superati al 100%**.


### BA. Riconoscimento Bologna vs Fuori Bologna, Toggle Casa ⇄ Trasferta e Tariffa Fissa / Forfait
- **Riconoscimento Automatico Bologna vs Fuori Bologna (CASA vs TRASFERTA)**:
  - **Problema Risolto**: Nel testo delle card del gestionale, il campo del luogo è etichettato come `"Sede:"` (es. *"Sede: Grand Hotel Rimini"* o *"Sede evento: Fiera di Milano"*). Il controllo precedente cercava genericamente `loc.includes("sede")` sul testo grezzo, scambiando ogni evento per lavoro "in sede" (Casa) e azzerando le trasferte.
  - **Nuova Logica di Parsing**:
    - Estrazione mirata del luogo isolando la vera città/provincia da `location`, `fullLocation` o righe d'indirizzo del DOM.
    - Se la location indica espressamente **Bologna**, **Villanova di Castenaso**, sigla provincia **BO** o *"Sede/Uffici Videorent"*, viene classificata come **🏠 CASA**.
    - Se la location indica qualsiasi altra città o provincia diversa da BO (es. Rimini/RN, Milano/MI, Roma/RM, Firenze/FI, Ferrara/FE, Reggio Emilia/RE, Parma/PR, ecc.), viene classificata automaticamente come **🚗 TRASFERTA**.
- **Pulsante Dinamico `Casa ⇄ Trasferta` in *"✏️ Modifica Giornate"***:
  - Nel menu di modifica singola giornata (`renderEditItemMenu`), aggiunto il pulsante dinamico per invertire tipologia con 1 tocco:
    - Se la giornata è Casa: compare **`🚗 Sposta in Trasferta`**.
    - Se la giornata è Trasferta: compare **`🏠 Sposta in Casa`**.
  - Alla pressione (`EDIT_SWITCH_TYPE`), il bot sposta atomicamente la registrazione tra le sezioni del foglio (Casa righe 3–26 ⇄ Trasferta righe 31–54), preservando orari, data, ID commessa, extra e ricalcolando il compenso con le formule corrette.
- **Supporto Tariffa Fissa / Forfait (Slegata dalle Ore di Lavoro)**:
  - **Esigenza Risolta**: Possibilità di registrare giornate con compenso forfettario concordato (es. 200 €, 250 €, 300 €) mantenendo gli orari reali di lavoro (es. 08:00 - 20:00) per la fedeltà del foglio presenze e consuntivo orario.
  - **Architettura Cella F**: Scrive l'importo concordato direttamente nella cella del compenso (Colonna F), preservando le formule native di somma del foglio (`=SUM`, `=SUMIF`) e lasciando pulita la colonna Extra (G) per le sole spese/diarie effettive.
  - **Comandi e Controlli Telegram**:
    - Tasto **`💵 Tariffa Fissa`** nel menu modifica: consente di digitare l'importo o selezionare tagli rapidi (`150 €`, `200 €`, `250 €`, `300 €`).
    - Quando attiva, il riepilogo evidenzia: `• Compenso: *250 €* 🔒 (Tariffa Fissa)`.
    - Tasto **`🔄 Ripristina Calcolo`**: ripristina all'istante la formula automatica standard copiandola dalla riga modello con `copyTo(..., PASTE_FORMULA)`.

### BB. Gestione Contestuale Separa/Unisci Evento (ID Commessa vs Bot Puro Standalone)
- **Problema & Necessità Risolta**:
  - Quando un utente utilizza l'estensione Userscript su Videorent, le giornate hanno sempre l'**ID Commessa** registrato in Colonna A (es. `#42576`). In questo scenario, le formule native del foglio (Colonna H `=SUMIF(A$...:A$...; A; F$...:F$...)`) raggruppano e sommano automaticamente tutte le giornate con lo stesso ID, mostrando il totale cumulativo sulla commessa.
  - Pertanto, i pulsanti fisici di spostamento righe (`✂️ Separa da Evento Ieri` / `🔗 Unisci a Evento Ieri`) risultavano fuorvianti e inutili per commesse con ID, mentre restano **fondamentali e indispensabili** per chi usa il bot in modalità standalone ridotta (senza estensione o per lavori privati fuori gestionale) dove la contiguità fisica delle righe è l'unico modo per collegare o separare i giorni.
- **Logica Contestuale Intelligente (Idea 1)**:
  - **Giornate con ID Commessa (`item.jobId` presente)**:
    - Nel testo della card viene mostrato chiaramente: `• Commessa: *#42576* 🔗 _(Collegata per ID)_`.
    - I pulsanti `✂️ Separa da Evento Ieri` e `🔗 Unisci a Evento Ieri` vengono categoricamente **nascosti**.
    - Nel callback `EDIT_TOGGLE_EVENT` e in `toggleEventConnection`, una guardia di sicurezza impedisce lo spostamento fisico e avvisa l'utente che il raggruppamento e il totale sono già gestiti dal foglio tramite ID.
    - `getPreviousDayStatus` salta all'istante la lettura del foglio (0 ms overhead).
  - **Giornate senza ID Commessa (`!item.jobId`, Bot Standalone / Lavori Privati)**:
    - Mostra il testo classico: `• Evento: *🔗 Collegato al giorno prima*` (oppure `Separato`).
    - Mostra i pulsanti operativi `✂️ Separa da Evento Ieri` / `🔗 Unisci a Evento Ieri` quando esiste una registrazione nel giorno precedente.
    - Esegue il batch shifting istantaneo delle righe nel foglio di calcolo con formula preservation.
- **Suite di Test Estesa a 414 Scenari (`test_suite.js`)**:
  - Aggiunto Gruppo 31 (16 test automatici end-to-end completi), portando la suite a **414/414 test superati al 100%**.

### BC. Ripristino Suggerimenti Orario di Chiusura Turno Live (`isLiveClosing` su `now`)
- **Problema Riscontrato**:
  - Quando l'utente premeva "🔴 Termina Giornata" / "Chiudi Turno" per terminare il proprio turno in tempo reale (es. alle 19:10), il bot proponeva bottoni con orari fermi alle ore canoniche della giornata (es. dalle 16:00 alle 18:00), completamente sfasati e incongruenti rispetto all'orario attuale.
  - *Causa Radice*: In `getEndTimeSuggestions(startTime)`, la presenza dell'orario di inizio (es. 08:00) scavalcava categoricamente il calcolo basato sull'orario attuale `now`, imponendo sempre i suggerimenti forfettari a 8h-10h dall'inizio (+8h -> 16:00, +8.5h -> 16:30, +9h -> 17:00, +9.5h -> 17:30, +10h -> 18:00) anche per la chiusura di turni attivi che si protraevano fino a sera o terminavano a qualsiasi orario reale.
- **Risoluzione Implementata**:
  - **Differenziazione Contestuale `isLiveClosing`**:
    - In `promptEndTime(chatId, messageId, startTime, isLiveClosing)` e `getEndTimeSuggestions(startTime, isLiveClosing)`, il calcolo a +8h..+10h è stato rigorosamente confinato ai soli **inserimenti manuali/retroattivi** (`state.mode === "MANUAL_ENTRY"`), dove la data è passata e l'orario attuale `now` non ha attinenza col turno.
    - Quando si chiude un turno attivo (`state.mode === "CLOSE_DAY"`, pulsanti "Termina Giornata", "Chiudi Turno", "Chiudi Turno Adesso" da promemoria 14h), il bot calcola ora sempre e categoricamente i suggerimenti asimmetrici attorno all'**orario attuale reale** (-30m, 0m, +30m, +60m, +90m), consentendo all'utente di chiudere con 1 tocco (es. alle 19:10 propone 18:30, 19:00, 19:30, 20:00, 20:30).
- **Suite di Test Estesa a 417 Scenari (`test_suite.js`)**:
  - Aggiunti test di regressione nel Gruppo 26 per validare che la chiusura live calcoli sempre i suggerimenti su `now` e preservi contestualmente i suggerimenti per inserimenti manuali (**417/417 test superati al 100%**).

### BD. Fast-Path Presenza in Tempo Reale, Intercettazione Accetta Lavoro & Esecuzione Secondaria Backfill
- **Problemi Riscontrati**:
  - **Mancata Notifica Telegram all'Accettazione Turno**: quando l'utente o un collega cliccava "Accetta lavoro" su Videorent, non arrivava alcuna notifica su Telegram; il sistema richiedeva di ricaricare manualmente la pagina web (`F5`) per rilevare il nuovo stato.
  - *Cause Radice*:
    1. *Match Rigido dei Pulsanti*: `extractCardCalendarData` e `handleImmediatePresenceConfirmation` cercavano solo `"accetta presenza"` e `"accetta lavoro"`. Se il bottone recitava semplicemente `"Accetta"`, `"Accetta incarico"`, conteneva icone SVG o si apriva una modale/popup di conferma di Bootstrap, l'evento click veniva ignorato e la card veniva considerata già confermata fin dal primo caricamento, bloccando la successiva notifica.
    2. *Chiamate AJAX di Videorent Non Monitorate*: al click su "Accetta", il gestionale comunica col suo server in background senza ricaricare la pagina; lo userscript non ascoltava tali chiamate di rete.
    3. *Nessun Controllo a Pagina Aperta*: la sincronizzazione avveniva solo all'apertura iniziale della pagina, senza scansione o refresh a scheda aperta.
    4. *Collo di Bottiglia su Apps Script*: prima di processare Calendar e inviare la notifica Telegram, `syncVideorentEventsToCalendar` eseguiva `backfillMissingJobIdsInSheet`, aprendo e scansionando sincronicamente il Google Sheet (`SpreadsheetApp.openById`), introducendo un ritardo di 2-3 secondi.
- **Risoluzioni Implementate**:
  - **Intercettazione Universale dei Pulsanti & Modali**:
    - `isPresenceConfirmButton` supporta tutte le varianti (`accetta`, `accetta lavoro`, `accetta incarico`, `accetta turno`, `prendi in carico`, `conferma`, popup di conferma modali).
    - Memorizzazione contestuale dell'ultima commessa interagita per catturare con certezza i click all'interno delle finestre modali di conferma.
  - **Hook di Rete Trasparente (`installNetworkInterceptors`)**:
    - Intercetta in background sia `XMLHttpRequest` che `window.fetch` nativi di Videorent: appena il server risponde con successo alla chiamata di presenza, avvia immediatamente `checkFastCardStateChanges` (0 ms).
  - **Monitoraggio in Tempo Reale (`checkFastCardStateChanges`)**:
    - Snapshot in memoria (`cardStateSnapshot`): rileva all'istante qualunque cambio di stato di presenza, orari o conferma.
    - Ascolto eventi di sistema `visibilitychange` (rientro sulla scheda da altra finestra/app) e `focus`, con battito cardiaco leggero di background ogni 30s.
  - **Fast-Path Notifica Telegram & Flag `justConfirmed`**:
    - Invia immediatamente la sola commessa modificata con flag `justConfirmed: true` (`isFullCatalog: false`), garantendo l'invio della notifica Telegram `✅ Videorent - Presenza Confermata!` in < 0.5s senza iframe lenti.
    - Funzionamento multi-utente collaudato: la notifica viene recapitata sia a Simone che ai colleghi sul proprio chat ID e con il nome del proprio calendario dedicato.
  - **Backfill Foglio come Job Secondario (`backfillMissingJobIdsInSheet`)**:
    - Rimossa l'apertura sincrona e bloccante del foglio Google dall'inizio del sync.
    - Calendario Google e messaggi Telegram prioritari al 100%. Il backfill retroattivo nel foglio gira ora esclusivamente a fine processo e solo su cataloghi completi, azzerando qualsiasi rallentamento.
- **Test Suite Estesa a 440 Scenari (`test_suite.js`):**
  - Aggiunto il Gruppo 33 che certifica l'intercettazione dei pulsanti, la notifica Fast-Path istantanea su Telegram per Simone e colleghi, e l'esecuzione secondaria non bloccante di `backfillMissingJobIdsInSheet` (**440/440 test superati al 100%**).

### BE. API Nativa Videorent, Supporto Calendario FullCalendar, Digest Notifiche e Adaptive Polling (v3.12.0)
- **Problemi e Opportunità Rilevati**:
  - **Lentezza scraping DOM & Iframe**: il parser dipendeva dal parsing del DOM delle card HTML e su alcune pagine ricorreva a iframe nascosti, richiedendo tempo e risorse.
  - **Presenza di API Interna Nativa Videorent**: l'ispezione della rete con `/browser` ha rivelato l'esistenza dell'endpoint nativo `https://g.videorent.it/my-planner.php?action=events`, che restituisce in formato JSON pulito l'intero catalogo commesse (450+ eventi) con ID, date, orari, stato (`presenza_confirmed: 0/1`, `stato`), compenso, troupe, cliente, cartella Drive e note, senza alcun scraping HTML.
  - **Supporto Modalità Calendario (`my-planner.php`)**: diversi colleghi preferiscono consultare il planner nella vista calendario griglia (FullCalendar). Mancava il supporto in tale vista, dove la griglia è sensibile e non deve essere alterata, ma le card evento aprono il popup modale nativo `#eventModal`.
  - **Spam Notifiche Telegram su Import Massivi**: quando più commesse venivano create, modificate o rimosse simultaneamente dal gestionale, il bot inviava raffiche di messaggi individuali anziché raggrupparli.
  - **Consumo Risorse su Schede Inattive**: polling continuativo a intervallo fisso anche quando la scheda del browser rimaneva aperta in background o minimizzata per ore.
  - **Finestra di Ricerca Calendar Eccessiva per Commesse Singole**: la ricerca eventi su Google Calendar utilizzava un buffer largo $\pm 48\text{h}$ anche per Fast-Path a singolo evento.
- **Risoluzioni Implementate**:
  - **Integrazione API Nativa con Fallback Trasparente (`fetchVideorentEventsAPI`)**:
    - Lo userscript interroga direttamente `my-planner.php?action=events`. I dati vengono normalizzati e processati all'istante in memoria (< 0.1s).
    - Se l'endpoint nativo restituisce un errore HTTP o un formato non valido, lo script degrada automaticamente e silenziosamente al tradizionale scraping DOM / iframe, garantendo continuità assoluta e zero rotture.
  - **Supporto Pieno a FullCalendar (`my-planner.php`)**:
    - Esteso il caricamento dello script su `my-planner.php` (vista calendario FullCalendar).
    - La griglia visuale del calendario viene preservata al 100% senza toccare layout o stili.
    - Quando l'utente clicca un evento, lo script intercetta l'apertura del modale Bootstrap `#eventModal` e inietta un badge compatto ed elegante con lo stato del foglio, i compensi e il pulsante per compilare.
    - La sincronizzazione in tempo reale con Google Calendar e Telegram è attiva anche in modalità calendario.
  - **Digest Notifiche Telegram Raggruppate (Anti-Spam)**:
    - Quando si verificano $\ge 2$ variazioni simultanee, il bot raggruppa gli alert in un unico messaggio digest elegante:
      - `⚠️ *Videorent - N Nuovi Lavori da Confermare!*`
      - `🚨 *Videorent - Variazioni Rilevate su N Commesse!*`
      - `🚨 *Videorent - N Lavori Annullati o Rimossi!*`
      - `✅ *Videorent - N Presenze Confermate!*`
    - Elimina il sovraccarico di notifiche preservando la massima chiarezza e leggibilità.
  - **Finestra Mirata Google Calendar Fast-Path ($\pm 2\text{h}$)**:
    - Per aggiornamenti a singolo evento (`events.length === 1`), la query a `CalendarApp.getEvents` usa un buffer mirato a $\pm 2$ ore dall'evento anziché $pm 48$ ore, riducendo i tempi di elaborazione Google Calendar a ~30ms.
  - **Adaptive Polling per Risparmio Energetico**:
    - Tracciamento attività utente (`mousemove`, `keydown`, `click`, `scroll`) con `lastUserActivity`.
    - Quando l'utente interagisce con la pagina o la scheda è visibile, il battito cardiaco esegue controlli rapidi ogni 30s.
    - Se la scheda è nascosta (`document.visibilityState === 'hidden'`) o l'utente è inattivo da oltre 15 minuti, l'intervallo si dilata automaticamente a 2.5 minuti, azzerando il consumo di CPU e batteria.
  - **Comando `/status` e Diagnostica Discreta nel Bot**:
    - Aggiunto il comando `/status` (o `stato`) che mostra una card riassuntiva pulita con lo stato del sistema (Userscript v3.12.0 connesso, Calendar attivo con nome calendario, foglio connesso e giornate del mese).
    - Aggiunta la riga discreta `📡 Stato Sistema` direttamente all'interno del menu `⚙️ Impostazioni`, senza aggiungere alcun pulsante ingombrante nella schermata Home.
- **Test Suite Estesa a 466 Scenari (`test_suite.js`):**
  - Aggiunto il Gruppo 34 che certifica i 4 digest di notifica, la diagnostica `/status`, la finestra ottimizzata $\pm 2\text{h}$, la presenza delle funzioni native e l'adaptive polling (**466/466 test superati al 100%**).

### BF. Modulo Videorent Cloud: Autenticazione Diretta, Notifiche a PC Spento e Compilazione da Telegram (v3.13.0)
- **Problema & Necessità di Emancipazione dal PC**:
  - Prima di questa versione, l'unico modo per sincronizzare gli eventi Videorent e rilevare nuovi lavori o variazioni era tenere una scheda del browser aperta su PC con l'estensione Tampermonkey attiva.
  - Inoltre, la compilazione dei consuntivi (salvataggio slot orari e totale compenso a gestionale) richiedeva necessariamente di accedere al planner da computer e cliccare sui pulsanti verdi dello script.
- **Risoluzioni Implementate nel Backend Bot (`google_apps_script_bot.js`)**:
  - **Autenticazione Cloud Senza Dipendenze (`videorentLogin`)**:
    - Il bot effettua chiamate HTTP POST native tramite `UrlFetchApp` verso `https://g.videorent.it/login.php` trasmettendo `u_name` e `password`.
    - Gestisce la sessione estraendo il cookie `PHPSESSID` e lo valida verificando che l'endpoint `my-planner.php?action=events` risponda con codice 200 e payload JSON valido.
  - **Wizard Telegram con Cancellazione Sicura Password**:
    - Attivabile con `/videorent`, `/login_vr` o dal menu `⚙️ Impostazioni` ➔ `☁️ Videorent Cloud`.
    - Il bot guida l'utente inserendo prima lo username e poi la password.
    - Appena ricevuto il messaggio contenente la password, il bot invoca immediatamente `deleteMessage(chatId, messageId)`, eliminandolo all'istante dalla chat Telegram per non lasciarne traccia.
    - Le credenziali vengono memorizzate in modo isolato e sicuro in `PropertiesService`.
  - **Session Caching ad Alte Prestazioni (`getVideorentSessionForUser`)**:
    - La sessione `PHPSESSID` viene memorizzata in `CacheService` per 2 ore (7200s), evitando login ripetuti. Se il portale restituisce un redirect 302 a `index.php`, il backend invalida automaticamente la cache e ri-autentica l'utente in modo trasparente.
  - **Notifiche a PC Spento & Background Cron Adattivo (`runVideorentCloudCron`)**:
    - **Schedulazione Adattiva Orari Ufficio (Lun-Sab 08:00–17:00 fuso Europa/Roma)**:
      - **In orario ufficio** (Lun–Sab 08:00–17:00): esecuzione attiva ad alta frequenza ogni **15 minuti**, garantendo massima reattività per nuove assegnazioni, variazioni o conferme.
      - **Fuori orario ufficio** (dopo le 17:00, di notte e l'intera Domenica): l'esecuzione viene automaticamente dilatata a **1 volta ogni 60 minuti** verificando il delta temporale `VR_CLOUD_LAST_OFFHOURS_RUN`. Le invocazioni intermedie del trigger escono istantaneamente a 0 ms senza inviare chiamate di rete a vuoto, azzerando il carico sui server e risparmiando oltre il 75% di quote API.
    - Il bot interroga l'API di Videorent, confronta lo stato degli eventi e aggiorna Google Calendar, inviando notifiche Telegram per nuovi lavori da confermare, variazioni o cancellazioni anche se l'utente ha il computer spento da giorni.
    - Possibilità di forzare un controllo istantaneo su richiesta con il comando `/check_vr` o dal menu cloud.
  - **Compilazione Consuntivi Direttamente da Telegram (`/compila` e `compileJobToVideorent`)**:
    - Il comando `/compila` scansiona il Google Sheet del mese e rileva le commesse che hanno orari e compenso calcolati, ma che su Videorent risultano ancora a `invoice_amount === 0`.
    - Mostra una tastiera inline interattiva per compilare una singola commessa (`⚡ Compila #42576 (220 €)`) o tutte insieme (`⚡ Compila TUTTE`).
    - Per ogni commessa, il bot invia le chiamate POST ufficiali a Videorent (`action=save_slot` per ciascun giorno e `action=save_invoice` con il totale spettante calcolato da `calculateEntryPay`).
    - Risponde in tempo reale con la conferma di avvenuta registrazione ufficiale a gestionale.
- **Test Suite Estesa a 484 Scenari (`test_suite.js`):**
  - Aggiunto il Gruppo 35 che collauda l'intero ciclo di vita: login, session caching, cancellazione sicura della password, compilazione slot/fattura, comando `/compila` ed esecuzione cron (**484/484 test superati al 100%**).

### BG. Correzione Calcolo Totale Compenso nel Modale FullCalendar (`#eventModal`) (v3.12.1)
- **Causa del Bug Riscontrato**:
  - All'apertura del modale Bootstrap `#eventModal` nella vista Calendario FullCalendar, il badge visualizzava sempre `Tot: 0 €` nonostante la presenza di giornate registrate e l'importo corretto presente a gestionale.
  - L'aggregazione utilizzava erroneamente `parseFloat(it.totalPay)` anziché `parseFloat(it.dayPay)` (che è la proprietà ufficiale restituita dal backend `readSectionForSync` e utilizzata in `analyzeEventCoverage` nella vista a lista).
- **Risoluzione Implementata (`videorent_sync.user.js` & `google_apps_script_bot.js`)**:
  - Calcolo del compenso totale aggregato tramite `(parseFloat(it.dayPay) || parseFloat(it.totalPay) || 0)`, identico alla logica della vista a lista.
  - Rilevamento automatico dell'eventuale importo già fatturato a gestionale dal campo input `invoice_amount` del modale.
  - Sincronizzazione dinamica dello stato:
    - Se l'evento è già compilato su Videorent e coincide con il foglio: badge verde con `Tot: X € (Compilato) ✅`.
    - Se è rilevata una discrepanza tra Videorent e Foglio Google: badge giallo di avviso `⚠️ Differenza: VR X € vs Foglio Y €`.
    - Se non è ancora compilato su Videorent: badge verde con `Tot: Y € ✅`.


### BH. Protezione Eventi Multi-Giorno su Compilazione Videorent Cloud (v3.13.1)
- **Contesto ed Esigenza Operativa**:
  - Negli eventi di più giorni consecutivi (es. fiera o produzione su 3 giorni dal 01/10 al 03/10), un utente potrebbe voler registrare a fine turno del primo giorno gli orari effettivi svolti su Videorent, prima che l'evento sia completamente terminato.
  - In precedenza, la compilazione salvava subito sia gli slot orari sia l'importo totale calcolato fino a quel momento come fattura, scrivendo una cifra parziale e provvisoria a gestionale (es. 250 € anziché 750 €).
- **Risoluzione e Logica di Protezione a Due Fasi (`compileJobToVideorent` & `getUncompiledJobsForUser`)**:
  - **Calcolo Date Richieste (`getEventDatesCloud`)**: ricava la lista esatta delle date di calendario comprese tra `start` ed `end` dell'evento da Videorent.
  - **Verifica Copertura Integrale (`isFullyCovered`)**: verifica se tutte le date previste dall'evento sono registrate e concluse nel Google Sheet.
  - **Salvataggio Condizionale dell'Importo Fattura**:
    - Gli **slot orari** (`action=save_slot`) vengono SEMPRE salvati per tutte le giornate caricate nel foglio.
    - L'**importo fattura** (`action=save_invoice`) viene inviato **SOLO SE** `isFullyCovered === true` (tutti i giorni sono completati). Se l'evento è ancora in corso (es. Giorno 1 di 3), il campo compenso rimane non fatturato su Videorent (`invoice_amount` resta invariato), scongiurando importi parziali errati.
  - **Feedback Visivo Telegram Personalizzato**:
    - Nel menu `/compila`: se l'evento è in corso mostra il pulsante `⏳ Salva Orari Parziali #ID (X/Y gg)` con etichetta `⏳ (In corso X/Y gg)`. Se completo mostra `⚡ Compila #ID (Tot €) ✅ (Completa)`.
    - Al clic: se parziale notifica con `⏳ Orari Parziali Registrati su Videorent!` evidenziando che il compenso è provvisorio e sarà fatturato a chiusura evento; se completo notifica con `✅ Commessa #ID Compilata con Successo!`.
- **Test Suite Dedicata (`test_suite.js`)**:
  - Aggiunto il test 35.7 nel Gruppo 35 che collauda sia la fase parziale (orari salvati, zero chiamate `save_invoice`) sia la fase conclusiva (tutti gli orari e compenso totale registrati a gestionale). Totale suite: **498/498 test superati al 100%**.

### BI. Nomenclatura Calendar [Da Conf.], Colori Personalizzabili dal Bot, Cron a 2 Fasi e Auto-Import Foglio (v3.14.0)
- **Nomenclatura Calendar Compatta per Smartphone (Opzione C)**:
  - Sostituito il vecchio suffisso con punto di domanda (`Videorent?` e `Videorent Roma?`) con il tag chiaro e compatto:
    - Bologna Confermato: `Videorent`
    - Bologna Non Confermato: `Videorent [Da Conf.]`
    - Fuori Bologna Confermato: `Videorent <Città>`
    - Fuori Bologna Non Confermato: `Videorent <Città> [Da Conf.]`
  - Risolve l'ambiguità per i colleghi e garantisce leggibilità ottimale sui widget smartphone senza troncamenti fastidiosi.
- **Colori Google Calendar Dinamici & Personalizzabili dal Bot**:
  - Gli eventi sincronizzati su Calendar impostano dinamicamente il colore dell'evento tramite l'API Google Apps Script (`primaryCalEvent.setColor`).
  - Colori di default: **Basilico / Verde (10)** per i lavori confermati e **Banana / Giallo (5)** per i lavori in forse. Non appena un lavoro passa a confermato, il titolo e il colore passano istantaneamente a Verde!
  - **Menu Dedicato nel Bot (`📅 Gestione Calendar ➔ 🎨 Colori Eventi Calendar`)**:
    - Permette a ciascun collaboratore di personalizzare individualmente il colore sia per i lavori confermati sia per quelli da confermare, scegliendo tra **tutti i 24 colori completi di Google Calendar** (Cacao, Fenicottero, Pomodoro, Mandarino, Zucca, Mango, Eucalipto, Basilico, Pistacchio, Avocado, Cedro, Banana, Salvia, Pavone, Cobalto, Mirtillo, Lavanda, Glicine, Grafite, Betulla, Barbabietola, Ciliegio, Uva, Ametista) con interfaccia minimale ed essenziale.
    - **Aggiornamento a Caldo Retroattivo Immediato (`updateCalendarEventsColorsLive`)**: Appena l'utente seleziona un nuovo colore o ripristina i default dal bot Telegram, il bot aggiorna all'istante il colore di tutti gli eventi già presenti e futuri su Google Calendar (`calEvent.setColor(...)`), senza dover attendere nuove modifiche orarie.
- **Cron a 2 Fasi & Rilevamento Squadra Condivisa Simmetrico**:
  - Risolto il problema delle notifiche asimmetriche tra colleghi:
    - **Fase 1 (Raccolta Globale)**: Il bot scansiona prima il planner Videorent di tutti i colleghi abilitati e registra la mappa commesse (`updatePlannerCatalogMap`).
    - **Fase 2 (Incrocio e Notifiche Simultanee)**: Genera gli eventi Google Calendar e le notifiche Telegram valorizzando la squadra (`👥 Lavori con: [Nome Collega]`), garantendo che entrambi ricevano subito la notifica con il nome del collega.
  - Standardizzato il vocabolario: eliminato il termine "Troupe", sostituito con "Squadra" o "Lavori con:".
- **Auto-Importazione Commesse Passate nel Foglio a PC Spento**:
  - Nel cron periodico e nel controllo manuale (`/check_vr`), il bot analizza gli eventi conclusi nel passato (`date < today`) aventi orari o fattura e li importa automaticamente nel Google Sheet personale dell'utente tramite `autoImportVideorentJobs` e `backfillMissingJobIdsInSheet`, azzerando la necessità di aprire il gestionale da PC per aggiornare il foglio.
- **Test Suite Dedicata Estesa a 533 Scenari (`test_suite.js`):**
  - Aggiornati i test di nomenclatura Gruppo 10 (`[Da Conf.]` e colori Verde/Giallo) e aggiunti i test 35.8 e 35.9 per la gestione colori, il cron a 2 fasi, l'aggiornamento retroattivo a caldo e tutti i 24 colori completi Calendar.

### BJ. Hub Riepilogo Unificato (Settimana, Mese Corrente, Mese Precedente) e Compila Gestionale in Home
- **Riorganizzazione Menu Principale (Home)**:
  - Al posto di `📅 Settimana` e `📊 Riepilogo Mese`, la riga centrale della Home è ora organizzata con:
    - `[ 📊 Riepilogo ]` (`MENU_SUMMARY_HUB`): Apre il nuovo Hub unificato dei consuntivi.
    - `[ ⚡ Compila Gestionale ]` (`MENU_COMPILE_JOBS`): Accesso diretto e immediato alla verifica e compilazione delle commesse su Videorent direttamente dalla schermata Home.
- **Hub Riepilogo Unificato (`renderSummaryHub`)**:
  - Unico punto di accesso per consultare l'attività lavorativa:
    - `[ 📅 Settimana Corrente ]` (`MENU_WEEK_VIEW`): Mostra il mini-calendario dei 7 giorni della settimana.
    - `[ 📊 Mese Corrente (<NomeMese>) ]` (`MENU_MONTH_STATS_CUR`): Mostra statistiche, giornate Casa/Trasferta, ore ed extra del mese attuale.
    - `[ 🗓️ Mese Precedente (<NomeMese>) ]` (`MENU_MONTH_STATS_PREV`): Mostra il consuntivo del mese precedente calcolato istantaneamente dalla scheda del mese precedente.
    - `[ 📄 Esporta PDF Consuntivo ]` (`MENU_EXPORT_PDF`): Genera ed invia il PDF consuntivo del mese.
    - `[ 🔙 Torna alla Home ]` (`CMD_CANCEL`): Ritorno immediato in-place alla Home.
- **Navigazione Fluida a Due Vie**:
  - Dalla vista settimana e dalla vista mese, la pulsantiera include sia `[ 🔙 Torna a Riepilogo ]` per cambiare periodo sia `[ 🏠 Home ]` per tornare al menu principale.
  - Comandi diretti da chat: digitando `/riepilogo` o `riepilogo` si accede all'Hub; digitando `/compila` o `compila` si apre la compilazione commesse.
- **Test Suite Estesa a 552 Scenari (`test_suite.js`):**
  - Aggiunto il Gruppo 36 con 13 test specifici per l'Hub Riepilogo e il Gruppo 37 con 19 test dedicati per l'Hub Tariffe, l'Hub Connessione Gestionale e la pulizia testi (**552/552 test superati al 100%**).

### BK. Hub Gestione Tariffe, Hub Connessione Gestionale e Pulizia Testi Telegram
- **Compilazione Parziale Multi-Giorno**:
  - Eliminata la didascalia verbosa su "Protezione Eventi Multi-Giorno" e "logica ufficiale".
  - Sostituita con testo pulito e minimale:
    `⏳ *Orari parziali salvati (#<ID>)*`
    `• Giornate: *X/Y* (<date>)`
    `• Totale provvisorio: *<tot> €* (compilo importo a fine evento)`
- **Hub Unificato Gestione Tariffe (`MENU_RATES_HUB` / `renderRatesHub`)**:
  - Tutte le configurazioni economiche (Casa Base, Durata standard K4, Straordinario, Mezza giornata, Trasferta Base, Durata standard K32, Straordinario, Mezza giornata, Regola Weekend/Festivi) sono ora raccolte in un unico Hub centralizzato.
  - Consente di visualizzare e modificare tutti i parametri concordati nel wizard iniziale con navigazione fluida verso i singoli sottomenu e ritorno a Impostazioni.
- **Hub Connessione Gestionale (`MENU_VR_CONNECTION_HUB` / `renderVideorentConnectionHub`)**:
  - Presenta in modo chiaro e intuitivo le due modalità di connessione del bot:
    1. **Videorent Cloud (Consigliato)**: Login diretto da Telegram, notifiche e compilazione da smartphone anche a PC spento.
    2. **Estensione Browser PC (Tampermonkey)**: Compilazione con 1 clic direttamente su `g.videorent.it`.
  - Entrambi i sottomenu ritornano coerentemente all'Hub di connessione.
- **Menu Impostazioni Super Compatto (`renderSettings`)**:
  - Ridotto da 8 a sole 4 righe pulite:
    `[ 💰 Gestione Tariffe ]`
    `[ ⚡ Connessione Gestionale ]`
    `[ 📅 Google Calendar ]`
    `[ 📊 Apri Foglio Google ]    [ 🔙 Torna alla Home ]`
- **Pulizia Radicale Testi Telegram**:
  - Eliminata la didascalia didattica sulle impostazioni dell'app Google Calendar sullo smartphone in `renderCalendarMenu`.
  - Messaggio commesse completate ridotto all'essenziale: `✅ *Tutte le commesse risultano compilate.*`.
  - Tastiera di fine turno `completionKeyboard` aggiornata con `[ 📊 Riepilogo ]` al posto del vecchio `Riepilogo Mese`.

---

### BL. Risoluzione Falsi Positivi Odonimi Stradali Senza Hardcoding (v3.16.0)
- **Problema Riscontrato**: Eventi svoltisi in vie intitolate a città italiane (es. `via trattati di roma 2/4` a Villanova di Castenaso, `Via Roma` a Firenze o `Via Bologna` a Milano) venivano erroneamente classificati come trasferta nella città omonima (`Videorent Roma`).
- **Causa Radice**: La scansione dell'elenco `otherCities` analizzava l'intera stringa grezza dell'indirizzo senza discriminare i nomi di strada dai comuni effettivi di destinazione.
- **Risoluzione Rigorosa Senza Hardcoding (`detectCityAndIsBologna` & `stripStreetOdonyms`)**:
  1. **Nessuna Location o Cliente Hardcoded**: La determinazione è 100% oggettiva e basata unicamente sull'indirizzo e comune fornito. Nessun bypass sul nome del cliente o della sede (es. un lavoro con cliente Videorent che si svolge a Roma o Milano viene correttamente classificato come trasferta a Roma o Milano).
  2. **Priorità Campo Comune/Città (`cit`)**: Se il database Videorent valorizza il comune (es. `Villanova di Castenaso` o `Castenaso`), questo ha precedenza assoluta sull'indirizzo civico.
  3. **Filtro Odonimi Stradali (`stripStreetOdonyms`)**: Rimozione difensiva dei prefissi toponomastici (`via`, `viale`, `corso`, `piazza`, `piazzale`, `largo`, `vicolo`, `strada`, `borgo`) seguiti dal nome via e numero civico prima della ricerca delle città di trasferta (es. `via trattati di roma 2/4` viene esclusa come via e resta `Villanova di Castenaso` ➔ Bologna; `Via Roma 15` a Firenze resta `Firenze` ➔ Trasferta; `Via Bologna 10` a Milano resta `Milano` ➔ Trasferta).
  4. **Unificazione Logica di Assegnazione**: `executeCompileJob` delega interamente a `detectCityAndIsBologna`, garantendo coerenza al 100% tra titoli Google Calendar e tipologia turno Casa/Trasferta nel foglio.

---

### BM. Telegram Mini App (Web App), Wizard Onboarding a 6 Step e Lettore Manuale Integrato (v3.20)
- **Telegram Mini App (TMA)**:
  - Nuova interfaccia grafica moderna accessibile da Telegram tramite il pulsante nativo *"Menu / Apri App"* o con il comando `/app`.
  - Repository pubblico dedicato `c:\!VR\CalcoloTariffaVR-app` (ospitato su GitHub Pages), contenente esclusivamente il codice client frontend (`index.html`) e privo di dati sensibili, token bot, password o ID fogli. L'autenticazione è convalidata crittograficamente sul backend Google Apps Script tramite `Telegram.WebApp.initData` (HMAC-SHA256).
  - Gestione operativa turni live (Avvia / Termina turno con calcolo automatico), registrazione retroattiva con supporto a compensi concordati fissi (`isFixedRate`), calendario mensile a schede con modifica/cancellazione voci e ricalcolo matriciale immediato.
- **Nuovo Wizard di Onboarding Multi-Step a 6 Passaggi**:
  - Esperienza utente unificata e perfettamente allineata tra Bot Telegram e Mini App:
    1. *Casa*: Compenso base (€), durata giornata standard (8h, 9h, 10h o inserimento), mezza giornata (Sì con soglia ore e compenso, oppure Sempre intera con ore a 0), straordinario orario (€/h).
    2. *Trasferta*: Compenso base fuori sede (€), durata giornata standard, mezza giornata, straordinario orario.
    3. *Regola Weekend e Festivi*: Selezione tra 3 modalità (*Tutti*, *Solo Domenica e Festivi*, *Disattivata*).
    4. *Videorent Gestionale*: Selezione tra Videorent Cloud (con verifica credenziali in tempo reale tramite azione `miniapp_test_vr_login`), Tampermonkey PC, oppure Salta.
    5. *Google Calendar*: Inserimento email Gmail per sincronizzazione automatica a calendario o Salta.
    6. *Riepilogo & Attivazione*: Scheda riassuntiva di tutti i parametri e attivazione con 1 tocco.
  - **Modalità Prova Sandbox per Simone**: Pulsante *"🧪 Prova Esperienza Wizard Collega"* nella schermata Gestione Colleghi (e comando `/testwizard` in chat), che permette di testare tutti i 6 passaggi partendo da campi vuoti senza intaccare le tariffe reali né toccare il turno odierno.
- **Riconciliazione Nomenclatura Tariffe (Zero Undefined)**:
  - Mappatura simmetrica di tutti gli alias in `getRatesFromSheet` (`trasfBase` ⇄ `trasfertaBase`, `trasfExtra` ⇄ `trasfertaStraord`, `casaExtra` ⇄ `casaStraord`, `casaShortPay` ⇄ `casaShort`) e fallback difensivi nel frontend, azzerando qualsiasi valore `undefined` nella scheda dettagli collega.
- **Notifica Credenziali Errate Una Tantum**:
  - Modificato il monitoraggio credenziali Videorent: in caso di password modificata sul portale, viene inviata una sola notifica di cortesia al collaboratore interessato (senza reiterazioni ogni 24h).
- **Lettore Manuale Utente Integrato nelle Impostazioni della Mini App**:
  - Aggiunta della sottosezione *📖 Manuale & Guida Utente* in `screen-settings` con modal lettore a tutto schermo `#modal-manual-reader`, navigazione rapida a pillole per capitoli, supporto completo Dark/Light mode e link diretto per consultare o scaricare il file PDF ufficiale `Manuale_Utente_Videorent.pdf`.

### BN. Input Orari Rapido nella Mini App (5 Pill Contestuali + Campo Numerico Diretto, Zero Orologio a Disco)
- **Problema**: L'utilizzo di `<input type="time">` nei modal di avvio/chiusura turno e inserimento/modifica giornate costringeva lo smartphone ad aprire l'orologio nativo del sistema operativo (scelta ore con lancette, poi scelta minuti, poi conferma), rallentando l'inserimento quotidiano per orari che sono sistematicamente arrotondati all'ora o alla mezz'ora (`:00` o `:30`).
- **Risoluzione Implementata (Stesso Approccio Intuitivo della Chat Telegram)**:
  - **Zero Orologio a Disco**: Eliminato completamente `<input type="time">` e sostituito con `<input type="text" inputmode="decimal">`. Su mobile apre direttamente la tastiera numerica senza mai sollevare il selettore a lancette.
  - **5 Pillole Rapide Contestuali (`.time-pills-row`)**:
    - **Inizio Turno (Live)**: 5 pillole asimmetriche calcolate sull'orario corrente arrotondato (`-90m`, `-60m`, `-30m`, `curr`, `+30m`). Con 1 tocco l'orario viene impostato.
    - **Fine Turno (Live)**: 5 pillole asimmetriche (`-30m`, `curr`, `+30m`, `+60m`, `+90m`).
    - **Registrazione Retroattiva e Modifica**:
      - *Inizio*: 5 orari di convocazione mattutini reali (`07:00`, `07:30`, `08:00`, `08:30`, `09:00`, con `08:00` preselezionato).
      - *Fine a salire*: 5 opzioni a salire dalle 17:00 in poi (`17:00`, `17:30`, `18:00`, `18:30`, `19:00`, con `17:00` preselezionato; se l'inizio + 9h eccede le 17:00, parte dinamicamente dalla fine turno).
  - **Digitazione Diretta e Normalizzazione Automatica (`normalizeInputTime`)**:
    - Il collaboratore può toccare una pillola oppure digitare al volo nel campo sottostante cifre semplici (es. `8` ➔ `08:00`, `8.30` o `8:30` ➔ `08:30`, `19` ➔ `19:00`, `19.30` ➔ `19:30`).
    - Alla digitazione o su `blur`, il campo viene formattato automaticamente e la pillola corrispondente si illumina (`.selected`).
  - **Zero Pulsanti Ridondanti**: Nessun tasto incrementale `[-30m]/[+30m]`; l'interfaccia resta pulita, immediata e veloce.

### BO. Risoluzione Definitiva Loop Notifiche su Commesse Multi-Slot / Discontinue (v3.21)
- **Problema Riscontrato**: Ferrari (e altri colleghi) hanno segnalato l'arrivo ripetuto ogni 15 minuti di notifiche Telegram `🚨 Videorent - Variazione Date/Orari!` con il medesimo evento (es. commessa `#42283 – Format Buzzer System Quiz - Reggio Emilia`), sia ad apertura turno che durante la giornata.
- **Causa Radice**:
  - Alcune commesse su Videorent (`my-planner.php?action=events`) hanno sessioni di lavoro discontinue su più slot (es. Slot 1 per allestimento il 02/11/2026 09:00–18:00 e Slot 2 per evento il 04/11/2026–06/11/2026) condividendo lo stesso ID numerico (`ev.id = "42283"`).
  - La sincronizzazione verso Google Calendar (`syncVideorentEventsToCalendar`) utilizzava l'ID commessa secco come chiave della mappa `syncMap[evIdStr]`.
  - Ad ogni ciclo di cron (ogni 15 minuti), lo Slot 1 confrontava la sua data (`02/11`) con l'ultimo valore memorizzato dallo Slot 2 (`04/11`), rilevando una falsa variazione e inviando l'alert Telegram. Subito dopo nello stesso ciclo, lo Slot 2 sovrascriveva la chiave con `04/11`. Al ciclo successivo il processo si ripeteva all'infinito.
  - Inoltre, la logica di pulizia duplicati considerava il secondo evento su Google Calendar come un clone spurio e lo cancellava a turno.
- **Risoluzione Implementata**:
  - **Raggruppamento Eventi per Commessa (`eventsByJob`)**: Gli eventi in ingresso vengono raggruppati per `cleanJobId`.
  - **Identificatori Univoci di Slot**:
    - Per commesse standard a slot singolo, la chiave rimane `cleanJobId` (retrocompatibilità assoluta 100%).
    - Per commesse multi-slot discontinue, le chiavi vengono indicizzate univocamente: `cleanJobId + "_" + (slotIndex + 1)` (es. `42283_1`, `42283_2`).
  - **Matching Intelligente a 2 Fasi tra Slot e Google Calendar**:
    - *Fase 3a*: Matching esatto sull'orario di inizio (tolleranza $\le 60$s) per identificare subito lo slot rimasto invariato.
    - *Fase 3b*: Matching per vicinanza temporale su finestre fino a 14 giorni per gli slot spostati o ripianificati.
  - **Gestione Eventi Multi-Slot su Calendar**: Assegna ogni slot al rispettivo evento Google Calendar; la cancellazione di eventi in eccesso scatta solo se il numero di eventi su Google Calendar supera effettivamente il numero di slot reali.
  - **Protezione Pruning e Limite Proprietà (9KB)**: Inclusa logica di fallback try/catch con pruning automatico degli eventi passati se la dimensione di `PropertiesService` si avvicina al limite di sistema.
  - **Identificazione Pura ID in Colonna A (`findJobIdForDate`)**: Garantita la rimozione del suffisso interno di slot (es. `42283_1` ➔ `#42283`) all'apertura del turno e nel backfill catalogo, preservando l'integrità relazionale di Colonna A.
  - **Compilazione Protetta Multi-Slot & Chiusura Mirata all'Ultimo Giorno (`compileJobToVideorent`, `getJobDiscrepancyInfo`)**:
    - *Giorni intermedi in corso (`today < lastDate`)*: salva regolarmente gli orari dello slot appena concluso (`save_slot`), ma trattiene il compenso a fattura (`isFullyCovered: false`) e **non invia alcun prompt di chiusura a fine turno**, azzerando il disturbo durante commesse lunghe.
    - *Ultimo giorno programmato o commessa conclusa (`today >= lastDate`)*: se tutte le date sono coperte, compila normalmente; se mancano date a foglio, attiva esclusivamente all'ultimo giorno la possibilità di chiudere il totale fatturato maturato con azione esplicita `VR_FORCE_INVOICE_`, con messaggio pulito e oggettivo privo di metacommenti.
  - **Copertura Test End-to-End**: Aggiunto **Gruppo 46** in `test_suite.js` con collaudo loop notifiche (cicli 2 e 3), variazioni reali, purezza ID Commessa Colonna A, assenza prompt nei giorni intermedi, attivazione all'ultimo giorno e forzatura fattura a fine evento (**741/741 test passati**).

### BQ. Sicurezza Repository Pubblico Mini App (CalcoloTariffaVR-app) e Prevenzione Leak
- **Separazione Rigorosa dei Repository**:
  - Backend/Logica: `c:\!VR\App tariffa` (`SxS46/CalcoloTariffaVR`), contiene il backend Apps Script, le automazioni e la configurazione Clasp.
  - Frontend Mini App: `c:\!VR\CalcoloTariffaVR-app` (`SxS46/CalcoloTariffaVR-app`), repository pubblico ospitato su GitHub Pages per Telegram Web App.
- **Politica Zero-Secret per il Frontend**:
  - Nel repository pubblico sono ammessi esclusivamente i file client statici (`index.html`, `sw.js`, `manifest.json`, `Manuale_Utente_Videorent.pdf`).
  - È bandito qualsiasi token bot, password gestionale, spreadsheetId di fogli Google, syncKey esadecimale o dato personale/contabile. L'autenticazione verso il backend avviene unicamente via HMAC crittografico nativo di Telegram (`Telegram.WebApp.initData`).

### BP. Unificazione Terminologica "Lavoro" / "Lavori" e Revisione Manuale Utente v3.21
- **Unificazione Terminologica Globale**:
  - Sostituito in tutti i testi visibili all'utente (bot Telegram, descrizioni eventi Google Calendar, notifiche, modali ed etichette della Telegram Mini App, userscript Tampermonkey) il termine "commessa / commesse" con "lavoro / lavori".
  - Mantenuta piena retrocompatibilità per regex e lookup parser (`/(?:commessa|id|rif\.?|videorent|lavoro)\s*[:--]?\s*(\d{4,7})/gim`) e per le etichette DOM di Videorent (`accetta commessa` e `accetta lavoro`).
  - Aggiornate le asserzioni della test suite (`test_suite.js`), confermando 776/776 test passati con successo (100% pass).
- **Aggiornamento e Riorganizzazione Manuale Utente (`manuale.html` e `Manuale_Utente_Videorent.pdf`)**:
  - Elevata la versione della guida collaboratori a **v3.21 (Novembre 2026)**.
  - Struttura ampliata a **8 pagine A4** bilanciate senza traboccamenti:
    - *Pagina 1 (Copertina)*: Versione v3.21, sintesi caratteristiche chiave (Mini App, Cloud, Calendar con 24 colori e troupe, Pill orari, Conferma/Rifiuto, Scheda Materiali).
    - *Pagina 2 (Capitolo 1 - Primo Accesso e Abilitazione)*: Flusso rapido in 3 step, isolamento dei dati contabili, modalità Bot e Mini App.
    - *Pagina 3 (Capitolo 2 - Il Nuovo Wizard a 6 Step e Gestione Tariffe)*: Tabella completa dei 6 passaggi, flessibilità mezza giornata e modifica tariffe.
    - *Pagina 4 (Capitolo 3 - Utilizzo Quotidiano: Bot e Mini App)*: Nuove 5 pill contestuali orari (-60m, -30m, Now, +30m, +60m), promemoria automatico discreto delle 14 ore e tabella funzioni pulsanti Home.
    - *Pagina 5 (Capitolo 4 - Google Calendar sul Telefono & Troupe)*: Nomenclatura Bologna/Fuori Bologna, elenco colleghi di squadra (`👥 Lavori con:`) e personalizzazione dei 24 colori Google.
    - *Pagina 6 (Capitolo 5 - Convocazioni, Presenza Tecnico & Materiali)*: Nuovo capitolo dedicato a:
      - Conferma immediata e Rifiuto con motivazione ufficiale inoltrata alla logistica.
      - Scheda tecnica materiali e sale con barra di ricerca in tempo reale (per apparato o sala).
      - Gestione avanzata dei lavori multi-slot e date discontinue.
    - *Pagina 7 (Capitolo 6 - Compilazione Gestionale: Cloud & PC)*: Videorent Cloud con compilazione 1-tap a PC spento, estensione PC Tampermonkey con bulk sync, e riconoscimento automatico trasferta per location.
    - *Pagina 8 (Capitolo 7 - Consuntivo PDF Ufficiale & FAQ)*: Esportazione documento contabile A4 e risposte alle domande frequenti.
  - Rigenerato `Manuale_Utente_Videorent.pdf` con Chrome headless (2.1 MB).
  - **Integrazione Manuale nel Bot Telegram**:
    - Aggiunto il pulsante `[📖 Manuale e Guida Utente]` nel menu principale delle Impostazioni (`renderSettingsMenu`).
    - Introdotto il sottomenu `renderManualMenu`: riepilogo dei capitoli e opzioni per scaricare istantaneamente il PDF in chat (`CMD_DOWNLOAD_MANUAL`) o aprirlo nella Telegram Mini App.
    - Aggiornato `USER_MANUAL_FILE_ID` con il file ID nativo di Telegram del nuovo PDF v3.21.
- **Unificazione Terminologica "Tecnico / Tecnici" & Inversione Pulsanti Home**:
  - Rimosso categoricamente il termine "collaboratore / collaboratori" sostituito ovunque con "tecnico / tecnici" in `manuale.html`, nel PDF e nei testi del bot.
  - Invertita la disposizione dei pulsanti nella Home del bot (`getHomeKeyboard`): `[✏️ Modifica Giornate]` posizionato subito sotto `[🟢 Inizia Giornata]`, mentre `[📝 Inserisci Passata]` è stato spostato in basso affiancato a `[⚙️ Impostazioni]`.

---

## 4. Automazione del Flusso di Sviluppo (Zero Copia-Incolla)

### Tampermonkey (Sviluppo Diretto da File Locale via `@require`)
1. Nelle estensioni del browser (Chrome/Edge) abilitare *"Consenti l'accesso agli URL dei file"* per Tampermonkey.
2. Nelle impostazioni avanzate di Tampermonkey, sotto **Esterni (Externals)**, impostare **Intervallo di aggiornamento** su **Sempre (Always)**.
3. In Tampermonkey è installato solo `tampermonkey_loader.user.js` con la direttiva:
   `// @require file:///C:/!VR/App%20tariffa/videorent_sync.user.js`
4. **Risultato**: Qualsiasi modifica salvata nel file locale `videorent_sync.user.js` viene ricaricata ed eseguita immediatamente al primo F5 su Videorent.

### Google Apps Script (Deploy con Clasp)
- **Script ID**: `1TvLc26nwaR3M6ILw5glXenpDZLfa-zy48kXbH2fRjflhGGMWCKO5m-eu`
- **Deployment ID Produzione**: `AKfycbxZZMRCEy3OmmuQht3ABI3_o4YOFfrITiWRnQxQz4hL9rU5LG2zT6HXLmgmVOwQMw`
- **Comando di Deploy automatico**:
   Basta eseguire (o fare doppio clic su) `aggiorna_bot.bat`:
   ```cmd
   node test_suite.js
   npx @google/clasp push
   npx @google/clasp deploy -i AKfycbxZZMRCEy3OmmuQht3ABI3_o4YOFfrITiWRnQxQz4hL9rU5LG2zT6HXLmgmVOwQMw -d "Aggiornamento automatico"
   ```
- **Risultato**: Esegue il collaudo di tutti i 117 test, carica il codice e aggiorna la versione di produzione della Web App e del Bot Telegram in ~5 secondi senza cambiare l'URL pubblico.

### Repository GitHub (`SxS46/CalcoloTariffaVR`)
- Repository remoto: `https://github.com/SxS46/CalcoloTariffaVR.git`
- Branch principale: `main`
- Sincronizzazione tra computer portatile e PC fisso tramite `git push` e `git pull`.

---

## 5. Regole Obbligatorie per Modifiche Future (Per Agenti e Sviluppatori)

1. **COMMIT & PUSH SEMPRE OBBLIGATORIO**:
   - **Ogni volta che viene effettuata qualsiasi modifica a qualsiasi file del progetto, l'agente DEVE eseguire immediatamente `git add -A`, creare un commit descrittivo ed eseguire `git push origin main`.** Il repository remoto GitHub deve essere sempre allineato al 100%.
2. **Validazione Sintassi**:
   - Se si modifica `videorent_sync.user.js`, testare sempre la sintassi con `node -c "videorent_sync.user.js"`.
   - Se si modifica `google_apps_script_bot.js`, testare la sintassi con `node -c "google_apps_script_bot.js"`.
3. **Deploy Apps Script**:
   - Dopo aver modificato `google_apps_script_bot.js`, eseguire sempre `aggiorna_bot.bat` (oppure `node test_suite.js && npx @google/clasp push && npx @google/clasp deploy -i AKfycbxZZMRCEy3OmmuQht3ABI3_o4YOFfrITiWRnQxQz4hL9rU5LG2zT6HXLmgmVOwQMw -d "<nota>"`).
4. **Struttura Dati Google Sheets (8 Colonne)**:
   - **Colonna A**: `ID Commessa` (es. `#42576`, assegnato automaticamente dal Bot attingendolo dal catalogo Videorent o backfill dallo userscript; vuoto per turni extra/privati).
   - **Colonna B**: `Data`
   - **Colonna C**: `Inizio`
   - **Colonna D**: `Fine`
   - **Colonna E**: `Ore`
   - **Colonna F**: `Compenso`
   - **Colonna G**: `Extra`
   - **Colonna H**: `Consecutivo / Totale Evento` (formula aggregata nativa con punto e virgola `;` che somma tutte le giornate con lo stesso ID e mostra il totale solo sull'ultima riga dell'evento; per i turni senza ID mostra il compenso del singolo turno).
   - **Colonna K**: `Tariffe` (`K2:K6` Casa, `K30:K34` Trasferta).
   - Mantenere la sincronizzazione sequenziale con ritardo minimo di 60-120ms per le chiamate a Videorent.
5. **Attesa di Approvazione Esplicita su Piani e Domande**:
   - **Ogni volta che viene generato o aggiornato un `implementation plan` oppure ci sono domande per Simone, l'agente NON deve procedere autonomamente: DEVE fermarsi e attendere sempre la risposta o l'approvazione esplicita di Simone prima di iniziare a modificare il codice.**
6. **COLLAUDO AUTONOMO OBBLIGATORIO AD OGNI SINGOLA MODIFICA (`node test_suite.js`)**:
   - **L'agente DEVE eseguire autonomamente la suite di test completa (`node test_suite.js`) dopo ogni singola modifica al codice, senza attendere che sia l'utente a chiederlo.**
   - Tutti i test devono passare categoricamente al 100% (145/145 test attuali). Non rilasciare, committare o deployare mai codice che fallisca i test.
7. **FUNZIONAMENTO E TEST MULTI-UTENTE / COLLEGA OBBLIGATORIO**:
   - **Qualsiasi modifica architetturale, tariffaria, di apertura/chiusura turno o di sincronizzazione DEVE funzionare perfettamente ed essere testata anche nello scenario in cui un collega utilizzi il software.**
   - Verificare categoricamente:
     - Isolamento assoluto dei fogli di calcolo di ciascun utente (`cleanSpreadsheetForNewUser`, isolamento drive ID).
     - Tariffe personali del collega (il collega ha le sue tariffe configurate, non quelle di Simone).
     - Userscript Tampermonkey dedicato per il collega (`generateUserScriptContent`) con prefisso cache locale isolato per chiave sync (`vr_sheet_sync_cache_<prefix>`).
     - Zero interferenze o cross-talk tra Simone e i colleghi.
8. **PROATTIVITÀ NELL'OTTIMIZZAZIONE ARCHITETTURALE (Zero Cicli Inutili & Standby di Default)**:
   - **Tutte le ottimizzazioni di performance, persistenza e riduzione dei cicli DEVONO essere applicate subito alla prima stesura del codice, senza attendere che sia l'utente a richiederle.**
   - **Checklist Proattiva Obbligatoria per ogni Modifica**:
     1. *Idempotenza & Persistenza Stato*: memorizzare sempre firme/hash su storage persistente (`localStorage` / `CacheService` / `PropertiesService`). Nessun flag volatile solo in RAM che provochi la riesecuzione di calcoli o chiamate di rete al reload (F5).
     2. *Disaccoppiamento UI vs Sincronizzazione*: mai annidare chiamate di rete o trigger di sync all'interno di funzioni grafiche di rendering/decorazione (`processCards`).
     3. *Event-Driven & Standby Immediato (IDLE)*: una volta terminato il lavoro, il sistema deve entrare categoricamente in modalità riposo silenzioso (0% CPU, zero chiamate di rete), risvegliandosi esclusivamente tramite eventi specifici (`MutationObserver`, click utente o timer discreti a bassa frequenza).
     4. *Caching Intelligente con TTL*: qualsiasi richiesta di background (iframe, multi-pagina, API) deve avere un controllo di freschezza (TTL 5–10 min) per evitare richieste ridondanti.

9. **DIVIETO ASSOLUTO DI INVENTARE O FABBRICARE DATI (Zero Mock / Zero Fallback Fittizi)**:
   - **È categoricamente vietato all'agente e al software inventare, simulare o fabbricare qualsiasi dato reale (date, importi, compensi, orari, commesse o schede mensili).**
   - Qualsiasi dato deve provenire esclusivamente ed autenticamente da una fonte reale (il gestionale Videorent o le registrazioni dell'utente nel Google Sheet).
   - Se un dato non è presente o non è noto, il sistema NON deve MAI inventarlo né tramite fallback arbitrari né assemblando estremi orari eterogenei: deve essere registrato o mostrato come non disponibile/incompleto, senza alcuna alterazione o finzione.

10. **ESCLUSIONE RIGIDA FORMAZIONE NON RETRIBUITA & EVENTI A COMPENSO ZERO (Regola Storico Non Compilato)**:
   - **Possono capitare eventi o giornate di formazione non retribuite o commesse senza compenso (`invoiceAmount <= 0`).**
   - Tali eventi vanno **tassativamente lasciati NON COMPILATI nello storico** su Videorent.
   - **NON DEVONO MAI entrare nel Foglio di Calcolo Google Sheet** (né con orari, né con importi fittizi o tariffe base di default).
   - `isCompiled` richiede categoricamente `invoiceAmount > 0` ed esclude rigorosamente i lavori futuri.
   - `autoImportVideorentJobs` scarta categoricamente all'inizio del ciclo qualsiasi commessa con `invoiceAmount <= 0`.
   - Rimossa qualsiasi assegnazione di tariffa base (`rates.casaBase` / `rates.trasfertaBase`) come fallback per commesse a compenso zero.

11. **PROTEZIONE DOPPIA RIGA ODIERNA & PRIORITÀ ORARI REALI (Anti-Ghost & Riutilizzo Segnaposto)**:
   - **Problema Risolto (07/10)**: Se nel foglio era presente una riga segnaposto con data odierna ma senza orari (da auto-import preventivo o consultazione anticipata gestionale), la successiva registrazione da Telegram inseriva una seconda riga odierna anziché riutilizzare la prima. Nello Userscript, `find()` selezionava la prima riga priva di orari, marcando il turno come pendente e ignorando la seconda riga valida.
   - **Risoluzione a 3 Livelli**:
     1. *Riutilizzo Segnaposto in `insertChronologicalEntry`*: se nel foglio è già presente una riga per la stessa data con orari vuoti (`sVal === ""` e `eVal === ""`), il bot la riutilizza direttamente senza scalare le righe e senza duplicati.
     2. *Priorità Assoluta Orari Validi in `analyzeEventCoverage`*: `find()` cerca prioritariamente match con orari compilati (`match.start && match.end`), impedendo a eventuali righe vuote di oscurare un turno già registrato.
     3. *Filtro Preventivo Auto-Import*: `checkAndAutoImportMissingJobs` e `autoImportVideorentJobs` scartano categoricamente qualsiasi commessa priva di consuntivazione o di orari validi per la data odierna o futura, aggiornando solo il catalogo.

12. **DURATA GIORNATA INTERA TECNICO (CELLE NATIVE K4 E K32)**:
    - **Personalizzazione Durata Standard**: Ciascun tecnico concorda quante ore dura la sua giornata intera (es. 8h, 9h, 10h) prima che scattino gli straordinari.
    - **Wizard Colleghi**: Domanda esplicita per Casa e per Trasferta con bottoni rapidi (8h, 9h, 10h) e digitazione chat.
    - **Google Sheet**: Memorizzato nelle celle `K4` (Casa) e `K32` (Trasferta) in formato frazione giorno (`ore / 24`) con formato orario `"hh:mm"`, rispettando le formule del foglio native.
    - **Menu Impostazioni**: Modificabile in qualunque momento da Simone e dai colleghi da `🏠 Configura Casa` e `🚗 Configura Trasferta`.


13. **SYNC CALENDAR ESCLUSIVAMENTE PER EVENTI FUTURI (ZERO EVENTI PASSATI)**:
    - **Regola Ferrea**: La sincronizzazione su Google Calendar (sia da Userscript che da Bot Cloud) opera **SOLO ED ESCLUSIVAMENTE sugli eventi futuri o odierni** (`endDate >= startOfToday`).
    - Gli eventi passati (conclusi prima della data odierna) non vengono **mai sincronizzati, né cercati, né modificati né eliminati** su Google Calendar.
    - Zero notifiche Telegram per commesse passate o archivi storici.
    - Finestra di ricerca su Calendar calcolata unicamente su eventi futuri (azzerando le query storiche pluriennali).
    - **Location Detection a 5 Parametri**: `detectCityAndIsBologna` ispeziona `locationStr`, `venueStr`, `cityStr`, `headingStr` e `clientStr`. Rileva tutte le location bolognesi storiche (Royal Hotel Carlton, Starhotel, Savoia, Duse, Fiera, Sede, Fermo materiale, ecc.) assegnando Bologna (`isBologna = true`); se sono menzionate altre città esterne (Roma, Milano, Rimini, ecc.) estrae la città specifica.
    - **Nomenclatura Opzione C**: `Videorent` / `Videorent [Da Conf.]` a Bologna; `Videorent <Città>` / `Videorent <Città> [Da Conf.]` fuori Bologna. Nessun punto interrogativo (`?`).
    - **Colori**: Verde (`10`) per confermati, Giallo (`5`) per da confermare, personalizzabili liberamente tramite bot.
    - **Anti-Spam Telegram**: Nessuna notifica inviata per eventi passati, con circuit breaker protettivo di massimo 5 notifiche per esecuzione.

14. **AUTO-IMPORT NEL FOGLIO SOLO PER L'ANNO IN CORSO (DIVIETO ANNI PRECEDENTI)**:
    - **Regola Ferrea**: Sia per l'auto-import da Userscript (`scanStoricoForMissingJobs`) sia per il Bot Cloud (`convertCloudEventsToImportJobs` e `autoImportVideorentJobs`), vengono elaborate e scritte nel Google Sheet **SOLO ED ESCLUSIVAMENTE le commesse dell'ANNO IN CORSO** (`dateObj.getFullYear() === currentYear`).
    - Gli anni precedenti (2019, 2020, 2021, 2022, 2023, 2024, 2025) vengono **categoricamente scartati a monte**: non generano schede di archivio storico e non creano righe nel foglio.
    - La logica tra Userscript Tampermonkey e API JSON Cloud del Bot è identica e simmetrica al 100%.

15. **DIVIETO ASSOLUTO DI UTILIZZARE INIZIO E FINE EVENTO DI GESTIONALE PER INVENTARE GIORNATE O CALCOLARE TARIFFE NEL FOGLIO**:
    - **Principio Inviolabile**: Gli orari `start` ed `end` restituiti dal gestionale Videorent rappresentano unicamente la finestra temporale complessiva dell'evento/commessa del cliente.
    - **Divieto Ferreo di Fabbricazione**: È categoricamente vietato al Bot, al Cloud Cron o a qualsiasi script prendere gli estremi di inizio/fine evento dal gestionale per calcolare compensi, interpolare date intermedie, inventare orari o inserire/proporre giornate lavorative nel Google Sheet.
    - **Il Google Sheet è scritto e deciso unicamente dal tecnico**: Soltanto le registrazioni effettive del tecnico (apertura/chiusura turno live da Telegram o inserimento manuale nel foglio) hanno valore per determinare i giorni e le ore lavorate.
    - `convertCloudEventsToImportJobs` non deve fabbricare alcuno slot o compenso fittizio.
    - Il flusso di compilazione opera esclusivamente dal Foglio Google verso Videorent (`/compila`), MAI e poi mai il contrario con orari fittizi o presunti.

16. **UI TELEGRAM MINIMALE ED ESSENZIALE (ZERO TESTI RIDONDANTI O OVVITÀ)**:
    - **I testi dell'interfaccia Telegram del bot devono essere asciutti, minimali e strettamente utili.**
    - **VIETATO esplicitare spiegazioni ovvie, didascalie ridondanti o tecnicismi interni** (es. vietato scrivere frasi come "Tocca uno degli 11/24 colori ufficiali di Google Calendar", "Puoi personalizzare i colori...", "Tocca un pulsante per configurare...", "Seleziona un'operazione:", ecc.).
    - L'utente comprende intuitivamente i pulsanti e la schermata: mostrare solo il dato, lo stato o la domanda indispensabile con zero rumore visivo.

17. **SEMPLIFICAZIONI ONBOARDING, LINK 1-CLIC GOOGLE CALENDAR E ACCESSO DIRETTO FOGLIO (ZERO ATTRITO)**:
    - **Link Diretto 1-Clic Google Calendar**:
      - Sia nel menu del bot Telegram del collega (`renderCalendarMenu`) sia nella scheda amministratore di Simone (`renderAdminUserDetail`), è presente il pulsante e link diretto `https://calendar.google.com/calendar/render?cid=<calendarId>`.
      - Aprendo questo link da qualsiasi dispositivo (browser o mobile), Google Calendar apre direttamente il popup "Aggiungere calendario? Vuoi aggiungere il calendario Videorent - <Nome>?" con conferma immediata in 1 solo tap, azzerando le configurazioni manuali e le attese.
      - Risoluzione difensiva immediata: se `u.calendarId` non è ancora memorizzato nell'anagrafica, viene risolto e salvato all'istante tramite `getOrCreateVideorentCalendar`.
    - **Accesso Foglio Google Senza Richiesta Permessi**:
      - La funzione `makeSpreadsheetViewableWithLink(sheetId)` imposta `ANYONE_WITH_LINK` in sola visualizzazione (`DriveApp.Permission.VIEW`).
      - Viene applicata automaticamente sia alla clonazione iniziale del foglio al momento dell'approvazione del collega, sia durante la manutenzione oraria automatica per tutti i colleghi esistenti. Nessun bisogno per Simone di approvare richieste di accesso o condividere permessi manuali via Drive.
    - **Enforcement Data Passata su Registrazione Manuale Retroattiva**:
      - Se durante un inserimento manuale retroattivo l'utente inserisce giorno e mese senza specificare l'anno (es. `28/10`), e la data ricade nel futuro rispetto a oggi, `parseDateString` scala automaticamente all'anno precedente (`now.getFullYear() - 1`). È comunque sempre possibile specificare l'anno in modo esplicito (`GG/MM/AA` o `GG/MM/AAAA`).

18. **MONITORAGGIO SALUTE GESTIONALE VIDEORENT, RATE LIMITER & ALERT MIRATI (`recordVideorentHealth`)**:
    - **Distinzione Netta Guasto Sistemico vs Credenziali Utente**:
      - *Guasto Sistemico del Portale*: HTTP 404, HTTP 500/502/503/504, connection refused, timeout di rete, o risposta 200 priva del JSON/array eventi atteso (es. pagina di manutenzione o cambio radicale di framework). Gestito a livello centrale per Simone (Admin).
      - *Credenziali Singolo Collega Non Valide*: errata password o sessione non generata su login singolo. Zero spam per Simone; notifica inviata direttamente al singolo collega con rate limiter di massimo 1 avviso ogni 24 ore e pulsante diretto `[ ☁️ Aggiorna Password Videorent ]`.
    - **Filtro Anti-Falso Allarme (Doppia Conferma)**:
      - Un singolo errore transitorio isolato (micro-disconnessione o timeout di pochi secondi) non invia allarmi.
      - L'allarme scatta **solo dopo 2 controlli falliti consecutivi** (`VR_OUTAGE_CONSECUTIVE_FAILS >= 2`, equivalenti a 15-30 minuti di disservizio persistente confermato).
    - **Rate Limiter Anti-Spam su Simone (Admin)**:
      - In caso di disservizio prolungato, l'allarme viene inviato a Simone con cooldown di **almeno 4 ore** (`cooldownMs = 4 * 60 * 60 * 1000`).
    - **Auto-Recovery Notification (Notifica di Ripristino)**:
      - Non appena il gestionale torna a rispondere con successo (HTTP 200 e JSON valido), se era attivo uno stato di outage (`VR_OUTAGE_IS_ACTIVE === "true"`), il bot invia a Simone una notifica di ripristino:
        `🟢 Connessione Gestionale Ripristinata: il portale Videorent è nuovamente raggiungibile e risponde correttamente alle richieste.`
      - Il contatore fallimenti viene resettato a 0.

19. **WIZARD ONBOARDING FLUIDO, RIPRESA AL TOCCO E STEP GESTIONALE / CALENDAR INTEGRATI**:
    - **Ripresa Automatica al Tocco (`resumeOrStartWizard`)**:
      - Quando un collega è in stato `APPROVED_PENDING_WIZARD` (dopo approvazione o reset), se tocca qualsiasi pulsante precedente o apre il bot, il bot non mostra più solo il tooltip toast:
      - Il messaggio viene aggiornato immediatamente sul posto visualizzando la schermata esatta del passaggio del wizard a cui si trova l'utente (o avviandolo da capo se non iniziato). Nessun bisogno di digitare `/start`.
    - **Sequenza Wizard Fluida e Integrata**:
      - *Passo 1*: Configurazione Tariffe Casa (Base, Durata, Mezza Giornata, Straordinario).
      - *Passo 2*: Configurazione Tariffe Trasferta (Base, Durata, Mezza Giornata, Straordinario).
      - *Passo 3*: Regola Weekend & Festivi (Sab+Dom / Solo Dom / Disattivata) ➔ Salva subito le tariffe nel foglio Google (`saveWizardRatesToSheet`).
      - *Passo 4*: **Collegamento Gestionale Videorent** (`renderWizardStepVideorent`):
        - Spiega le 2 modalità proponendo **Videorent Cloud (Consigliato)** come principale.
        - Opzione `[ ☁️ Accedi con Videorent Cloud ]`: chiede username e password, valida l'accesso con test login, cifra le credenziali e abilita il cloud sync.
        - Opzione `[ 💻 Usa Estensione PC (Tampermonkey) ]`: mostra link installazione 1-clic e pulsante Continua.
        - Opzione `[ ⏭️ Salta questo passaggio ]`: prosegue al passo successivo.
      - *Passo 5*: **Sincronizzazione Google Calendar** (`renderWizardStepCalendar`):
        - Chiede se sincronizzare i turni con l'agenda Google del telefono.
        - Opzione `[ 📧 Sì, Collega la mia Email Google ]`: chiede la Gmail, condivide il calendario secondario (`shareVideorentCalendar`), e genera subito il pulsante 1-clic **`➕ Aggiungi al mio Google Calendar`** (`addCalUrl`) con istruzioni chiare per attivare la spunta "Sincronizza" nelle impostazioni dell'app.
        - Opzione `[ ⏭️ Salta questo passaggio ]`: conclude direttamente.
      - *Completamento*: L'utente tocca `[ 🚀 Entra nel Bot e Inizia ]` e approda direttamente nella schermata Home operativa pronta all'uso.

20. **MODALITÀ TEST WIZARD PER ADMIN, PRESERVAZIONE TURNO ATTIVO E PULIZIA MESSAGGI (1-MESSAGE UI)**:
    - **Test User Experience per Simone (`/testwizard`)**:
      - Simone può testare in prima persona l'esatto wizard vissuto da un collega appena approvato senza alterare le proprie tariffe o il proprio stato nel database (`isTestWizard: true`).
      - Accessibile digitando il comando `/testwizard` o dal pannello `👥 Gestione Colleghi`.
      - Pulsante sempre visibile `[ 🚪 Esci dalla Prova ]` in ogni schermata per tornare istantaneamente alla Home.
    - **Preservazione Turno Attivo (`stashedActiveDay`)**:
      - All'avvio del wizard (sia test che reale), l'eventuale turno odierno in corso (`state.activeDay`) non viene cancellato ma viene salvato in `state.stashedActiveDay`.
      - All'uscita dal test o al termine del wizard, il turno attivo viene automaticamente ripristinato, evitando che il bot si "dimentichi" la giornata in corso.
    - **Interfaccia Pulita a Singolo Messaggio (`sendWizardStep`)**:
      - Durante i passaggi che richiedono digitazione testo (es. tariffe, credenziali Videorent, email Google), ogni nuovo prompt del bot elimina la domanda precedente (`state.lastPromptMessageId`).
      - In combinazione con l'eliminazione immediata dei messaggi inviati dall'utente, la chat rimane sempre pulita con esattamente 1 messaggio attivo, senza l'accumulo di decine di domande orfane.

21. **REGOLA FONDAMENTALE DI PARITÀ TOTALE MINI APP TELEGRAM ⇄ BOT TELEGRAM**:
    - **Principio di Specularità Assoluta**:
      La Telegram Mini App (TMA) grafica deve essere sempre uno **specchio funzionale al 100%** del Bot Telegram. Qualsiasi funzionalità, impostazione o controllo presente nel bot DEVE essere fruibile e presente con pari completezza anche nella Mini App, e viceversa:
      - **Gestione Google Calendar**: Toggle ON/OFF sincronizzazione, inserimento/modifica email Google (`calendarEmail`), condivisione automatica calendario, link 1-clic `addCalUrl`.
      - **Collegamento Gestionale Videorent**:
        - *Videorent Cloud*: visualizzazione stato connessione, form per inserimento/aggiornamento credenziali (`vrUsername` e `vrPassword`), pulsante di compilazione immediata su Videorent.
        - *Estensione PC (Tampermonkey)*: link nativo 1-clic di installazione Userscript (`videorent_sync.user.js`), pulsante per copiare la chiave Sync (`syncKey`), download del file `.user.js`.
      - **Gestione Tariffe Personali**: Base Casa, Base Trasferta, Mezza Giornata (ore e compenso per entrambe), Straordinari orari, Regola Weekend & Festivi (`ALL`, `SUNDAY_AND_HOLIDAYS`, `NONE`).
      - **Consultazione Documenti e Foglio**: Visualizzazione Foglio Google personale su Drive, generazione e consultazione Consuntivo Mensile (con opzioni Stampa/Salva PDF ed invio diretto del file PDF su chat Telegram via bot).
      - **Navigazione Storico Mensile**: Navigazione mesi precedenti e successivi sia nel riepilogo statistiche che nell'elenco delle giornate registrate, con ordinamento cronologico **decrescente** (giornate più recenti in alto).
      - **Pannello Admin per Simone**: Approvazione, sospensione e riavvio wizard dei colleghi accessibile da entrambe le interfacce.

22. **DESIGN RESPONSIVE, ACCESSIBILITÀ FONT GRANDI E ALLINEAMENTO UX (MINI APP)**:
    - **Allineamento Verticale Righe Form (`form-grid`)**:
      - Quando due campi sono affiancati (es. Base e Straordinario, Inizio e Fine, Mese e Anno), le etichette (`.form-label`) possono andare a capo a causa di font ingranditi dall'utente (accessibilità Android/iOS) o schermi stretti (≤ 360px).
      - Per prevenire qualsiasi disallineamento verticale dei campi di input, le etichette hanno `min-height: 2.2em; display: flex; align-items: flex-end;` e i controlli `.form-input, .form-select` hanno `margin-top: auto; width: 100%;`. In questo modo entrambi gli input si ancorano sempre alla stessa identica linea orizzontale inferiore, senza scalini.
      - Le etichette devono essere sintetiche e bilanciate: es. `🏠 Casa Base (€)`, `🚗 Trasferta Base (€)`, `Straordinario (€)`.
    - **Prevenzione Collisioni Flex e Bottoni (`.btn`, `.setting-row`)**:
      - `.btn` ha `flex: 0 0 auto` di default. Soltanto all'interno di `.btn-row` i pulsanti condividono lo spazio equamente (`.btn-row .btn { flex: 1 1 0; min-width: 0; }`).
      - Nelle righe con testo e pulsante (es. "Stato Sincronizzazione" e bottone "Attiva/Disattiva"), utilizzare `.setting-row` con `gap: 12px;`, testo con `flex: 1 1 auto; min-width: 0;` e bottone con `flex: 0 0 auto;`. In questo modo il pulsante non può mai collidere né sovrapporsi al testo.
    - **Supporto Risoluzioni Piccole e Font Accessibilità**:
      - Testare sempre il layout su schermi stretti (fino a 320px–360px) e con font di sistema scalati a 1.2x–1.5x.
      - Utilizzare `clamp()` su testi grandi (`shift-time`, `stat-value`), `word-break: break-word` e `min-width: 0` sui container flessibili per evitare overflow orizzontali o barre di scorrimento spurie.

23. **CONFERMA OBBLIGATORIA PER QUALSIASI AZIONE DISTRUTTIVA, ANNULLAMENTO O DISCONNESSIONE**:
    - **Regola Generale**: Qualsiasi operazione che comporti perdita di dati, cancellazione, annullamento di turni o modifiche, disattivazione o disconnessione di account/servizi DEVE richiedere esplicita conferma all'utente prima di essere eseguita. Zero azioni distruttive a clic singolo.
    - **Nella Mini App (`index.html`)**:
      - Utilizzare la funzione asincrona `askConfirmation(message)`, che sfrutta la finestra di dialogo modale nativa Telegram `Telegram.WebApp.showConfirm(message, callback)` con fallback trasparente su `window.confirm`.
      - Azioni tassativamente protette:
        - *Annulla Turno Aperto* (`confirmCancelShift`): avviso perdita dati del turno odierno.
        - *Elimina Registrazione dal Foglio* (`confirmDeleteEntry`): avviso cancellazione definitiva riga con indicazione della data.
        - *Disconnetti Videorent Cloud* (`disconnectVrCloud`): avviso interruzione notifiche e compilazione automatica.
        - *Disattiva Sincronizzazione Google Calendar* (`toggleCalendarSync`): richiesta di conferma solo quando si passa da attiva a disattivata.
        - *Annulla Modifica Credenziali Cloud* (`cancelVrCloudEdit`): richiesta conferma se i campi sono stati digitati/modificati.
        - *Azioni Admin Colleghi* (`adminToggleSuspend`, `adminApproveUser`): conferma per sospensione, riattivazione o approvazione nuovo collega.
    - **Nel Bot Telegram (`google_apps_script_bot.js`)**:
      - Mai eseguire cancellazioni o disattivazioni su tocco singolo: inviare prima una card interattiva con messaggio di avviso e due pulsanti inline:
        1. Conferma esplicita (es. `🔴 Sì, Disattiva`, `🗑️ Sì, Disconnetti`, `🚫 Sì, Sospendi`, `🔄 Sì, Rigenera`).
        2. Annullamento/ritorno sicuro (es. `🔙 Annulla, Mantieni Attivo`).
      - Applicato sistematicamente a: `SETTING_TOGGLE_CALENDAR_SYNC` (disattivazione), `VR_DISCONNECT`, `ADMIN_SUSPEND_`, `ADMIN_REGENKEY_`, `ADMIN_RESETWIZARD_`, `CMD_CANCEL_DAY`, `DELETE_ENTRY_`, `ADMIN_DELETECONFIRM_`.

24. **DIVIETO DI TRAVASO DI ISTRUZIONI DI SVILUPPO, GIUSTIFICAZIONI O COMMENTI DIETRO LE QUINTE NEI TESTI RIVOLTI ALL''UTENTE E NEL MANUALE**:
    - **Separazione Netta tra Conversazione e Prodotto Finale**: Le motivazioni, le istruzioni o i requisiti espressi da Simone nella chat (es. retroscena organizzativi, spiegazioni sul perché non mostrare dati, ecc.) servono unicamente all''agente per comprendere l''architettura da implementare.
    - **Divieto Assoluto di Metacommento nei Testi User-Facing**:
      - NON inserire mai nei messaggi del bot, nella Mini App, nei banner, nei tooltip o nelle etichette frasi esplicative tipo ''per tutelare la tua privacy'', ''per motivi di riservatezza'', ''dati 100% riservati'', ''su richiesta dell''amministratore'' o simili giustificazioni.
      - L''interfaccia deve presentarsi in modo naturale, pulito, professionale e oggettivo, esattamente come farebbe una comune applicazione commerciale (es. mostrare direttamente i dati, le opzioni e le funzioni senza spiegare il perché delle scelte di design).
    - **Nel Manuale e nella Documentazione Tecnica (AGENTS.md)**:
      - Documentare le funzionalità dal punto di vista tecnico e operativo (es. isolamento dati multi-utente, sincronizzazione oraria, gestione credenziali, parametri calcolo) senza citare opinioni discorsive, retroscena della chat o aneddoti conversazionali.

25. **RIGIDA COERENZA VISIVA E DIVIETO DI MODIFICHE ARBITRARIE ALLO STILE GRAFICO (MINI APP & BOT)**:
    - **Principio di Immutabilità del Design System**: Lo stile grafico dell'applicazione, delle card e dei componenti è definitivo, collaudato e non deve essere MAI alterato con personalizzazioni ad-hoc, stili estemporanei o deviazioni creative. Ogni nuovo componente o sezione aggiunta deve essere visivamente indistinguibile da quelli esistenti.
    - **Struttura Tassativa delle Card (`.card`)**:
      - Qualsiasi raggruppamento o scheda (in Home, Riepilogo, Impostazioni, Modal, Wizard) deve utilizzare ESCLUSIVAMENTE la classe standard `<div class="card">`.
      - **Divieto Assoluto di Bordi Colorati Asimmetrici**: È tassativamente vietato aggiungere bordi evidenziati su un solo lato (es. `border-left: 4px solid var(--accent);` o simili). Il bordo deve essere sempre e solo il contorno perimetrale uniforme standard `1px solid var(--border)` con `border-radius: 12px;` e sfondo scuro `var(--card-bg)`.
    - **Tipografia e Spaziature Uniformi**:
      - **Titoli di Sezione**: Utilizzare sempre la classe standard `<div class="shift-title" style="margin-bottom:12px;">` con icona tematica seguita da testo sintetico e sobrio.
      - **Testo Descrittivo**: Sempre con classe o stile `font-size: 13px; color: var(--text-muted); margin-bottom: 14px;` e `line-height: 1.4;`. Nessun testo prolisso o impaginato in modo eterogeneo.
    - **Regole Tassative per i Pulsanti (`.btn`, `.btn-row`)**:
      - Quando due o più pulsanti sono affiancati, devono essere sempre racchiusi in `<div class="btn-row">`.
      - **Divieto di Dimensioni Asimmetriche**: Non applicare MAI proprietà `flex: 1.4`, larghezze arbitrarie in pixel o stili inline che sbilanciano i pulsanti. La regola CSS `.btn-row .btn { flex: 1 1 0; min-width: 0; }` garantisce proporzioni, altezza e allineamento perfetti e speculari.
      - **Pulsanti a Larghezza Piena**: Quando un'azione secondaria o di supporto occupa l'intera riga, utilizzare `width: 100%;` con classe `.btn-outline` o `.btn-primary` e margine standard (`margin-bottom: 10px;` o `margin-top: 8px;`).
      - **Classi Colore Ammesse**:
        - `.btn-primary`: Azione principale della card (blu/accent).
        - `.btn-outline`: Azione secondaria, download o consultazione (sfondo trasparente con bordo).
        - `.btn-success`: Azione confermativa o di invio (verde).
        - `.btn-danger`: Azione distruttiva protetta da conferma (rosso tenue).
    - **Palette Colori Ufficiale e Variabili CSS (Zero Colori Hardcoded)**:
      - Utilizzare esclusivamente le variabili CSS globali definite nel tema scuro:
        - Sfondo app: `var(--bg)` (`#0f172a`)
        - Sfondo card: `var(--card-bg)` (`#1e293b`)
        - Bordi: `var(--border)` (`#334155`)
        - Testo principale: `var(--text)` (`#f8fafc`)
        - Testo secondario: `var(--text-muted)` (`#94a3b8`)
        - Colore primario (accent): `var(--accent)` (`#0ea5e9`)
        - Verde successo: `var(--success)` (`#10b981`)
        - Giallo avviso: `var(--warning)` (`#f59e0b`)
        - Rosso pericolo: `var(--danger)` (`#ef4444`)
      - Nessun colore esadecimale o rgba() arbitrario hardcoded negli stili inline di nuove card o pulsanti.

26. **SEZIONE NATIVA EVENTI NELLA TELEGRAM MINI APP & MONITORAGGIO ANNUALE**:
    - **Nomenclatura Tassativa**: Utilizzare sempre il termine **"Eventi"** (mai "Commesse") in tutta l'interfaccia utente (tab bar, schermate, titoli).
    - **Architettura 100% Nativa (Zero Webview)**: Trasposizione completa e performante delle risorse di Videorent (`g.videorent.it`) nei componenti nativi della Mini App:
      - **Tab Bar a 4 Sezioni**: `Home` (Dashboard) | `Eventi` | `Riepilogo` | `Impostazioni`.
      - **Segmented Control Prossimi / Storico**: Eventi futuri ordinati per data crescente (il più vicino in alto); storico passato ordinato per data decrescente (il più recente in alto).
      - **Modal Dettaglio Evento**:
        - Dati generali commessa, data/orari slot, location, link rapido a Google Drive (se presente).
        - **Note di Produzione (`noteSchedaLavoro`)**: Evidenziate con card dedicata per avvisi logistici/tecnici.
        - **Materiali & Attrezzature**: Tabella suddivisa per Sala/Ambiente e Categoria con quantitativi.
        - **Note del Tecnico Interattive**: Possibilità di leggere le note della commessa, aggiungere nuove note da Mini App verso Videorent ed eliminare le proprie note.
        - **Conferma Presenza**: Tasto rapido per confermare la propria presenza su Videorent.
    - **Cache Persistente & Prefetch Automatico Scheda Odierna (Zero-Wait)**:
      - Architettura *Stale-While-Revalidate* in `localStorage`: liste ed eventi caricati a 0 ms anche offline.
      - **Prefetch Silenzioso all'Apertura (`prefetchTodayEventDetail`)**: Non appena la Mini App si avvia ed entra nella dashboard, scansiona in background il calendario di Videorent. Se rileva una commessa assegnata nella giornata odierna, ne scarica silenziosamente l'intera scheda tecnica (note di produzione, lista attrezzature preventivate e note del tecnico) memorizzandola in cache locale. Quando il tecnico tocca la commessa sul campo, si apre all'istante senza alcun tempo di attesa o richiesta di rete.
    - **Monitoraggio Annuale nel Riepilogo Mensile Telegram**:
      - In `renderMonthStats`, visualizzazione automatica del compenso cumulativo annuo in fondo al riepilogo:
        `📈 Totale anno: X €` (calcolato dinamicamente tramite `getYearTotalPay` sommando i compensi di tutte le schede mese dell'anno senza formule forfettarie).

27. **RIPARTIZIONE TARIFFE FISSE SENZA DECIMALI & RISOLUZIONE DISCREPANZE A 3 SCELTE**:
    - **Tariffe Storiche Senza Decimali (Quote Intere)**:
      - Quando il bot o lo userscript importano una commessa storica con compenso fisso complessivo ($T$) distribuito su più giornate ($N$), il totale viene ripartito a quote intere senza decimali: `base = Math.floor(T / N)` su ogni slot, e il resto esatto `T - (base * N)` viene accreditato sull'ultima giornata (es. 1000 € su 3 gg ➔ 333 €, 333 €, 334 €; 999 € su 3 gg ➔ 333 €, 333 €, 333 €).
      - Scrittura diretta in Colonna F come forfait (`isFixedRate: true`), slegata dalle ore.
    - **Risoluzione Discrepanze Gestionale vs Foglio**:
      - Se una riga sul Foglio ha già dati registrati diversi da Videorent, il gestionale NON sovrascrive mai in automatico.
      - Viene presentata all'utente una scelta esplicita a 3 opzioni:
        1. **Usa dati Foglio**: mantiene i valori del foglio e li sincronizza verso il portale.
        2. **Usa dati Gestionale**: aggiorna il foglio con orari e importo di Videorent.
        3. **Ignora e non chiedere più**: memorizza la discrepanza tra quelle ignorate (`addIgnoredDiscrepancy`) senza mai più proporla.

28. **ISOLAMENTO UX MINI APP, ZERO SPAM CHAT & LETTURA CERTIFICATA DAL FOGLIO**:
    - **Separazione Netta dei Canali UX**:
      - Le azioni operative della Mini App (`miniapp_start_shift`, `miniapp_cancel_shift`, `miniapp_end_shift`, `miniapp_save_wizard`) non inviano MAI notifiche di testo nella chat Telegram.
      - L'unica eccezione che invia messaggi in chat è l'azione richiesta esplicitamente dall'utente (es. *"Invia PDF su Telegram"*).
    - **Sincronizzazione Silenziosa Menù Chat (`updateChatHomeQuietly`)**:
      - Aggiorna il messaggio del menù Home della chat Telegram sul posto (`editMessage`) con lo stato attuale (`🟢 In Servizio` / `⚪ Standby`) e la tastiera corrispondente (`🔴 Termina Giornata`, `🗑️ Annulla Giornata` oppure `🟢 Inizia Giornata`), consentendo all'utente di aprire da Mini App e chiudere da chat (o viceversa) in perfetta continuità.
      - Se il messaggio non è modificabile o non esiste, invia il nuovo menù aggiornato con i pulsanti corretti.
    - **Riepilogo Chiusura Turno Nativo su Mini App (`modal-shift-summary`)**:
      - Alla chiusura del turno dalla Mini App si apre un modal nativo con tipologia, data, ore effettive, compenso base, extra e totale spettante, con pulsante **`OK`**.
    - **Lettura Certificata dal Foglio Google (Zero Calcoli Fittizi)**:
      - Il client non effettua calcoli stimati o approssimati. Attende che il backend scriva le celle nel Foglio Google, forzi il ricalcolo nativo con `SpreadsheetApp.flush()` e rilegga direttamente i valori effettivi generati dalle formule di Colonna F (Compenso) e Colonna G (Extra). In questo modo il modal mostra al 100% solo ciò che il Foglio di calcolo ha realmente registrato.

29. **DIVIETO ASSOLUTO E CATEGORICO DI CALCOLI STIMATI, APPROSSIMATI, PRESUNTI O INVENTATI (REGOLA ZERO-ESTIMATION)**:
    - **Principio Fondamentale e Invalicabile**: L'assistente e qualsiasi componente software (Mini App, Bot Telegram, Userscript) non devono MAI, in nessuna circostanza e per nessun motivo, calcolare tariffe, ore o importi a livello di stima, congettura, fallback o approssimazione logica.
    - **Fonte di Verità Unica ed Esclusiva**:
      - L'unica fonte di verità ammessa per gli importi e le ore lavorate è **ciò che viene effettivamente scritto e calcolato dalle formule native del Google Sheet** (o dai dati reali scaricati dal gestionale Videorent).
    - **Divieto di Fallback "di Comodo"**:
      - Se per qualsiasi motivo (errore di rete, timeout, ritardo di sincronizzazione o anomalia del backend) un dato non arriva o non è stato confermato dal foglio, è **tassativamente vietato calcolarlo al volo nel client** per "tappare il buco" o mostrare qualcosa a video.
      - In caso di assenza del dato o mancata conferma del foglio, il sistema deve restituire **esclusivamente un errore bloccante esplicito** (es. *"Errore: registrazione non confermata dal foglio di calcolo"*), permettendo all'utente di verificare e intervenire sul foglio reale, senza mai mostrare cifre presunte.
    - **Validità Permanente**: Questa regola prevale su qualsiasi logica di resilienza frontend. Il software deve essere un fedele strumento contabile di controllo, mai un generatore di previsioni.

30. **SICUREZZA, PRIVACY E ZERO-LEAK NEL REPOSITORY PUBBLICO DELLA MINI APP (`CalcoloTariffaVR-app`)**:
    - **Natura del Repository**: Il repository `SxS46/CalcoloTariffaVR-app` è **PUBBLICO** su GitHub per consentire l'hosting su GitHub Pages richiesto dalla Telegram Mini App.
    - **DIVIETO CATEGORICO DI DATI SENSIBILI**:
      È severamente e permanentemente vietato inserire, committare o pushare nel repository `CalcoloTariffaVR-app` qualsiasi informazione privata o sensibile, tra cui:
      1. *Token Telegram*: né il token del bot Telegram (`TELEGRAM_TOKEN`), né il webhook secret (`TELEGRAM_WEBHOOK_SECRET`).
      2. *Credenziali Videorent*: nessuna password, username o cookie di sessione (`sessionId`).
      3. *Fogli di Calcolo Google*: nessun ID di fogli Google (`spreadsheetId`), formule contabili o dump di dati finanziari.
      4. *Chiavi Segrete*: nessuna chiave di sincronizzazione (`syncKey` esadecimale a 64 caratteri).
      5. *Dati Personali o Contabili*: nessun dato anagrafico di Simone o dei colleghi, compensi, fatturato o tariffe private.
    - **File Tassativamente Ammessi nel Repo Pubblico**:
      - `index.html`: codice client puro HTML/CSS/JavaScript. L'autenticazione verso le API backend deve avvenire **esclusivamente** tramite la firma crittografica nativa di Telegram (`Telegram.WebApp.initData`), senza chiavi cablate nel codice.
      - `sw.js`: service worker per il caching offline e PWA.
      - `manifest.json`: manifesto PWA con icone e colori tema.
      - `Manuale_Utente_Videorent.pdf`: guida utente pulita priva di credenziali.
    - **Procedura di Verifica Pre-Push**:
      - Prima di eseguire il push su `CalcoloTariffaVR-app`, verificare sempre che i file non contengano token, chiavi segrete o riferimenti a fogli privati.
      - Quando si aggiorna la Mini App, allineare i file client in `c:\!VR\CalcoloTariffaVR-app`, eseguire il commit e pushare sia sul repo principale (`CalcoloTariffaVR`) sia su quello pubblico (`CalcoloTariffaVR-app`).

31. **SCHEDA LAVORO MINI APP: RICERCA MATERIALI, ACCORDION SALE E TOGGLE ESPANDI/COMPRIMI**:
    - **Filtro e Ricerca Dinamica Materiali (`ev-materials-search-input`)**:
      - Barra di ricerca rapida con campo testuale e pulsante di azzeramento `✕`.
      - Ricerca istantanea su nome articolo, note tecniche o nome sala.
      - Auto-espansione automatica durante la ricerca: se la query è valorizzata, le sale con articoli trovati vengono aperte automaticamente per mostrare subito i risultati.
    - **Accordion per Sala con Collasso Intelligente (`toggleMaterialsRoom`)**:
      - Quando la lista materiali supera 6 articoli complessivi e contiene più di 1 sala, la prima sala parte aperta di default e le sale successive partono collassate per mantenere compatta la visualizzazione.
      - Ciascun header sala è un elemento cliccabile con indicatore dinamico (`▶` / `▼`) e badge con il conteggio articoli della sala.
    - **Barra di Comando Globale Sotto la Ricerca (`#ev-materials-toggle-row`)**:
      - Posizionata direttamente sotto il campo di ricerca.
      - Mostra a sinistra il conteggio totale sintetico (es. *18 articoli in 3 sale*).
      - Include a destra il pulsante di commutazione rapida `▼ Espandi tutto` / `▲ Comprimi tutto` (`toggleAllMaterialsRooms`), con testo che si adatta dinamicamente allo stato di apertura delle sale.
    - **Supporto PWA Service Worker (v7)**:
      - Service worker bumpato a `vr-app-shell-v7` con `./sw.js?v=7` per garantire l'aggiornamento immediato della cache offline nei client Telegram.

32. **NAVIGAZIONE NATIVA TELEGRAM BACK BUTTON & GESTIONE STORICO (ZERO CHIUSURE INVOLONTARIE)**:
    - **Problema Risolto**: Quando l'utente era dentro una scheda lavoro commessa (`modal-event-detail`) o in una sotto pagina/schermata secondaria (Eventi, Riepilogo, Impostazioni, Admin), premendo il tasto indietro dello smartphone (o la gesture Android o la freccia in alto), Telegram chiudeva bruscamente l'intera Mini App invece di tornare alla pagina precedente.
    - **Gestione Nativa Telegram `BackButton` (`Telegram.WebApp.BackButton`)**:
      - Quando l'utente apre un lavoro, un modale o naviga in una schermata secondaria, il sistema attiva nativamente il pulsante indietro con `tg.BackButton.show()`.
      - Telegram per Android/iOS/Desktop intercetta sia la freccia in testata sia il tasto hardware/gesture back di Android, inoltrando l'evento a `handleBackNavigation` anziché terminare la webapp.
    - **Stack Modali a Livelli Multipli (`modalStack`)**:
      - Se sono aperti modali annidati (es. Scheda lavoro ➔ Rifiuta lavoro con motivo), il primo tap su indietro chiude il modale superiore (`modal-reject-work`), lasciando aperta la scheda lavoro.
      - Il secondo tap chiude la scheda lavoro e torna alla schermata chiamante.
      - Si sincronizza bidirezionalmente con le chiusure da UI (pulsante `✕`, pulsante `Annulla`, tocco backdrop).
    - **Cronologia Schermate Anti-Ciclo (`screenHistory`)**:
      - Traccia i cambi di schermata secondaria (`events`, `summary`, `settings`, `admin`).
      - Pota automaticamente i percorsi ciclici in caso di ri-tap su schede precedenti.
      - Toccando "Dashboard" o tornando alla radice, la cronologia si resetta a `['dashboard']` e `tg.BackButton.hide()` viene invocato per consentire a Telegram la naturale chiusura della WebApp solo quando ci si trova effettivamente alla Home.
    - **Supporto PWA Service Worker (v15)**:
      - Cache aggiornata a `vr-app-shell-v15` con espansione sale materiali di default e haptic feedback protetto.

### BE. Gestione Dinamica Commesse Multi-Giorno: Accorciamenti, Estensioni e Riconferma Presenza
- **Lavori Accorciati (Rimozione Singoli Slot)**:
  - Fase 4d: confronta tutti i giorni noti della commessa con gli slot in arrivo da Videorent. Se uno slot viene eliminato (`!prev.matched`), il bot elimina all'istante l'evento associato da Google Calendar tramite il suo `calendarEventId` e notifica su Telegram `🚨 Videorent - Lavoro Annullato o Rimosso!` con date e orari del giorno cancellato.
  - Finestra di ricerca calendar espansa dinamicamente: include gli intervalli di tutti gli slot precedenti presenti in `syncMap`, garantendo la pulizia immediata anche in modalità fast-path su singolo slot.
- **Lavori Modificati che Richiedono Nuova Conferma (Re-confirmation Flow)**:
  - Quando Videorent modifica le date o gli orari di una commessa già confermata, resetta la presenza del tecnico (`presenza_confirmed = 0` / `isTechnicianConfirmed = false`).
  - Il bot riconosce che si tratta dello stesso lavoro precedentemente confermato (`isReconfirm: true`):
    - Invia l'alert dedicato `⚠️ Videorent - Nuovo Lavoro da Confermare (Riconferma Modifiche)!` evidenziando sia le date precedenti sia le nuove date.
    - Mostra il pulsante interattivo `[✅ Riconferma Presenza]` anziché il generico "Conferma Presenza".
- **Eventi che Aumentano di Giorni (Estensione Commessa)**:
  - Rileva l'aggiunta di ulteriori date a una commessa già registrata (`isAddedSlot: true` su commessa con slot storici esistenti).
  - Se il nuovo slot è già confermato: crea il nuovo evento su Google Calendar e notifica `➕ Videorent - Giorno Aggiunto alla Commessa!` senza trattarlo come un nuovo lavoro duplicato.
  - Se il nuovo slot richiede conferma: invia l'alert `⚠️ Videorent - Nuovo Lavoro da Confermare (Giorno Aggiunto)!` con il pulsante `[✅ Conferma Giorno]`.
- **Test Suite Dedicata (Gruppo 50)**:
  - 839 test su 839 superati con successo in 0.42s coprendo accorciamenti, estensioni, variazioni orarie e riconferme.

### BF. Trasparenza Completa Tariffa Weekend/Festivi per Lavoro a Casa (Mini App & Bot)
- **Problema Risolto**: Dalla Mini App, aprendo o registrando una giornata a Casa nel weekend (o in un giorno festivo), non era indicato da nessuna parte che per accordi veniva applicata la tariffa Trasferta maggiorata. L'utente non aveva evidenza del compenso speciale fino alla chiusura, dove il bonus compariva solo accorpato negli extra senza spiegazione.
- **Riconoscimento Istantaneo e Trasparenza nella Mini App (`index.html`)**:
  - **Pulsante Dashboard Principale**: Quando il turno è chiuso e oggi ricade nel weekend o festivo con regola attiva (`ALL` o `SUNDAY_AND_HOLIDAYS`), il pulsante si aggiorna dinamicamente in `🏠 Inizia Casa (Tariffa WE ⭐)`.
  - **Banner di Avviso nel Modale Inizia Turno (`#start-modal-weekend-alert`)**:
    - Titolo: `🏠 Inizia Casa (Tariffa WE ⭐)`.
    - Box informativo sfumato con icona ✨: evidenzia chiaramente il giorno (es. *Sabato*, *Domenica*, *Festivo*) e precisa che verrà applicata automaticamente la tariffa Trasferta (es. *250 € anziché 220 €*).
  - **Card Turno Attivo in Servizio (`#shift-card`)**:
    - Mostra il badge nel titolo: `In Turno • 🏠 Casa [🚗 Tariffa WE]`.
    - Compare il chip informativo permanente `#shift-weekend-badge`: `✨ Sabato: Tariffa Trasferta applicata (250 € base)`.
  - **Modale Chiudi Turno (`#end-modal-weekend-alert`)**:
    - Mostra l'alert prima del salvataggio: ricorda che alla giornata verrà applicata la tariffa Trasferta con maggiorazione automatica (+30 € negli extra).
  - **Riepilogo Turno Chiuso con Scorporo Limpido (`#modal-shift-summary`)**:
    - Riga dedicata: `✨ Maggiorazione Weekend (Trasferta): +30 €`.
    - Riga Extra / Rimborsi: mostra esclusivamente gli extra effettivi inseriti dall'utente (`s.userExtra`), eliminando ogni ambiguità contabile.
  - **Modali "Aggiungi Giornata" ed "Modifica Giornata"**:
    - Banner reattivo `#manual-modal-weekend-alert` e `#edit-modal-weekend-alert` che si attiva in tempo reale cambiando data o tipologia (Casa) in presenza di giorni festivi o weekend.
- **Backend ed Ecosistema Bot Telegram (`google_apps_script_bot.js`)**:
  - `miniapp_init`: arricchito con il payload `todayHolidayInfo` (nome festività, differenza tariffaria in € e regola attiva).
  - `miniapp_start_shift` e `handleShiftOpenStart`: persistono `isWeekendHolidayBonus: true` e `holidayName` dentro l'oggetto `activeDay` dell'utente.
  - `renderHome` del bot: evidenzia `🟢 *In Servizio: 🏠 Casa (Tariffa WE ⭐)*` e una riga descrittiva `✨ _Applicata tariffa Trasferta (Sabato)_`.
- **Test Suite Dedicata (Gruppo 51)**:
  - 862 test su 862 superati con successo in 0.48s a copertura dell'intero ciclo di vita del turno weekend da Mini App e Bot.



