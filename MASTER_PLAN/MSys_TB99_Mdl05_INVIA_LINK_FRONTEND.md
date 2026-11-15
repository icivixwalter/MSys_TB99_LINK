# MSys_TB99_Mdl05_INVIA_LINK_FRONTEND

Scheda del modulo `Msys_TB99_LINK_ Mdl05_INVIA_LINK_FRONTEND` del database `MDB\MSys_TB99_LINK.mdb`. Torna al [Master Plan](MSys_TB99_00_MASTER_PLAN.html).

Il modulo prende i collegamenti raccolti nel catalogo `Msys_TB99_LINK` e li **invia ai database front-end**, sovrascrivendo quelli esistenti. Esempio: `GE_MSys_DF01_COMANDI.mdb` riceve i link alle tabelle di `MSys_DF01.mdb`. Nel front-end arrivano solo collegamenti: le tabelle originali restano nel back-end.

Versione 1.2 del 15/11/2026. Le destinazioni stanno in `Msys_TB99_LINK_PATH_DB`, legata all'**oggetto**: `ID_OBJ` → `Msys_TB99_LINK_Object` → `Name` → `Msys_TB99_LINK`. Il modulo invia a ogni destinazione attiva i link delle tabelle usate dai suoi oggetti (lo stesso elenco della query `Qry99_05_SELECT_LINK_DA_INVIARE`) e scrive l'esito su tutte le righe della coppia tabella/destinazione. I percorsi sono divisi in `Database_s` (cartella) + `Database` (file). Copia del codice: `OBJECT\MODULES\Msys_TB99_LINK_ Mdl05_INVIA_LINK_FRONTEND.bas`.

## 1. Funzioni pubbliche

| Funzione | Parametri | Restituisce |
| --- | --- | --- |
| `POPOLA_LINK_FRONTEND_pf` | `bSilenzioso` | numero di assegnazioni inserite (-1 se errore) |
| `VERIFICA_LINK_FRONTEND_pf` | `sFrontEnd`, `bSilenzioso` | riepilogo `TOT=..;SIMULA_SOVRASCRITTO=..;...` |
| `INVIA_LINK_FRONTEND_pf` | `sFrontEnd`, `bSimulazione`, `bSilenzioso` | riepilogo `TOT=..;CREATO=..;SOVRASCRITTO=..;...` |

- `sFrontEnd`: vuoto per tutti i front-end attivi, oppure il percorso completo di un solo front-end.
- `bSimulazione = True`: apre i front-end in sola lettura e scrive gli esiti `SIMULA_*`, senza modificare niente.
- `bSilenzioso = True`: nessun `MsgBox`, per l'automazione e i test.

### 1.1 Esempi dalla finestra Immediata

```
? POPOLA_LINK_FRONTEND_pf()
? VERIFICA_LINK_FRONTEND_pf()
? INVIA_LINK_FRONTEND_pf()
? INVIA_LINK_FRONTEND_pf("C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_DF01\MDB\GE_MSys_DF01_COMANDI.mdb")
? INVIA_LINK_FRONTEND_pf("", True)        ' simulazione su tutti
```

## 2. Regole

| Regola | Contenuto |
| --- | --- |
| R1 | Nel front-end si creano **solo collegamenti**: le tabelle originali non si copiano mai. |
| R2 | Un link già presente nel front-end viene **cancellato e ricreato** (sovrascrittura). |
| R3 | Una **tabella locale** con lo stesso nome nel front-end **non viene mai toccata**: esito `TABELLA_LOCALE`. |
| R4 | Si inviano solo link del catalogo con `Verifica_s = OK` e database dentro `PROGETTI_MDB\MSYS\`. |
| R5 | Il front-end deve stare dentro `PROGETTI_MDB\MSYS\` e non può essere il database sorgente della tabella né `MSys_TB99_LINK.mdb`. |

## 3. Esiti

| Esito | Significato | Cosa fare |
| --- | --- | --- |
| `CREATO` | link nuovo creato nel front-end | niente |
| `SOVRASCRITTO` | link esistente cancellato e ricreato; in `Nota_s` il percorso precedente | niente |
| `SIMULA_CREATO`, `SIMULA_SOVRASCRITTO` | come sopra, ma in simulazione | lanciare l'invio reale |
| `TABELLA_LOCALE` | nel front-end c'è una tabella vera con quel nome | spostare i dati nel back-end, cancellare la tabella locale, inviare di nuovo |
| `SORGENTE` | il front-end è il database sorgente della tabella | togliere la riga di assegnazione |
| `NON_A_CATALOGO` | il nome non esiste in `Msys_TB99_LINK` | censire la tabella o togliere la riga |
| `CATALOGO_NON_OK` | link del catalogo non `OK` o fuori da `MSYS\` | ricostruire il catalogo (`RICOSTRUICI_LINK_TABELLE_pf`) |
| `SORGENTE_MANCANTE` | database sorgente o tabella sorgente non trovati | controllare `Database` nel catalogo |
| `FE_MANCANTE` | file front-end non trovato | correggere `Database_s` e `Database` della destinazione |
| `FE_NON_VALIDO` | front-end fuori da `MSYS\` o uguale a `MSys_TB99_LINK.mdb` | correggere `Database_s` e `Database` della destinazione |
| `FE_NON_APRIBILE` | front-end bloccato o danneggiato | chiudere il front-end e riprovare |
| `ERRORE` | errore durante la creazione del link; dettaglio in `Nota_s` e nel log | leggere il log |

## 4. Struttura del codice

Ogni punto ha la stessa etichetta nell'indice in testata e nel codice: si trova con Ctrl+F.

| Etichetta | Contenuto |
| --- | --- |
| `#01#` | costanti e variabili di modulo (`RADICE_MSYS`, nomi tabelle, log) |
| `#02#` | `INVIA_LINK_FRONTEND_pf` |
| `#02.01#` | preparazione: log, controllo tabelle, conferma utente |
| `#02.02#` | apertura elenco assegnazioni attive |
| `#02.03#` | ciclo sulle assegnazioni |
| `#02.03.01#` | cambio front-end: chiusura precedente, controlli, apertura |
| `#02.03.02#` | controlli sul link del catalogo |
| `#02.03.03#` | invio del link |
| `#02.03.04#` | registrazione esito |
| `#02.04#` | chiusura, statistiche, messaggio finale |
| `#02.99#` | gestione errori |
| `#03#` | `VERIFICA_LINK_FRONTEND_pf` |
| `#04#` | `POPOLA_LINK_FRONTEND_pf` |
| `#04.01#` | preparazione: log, nomi del catalogo |
| `#04.02#` | scansione dei database in `RADICE_MSYS` |
| `#04.03#` | lettura dei link di un database e inserimento assegnazioni |
| `#04.04#` | chiusura e messaggio finale |
| `#04.99#` | gestione errori |
| `#05#` | `InviaUnLink`: invio di un singolo link |
| `#05.01#` | ricerca della tabella nel front-end |
| `#05.02#` | tabella locale: stop (R3) |
| `#05.03#` | cancellazione del link esistente (R2) |
| `#05.04#` | creazione del nuovo link |
| `#06#` | controlli su file e database |
| `#06.01#` | `TabellaInSorgente` (con cache per database) |
| `#06.02#` | `PercorsoInRadice` |
| `#06.03#` | `PercorsoEscluso` (salta `AA_SALVATAGGI`, `OLD`, `BACKUP*`, `EXE`, `.git`) |
| `#06.04#` | `TabellaEsiste` |
| `#07#` | log: `#07.01#` `ApriLog`, `#07.02#` `ScriviLog`, `#07.03#` `ChiudiLog` |
| `#08#` | utilità: `#08.01#` `SqlTesto`, `#08.02#` `AggiornaEsito`, `#08.03#` `ElencaDatabase` |

## 5. Log

- File: `MDB\LOGS\AAAAMMGG_HHMMSS_TB99_Invio_Link_FrontEnd_Log.txt`, creato nella cartella del database.
- Una riga per evento, uguale nel file e nella finestra Immediata:

```
2026-11-15 09:04:48 | Msys_TB99_LINK_ Mdl05_INVIA_LINK_FRONTEND | Pos: 02.03.04 | OK | MSys_DF04_utile -> CREATO
```

- Esiti di riga: `INIZIO`, `INFO`, `OK`, `AVVISO`, `ERRORE`, `FINE`.
- In caso di errore il log scrive `codice fermo a #SS.PP.ss#` con numero e descrizione dell'errore.

## 6. Come censire le assegnazioni

`POPOLA_LINK_FRONTEND_pf` legge tutti i `.mdb` e `.accdb` di `PROGETTI_MDB\MSYS\`. Per ogni link già presente il cui nome è nel catalogo, aggancia la destinazione agli **oggetti** della tabella: inserisce in `Msys_TB99_LINK_PATH_DB` una riga (`ID_OBJ`, `Database_s`, `Database`) per ogni oggetto scelto con questa precedenza:

1. oggetti della tabella che si trovano già in quel database;
2. oggetti della tabella nei front-end `GE_*`;
3. tutti gli oggetti della tabella.

Altre regole:

- per i file `GE_*` e `MSYS_ATTIVA_GEST*` la riga è **attiva** (`Scel_b = True`); per gli altri è **disattivata**, da attivare a mano;
- una tabella senza oggetti, per esempio un separatore, non riceve destinazioni (avviso nel log);
- una riga già presente (stesso oggetto e stesso database) non viene duplicata.

Per portare in un database un oggetto che non ci è mai stato, si inserisce a mano una riga con `ID_OBJ` dell'oggetto, `Database_s` (cartella), `Database` (file) e `Scel_b = True`: il link della sua tabella verrà inviato con `INVIA_LINK_FRONTEND_pf`.

## 7. Test del 15/11/2026

Eseguiti su copie dei front-end in una cartella di prova dentro `MSYS\`, poi cancellata insieme alle righe di prova.

| Caso | Esito ottenuto |
| --- | --- |
| `MSys_DF01_COMANDI` in copia di `GE_MSys_DF01_COMANDI.mdb`, link verso `MSYS_OGGETTI` | `SOVRASCRITTO`, ora `MSYS\MSys_DF01\MDB\MSys_DF01.mdb`, 31 righe leggibili |
| `MSys_DF04_utile`, assente nel front-end | `CREATO`, 10 righe leggibili |
| `MSys_DF03_QUOTE` trasformata in tabella locale con 1 riga | `TABELLA_LOCALE`, tabella e valore `ORIGINALE` intatti |
| `NOME_INESISTENTE_XYZ` | `NON_A_CATALOGO` |
| front-end `GE_NON_ESISTE.mdb` | `FE_MANCANTE` |
| `MSys_DF04_utile` verso `MSys_DF04.mdb` (il suo sorgente), in simulazione | `SORGENTE` |
| secondo invio sulla stessa copia | 3 `SOVRASCRITTO`, 1 `NON_A_CATALOGO`, nessun errore |
| `POPOLA_LINK_FRONTEND_pf` due volte | 70 righe, poi 0 |
| `VERIFICA_LINK_FRONTEND_pf` su tutti i front-end reali | 57 `SIMULA_SOVRASCRITTO`, nessuna modifica |
