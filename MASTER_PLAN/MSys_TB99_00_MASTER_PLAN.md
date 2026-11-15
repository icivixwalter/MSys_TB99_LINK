# MSys_TB99 – Master Plan dei collegamenti

Mini-progetto **TB99_LINK** del super-progetto MSys. Il database `MDB\MSys_TB99_LINK.mdb` (codice progetto `MSys_TB99_LINK`, codice `@SHADOW_SYSTEM_CATALOG`) **concentra tutte le tabelle collegate del progetto**.

Per ogni tabella collegata sa tre cose:

1. **dove si trova la tabella fisica** (`Msys_TB99_LINK`);
2. **quali oggetti funzionano con la tabella** nel suo mini-progetto originale: query, maschere, moduli, macro, report che la usano e che si trovano **nello stesso database della tabella fisica** (`Msys_TB99_LINK_Object`);
3. **in quali database va esportato ogni oggetto**, e quindi quale collegamento serve in quel database (`Msys_TB99_LINK_PATH_DB`).

Le tre tabelle sono una **catena**: `Msys_TB99_LINK` → `Msys_TB99_LINK_Object` (campo `Name`) → `Msys_TB99_LINK_PATH_DB` (campo `ID_OBJ`).

**Perché esiste.** Serve a spostare gli oggetti da un mini-progetto a un altro con libertà e precisione, portando con sé anche il collegamento alla tabella, così l'oggetto funziona subito. Esempio: nel progetto master `MSYS_ATTIVA_GEST_OGGETTI.mdb` servono i menu. Il catalogo dice:

- dove sta la tabella `MSys_TB06_MENU_VOCI`;
- quali query, maschere e moduli del front-end `GE_MSys_TB06_MENU.mdb` la usano;
- che vanno esportati, insieme al link, in `MSYS_ATTIVA_GEST_OGGETTI.mdb`.

**Regole**
- Il catalogo contiene **solo** tabelle di database in `C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\`.
- Nei front-end arrivano **solo collegamenti**, mai le tabelle originali.

Aggiornato il 15/11/2026, dopo la ricostruzione degli oggetti nel database sorgente (capitolo 1.7). Contenuto:

| Tabella | Righe |
| --- | --- |
| `Msys_TB99_LINK` | 54 tabelle collegate in 17 database, tutte OK |
| `Msys_TB99_LINK_Object` | 137 oggetti, su 19 tabelle, ognuno nel database della sua tabella (12 database) |
| `Msys_TB99_LINK_PATH_DB` | 316 destinazioni oggetto → database, su 17 database (139 attive) |

## 1. Storico del 15/11/2026

### 1.1 Riallineamento dei percorsi

- I 43 link originali puntavano alla vecchia struttura `MSYS_OGGETTI`: dove la tabella esiste in `MSYS\`, il link è passato lì.
- 21 tabelle nuove sono state censite.
- La query `Qry99_70_INSERT` è stata corretta (puntava a una tabella inesistente).
- Il modulo `Mdl01` è stato corretto:
  - i log ora vanno in `MDB\LOGS\`;
  - il codice del mini-progetto non viene più sovrascritto.
- Backup: `MDB\BACKUP_CLAUDE_2026_11_15_08_42_28_RELINK\`.

### 1.2 Solo il progetto riscritto

Tolti 22 collegamenti verso `MSYS_OGGETTI` (DF10, DF14, TB02, TB03, TB04, TB05, TB40 PROG_MASTER, TB41). Le tabelle di origine non sono state toccate. Backup: `MDB\BACKUP_CLAUDE_2026_11_15_08_54_13_SOLO_MSYS\`.

### 1.3 Invio dei collegamenti ai front-end

Creato il modulo `Mdl05_INVIA_LINK_FRONTEND`, provato su copie in 9 casi, tutti superati (scheda [MSys_TB99_Mdl05_INVIA_LINK_FRONTEND](MSys_TB99_Mdl05_INVIA_LINK_FRONTEND.html)). **L'invio reale ai front-end non è ancora stato eseguito.** Backup: `MDB\BACKUP_CLAUDE_2026_11_15_09_01_20_MDL05\`.

### 1.4 Menu rinumerato TB06, TB07, TB08

Il mini-progetto menu (`MSYS\MSys_TB06_MENU\`) è passato da TB01/02/03 a TB06/07/08. Le sue 11 tabelle sono state aggiunte al catalogo. Dettagli nel master plan di TB06.

### 1.5 Ristrutturazione delle tre tabelle

| Tabella | Prima | Dopo |
| --- | --- | --- |
| `Msys_TB99_LINK` | 29 campi; `Database` = percorso completo | 13 campi; `Database` = solo file, `Database_s` = solo cartella, `Connect` = percorso completo; `Name` univoco |
| `Msys_TB99_LINK_Object` | 27 campi; 508 oggetti senza legame con la tabella che usano | 11 campi; 612 righe, ognuna lega un oggetto alla tabella che usa |
| `Msys_TB99_LINK_PATH_DB` | 15 campi; elenco dei database sorgente | 10 campi; destinazioni di ogni tabella collegata |
| `Msys_TB99_LINK_FRONTEND` | 70 destinazioni per il modulo `Mdl05` | **tolta**: le righe sono confluite in `PATH_DB` |

- **Relazioni:** create con integrità referenziale e aggiornamento a catena sul campo `Name`, da `Msys_TB99_LINK` a `Object` e a `PATH_DB`.
- **Query:** riscritte tutte, e aggiunte `Qry99_02`, `Qry99_03`, `Qry99_04` (capitolo 4).
- **Moduli adattati e compilati:**
  - `Mdl01` legge il percorso da `DB_COMPLETO`;
  - `Mdl04` scrive nei nuovi campi e non cancella più gli oggetti legati a una tabella;
  - `Mdl05` legge e scrive `PATH_DB`; verificato, 57 link come prima.
- **Backup:**
  - le vecchie 508 righe degli oggetti sono in `Msys_TB99_LINK_Object_OLD_2026_11_15`;
  - il database prima della modifica è in `MDB\BACKUP_CLAUDE_2026_11_15_21_03_23_RISTRUTTURA\`.

### 1.6 Catena LINK → Object → PATH_DB e campi ridondanti tolti

- **Legame delle destinazioni:** `Msys_TB99_LINK_PATH_DB` non è più legata alla tabella (`Name`) ma all'**oggetto** (`ID_OBJ`).
- **Aggancio delle 74 destinazioni** precedenti, una per tabella, agli oggetti della tabella, con questa precedenza:
  1. oggetti già presenti nel database di destinazione;
  2. oggetti dei front-end `GE_*`;
  3. tutti gli oggetti della tabella.

  Sono risultate 331 righe. Le 31 destinazioni di tabelle **separatore** (`..._}----@`, `..._{@@@...}`) non hanno oggetti e sono state tolte.
- **`POPOLA_LINK_FRONTEND_pf`:** ha poi aggiunto 117 destinazioni, cioè il front-end del menu `GE_MSys_TB06_MENU.mdb` (36, attive) e i back-end con link verso altri back-end (81, disattivate).
- **Campi tolti perché ridondanti:**
  - da `Object`: `Type` (uguale a `TIPO_OGGETTO_s`) e `COD_SUB_PROGETTO_s` (si ricava dalla cartella);
  - da `PATH_DB`: `Name` (si ricava dall'oggetto) e `TIPO_DB_s` (si ricava dal nome del file: `GE_*` = front-end).
- **Codice e strumenti adattati:**
  - `Mdl05`: elenca i link da inviare risalendo da `PATH_DB` alla tabella attraverso l'oggetto, e scrive l'esito su tutte le righe della coppia tabella/destinazione;
  - `Mdl04`: non scrive più i campi tolti;
  - `RIALLINEA` e `RICOSTRUISCI_..._3_CARICA`: quando ricarica gli oggetti, salva e riaggancia le destinazioni.
- **Prove:** compilazione riuscita; `VERIFICA` trova 36 link (i 32 di prima più le 4 tabelle del menu per il suo front-end); un secondo `POPOLA` aggiunge 0 righe.
- **Backup:** `MDB\BACKUP_CLAUDE_2026_11_15_21_43_28_CATENA\`.

### 1.7 Oggetti solo nel database della tabella (ricostruzione)

La regola precedente era sbagliata: collegava a una tabella tutti gli oggetti di tutti i database che la usano, compresi front-end e copie. La regola corretta è questa.

> **Gli oggetti di una tabella sono quelli che la usano e che si trovano nello stesso database della tabella fisica.**
> Esempio: `MSys_DF01_COMANDI` sta in `MSys_DF01.mdb`. I suoi oggetti sono le query `MSys_DF01_COMANDI_Qry01_01_SELECT_TUTTI` e `Msys_DF01_COMANDI_Qry01_02_GROUP_TIPO`, le maschere `Msys_DF01_Frm01_M01_GE_COMANDI` e `Msys_DF01_Frm01_S01_COMANDI`, e i moduli `Msys_DF01_Mdl01_COLLEGA_COMANDI` e `CLASSE_GESTIONE_OGGETTI_DATABASE`, tutti in `MSys_DF01.mdb`.

- **Analisi:** `RICOSTRUISCI_TB99_OGGETTI_2_ANALIZZA.py` ora tiene solo gli oggetti che stanno nel database sorgente della tabella. Il catalogo che riceve contiene nome della tabella e database sorgente.
- **Oggetti ricostruiti:** 137 (prima 612), su 19 tabelle. Il controllo "oggetti fuori dal database della tabella" dà 0.
- **Destinazioni ricostruite:** le 62 coppie tabella → database già presenti sono state ricreate per tutti gli oggetti della tabella, per un totale di 316 righe. È stata esclusa la destinazione uguale al database in cui l'oggetto si trova già.
- **`POPOLA_LINK_FRONTEND_pf`:** aggancia ogni nuova destinazione a tutti gli oggetti della tabella. Il primo caricamento è stato annullato perché `MSys_TB99_LINK.mdb` era aperto in Access; il modulo è stato caricato alle 22:39, a database chiuso. Un secondo `POPOLA` aggiunge 0 righe e `VERIFICA` trova 23 link.
- **Tabelle senza oggetti nel proprio database:** non ricevono destinazioni. Sono `MSys_DF01_COMANDI_TMP`, `MSYS_DF03_GROUP_QUOTE`, `MSys_DF04_utile`, `Msys_DF12_ObjProgetti`, le 4 tabelle di `MSys_TB40.mdb`, `MSys_TB07_MENU_FUNZIONI_NW`, `MSys_TB08_MENU_OGGETTI`, `MSys_TAB40_FORM_MASTER*` e `Msys_ESPORTA_ELENCO_OGGETTI_DB_TMP`. `MSys_TB40.mdb` e `MSys_DF12.mdb` non contengono nessun oggetto: le loro query e maschere stanno nei front-end.
- **Backup:** `MDB\BACKUP_CLAUDE_2026_11_15_22_14_41_OGGETTI_SORGENTE\`.

## 2. Struttura delle tre tabelle

```
  Msys_TB99_LINK                    1 riga = 1 tabella collegata e dove si trova
  PK ID_LINK    UNIVOCO Name
        |
        | Name   1 : N
        v
  Msys_TB99_LINK_Object             1 riga = 1 oggetto che usa la tabella e sta nel SUO database
  PK ID_OBJ     FK Name
        |
        | ID_OBJ 1 : N
        v
  Msys_TB99_LINK_PATH_DB            1 riga = 1 database in cui esportare l'oggetto
  PK ID_PATH    FK ID_OBJ

  Da una destinazione si risale all'oggetto (ID_OBJ) e alla tabella (Name):
  - oggetti da esportare in quel database  = query Qry99_04
  - link da creare in quel database        = query Qry99_05 (tabelle degli oggetti, senza doppioni)

  Relazioni con integrita referenziale; aggiornamento a catena su Name; nessuna cancellazione
  a catena: per togliere una tabella si tolgono prima le destinazioni, poi gli oggetti.
```

### 2.1 Msys_TB99_LINK – la tabella collegata e dove si trova

| Campo | Tipo | Significato | Esempio |
| --- | --- | --- | --- |
| `ID_LINK` | Contatore | chiave primaria | 12 |
| `Name` | Testo 255 | **nome della tabella collegata**, univoco; uguale nel database sorgente | `MSys_DF01_COMANDI` |
| `Type` | Intero | 6 = tabella collegata | 6 |
| `Connect` | Testo 255 | stringa di collegamento completa | `;DATABASE=C:\...\MSYS\MSys_DF01\MDB\MSys_DF01.mdb` |
| `Database` | Testo 255 | **solo il file** del database sorgente | `MSys_DF01.mdb` |
| `Database_s` | Testo 255 | **solo la cartella**, con `\` finale | `C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_DF01\MDB\` |
| `TIPO_OGGETTO_s` | Testo 20 | `TABLES_LINK`, per distinguerla da una tabella vera (`TABLE`) | `TABLES_LINK` |
| `COD_PROGETTO_s` | Testo 45 | progetto che contiene l'informazione | `MSys_TB99_LINK` |
| `COD_SUB_PROGETTO_s` | Testo 45 | mini-progetto proprietario della tabella | `MSys_DF01` |
| `Verifica_s` | Testo 50 | `OK`, `MISSING`, `INVALID_PATH`, `NOT_REGISTERED`, `NOT_CHECKED` | `OK` |
| `Scel_b` | Sì/No | riga scelta | |
| `DATAINS_d`, `DATAAGG_d` | Data | inserimento e ultimo aggiornamento | |

Percorso completo = `Database_s & Database`: le query lo espongono come `DB_COMPLETO`.

### 2.2 Msys_TB99_LINK_Object – gli oggetti che lavorano con la tabella

| Campo | Tipo | Significato | Esempio |
| --- | --- | --- | --- |
| `ID_OBJ` | Contatore | chiave primaria; la usa `PATH_DB` | 1234 |
| `Name` | Testo 255 | **tabella collegata usata** (legame con `Msys_TB99_LINK`) | `MSys_TB06_MENU_VOCI` |
| `NOME_OGGETTO_s` | Testo 255 | nome dell'oggetto | `MSys_TB06_Qry01_01_VOCI` |
| `TIPO_OGGETTO_s` | Testo 20 | `QUERY`, `FORMS`, `REPORTS`, `MACRO`, `MODULES` (codici di DF02) | `QUERY` |
| `Database` | Testo 255 | file del database in cui si trova l'oggetto: è sempre il database della tabella | `MSys_DF01.mdb` |
| `Database_s` | Testo 255 | cartella di quel database (il mini-progetto è la prima cartella sotto `MSYS\`) | `C:\...\MSYS\MSys_DF01\MDB\` |
| `Scel_b`, `DATAINS_d`, `DATAAGG_d` | | servizio | |

Un oggetto che usa più tabelle dello stesso database compare una volta per ogni tabella. Le righe con `Name` vuoto sono l'elenco grezzo degli oggetti scritto da `Mdl04`.

### 2.3 Msys_TB99_LINK_PATH_DB – i database in cui esportare l'oggetto

| Campo | Tipo | Significato | Esempio |
| --- | --- | --- | --- |
| `ID_PATH` | Contatore | chiave primaria | |
| `ID_OBJ` | Lungo | **oggetto** da portare (legame con `Msys_TB99_LINK_Object`) | 1234 |
| `Database` | Testo 255 | file del database di destinazione | `MSYS_ATTIVA_GEST_OGGETTI.mdb` |
| `Database_s` | Testo 255 | cartella del database di destinazione | `C:\...\MSYS\MSYS_ATTIVA_GEST\MDB\` |
| `Scel_b` | Sì/No | `True` = destinazione attiva | |
| `Esito_s`, `Nota_s` | Testo | esito dell'ultimo invio o verifica del link, e dettaglio | `SIMULA_SOVRASCRITTO` |
| `DATAINS_d`, `DATA_INVIO_d` | Data | inserimento e ultimo invio | |

## 3. Come si usa il catalogo per spostare oggetti

Esempio: portare i menu nel progetto master `MSYS_ATTIVA_GEST_OGGETTI.mdb`.

1. **Destinazioni.** In `Msys_TB99_LINK_PATH_DB` ci sono 22 righe verso `MSYS_ATTIVA_GEST_OGGETTI.mdb`, una per ogni oggetto del menu che sta nel database delle sue tabelle (`MSys_TB06_MENU_VOCI.mdb`, `MSys_TB07_MENU_FUNZIONI.mdb`). Sono **disattivate** (`Scel_b = False`, esito `DA_ATTIVARE`), perché il progetto master usa ancora i vecchi nomi `GEST_MENU_Tb01_...`.
2. **Cosa portare.** Con le righe attivate:
   - la query `Msys_TB99_LINK_Qry99_04_SELECT_DA_ESPORTARE` elenca gli oggetti da esportare e il database da cui prenderli: per il menu, le query `MSys_TB06_Qry…` e `MSys_TB07_Qry…`, le maschere `MSys_TB06_Frm01_*` e `MSys_TB07_Frm01_*` e la macro `COLLEGA_Mcr_MSys_TB06_MENU_VOCI`, tutte nei database delle tabelle;
   - la query `Msys_TB99_LINK_Qry99_05_SELECT_LINK_DA_INVIARE` elenca i link da creare.
3. **Esportare gli oggetti** dal front-end del menu al progetto master. Oggi si fa con le maschere di importazione ed esportazione (`UTILITA_Msys_Frm22`, `Frm25`). Un modulo che legga la query 04 ed esporti da solo è nel piano (capitolo 7, Fase 2).
4. **Inviare i link.** Attivare le righe (`Scel_b = True`) e lanciare `INVIA_LINK_FRONTEND_pf("C:\...\MSYS_ATTIVA_GEST\MDB\MSYS_ATTIVA_GEST_OGGETTI.mdb")`. Il link viene creato o sovrascritto; una tabella locale con lo stesso nome non viene toccata.

## 4. Query

| Query | Tipo | Uso |
| --- | --- | --- |
| `Msys_TB99_LINK_Qry99_01_SELECT_Archivio` | selezione (aggiornabile) | catalogo con `DB_COMPLETO`: lo legge `Mdl01` per ricostruire i link |
| `Msys_TB99_LINK_Qry99_01_SELECT_BackUp` | selezione | tabelle senza i separatori, per il backup (`Mdl03`) |
| `Msys_TB99_LINK_Qry99_01_SELECT_LUNGH_MaxMin` | totali | lunghezza minima, massima e media dei percorsi |
| `Msys_TB99_LINK_Qry99_02_SELECT_OGGETTI` | selezione | **tabella → oggetti che la usano** |
| `Msys_TB99_LINK_Qry99_03_SELECT_PATH_DB` | selezione | **catena completa**: destinazione, oggetto, tabella, esito dell'ultimo invio |
| `Msys_TB99_LINK_Qry99_04_SELECT_DA_ESPORTARE` | raggruppamento | **oggetti da esportare** in ogni destinazione attiva, senza doppioni |
| `Msys_TB99_LINK_Qry99_05_SELECT_LINK_DA_INVIARE` | raggruppamento | **link da creare** in ogni destinazione attiva (le tabelle dei suoi oggetti); è l'elenco che usa `Mdl05` |
| `Msys_TB99_LINK_Qry99_10_GROUP_COD_SUB_PROGETTO_s` | raggruppamento | un database sorgente per riga, con numero di tabelle (lo legge `Mdl04`) |
| `Msys_TB99_LINK_PATH_DB_Qry99_01_SELECT_DB*` | selezione | destinazioni attive con tabella e oggetto |
| `Msys_TB99_LINK_Qry99_70_INSERT_ArchivioTabelleLINK` | accodamento | aggiunge al catalogo le tabelle collegate di questo database non ancora censite, già con file e cartella divisi |
| `Msys_TB99_LINK_Qry99_80_DELETE_Archivio` | eliminazione | svuota il catalogo; si blocca se ci sono oggetti o destinazioni legati (integrità) |

## 5. Moduli

| Modulo | Funzioni pubbliche | Cosa fa |
| --- | --- | --- |
| `Msys_TB99_LINK_ Mdl01_RICOSTRUICI_LINK_TABELLE` | `RICOSTRUICI_LINK_TABELLE_pf`, `CancellaTuttiLinkEsistenti` | ricostruisce i link dentro `MSys_TB99_LINK.mdb` dal catalogo |
| `Msys_TB99_LINK_ Mdl03_ESPORTA_DATI_IN_BACKUP` | | copia in backup le tabelle del catalogo |
| `Msys_TB99_LINK_ Mdl04_IMPORTA_NOMI_OGGETTI` | `ImportaOggettiDaDatabaseEsterno` | elenca tutti gli oggetti dei database sorgente **senza** legarli a una tabella (`Name` vuoto); non cancella più le righe legate |
| `Msys_TB99_LINK_ Mdl05_INVIA_LINK_FRONTEND` | `POPOLA_…`, `VERIFICA_…`, `INVIA_LINK_FRONTEND_pf` | invia a ogni destinazione attiva i link delle tabelle usate dai suoi oggetti (scheda dedicata) |

### 5.1 Maschere del catalogo

| Maschera | Origine | Contenuto |
| --- | --- | --- |
| `Msys_TB99_LINK_FrmM02_GE_LINK` | – (non associata) | **maschera principale**, senza barre di scorrimento, circa 1.464 × 710 pixel (17.686 twip di larghezza). L'**intestazione fissa** è su una riga: titolo e pulsanti-linguetta **"1 - COLLEGAMENTI"** e **"2 - IMPORTA OGGETTI"** (`Cmd_Pg01`, `Cmd_Pg02`; quello attivo è in grassetto). Nel corpo c'è la struttura a schede `TabCtl_01` con 12 pagine, senza linguette proprie, su misura della UTILITA ridotta; le pagine da `Pg_03` a `Pg_12` sono nascoste |
| `UTILITA_Msys_Frm21_IMPORTA_OGGETTI` | – | in `Pg_02` (`SottoForm_IMPORTA`), a misura esatta e senza bordo. **Ridotta all'80%** in posizioni, misure e caratteri (21.770 × 9.807 → 17.416 × 7.845 twip; caratteri 11→9, 9→7, 8→7 pt), senza barre di scorrimento, selettori e pulsanti di spostamento. La riduzione vale solo per la copia dentro `MSys_TB99_LINK.mdb` |
| `Msys_TB99_LINK_FrmS05_LINK` | `Msys_TB99_LINK` (solo tabelle con oggetti) | foglio dati delle **tabelle collegate**: mini-progetto, nome, file, cartella, verifica, tipo, `Connect`, N. oggetti |
| `Msys_TB99_LINK_FrmS03_OGGETTI` | `Msys_TB99_LINK_Object` | foglio dati degli **oggetti della tabella selezionata** (nel suo database) |
| `Msys_TB99_LINK_FrmS04_PATH_DB` | `Msys_TB99_LINK_PATH_DB` | foglio dati delle **destinazioni dell'oggetto selezionato** |

```
  Msys_TB99_LINK_FrmM02_GE_LINK
  TB99 - TABELLE COLLEGATE, OGGETTI E DESTINAZIONI
  +-[ COLLEGAMENTI ]------------------------------------------------------------+
  | 1 - TABELLE COLLEGATE                 SottoForm_LINK    (FrmS05)            |
  |   MSys_DF01  MSys_DF01_COMANDI  MSys_DF01.mdb  C:\...\MSys_DF01\MDB\  OK  6 |
  |   MSys_DF02  MSys_DF02_TIPO_OGGETTO  ...                                    |
  | 2 - OGGETTI  SottoForm_OGGETTI (FrmS03)  | 3 - DESTINAZIONI  SottoForm_PATH_DB |
  |   FORMS  Msys_DF01_Frm01_M01_GE_...      |   (FrmS04)                         |
  |   QUERY  MSys_DF01_COMANDI_Qry01_01_...  |   MSYS_ATTIVA_GEST_OGGETTI.mdb     |
  |   ...                                    |   GE_MSys_DF01_COMANDI.mdb         |
  +-----------------------------------------------------------------------------+
  +-[ IMPORTA OGGETTI ]  Pg_02: UTILITA_Msys_Frm21_IMPORTA_OGGETTI ridotta all'80%, a misura esatta, senza barre
  Pg_03 ... Pg_12 nascoste
```

- **Catena:** le caselle nascoste `TXT_NAME` (`=[SottoForm_LINK].[Form]![Name_TXT]`) e `TXT_ID_OBJ` (`=[SottoForm_OGGETTI].[Form]![ID_OBJ_TXT]`) collegano le sottomaschere: `SottoForm_OGGETTI` su `TXT_NAME` → `Name`, `SottoForm_PATH_DB` su `TXT_ID_OBJ` → `ID_OBJ`. L'evento `Form_Current` di S05 e S03 aggiorna le sottomaschere che seguono.
- **Nuove righe:** un oggetto aggiunto nella 2 prende da solo la tabella selezionata; una destinazione aggiunta nella 3 prende da sola l'oggetto selezionato.
- **Costruzione:** le maschere sono state costruite il 15/11/2026 direttamente in Access, perché su questo PC non si possono importare da testo. Il loro testo esportato è in `OBJECT\FORMS\`.
- **Prova:** all'apertura, la 1 è su `MSys_DF01_COMANDI` (6 oggetti); la 2 mostra `Msys_DF01_Frm01_M01_GE_COMANDI` (`ID_OBJ` 1133); la 3 mostra `MSYS_ATTIVA_GEST_OGGETTI.mdb`.
- **Backup:** `MDB\BACKUP_CLAUDE_2026_11_15_22_51_39_MASCHERA_TAB\` e `MDB\BACKUP_CLAUDE_2026_11_15_23_00_18_MASCHERA_PAG2\`, `MDB\BACKUP_CLAUDE_2026_11_15_23_05_02_MASCHERA_LINGUETTE\`, `MDB\BACKUP_CLAUDE_2026_11_15_23_24_06_RIDUZIONE_80\`.

## 6. Strumenti (MASTER_PLAN\STRUMENTI\)

| File | Cosa fa |
| --- | --- |
| `RIALLINEA_TB99_LINK.ps1` | versione 2.0: verifica ogni tabella (`Verifica_s`), toglie i record fuori da `MSYS\` (prima i loro oggetti e destinazioni), ricostruisce i link fisici; non tocca `PATH_DB` |
| `RICOSTRUISCI_TB99_OGGETTI_1_ESPORTA.ps1` | esporta in testo query, maschere, report, macro e moduli di tutti i database di `MSYS\`, lavorando su copie |
| `RICOSTRUISCI_TB99_OGGETTI_2_ANALIZZA.py` | trova quali tabelle del catalogo usa ogni oggetto, direttamente o attraverso query e sottomaschere, e **tiene solo gli oggetti che stanno nel database della tabella** |
| `RICOSTRUISCI_TB99_OGGETTI_3_CARICA.ps1` | lancia l'analisi e ricarica `Msys_TB99_LINK_Object`; salva le destinazioni e le riaggancia ai nuovi `ID_OBJ` |
| `GENERA_HTML.bat` | rigenera gli `.html` da questi `.md` |

Per ricostruire gli oggetti dopo aver modificato i progetti, con `MSys_TB99_LINK.mdb` chiuso:

```
powershell -File RICOSTRUISCI_TB99_OGGETTI_1_ESPORTA.ps1 -Lavoro C:\TEMP\tb99_oggetti
powershell -File RICOSTRUISCI_TB99_OGGETTI_3_CARICA.ps1  -Lavoro C:\TEMP\tb99_oggetti
```

Limiti dell'analisi:
- un oggetto viene legato a una tabella se ne contiene il nome esatto;
- i nomi costruiti a pezzi nel codice VBA non vengono riconosciuti;
- le dipendenze tra database diversi non vengono seguite.

## 7. Piano dei prossimi passi

### 7.1 Fase 1 – Menu nel progetto master (l'esempio di questo progetto)

1. In `MSYS_ATTIVA_GEST_OGGETTI.mdb` togliere i vecchi oggetti e i link `GEST_MENU_Tb01_...`, `GEST_MENU_TB02_...`.
2. Esportare gli oggetti elencati dalla query 04 da `GE_MSys_TB06_MENU.mdb`.
3. Attivare le 4 destinazioni e lanciare `INVIA_LINK_FRONTEND_pf` sul progetto master.

### 7.2 Fase 2 – Esportazione automatica degli oggetti

Nuovo modulo `Mdl06_ESPORTA_OGGETTI`: per ogni destinazione attiva legge la query 04 ed esporta gli oggetti con `DoCmd.TransferDatabase acExport`, poi invia i link con `Mdl05`. Esiti e log come `Mdl05`.

### 7.3 Fase 3 – Primo invio reale dei link ai front-end

Chiudere i front-end, farne una copia, lanciare `VERIFICA_LINK_FRONTEND_pf` e poi `INVIA_LINK_FRONTEND_pf` sulle 57 destinazioni attive.

### 7.4 Fase 4 – Pulizia

- Togliere `Msys_TB99_LINK_Object_OLD_2026_11_15` quando non serve più.
- Togliere `tmp_ImportOggetti` e `Query1`.
- Togliere la tabella `Msys_TB99_}---@LINK`: è una copia dei progetti di TB01 con 24 righe.
- Le tabelle separatore non hanno oggetti e quindi nemmeno destinazioni: se servono nei front-end vanno create a mano.
- Tabelle a catalogo ma senza oggetti: `MSys_DF04_utile`, `MSys_TAB40_FORM_MASTER*`, `MSys_TB07_MENU_FUNZIONI_NW`, `MSys_TB08_MENU_OGGETTI`, `Msys_ESPORTA_ELENCO_OGGETTI_DB_TMP`. Decidere se servono.
- Registrare nel catalogo `Msys_TB02_TIPO_OGGETTI.MDB` e `Msys_TB03_ELENCO_ObjProgetti.mdb`, ora presenti in `MSYS\MSys_TB01_PROGETTI\MDB\`.

## 8. Catalogo completo dei collegamenti (54)

| Mini-progetto | Tabella | Database | Cartella (in `PROGETTI_MDB\MSYS\`) | Oggetti | Destinazioni dei suoi oggetti | Verifica |
| --- | --- | --- | --- | --- | --- | --- |
| `MSYS_ATTIVA_GEST` | `MSys_TAB40_}----------------------------------------@FORM_MASTER` | `MSYS_ATTIVA_GEST_OGGETTI.mdb` | `MSYS_ATTIVA_GEST\MDB\` | 0 | 0 | OK |
| `MSYS_ATTIVA_GEST` | `MSys_TAB40_FORM_MASTER` | `MSYS_ATTIVA_GEST_OGGETTI.mdb` | `MSYS_ATTIVA_GEST\MDB\` | 0 | 0 | OK |
| `MSYS_ATTIVA_GEST` | `MSys_TAB40_FORM_MASTER_TMP` | `MSYS_ATTIVA_GEST_OGGETTI.mdb` | `MSYS_ATTIVA_GEST\MDB\` | 0 | 0 | OK |
| `MSys_DF01` | `MSys_DF01_}----------------------------------------------------@` | `MSys_DF01.mdb` | `MSys_DF01\MDB\` | 0 | 0 | OK |
| `MSys_DF01` | `MSys_DF01_COMANDI` | `MSys_DF01.mdb` | `MSys_DF01\MDB\` | 6 | 12 | OK |
| `MSys_DF01` | `MSys_DF01_COMANDI_TMP` | `MSys_DF01.mdb` | `MSys_DF01\MDB\` | 0 | 0 | OK |
| `MSys_DF02` | `MSys_DF02_}----------------------------------------------------@` | `MSys_DF02.mdb` | `MSys_DF02\MDB\` | 0 | 0 | OK |
| `MSys_DF02` | `MSys_DF02_TIPO_OGGETTO` | `MSys_DF02.mdb` | `MSys_DF02\MDB\` | 4 | 28 | OK |
| `MSys_DF03` | `MSys_DF03_}----------------------------------------------------@` | `MSys_DF03.mdb` | `MSys_DF03\MDB\` | 0 | 0 | OK |
| `MSYS_DF03` | `MSYS_DF03_GROUP_QUOTE` | `MSys_DF03.mdb` | `MSys_DF03\MDB\` | 0 | 0 | OK |
| `MSys_DF03` | `MSys_DF03_QUOTE` | `MSys_DF03.mdb` | `MSys_DF03\MDB\` | 1 | 1 | OK |
| `MSys_DF04` | `MSys_DF04_}----------------------------------------------------@` | `MSys_DF04.mdb` | `MSys_DF04\MDB\` | 0 | 0 | OK |
| `MSys_DF04` | `MSys_DF04_utile` | `MSys_DF04.mdb` | `MSys_DF04\MDB\` | 0 | 0 | OK |
| `MSys_DF05` | `MSys_DF05_}----------------------------------------------------@` | `MSys_DF05.mdb` | `MSys_DF05\MDB\` | 0 | 0 | OK |
| `MSys_DF05` | `MSys_DF05_ATTRIBUTI_TABELLA` | `MSys_DF05.mdb` | `MSys_DF05\MDB\` | 1 | 1 | OK |
| `MSys_DF11` | `MSys_DF11_{@@@@@@@@@@@@@@@@@@@COD_PROGETTO@@@@@@@@@@@@@@@@@@@@@}` | `MSys_DF11.mdb` | `MSys_DF11\MDB\` | 0 | 0 | OK |
| `MSys_DF11` | `MSys_DF11_}----------------------------------------------------@` | `MSys_DF11.mdb` | `MSys_DF11\MDB\` | 0 | 0 | OK |
| `Msys_DF11` | `Msys_DF11_COD_PROGETTO` | `MSys_DF11.mdb` | `MSys_DF11\MDB\` | 5 | 45 | OK |
| `Msys_DF11` | `Msys_DF11_PROGETTI` | `MSys_DF11.mdb` | `MSys_DF11\MDB\` | 4 | 20 | OK |
| `Msys_DF12` | `Msys_DF12_}----------------------------------------------------@` | `MSys_DF12.mdb` | `MSys_DF12\MDB\` | 0 | 0 | OK |
| `Msys_DF12` | `Msys_DF12_ObjProgetti` | `MSys_DF12.mdb` | `MSys_DF12\MDB\` | 0 | 0 | OK |
| `MSys_DF13` | `MSys_DF13_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab` | `MSys_DF13.mdb` | `MSys_DF13\MDB\` | 0 | 0 | OK |
| `Msys_DF13` | `Msys_DF13_}----------------------------------------------------@` | `MSys_DF13.mdb` | `MSys_DF13\MDB\` | 0 | 0 | OK |
| `Msys_DF13` | `Msys_DF13_DLL_PROGETTO` | `MSys_DF13.mdb` | `MSys_DF13\MDB\` | 3 | 6 | OK |
| `Msys_TB01_GE` | `MSys_ESPORTA_}-------------------------------------------------@` | `GE_Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 0 | 0 | OK |
| `Msys_TB01_GE` | `Msys_ESPORTA_DB_EST` | `GE_Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 15 | 0 | OK |
| `Msys_TB01_GE` | `Msys_ESPORTA_DB_EST_TMP` | `GE_Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 9 | 0 | OK |
| `Msys_TB01_GE` | `Msys_ESPORTA_ELENCO_OGGETTI_DB` | `GE_Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 14 | 0 | OK |
| `Msys_TB01_GE` | `Msys_ESPORTA_ELENCO_OGGETTI_DB_TMP` | `GE_Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 0 | 0 | OK |
| `Msys_TB01` | `Msys_TB01_}---------------------------------------@PROGETTI` | `Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 0 | 0 | OK |
| `Msys_TB01` | `Msys_TB01_PROGETTI` | `Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 4 | 8 | OK |
| `Msys_TB01` | `PROGETTI_MSYS_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab` | `Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 0 | 0 | OK |
| `Msys_TB01` | `PROGETTI_Msys_TB01_}-------------------------------------------@` | `Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 0 | 0 | OK |
| `Msys_TB01` | `PROGETTI_Msys_TB01_PROJECT` | `Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 17 | 85 | OK |
| `Msys_TB01` | `PROGETTI_Msys_TB01_PROJECT_TMP` | `Msys_TB01_PROGETTI.mdb` | `MSys_TB01_PROGETTI\MDB\` | 19 | 57 | OK |
| `MSys_TB06` | `MSys_TB06_{@==================================================@}` | `MSys_TB06_MENU_VOCI.mdb` | `MSys_TB06_MENU\MDB\` | 0 | 0 | OK |
| `MSys_TB06` | `MSys_TB06_}-----------------------------------------------@` | `MSys_TB06_MENU_VOCI.mdb` | `MSys_TB06_MENU\MDB\` | 0 | 0 | OK |
| `MSys_TB06` | `MSys_TB06_MENU_VOCI` | `MSys_TB06_MENU_VOCI.mdb` | `MSys_TB06_MENU\MDB\` | 5 | 10 | OK |
| `MSys_TB06` | `MSys_TB06_MENU_VOCI_TMP` | `MSys_TB06_MENU_VOCI.mdb` | `MSys_TB06_MENU\MDB\` | 8 | 16 | OK |
| `MSys_TB07` | `MSys_TB07_{@==================================================@}` | `MSys_TB07_MENU_FUNZIONI.mdb` | `MSys_TB06_MENU\MDB\` | 0 | 0 | OK |
| `MSys_TB07` | `MSys_TB07_}-----------------------------------------------@` | `MSys_TB07_MENU_FUNZIONI.mdb` | `MSys_TB06_MENU\MDB\` | 0 | 0 | OK |
| `MSys_TB07` | `MSys_TB07_MENU_FUNZIONI` | `MSys_TB07_MENU_FUNZIONI.mdb` | `MSys_TB06_MENU\MDB\` | 4 | 12 | OK |
| `MSys_TB07` | `MSys_TB07_MENU_FUNZIONI_NW` | `MSys_TB07_MENU_FUNZIONI.mdb` | `MSys_TB06_MENU\MDB\` | 0 | 0 | OK |
| `MSys_TB07` | `MSys_TB07_MENU_FUNZIONI_TMP` | `MSys_TB07_MENU_FUNZIONI.mdb` | `MSys_TB06_MENU\MDB\` | 5 | 15 | OK |
| `MSys_TB08` | `MSys_TB08_}-----------------------------------------------@` | `MSys_TB08_MENU_OGGETTI.mdb` | `MSys_TB06_MENU\MDB\` | 0 | 0 | OK |
| `MSys_TB08` | `MSys_TB08_MENU_OGGETTI` | `MSys_TB08_MENU_OGGETTI.mdb` | `MSys_TB06_MENU\MDB\` | 0 | 0 | OK |
| `Msys_TB12` | `Msys_TB12_FILES` | `Msys_TB12_FILES.mdb` | `Msys_TB12_FILES\MDB\` | 5 | 0 | OK |
| `MSys_TB30` | `MSys_TB30_}----------------------------------------------------@` | `Msys_TB30_UTIL.mdb` | `Msys_TB30_UTIL\MDB\` | 0 | 0 | OK |
| `MSys_TB30` | `MSys_TB30_UTIL_MSysObjects` | `Msys_TB30_UTIL.mdb` | `Msys_TB30_UTIL\MDB\` | 8 | 0 | OK |
| `MSys_TB40N` | `MSys_TB40_}----------------------------------------------------@` | `MSys_TB40.mdb` | `MSys_TB40\MDB\` | 0 | 0 | OK |
| `MSys_TB40N` | `MSys_TB40_Tb01_PROGETTI` | `MSys_TB40.mdb` | `MSys_TB40\MDB\` | 0 | 0 | OK |
| `MSys_TB40N` | `MSys_TB40_Tb01_PROGETTI_TMP` | `MSys_TB40.mdb` | `MSys_TB40\MDB\` | 0 | 0 | OK |
| `MSys_TB40N` | `MSys_TB40_Tb02_PAGINE` | `MSys_TB40.mdb` | `MSys_TB40\MDB\` | 0 | 0 | OK |
| `MSys_TB40N` | `MSys_TB40_Tb02_PAGINE_TMP` | `MSys_TB40.mdb` | `MSys_TB40\MDB\` | 0 | 0 | OK |
