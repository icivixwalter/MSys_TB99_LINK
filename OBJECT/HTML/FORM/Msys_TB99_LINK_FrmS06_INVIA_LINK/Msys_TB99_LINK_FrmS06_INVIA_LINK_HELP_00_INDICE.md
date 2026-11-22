---
titolo: Indice della guida
menu: Indice
guida: Guida della form INVIO DEI LINK
sottotitolo: Form <b>Msys_TB99_LINK_FrmS06_INVIA_LINK</b> - database MSys_TB99_LINK - catalogo TB99_LINK
versione: 1.1 del 22/11/2026 (1.0 del 21/11/2026)
---
> **A cosa serve la form.** Crea nei **database di destinazione** le tabelle **collegate** (link) che servono agli oggetti importati (query, form, report). Ogni link punta al **database madre**, cioe' il database dove sta la tabella vera, preso dal catalogo `Msys_TB99_LINK`. Nel database MSys_TB99_LINK si scrivono soltanto gli esiti: le tabelle vere non vengono mai copiate.

## La sequenza di lavoro

1. Scegliere la **Destinazione** (un database oppure *(tutti)*) - [dettagli](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_01_FILTRI.md).
1. Decidere quali righe inviare con la casella **Attiva** nell'elenco, oppure con [Attiva tutte](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_02_ATTIVA_TUTTE.md), [Disattiva tutte](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_03_DISATTIVA.md), [Solo front end GE_](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_04_SOLO_GE.md).
1. Premere [[1) SIMULA|simula]]: nessun database viene modificato, ogni riga riceve un esito - [dettagli](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_05_SIMULA.md).
1. Controllare gli esiti: righe in <span style="color:#aa0000;font-weight:bold">rosso</span> nell'elenco, [Riepilogo](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_08_RIEPILOGO.md) nella casella a destra, [Stampa report](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_07_STAMPA.md).
1. Premere [[2) CREA I LINK|invia]]: si accende solo dopo una simulazione fatta sulle stesse scelte - [dettagli](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_06_CREA_LINK.md).

## Le pagine della guida

{{SCHEDE}}

## Le regole che la form rispetta

La form non ha regole proprie: richiama il modulo `Msys_TB99_LINK_ Mdl05_INVIA_LINK_FRONTEND` (funzione `INVIA_LINK_FRONTEND_pf`).

- Si creano solo **link**; le tabelle vere non vengono mai copiate.
- Un link gia' presente viene cancellato e ricreato, cosi' punta al percorso giusto (`SOVRASCRITTO`).
- Una tabella **vera** con lo stesso nome nel database di destinazione non viene toccata (`TABELLA_LOCALE`).
- Nel database madre della tabella non si crea il link (`SORGENTE`).
- Si lavora solo dentro l'albero `C:\CASA\...\PROGETTI_MDB\MSYS`; sono escluse le cartelle OLD, AA_SALVATAGGI, BACKUP, EXE e .git.

## Le tabelle usate

| Tabella | Contenuto |
|---|---|
| `Msys_TB99_LINK` | catalogo delle tabelle vere e del loro database madre |
| `Msys_TB99_LINK_Object` | gli oggetti (query, form, report) che usano ogni tabella |
| `Msys_TB99_LINK_PATH_DB` | i database dove gli oggetti sono usati: per ogni riga la scelta (`Scel_b` = colonna **Attiva**), l'esito e la nota |
