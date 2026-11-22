---
titolo: 2) CREA I LINK
menu: 2) CREA I LINK
descrizione: la creazione vera dei link
---
**Dove si trova:** nella riga dei pulsanti, [[2) CREA I LINK|invia]] (in rosso). Il punto [[?]] sotto il pulsante apre un aiuto breve; da li si arriva a questa pagina.

> [!ATTENZIONE] **Questo pulsante modifica i database di destinazione.** I link vengono creati **nel database di destinazione** (colonna Destinazione) e puntano al **database madre** della tabella. Nel database MSys_TB99_LINK si scrivono solo gli esiti.

## Quando si accende

- Solo dopo [[1) SIMULA|simula]], e solo se le scelte non sono cambiate.
- Si spegne se si cambia la destinazione, una casella **Attiva**, o si usa Attiva tutte / Disattiva tutte / Solo front end GE_.
- La riga di stato sotto i pulsanti dice sempre perche' e' spento o acceso.

## Cosa fa

1. Ricontrolla che le scelte siano le stesse della simulazione.
1. Chiede conferma indicando quanti database, quante righe attive e quante righe da controllare (che verranno saltate).
1. Richiama `INVIA_LINK_FRONTEND_pf(destinazione, bSimulazione = False)` del modulo Mdl05.
1. Per ogni riga ok: se il link esiste lo cancella e lo ricrea (`SOVRASCRITTO`), altrimenti lo crea (`CREATO`).
1. Le righe da controllare vengono saltate e mantengono il loro esito.
1. Mette il riepilogo nella casella a destra, aggiorna l'elenco e spegne di nuovo il pulsante.

## Cosa non fa mai

- Non copia tabelle e non tocca le tabelle vere (`TABELLA_LOCALE`).
- Non crea il link nel database madre della tabella (`SORGENTE`).
- Non lavora fuori dall'albero `C:\CASA\...\PROGETTI_MDB\MSYS`.

## Prima di premerlo

- Chiudere i database di destinazione aperti da altri utenti: un database bloccato da' l'esito `FE_NON_APRIBILE`.
- E' buona norma avere una copia di salvataggio dei front end.

## Dopo

Stampare il report con [Stampa report](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_07_STAMPA.md) per conservare cosa e' stato fatto. Il log dell'operazione e' nella cartella `LOGS` accanto al database. Per un nuovo invio rifare 1) SIMULA.
