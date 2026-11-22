---
titolo: 1) SIMULA
menu: 1) SIMULA
descrizione: prova senza modificare nulla
---
**Dove si trova:** nella riga dei pulsanti, [[1) SIMULA|simula]]. Il punto [[?]] sotto il pulsante apre un aiuto breve; da li si arriva a questa pagina.

> [!OK] **La simulazione non modifica nessun database di destinazione.** Scrive soltanto l'esito `SIMULA_*` e la nota nella tabella `Msys_TB99_LINK_PATH_DB`.

## Cosa fa

1. Prende le righe con **Attiva = si** della destinazione scelta; se non ce n'e' nessuna avvisa e si ferma.
1. Richiama `INVIA_LINK_FRONTEND_pf(destinazione, bSimulazione = True)` del modulo Mdl05.
1. Per ogni riga controlla: il database di destinazione esiste e si apre? la tabella e' a catalogo? il database madre esiste ed e' nell'radice C:\CASA\...\MSYS? c'e' gia' un link o una tabella vera con lo stesso nome?
1. Scrive l'esito: `SIMULA_CREATO` (il link verrebbe creato), `SIMULA_SOVRASCRITTO` (il link c'e' e verrebbe rifatto) oppure un esito da controllare.
1. Aggiorna l'elenco (verde = ok, rosso = da controllare) e mette il riepilogo nella casella a destra.
1. Ricorda le scelte su cui e' stata fatta (l'"impronta") e accende [[2) CREA I LINK|invia]].

## Come leggere il risultato

- **Righe in verde:** verranno create o rifatte.
- **Righe in rosso:** non verranno create; leggere la nota e la pagina [Esiti](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_10_ESITI.md).
- **SIMULA_SOVRASCRITTO** con nota "prima: Y:\..." o "prima: c:\..." vuol dire che oggi il link punta a un percorso vecchio: e' proprio cio' che l'invio corregge.
- Il filtro *solo righe da controllare* mostra subito i problemi.

## Durata

La simulazione apre ogni database di destinazione in rete: con molti database puo' durare qualche minuto. Il cursore a clessidra indica che sta lavorando.

## Dopo

Controllare con [Riepilogo](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_08_RIEPILOGO.md) e [Stampa report](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_07_STAMPA.md), poi premere [2) CREA I LINK](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_06_CREA_LINK.md).
