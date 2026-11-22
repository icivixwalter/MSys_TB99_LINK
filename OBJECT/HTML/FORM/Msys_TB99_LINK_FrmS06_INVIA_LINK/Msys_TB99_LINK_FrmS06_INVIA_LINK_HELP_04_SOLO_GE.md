---
titolo: Solo front end GE_
menu: Solo front end GE_
descrizione: la scelta consigliata
---
**Dove si trova:** nella riga dei pulsanti, [[Solo front end GE_]]. Il punto [[?]] sotto il pulsante apre un aiuto breve; da li si arriva a questa pagina.

## Cosa fa

Mette **Attiva = si** sulle righe dei database il cui nome comincia con `GE_` o `MSYS_ATTIVA_GEST` (i front end) e **Attiva = no** su tutte le altre.

- Agisce sulla **Destinazione** scelta in alto; con *(tutti)* agisce su tutti i database.
- Non tiene conto dei filtri *solo righe attive* / *solo righe da controllare*: cambia anche le righe che non si vedono.
- Scrive subito `Scel_b = True solo per GE_* e MSYS_ATTIVA_GEST*` nella tabella `Msys_TB99_LINK_PATH_DB`. Non tocca nessun database di destinazione.

## Quando usarlo

**E' la scelta consigliata.** I link servono ai front end, dove gli oggetti importati usano tabelle che stanno in altri database. Gli altri database contengono di solito le tabelle vere e non hanno bisogno di link.

> [!ATTENZIONE] Se c'era una simulazione, cambiando le scelte [[2) CREA I LINK|invia]] si spegne: rifare [[1) SIMULA|simula]].

## Dopo

Controllare l'elenco, poi premere [1) SIMULA](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_05_SIMULA.md).
