---
titolo: Attiva tutte
menu: Attiva tutte
descrizione: includere tutte le righe della destinazione
---
**Dove si trova:** nella riga dei pulsanti, [[Attiva tutte]]. Il punto [[?]] sotto il pulsante apre un aiuto breve; da li si arriva a questa pagina.

## Cosa fa

Mette **Attiva = si** su tutte le righe.

- Agisce sulla **Destinazione** scelta in alto; con *(tutti)* agisce su tutti i database.
- Non tiene conto dei filtri *solo righe attive* / *solo righe da controllare*: cambia anche le righe che non si vedono.
- Scrive subito `Scel_b = True` nella tabella `Msys_TB99_LINK_PATH_DB`. Non tocca nessun database di destinazione.

## Quando usarlo

Quando si vogliono inviare tutti i link di un database, ad esempio dopo aver scelto un solo database come destinazione. Con *(tutti)* attiva anche i database che non sono front end: di solito e' meglio [Solo front end GE_](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_04_SOLO_GE.md).

> [!ATTENZIONE] Se c'era una simulazione, cambiando le scelte [[2) CREA I LINK|invia]] si spegne: rifare [[1) SIMULA|simula]].

## Dopo

Controllare l'elenco, poi premere [1) SIMULA](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_05_SIMULA.md).
