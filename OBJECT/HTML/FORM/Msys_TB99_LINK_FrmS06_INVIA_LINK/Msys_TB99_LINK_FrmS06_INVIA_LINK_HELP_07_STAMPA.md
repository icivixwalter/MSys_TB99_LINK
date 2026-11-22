---
titolo: Stampa report
menu: Stampa report
descrizione: il documento completo da conservare
---
**Dove si trova:** nella riga dei pulsanti, [[Stampa report]]. Il punto [[?]] sotto il pulsante apre un aiuto breve; da li si arriva a questa pagina.

## Cosa fa

Apre in anteprima il report `Msys_TB99_LINK_RptS06_INVIA_LINK` con le righe che si vedono nell'elenco: stessa destinazione e stessi filtri (*solo righe attive*, *solo righe da controllare*).

## Com'e' fatto il report

- **Prima pagina:** lo stesso riepilogo della casella di testo (totali per esito, una riga per database, righe da controllare).
- **Pagine seguenti:** un gruppo per ogni database di destinazione, con il percorso.
- Per ogni riga: Attiva, tabella collegata, database madre, numero di oggetti, esito (in rosso se da controllare), nota (per i link rifatti: dove puntava prima).
- In fondo a ogni gruppo: quanti link e quanti da controllare. A pie' di pagina: data di stampa e numero di pagina.
- Pagina orizzontale; i percorsi lunghi sono accorciati in `H:\..\` e `Y:\..\`.

## Quando usarlo

E' il documento completo da conservare: dopo la simulazione, per controllare riga per riga con calma; dopo la creazione, come traccia di cio' che e' stato fatto. Dall'anteprima si puo' stampare o salvare in PDF.

> Per una stampa breve con i soli problemi: spuntare *solo righe da controllare* e poi premere Stampa report.
