---
titolo: Destinazione e filtri
menu: Destinazione e filtri
descrizione: quale database e quali righe vedere
---
**Dove si trova:** in alto: casella **Destinazione** e caselle **solo righe attive** e **solo righe da controllare**. Il punto [[?]] sotto il pulsante apre un aiuto breve; da li si arriva a questa pagina.

## Destinazione

Sceglie il database su cui lavorare. Con **(tutti)** la form lavora su tutti i database dell'elenco.

La destinazione vale per tutta la form: l'elenco, i pulsanti Attiva/Disattiva/Solo GE_, la simulazione, la creazione dei link, il riepilogo e il report riguardano solo il database scelto.

> [!ATTENZIONE] Cambiare la destinazione dopo una simulazione spegne [[2) CREA I LINK|invia]]: la simulazione va rifatta sulla nuova destinazione.

## solo righe attive

Mostra nell'elenco solo le righe con **Attiva** = si, cioe' quelle che verranno simulate e inviate. Cambia solo cio' che si vede: non cambia le scelte.

## solo righe da controllare

Mostra solo le righe con un esito da controllare (in rosso). Utile dopo la simulazione per vedere subito i problemi. Vale anche per il report: [[Stampa report]] stampa le righe che si vedono con i filtri attuali.

## L'elenco

Una riga per ogni coppia **tabella + database di destinazione**. Le colonne:

| Colonna | Significato |
|---|---|
| Attiva | si = la riga viene simulata e inviata; si cambia con un clic |
| Destinazione | il database dove nasce il link |
| Tabella | la tabella da collegare |
| Database_madre | il database dove sta la tabella vera (catalogo Msys_TB99_LINK) |
| Oggetti | quanti oggetti della destinazione usano la tabella |
| Esito | risultato dell'ultima simulazione o creazione - [vedi gli esiti](Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_10_ESITI.md) |
| Nota | spiegazione dell'esito; per i link rifatti indica dove puntava prima |

Il clic su **Attiva** salva subito la scelta nella tabella `Msys_TB99_LINK_PATH_DB` (campo `Scel_b`), per tutte le righe di quella coppia tabella + destinazione.
