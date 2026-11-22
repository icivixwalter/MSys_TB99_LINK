---
titolo: Indice della guida
menu: Indice
guida: Guida della form CATALOGO DEI COLLEGAMENTI
sottotitolo: Form <b>Msys_TB99_LINK_FrmM02_GE_LINK</b> - database MSys_TB99_LINK - catalogo TB99_LINK
versione: 1.1 del 22/11/2026
---
> **A cosa serve la form.** E' la maschera principale del catalogo TB99: mostra le **tabelle collegate** del
> progetto, gli **oggetti** (query, form, report, moduli) che le usano nel loro database e i **database di
> destinazione** in cui ogni oggetto va esportato. Dalla seconda linguetta si importano gli oggetti di un
> mini-progetto esterno.

## Le linguette

La form e' un controllo a schede (`TabCtl_01`) con 12 pagine; oggi se ne vedono 2. Ad ogni cambio di linguetta
le sottomaschere della pagina scelta vengono aggiornate (rilette dalle tabelle).

| Linguetta | Pagina | Cosa contiene |
|---|---|---|
| **COLLEGAMENTI** | `Pg_01` | la catena tabella -> oggetti -> destinazioni - [dettagli](Msys_TB99_LINK_FrmM02_GE_LINK_HELP_01_COLLEGAMENTI.md) |
| **IMPORTA OGGETTI** | `Pg_02` | la maschera di importazione degli oggetti - [dettagli](Msys_TB99_LINK_FrmM02_GE_LINK_HELP_02_IMPORTA_OGGETTI.md) |
| *(nascoste)* | `Pg_03` ... `Pg_12` | pagine pronte per usi futuri |

## Misura della form

La form si adatta alla finestra di Access: il controllo a schede arriva fino ai bordi e le sottomaschere
occupano tutta la pagina (su `Pg_01` le tre sottomaschere mantengono le loro proporzioni). Non scende mai sotto
la misura di progetto, cosi' nessun contenuto viene tagliato, e non compaiono barre di scorrimento.
Anche la maschera della linguetta **IMPORTA OGGETTI** allarga il suo elenco e le sue caselle di testo.

## Il pulsante ?

Il pulsante [[?]] in alto a destra apre la pagina di questa guida che riguarda la linguetta attiva.

## Le pagine della guida

{{SCHEDE}}

## Le tabelle usate

| Tabella | Contenuto |
|---|---|
| `Msys_TB99_LINK` | catalogo delle tabelle collegate e del loro database madre |
| `Msys_TB99_LINK_Object` | gli oggetti che usano ogni tabella, nello stesso database della tabella vera (legame `Name`) |
| `Msys_TB99_LINK_PATH_DB` | i database in cui esportare ogni oggetto (legame `ID_OBJ`) |
