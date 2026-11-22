---
titolo: Linguetta COLLEGAMENTI
menu: COLLEGAMENTI
descrizione: tabella collegata, i suoi oggetti e le loro destinazioni
---
**Dove si trova:** prima linguetta, pagina `Pg_01`. E' la pagina che si vede all'apertura della form.

## La catena delle tre sottomaschere

1. **In alto - TABELLE COLLEGATE** (`SottoForm_LINK` = `Msys_TB99_LINK_FrmS05_LINK`):
   le tabelle del catalogo `Msys_TB99_LINK` che hanno almeno un oggetto nel loro database.
2. **In basso a sinistra - OGGETTI** (`SottoForm_OGGETTI` = `Msys_TB99_LINK_FrmS03_OGGETTI`):
   gli oggetti della tabella selezionata in alto (legame `Name`).
3. **In basso a destra - DESTINAZIONI** (`SottoForm_PATH_DB` = `Msys_TB99_LINK_FrmS04_PATH_DB`):
   i database in cui esportare l'oggetto selezionato a sinistra (legame `ID_OBJ`).

Selezionando una riga in alto cambiano gli oggetti; selezionando un oggetto cambiano le destinazioni.

> [!NOTA] Gli oggetti di una tabella sono **solo** quelli che stanno nello **stesso database della tabella vera**
> (es. `MSys_DF01_COMANDI` in `MSys_DF01.mdb` -> le sue query, form e moduli in `MSys_DF01.mdb`),
> non quelli dei front end o di altri database.

## Le colonne

### Tabelle collegate

| Colonna | Significato |
|---|---|
| Mini-progetto | codice del mini-progetto (`COD_SUB_PROGETTO_s`) |
| Tabella collegata | nome della tabella (`Name`) |
| Database (file) / Cartella | il database madre, dove sta la tabella vera |
| Verifica | esito del controllo del collegamento |
| N. oggetti | quanti oggetti usano la tabella |
| Connect, Type, ID tabella | dati tecnici del collegamento |

### Oggetti

| Colonna | Significato |
|---|---|
| Tipo oggetto / Nome oggetto | query, form, report, modulo e il suo nome |
| Database (file) / Cartella | il database in cui si trova l'oggetto |
| Scelto | oggetto scelto per l'esportazione |

### Destinazioni

| Colonna | Significato |
|---|---|
| Attiva | destinazione usata dall'invio (`Scel_b`) |
| Database destinazione / Cartella | il database in cui esportare l'oggetto |
| Esito ultimo invio, Nota, Data invio | risultato dell'ultimo invio dei link |

## Aggiornamento

Tornando su questa linguetta le tre sottomaschere vengono rilette dalle tabelle e l'elenco riparte dalla prima
tabella collegata.
