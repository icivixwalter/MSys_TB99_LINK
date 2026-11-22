# HELP DEGLI OGGETTI - regole

Ogni oggetto del progetto (form, query, report, modulo, tabella, macro) ha il suo help in una cartella propria:

```
OBJECT\HTML\
    GENERA_HELP_HTML.bat          rigenera tutte le pagine .html dai .md
    INDICE_HELP.html              indice di tutti gli help (generato)
    LEGGIMI_HELP.md               queste regole
    _STRUMENTI\                   genera_help.py + HELP_stile.css (foglio di stile comune)
    FORM\<NomeOggetto>\           <NomeOggetto>_HELP_00_INDICE.md   + .html generato
                                  <NomeOggetto>_HELP_01_<CHIAVE>.md + .html generato ...
    QUERY\  REPORT\  MODULE\  TABLE\  MACRO\   stessa regola
```

- **Si modifica solo il `.md`**: l'`.html` viene riscritto da `GENERA_HELP_HTML.bat` (doppio clic).
- Il nome di ogni file inizia con il **nome della cartella** (= nome dell'oggetto), poi `_HELP_`, due cifre e una chiave.
  La pagina `00` e' l'indice dell'oggetto; il numero decide l'ordine del menu.
- Nella cartella si possono mettere anche allegati (pdf, immagini) e richiamarli con un link.
- Da Access le pagine si aprono con il modulo `Msys_TB99_LINK_ Mdl07_HELP`:
  `HELP_Apri_pf "FORM", Me.Name, "01_COLLEGAMENTI"` (senza pagina apre la `00_INDICE`).

## Il file .md

```
---
titolo: Destinazione e filtri            titolo della pagina
menu: Destinazione e filtri              voce del menu in alto (se manca = titolo)
descrizione: quale database scegliere    testo della scheda nell'indice
guida: Guida della form INVIO DEI LINK   solo pagina 00: titolo di tutte le pagine
sottotitolo: Form <b>...</b>             solo pagina 00: riga sotto il titolo
versione: 1.0 del 22/11/2026             solo pagina 00
---
## Titolo di paragrafo
### Sottotitolo

Testo normale, **grassetto**, *corsivo*, `codice`, [link a un'altra pagina](Nome_HELP_01_X.md).
Pulsante della form: [[Stampa report]], colorato: [[1) SIMULA|simula]] [[2) CREA I LINK|invia]]

- elenco puntato
1. elenco numerato

| Colonna | Significato |
|---------|-------------|
| {r}ROSSO | cella in rosso |  ({v} = verde)

> [!ATTENZIONE] riquadro rosso   ( [!OK] verde, [!NOTA] o niente = blu )
> le righe seguenti del riquadro iniziano con >

{{SCHEDE}}   (pagina 00) schede che portano a tutte le altre pagine
```

Le righe che iniziano con `<` sono html scritto a mano e passano cosi' come sono.
