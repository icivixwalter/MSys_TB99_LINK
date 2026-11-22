---
titolo: Esiti
menu: Esiti
descrizione: cosa significano e cosa fare
---
L'esito di ogni riga e' nella colonna **Esito** dell'elenco e nel report. In verde gli esiti buoni, in rosso quelli da controllare.

| Esito | Significato | Cosa fare |
|---|---|---|
| DA_INVIARE | riga non ancora simulata ne' inviata | premere 1) SIMULA |
| {v}SIMULA_CREATO | il link verrebbe creato | niente: verra' creato con 2) CREA I LINK |
| {v}SIMULA_SOVRASCRITTO | il link c'e' gia' e verrebbe rifatto; la nota dice dove punta oggi | niente: verra' ricreato sul database madre giusto |
| {v}CREATO | link creato | - |
| {v}SOVRASCRITTO | link ricreato | - |
| {r}TABELLA_LOCALE | nel database di destinazione c'e' una tabella vera con lo stesso nome: non viene toccata | decidere a mano se la tabella locale va tenuta; altrimenti disattivare la riga |
| {r}SORGENTE | la destinazione e' il database madre della tabella: niente link | disattivare la riga |
| {r}NON_A_CATALOGO | la tabella non c'e' in `Msys_TB99_LINK` | aggiungerla al catalogo o disattivare la riga |
| {r}CATALOGO_NON_OK | il database madre e' fuori dall'radice C:\CASA\...\MSYS oppure Verifica_s non e' OK | correggere il database madre nel catalogo o disattivare la riga |
| {r}SORGENTE_MANCANTE | il database madre o la tabella madre non si trovano | controllare il percorso nel catalogo |
| {r}FE_MANCANTE | il database di destinazione non esiste | controllare il percorso in PATH_DB |
| {r}FE_NON_VALIDO | destinazione fuori dall'radice C:\CASA\...\MSYS o uguale a questo database | disattivare la riga |
| {r}FE_NON_APRIBILE | destinazione bloccata (aperta da qualcuno) o danneggiata | chiuderla e rifare la simulazione |
| {r}ERRORE | errore imprevisto | leggere la nota e il log nella cartella LOGS |
