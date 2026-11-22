Option Compare Database
Option Explicit

'==================================================================================================
' INTESTAZIONE
'==================================================================================================
' MODULO     : Msys_TB99_LINK_ Mdl06_FORM_INVIA_LINK
' PROGETTO   : LLPP_DF99_LINK  -  catalogo TB99_LINK
' VERSIONE   : 1.0  del 21/11/2026
' FAQ        : @TB99@FORM@INVIA_(supporto alla form di simulazione e invio dei link ai front end)
'
' SCOPO
'   Funzioni usate dalla form Msys_TB99_LINK_FrmS06_INVIA_LINK e dal report
'   Msys_TB99_LINK_RptS06_INVIA_LINK. Non cambia le regole del modulo
'   "Msys_TB99_LINK_ Mdl05_INVIA_LINK_FRONTEND": lo richiama soltanto.
'
' SEQUENZA DI LAVORO (quella della form)
'   1) scegliere le righe (tabella da collegare + database di destinazione) con ATTIVA = Si/No
'   2) SIMULA  -> INVIA_LINK_FRONTEND_pf(fe, bSimulazione:=True)  : esiti SIMULA_*, nessuna modifica
'   3) controllare esiti, riepilogo nella casella di testo e report di stampa
'   4) CREA I LINK -> INVIA_LINK_FRONTEND_pf(fe, bSimulazione:=False) : crea i link nei DATABASE DI
'      DESTINAZIONE, che puntano al DATABASE MADRE della tabella (catalogo Msys_TB99_LINK)
'   Il pulsante 4) si abilita solo se la simulazione e' stata fatta sulle stesse scelte (S06_Firma).
'
' FUNZIONI PUBBLICHE
'   S06_Esegui        simulazione o invio, restituisce il riepilogo in testo
'   S06_Testo         riepilogo in testo degli esiti attuali (casella di testo)
'   S06_Firma         "impronta" delle scelte attive: se cambia, la simulazione va rifatta
'   S06_CambiaScelta  attiva/disattiva una riga (tutte le righe PATH_DB della coppia)
'   S06_ImpostaTutte  TUTTE / NESSUNA / SOLO_FE (solo front end: H: GE_* e GEST_*, C:\CASA GE_* e MSYS_ATTIVA_GEST*)
'   S06_Casa          True se il database e' la versione di casa (C:\CASA\...\MSYS): cambia radice e regola front end
'   S06_Problemi      numero di righe attive con esito da controllare
'   S06_Aiuto         testo di aiuto
'   S06_Where         filtro SQL comune a form e report
'   S06_Help          aiuto breve di un pulsante (punti "?") + apertura della pagina HTML
'   S06_HelpCartella  cartella delle pagine HTML (OBJECT\HTML\FORM\Msys_TB99_LINK_FrmS06_INVIA_LINK, modulo Mdl07_HELP)
'
' VERSIONE 1.1 del 21/11/2026: punti "?" sotto i pulsanti e guida HTML
' VERSIONE 1.2 del 21/11/2026: riconosce l'ambiente (H: ufficio / C:\CASA): radice, front end, cartella guida
' VERSIONE 1.3 del 22/11/2026: guida HTML nella cartella dell'oggetto OBJECT\HTML\FORM\Msys_TB99_LINK_FrmS06_INVIA_LINK (file .md + .html)
'
' INDICE
'   #01# costanti   #02# esecuzione   #03# riepilogo testo   #04# scelte   #05# aiuto
'   #07# aiuto dei pulsanti   #06# utilita
'==================================================================================================

'==================================================================================================
' #01# COSTANTI
'==================================================================================================
Public Const S06_QUERY As String = "Msys_TB99_LINK_Qry99_06_SELECT_INVIO_LINK"
Private Const ESITI_OK As String = "|SIMULA_CREATO|SIMULA_SOVRASCRITTO|CREATO|SOVRASCRITTO|"
'// form della guida e cartella fissa delle pagine HTML (usata se non si trovano accanto al database)
Private Const S06_FORM As String = "Msys_TB99_LINK_FrmS06_INVIA_LINK"
Private Const HELP_ASSOLUTA As String = "H:\LAVORI_PUBBLICI_FINO_AL_30_09_2019\Valter\OPERE_PUBBLICHE\96_ARCHIVI_LLPP_MDB\25_GESTIONE_LLPP_H\ARCHIVI_ACCESS_LLPP_H\LLPP_DF99_LINK\OBJECT\HTML\FORM\Msys_TB99_LINK_FrmS06_INVIA_LINK\"


'==================================================================================================
' #02# ESECUZIONE: simulazione (bSimula = True) o creazione dei link (bSimula = False)
'==================================================================================================
Public Function S06_Esegui(ByVal sFE As String, ByVal bSimula As Boolean) As String
    Dim sRiep As String
    If sFE = "(tutti)" Then sFE = ""
    DoCmd.Hourglass True
    sRiep = INVIA_LINK_FRONTEND_pf(sFE, bSimula, True)
    DoCmd.Hourglass False
    S06_Esegui = S06_Testo(sFE, IIf(bSimula, "SIMULAZIONE (nessuna modifica ai database)", "CREAZIONE DEI LINK ESEGUITA"), sRiep)
End Function


'==================================================================================================
' #03# RIEPILOGO IN TESTO
'==================================================================================================
Public Function S06_Testo(ByVal sFE As String, Optional ByVal sTitolo As String = "SITUAZIONE ATTUALE", _
                          Optional ByVal sRiep As String = "") As String
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim t As String, sW As String, n As Long, nProb As Long

    If sFE = "(tutti)" Then sFE = ""
    Set db = CurrentDb
    sW = S06_Where(sFE, True, False)

    t = "=== " & sTitolo & " ===" & vbCrLf
    t = t & "Data: " & Format(Now, "dd/mm/yyyy hh:nn") & vbCrLf
    t = t & "Destinazione: " & IIf(sFE = "", "tutti i database con righe attive", sFE) & vbCrLf
    If Left$(sRiep, 6) = "ERRORE" Then t = t & vbCrLf & "!!! ERRORE nel modulo di invio: " & sRiep & " (vedere il log)" & vbCrLf
    If sRiep = "ANNULLATO" Then t = t & vbCrLf & "Operazione annullata." & vbCrLf

    '// esiti delle righe attive
    t = t & vbCrLf & "ESITI (righe attive = tabella + database di destinazione)" & vbCrLf
    Set rs = db.OpenRecordset("SELECT Nz(ESITO,'(nessuno)') AS E, Count(*) AS N FROM " & S06_QUERY & sW & " GROUP BY Nz(ESITO,'(nessuno)') ORDER BY Count(*) DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        t = t & "  " & Left$(rs!e & Space(22), 22) & Right$(Space(6) & rs!n, 6) & "   " & SpiegaEsito(rs!e & "") & vbCrLf
        n = n + rs!n
        If InStr(ESITI_OK, "|" & rs!e & "|") = 0 And rs!e <> "DA_INVIARE" Then nProb = nProb + rs!n
        rs.MoveNext
    Loop
    rs.Close
    t = t & "  " & Left$("TOTALE" & Space(22), 22) & Right$(Space(6) & n, 6) & vbCrLf

    '// per database di destinazione
    t = t & vbCrLf & "PER DATABASE DI DESTINAZIONE (dove nasce il link)" & vbCrLf
    Set rs = db.OpenRecordset("SELECT FILE_DEST, Count(*) AS N, " & _
        "Sum(IIf(ESITO In ('SIMULA_CREATO','CREATO'),1,0)) AS NC, " & _
        "Sum(IIf(ESITO In ('SIMULA_SOVRASCRITTO','SOVRASCRITTO'),1,0)) AS NS, " & _
        "Sum(IIf(ESITO In ('SIMULA_CREATO','CREATO','SIMULA_SOVRASCRITTO','SOVRASCRITTO','DA_INVIARE') Or ESITO Is Null,0,1)) AS NP " & _
        "FROM " & S06_QUERY & sW & " GROUP BY FILE_DEST ORDER BY FILE_DEST", dbOpenSnapshot)
    Do While Not rs.EOF
        t = t & "  " & Left$(rs!FILE_DEST & Space(42), 42) & " link " & Right$(Space(4) & rs!n, 4) & _
                "  nuovi " & Right$(Space(4) & rs!NC, 4) & "  rifatti " & Right$(Space(4) & rs!NS, 4) & _
                IIf(rs!NP > 0, "  DA CONTROLLARE " & rs!NP, "") & vbCrLf
        rs.MoveNext
    Loop
    rs.Close

    '// righe da controllare
    If nProb > 0 Then
        t = t & vbCrLf & "DA CONTROLLARE - non verranno create (" & nProb & " righe, raggruppate per tabella)" & vbCrLf
        Set rs = db.OpenRecordset("SELECT TABELLA, ESITO, Count(*) AS ND, First(FILE_DEST) AS ES, First(DB_MADRE) AS MADRE FROM " & S06_QUERY & sW & _
            " AND Not (ESITO In ('SIMULA_CREATO','CREATO','SIMULA_SOVRASCRITTO','SOVRASCRITTO','DA_INVIARE') Or ESITO Is Null)" & _
            " GROUP BY TABELLA, ESITO ORDER BY ESITO, Count(*) DESC, TABELLA", dbOpenSnapshot)
        Do While Not rs.EOF
            t = t & "  " & Left$(rs!Tabella & Space(36), 36) & " " & Left$(rs!Esito & Space(18), 18) & _
                    IIf(rs!ND = 1, " in " & rs!ES, " in " & rs!ND & " database (es. " & rs!ES & ")")
            If rs!Esito = "CATALOGO_NON_OK" Then t = t & "  - madre fuori radice " & RadiceCorta() & " " & NomeFile(rs!MADRE & "")
            t = t & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
        t = t & "  (dettaglio riga per riga nel report di stampa)" & vbCrLf
    End If

    t = t & vbCrLf & "Log: " & UltimoLog() & vbCrLf
    S06_Testo = t
End Function

Private Function SpiegaEsito(ByVal e As String) As String
    Select Case e
        Case "DA_INVIARE":          SpiegaEsito = "non ancora simulato/inviato"
        Case "SIMULA_CREATO":       SpiegaEsito = "il link verrebbe CREATO"
        Case "SIMULA_SOVRASCRITTO": SpiegaEsito = "il link esiste: verrebbe RICREATO"
        Case "CREATO":              SpiegaEsito = "link creato"
        Case "SOVRASCRITTO":        SpiegaEsito = "link ricreato"
        Case "TABELLA_LOCALE":      SpiegaEsito = "nel database c'e' una tabella vera: non toccata"
        Case "SORGENTE":            SpiegaEsito = "e' il database madre della tabella: niente link"
        Case "NON_A_CATALOGO":      SpiegaEsito = "tabella non presente in Msys_TB99_LINK"
        Case "CATALOGO_NON_OK":     SpiegaEsito = "madre fuori radice " & RadiceCorta() & " o Verifica_s non OK"
        Case "SORGENTE_MANCANTE":   SpiegaEsito = "database madre o tabella madre non trovati"
        Case "FE_MANCANTE":         SpiegaEsito = "database di destinazione non trovato"
        Case "FE_NON_VALIDO":       SpiegaEsito = "destinazione fuori radice " & RadiceCorta() & " o uguale a questo database"
        Case "FE_NON_APRIBILE":     SpiegaEsito = "destinazione bloccata o danneggiata"
        Case "ERRORE":              SpiegaEsito = "errore: vedere la nota"
        Case Else:                  SpiegaEsito = ""
    End Select
End Function

'// percorso dell'ultimo log di invio/verifica nella cartella LOGS
Private Function UltimoLog() As String
    Dim s As String, f As String, ult As String
    On Error Resume Next
    s = CurrentProject.Path & "\LOGS\"
    f = Dir(s & "*TB99_Invio_Link_FrontEnd_Log.txt")
    Do While Len(f) > 0
        If f > ult Then ult = f
        f = Dir()
    Loop
    UltimoLog = IIf(ult = "", "(nessun log)", s & ult)
End Function


'==================================================================================================
' #04# SCELTE (ATTIVA = Scel_b su tutte le righe PATH_DB della coppia destinazione + tabella)
'==================================================================================================
Public Function S06_Firma(ByVal sFE As String) As String
    Dim rs As DAO.Recordset
    If sFE = "(tutti)" Then sFE = ""
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) AS N, Sum(Len(DB_DESTINAZIONE)*3+Len(TABELLA)) AS K, Sum(NRO_OGGETTI) AS O FROM " & S06_QUERY & S06_Where(sFE, True, False), dbOpenSnapshot)
    S06_Firma = sFE & "|" & Nz(rs!n, 0) & "|" & Nz(rs!k, 0) & "|" & Nz(rs!O, 0)
    rs.Close
End Function

Public Sub S06_CambiaScelta(ByVal sDest As String, ByVal sTabella As String, ByVal bValore As Boolean)
    CurrentDb.Execute "UPDATE Msys_TB99_LINK_PATH_DB SET Scel_b = " & IIf(bValore, "True", "False") & _
        " WHERE Database_s & [Database] = " & Txt(sDest) & _
        " AND ID_OBJ IN (SELECT ID_OBJ FROM Msys_TB99_LINK_Object WHERE Name = " & Txt(sTabella) & ")", dbFailOnError
End Sub

Public Function S06_ImpostaTutte(ByVal sModo As String, ByVal sFE As String) As Long
    Dim sSet As String, sW As String, db As DAO.Database
    If sFE = "(tutti)" Then sFE = ""
    Select Case sModo
        Case "TUTTE":   sSet = "True"
        Case "NESSUNA": sSet = "False"
        Case "SOLO_FE"
            '// stessa regola FRONTEND del modulo Mdl05 dell'ambiente
            If S06_Casa() Then
                sSet = "(UCase([Database]) Like 'GE[_]*' Or UCase([Database]) Like 'MSYS[_]ATTIVA[_]GEST*')"
            Else
                sSet = "(UCase([Database]) Like 'GE[_]*' Or UCase([Database]) Like 'GEST[_]*')"
            End If
        Case Else: Exit Function
    End Select
    If sFE <> "" Then sW = " WHERE Database_s & [Database] = " & Txt(sFE)
    Set db = CurrentDb
    db.Execute "UPDATE Msys_TB99_LINK_PATH_DB SET Scel_b = " & sSet & sW, dbFailOnError
    S06_ImpostaTutte = db.RecordsAffected
End Function

Public Function S06_Problemi(ByVal sFE As String) As Long
    Dim rs As DAO.Recordset
    If sFE = "(tutti)" Then sFE = ""
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM " & S06_QUERY & S06_Where(sFE, True, False) & _
        " AND Not (ESITO In ('SIMULA_CREATO','CREATO','SIMULA_SOVRASCRITTO','SOVRASCRITTO','DA_INVIARE') Or ESITO Is Null)", dbOpenSnapshot)
    S06_Problemi = rs.Fields(0).Value
    rs.Close
End Function

'// filtro comune: destinazione, solo attive, solo da controllare
Public Function S06_Where(ByVal sFE As String, ByVal bSoloAttive As Boolean, ByVal bSoloProblemi As Boolean) As String
    Dim w As String
    If sFE = "(tutti)" Then sFE = ""
    w = " WHERE 1=1"
    If sFE <> "" Then w = w & " AND DB_DESTINAZIONE = " & Txt(sFE)
    If bSoloAttive Then w = w & " AND ATTIVA = True"
    If bSoloProblemi Then w = w & " AND Not (ESITO In ('SIMULA_CREATO','CREATO','SIMULA_SOVRASCRITTO','SOVRASCRITTO','DA_INVIARE') Or ESITO Is Null)"
    S06_Where = w
End Function


'==================================================================================================
' #05# AIUTO
'==================================================================================================
Public Function S06_Aiuto() As String
    Dim t As String
    t = "=== AIUTO - INVIO DEI LINK DEL CATALOGO TB99 AI DATABASE ===" & vbCrLf & vbCrLf
    t = t & "COSA FA" & vbCrLf
    t = t & "  Crea nei DATABASE DI DESTINAZIONE (colonna Destinazione) le tabelle COLLEGATE che servono" & vbCrLf
    t = t & "  agli oggetti importati (query, form, report). Il link punta al DATABASE MADRE della tabella" & vbCrLf
    t = t & "  (colonna Database madre, dal catalogo Msys_TB99_LINK). In questo database si" & vbCrLf
    t = t & "  scrivono solo gli esiti. Le tabelle vere non vengono mai copiate." & vbCrLf & vbCrLf
    t = t & "COME SI FA" & vbCrLf
    t = t & "  1) Destinazione: scegliere un database oppure (tutti)." & vbCrLf
    t = t & "  2) Nell'elenco cliccare la casella ATTIVA per includere/escludere una riga." & vbCrLf
    t = t & "     Pulsanti rapidi: Attiva tutte, Disattiva tutte, Solo front end GE_ (consigliato)." & vbCrLf
    t = t & "  3) Premere 1) SIMULA: nessun database viene modificato; ogni riga riceve un esito SIMULA_*" & vbCrLf
    t = t & "     e qui compare il riepilogo. Il pulsante STAMPA prepara il report da stampare." & vbCrLf
    t = t & "  4) Controllare le righe in rosso (DA CONTROLLARE): non verranno create." & vbCrLf
    t = t & "  5) Premere 2) CREA I LINK: si attiva solo dopo una simulazione fatta sulle STESSE scelte." & vbCrLf
    t = t & "     Se si cambia una scelta, la simulazione va rifatta." & vbCrLf & vbCrLf
    t = t & "REGOLE (modulo Mdl05_INVIA_LINK_FRONTEND)" & vbCrLf
    t = t & "  - un link gia' presente viene cancellato e ricreato (SOVRASCRITTO)" & vbCrLf
    t = t & "  - una tabella VERA con lo stesso nome non viene toccata (TABELLA_LOCALE)" & vbCrLf
    t = t & "  - nel database madre della tabella non si crea il link (SORGENTE)" & vbCrLf
    t = t & "  - si lavora solo dentro la radice " & RadiceCorta() & vbCrLf & vbCrLf
    t = t & "ESITI" & vbCrLf
    t = t & "  SIMULA_CREATO / SIMULA_SOVRASCRITTO  = cosa succederebbe (simulazione)" & vbCrLf
    t = t & "  CREATO / SOVRASCRITTO                = fatto" & vbCrLf
    t = t & "  TABELLA_LOCALE, SORGENTE, FE_MANCANTE, FE_NON_APRIBILE, SORGENTE_MANCANTE," & vbCrLf
    t = t & "  NON_A_CATALOGO, CATALOGO_NON_OK, ERRORE = da controllare (vedere la nota)" & vbCrLf & vbCrLf
    t = t & "LOG: cartella LOGS accanto a questo database." & vbCrLf
    S06_Aiuto = t
End Function


'==================================================================================================
' #07# AIUTO DEI PULSANTI: punto "?" sotto ogni pulsante -> aiuto breve + pagina HTML completa
'      pagine in  OBJECT\HTML\FORM\Msys_TB99_LINK_FrmS06_INVIA_LINK\Msys_TB99_LINK_FrmS06_INVIA_LINK_HELP_<nn>_<CHIAVE>.html
'      (si modificano i file .md e si rigenerano gli html con OBJECT\HTML\GENERA_HELP_HTML.bat)
'==================================================================================================
'// cartella degli help: prima accanto al database (modulo Mdl07_HELP), poi il percorso fisso su H:
Public Function S06_HelpCartella() As String
    Dim p As String
    On Error Resume Next
    p = HELP_Cartella_pf("FORM", S06_FORM)
    If Len(Dir(p & "*.html")) = 0 Then p = HELP_ASSOLUTA
    S06_HelpCartella = p
End Function

Private Function HelpFile(ByVal sChiave As String) As String
    Select Case sChiave
        Case "FILTRI":       HelpFile = "01_FILTRI"
        Case "ATTIVA_TUTTE": HelpFile = "02_ATTIVA_TUTTE"
        Case "DISATTIVA":    HelpFile = "03_DISATTIVA"
        Case "SOLO_GE":      HelpFile = "04_SOLO_GE"
        Case "SIMULA":       HelpFile = "05_SIMULA"
        Case "INVIA":        HelpFile = "06_CREA_LINK"
        Case "STAMPA":       HelpFile = "07_STAMPA"
        Case "TESTO":        HelpFile = "08_RIEPILOGO"
        Case "AIUTO":        HelpFile = "00_INDICE"
        Case "ESITI":        HelpFile = "10_ESITI"
        Case Else:           HelpFile = "00_INDICE"
    End Select
    HelpFile = S06_HelpCartella() & S06_FORM & "_HELP_" & HelpFile & ".html"
End Function

Private Function HelpBreve(ByVal sChiave As String, ByRef sTitolo As String) As String
    Dim t As String
    Select Case sChiave
        Case "FILTRI"
            sTitolo = "Destinazione e filtri"
            t = "DESTINAZIONE: il database su cui lavorare, oppure (tutti)." & vbCrLf & _
                "Vale per elenco, pulsanti, simulazione, creazione, riepilogo e report." & vbCrLf & vbCrLf & _
                "solo righe attive: mostra solo le righe con Attiva = si." & vbCrLf & _
                "solo righe da controllare: mostra solo gli esiti in rosso." & vbCrLf & _
                "I filtri cambiano solo cio' che si vede, non le scelte." & vbCrLf & vbCrLf & _
                "Cambiare destinazione spegne 2) CREA I LINK."
        Case "ATTIVA_TUTTE"
            sTitolo = "Attiva tutte"
            t = "Mette Attiva = si su TUTTE le righe della destinazione scelta" & vbCrLf & _
                "(anche quelle nascoste dai filtri)." & vbCrLf & vbCrLf & _
                "Scrive solo la scelta (Scel_b) in Msys_TB99_LINK_PATH_DB:" & vbCrLf & _
                "nessun database di destinazione viene toccato." & vbCrLf & vbCrLf & _
                "Con (tutti) e' meglio usare 'Solo front end GE_'."
        Case "DISATTIVA"
            sTitolo = "Disattiva tutte"
            t = "Mette Attiva = no su TUTTE le righe della destinazione scelta." & vbCrLf & vbCrLf & _
                "Serve per ripartire da zero e attivare a mano solo le righe" & vbCrLf & _
                "che servono (clic sulla casella Attiva)." & vbCrLf & vbCrLf & _
                "Nessun database di destinazione viene toccato."
        Case "SOLO_GE"
            sTitolo = "Solo front end GE_ (consigliato)"
            t = "Attiva solo le righe dei database che iniziano con " & FrontEndTesto() & vbCrLf & _
                "(i front end) e disattiva tutte le altre." & vbCrLf & vbCrLf & _
                "E' la scelta consigliata: i link servono ai front end." & vbCrLf & vbCrLf & _
                "Nessun database di destinazione viene toccato."
        Case "SIMULA"
            sTitolo = "1) SIMULA"
            t = "Prova l'invio sulle righe attive SENZA MODIFICARE i database." & vbCrLf & vbCrLf & _
                "Ogni riga riceve un esito:" & vbCrLf & _
                "  SIMULA_CREATO = il link verrebbe creato" & vbCrLf & _
                "  SIMULA_SOVRASCRITTO = il link c'e' e verrebbe rifatto" & vbCrLf & _
                "  in rosso = da controllare, verra' saltata" & vbCrLf & vbCrLf & _
                "Il riepilogo compare nella casella a destra." & vbCrLf & _
                "Dopo la simulazione si accende 2) CREA I LINK."
        Case "INVIA"
            sTitolo = "2) CREA I LINK"
            t = "CREA DAVVERO i link nei DATABASE DI DESTINAZIONE." & vbCrLf & _
                "Ogni link punta al database madre della tabella." & vbCrLf & vbCrLf & _
                "Si accende solo dopo 1) SIMULA fatta sulle stesse scelte;" & vbCrLf & _
                "chiede conferma prima di partire." & vbCrLf & vbCrLf & _
                "I link esistenti vengono ricreati; le tabelle vere non vengono" & vbCrLf & _
                "toccate; le righe in rosso vengono saltate."
        Case "STAMPA"
            sTitolo = "Stampa report"
            t = "Apre in anteprima il report completo, con le righe che si" & vbCrLf & _
                "vedono nell'elenco (stessa destinazione e stessi filtri)." & vbCrLf & vbCrLf & _
                "Prima pagina: riepilogo. Poi un gruppo per database di" & vbCrLf & _
                "destinazione con tabella, database madre, esito e nota." & vbCrLf & vbCrLf & _
                "E' il documento da conservare (si puo' salvare in PDF)."
        Case "TESTO"
            sTitolo = "Riepilogo"
            t = "Scrive nella casella a destra la situazione attuale delle" & vbCrLf & _
                "righe attive: totali per esito, una riga per database," & vbCrLf & _
                "righe da controllare raggruppate per tabella, ultimo log." & vbCrLf & vbCrLf & _
                "Non simula e non modifica nulla."
        Case "AIUTO"
            sTitolo = "? Aiuto"
            t = "Scrive nella casella a destra la guida breve della form." & vbCrLf & vbCrLf & _
                "Il punto ? sotto ogni pulsante spiega quel pulsante;" & vbCrLf & _
                "la guida completa e' in pagine HTML (OBJECT\HTML\FORM\" & S06_FORM & ")."
        Case Else
            sTitolo = "Aiuto"
            t = "Guida della form di invio dei link."
    End Select
    HelpBreve = t
End Function

'// chiamata dai punti "?" della form
Public Sub S06_Help(ByVal sChiave As String)
    Dim sTit As String, t As String, f As String
    t = HelpBreve(sChiave, sTit)
    f = HelpFile(sChiave)
    If Len(Dir(f)) > 0 Then
        If MsgBox(t & vbCrLf & vbCrLf & "Aprire la guida completa (pagina HTML)?", vbQuestion + vbYesNo + vbDefaultButton2, "Aiuto - " & sTit) = vbYes Then
            S06_ApriHtml f
        End If
    Else
        MsgBox t & vbCrLf & vbCrLf & "(guida HTML non trovata: " & f & ")", vbInformation, "Aiuto - " & sTit
    End If
End Sub

Public Sub S06_ApriHtml(ByVal f As String)
    On Error Resume Next
    Shell "explorer.exe """ & f & """", vbNormalFocus
    If Err.Number <> 0 Then
        Err.Clear
        Application.FollowHyperlink f
    End If
    If Err.Number <> 0 Then MsgBox "Impossibile aprire la guida: " & f, vbExclamation, "Aiuto"
End Sub


'==================================================================================================
' #06# UTILITA
'==================================================================================================
'// ambiente: versione di casa (C:\CASA\...\MSYS) oppure ufficio (albero H:)
Public Function S06_Casa() As Boolean
    S06_Casa = (InStr(1, CurrentProject.Path, "C:\CASA\", vbTextCompare) = 1)
End Function

Private Function RadiceCorta() As String
    RadiceCorta = IIf(S06_Casa(), "C:\CASA\...\PROGETTI_MDB\MSYS", "H:\...\ARCHIVI_ACCESS_LLPP_H")
End Function

Private Function FrontEndTesto() As String
    FrontEndTesto = IIf(S06_Casa(), "GE_ o MSYS_ATTIVA_GEST", "GE_ o GEST_")
End Function

Private Function NomeFile(ByVal p As String) As String
    NomeFile = Mid$(p, InStrRev(p, "\") + 1)
End Function

Private Function Txt(ByVal s As String) As String
    Txt = "'" & Replace(s, "'", "''") & "'"
End Function