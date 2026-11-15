Attribute VB_Name = "Msys_TB99_LINK_ Mdl05_INVIA_LINK_FRONTEND"
Option Compare Database
Option Explicit

'==================================================================================================
' INTESTAZIONE
'==================================================================================================
' MODULO     : Msys_TB99_LINK_ Mdl05_INVIA_LINK_FRONTEND
' PROGETTO   : MSys  -  mini-progetto TB99_LINK  (codice @SHADOW_SYSTEM_CATALOG)
' DB         : C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_TB99_LINK\MDB\MSys_TB99_LINK.mdb
' PATH_LOG   : <cartella del database>\LOGS\AAAAMMGG_HHMMSS_TB99_Invio_Link_FrontEnd_Log.txt
' VERSIONE   : 1.0  del 15/11/2026
' FAQ        : @TB99@INVIA@LINK_(invia i collegamenti del catalogo TB99 ai database front end)
'
' SCOPO
'   MSys_TB99_LINK.mdb raccoglie TUTTI i collegamenti del progetto MSys (tabella Msys_TB99_LINK).
'   Questo modulo INVIA i collegamenti previsti a ogni database front end, sovrascrivendo
'   quelli esistenti. Esempio: GE_MSys_DF01_COMANDI.mdb riceve i link alle tabelle di MSys_DF01.mdb.
'
' REGOLE
'   R1  Nel front end si creano SOLO collegamenti: le tabelle originali non si copiano mai.
'   R2  Un link gia presente nel front end viene cancellato e ricreato (sovrascrittura).
'   R3  Una TABELLA LOCALE con lo stesso nome nel front end NON viene mai toccata:
'       esito TABELLA_LOCALE, da sistemare a mano.
'   R4  Si inviano solo link del catalogo con Verifica_s = OK e database dentro RADICE_MSYS.
'   R5  Il front end deve stare dentro RADICE_MSYS e non puo essere il database sorgente
'       della tabella (esito SORGENTE) ne questo database (MSys_TB99_LINK.mdb).
'
' TABELLE USATE
'   Msys_TB99_LINK            catalogo: Name, Database (sorgente), Verifica_s
'   Msys_TB99_LINK_Object     oggetti che usano la tabella (legame Name con Msys_TB99_LINK)
'   Msys_TB99_LINK_PATH_DB    destinazioni dell'oggetto (legame ID_OBJ con Msys_TB99_LINK_Object):
'                             il link da inviare e la tabella (Name) degli oggetti con destinazione attiva
'                             Scel_b = True -> riga attiva; Esito_s, Nota_s, DATA_INVIO_d = ultimo invio
'   Query di controllo        Msys_TB99_LINK_Qry99_03_SELECT_PATH_DB
'   NOTA 15/11/2026: catena Msys_TB99_LINK -> Msys_TB99_LINK_Object -> Msys_TB99_LINK_PATH_DB;
'                    i percorsi sono divisi in Database_s (cartella) + Database (file);
'                    l'esito dell'invio va su tutte le righe PATH_DB della coppia tabella/destinazione
'
' FUNZIONI PUBBLICHE
'   POPOLA_LINK_FRONTEND_pf   censisce i link previsti leggendo i link gia presenti nei database
'   INVIA_LINK_FRONTEND_pf    invia (o simula) i link previsti ai front end
'   VERIFICA_LINK_FRONTEND_pf come INVIA ma in simulazione: non modifica nessun database
'
' ESITI (campo Esito_s)
'   CREATO              link nuovo creato nel front end
'   SOVRASCRITTO        link esistente cancellato e ricreato
'   SIMULA_CREATO / SIMULA_SOVRASCRITTO   stesso esito in simulazione
'   TABELLA_LOCALE      nel front end c'e una tabella vera con quel nome (R3)
'   SORGENTE            il front end e il database sorgente della tabella (R5)
'   NON_A_CATALOGO      il nome non esiste in Msys_TB99_LINK
'   CATALOGO_NON_OK     link del catalogo con Verifica_s diversa da OK o fuori RADICE_MSYS (R4)
'   SORGENTE_MANCANTE   database sorgente o tabella sorgente non trovati
'   FE_MANCANTE         file front end non trovato
'   FE_NON_VALIDO       front end fuori RADICE_MSYS o uguale a questo database (R5)
'   FE_NON_APRIBILE     front end non apribile (bloccato o danneggiato)
'   ERRORE              errore durante la creazione del link (dettaglio in Nota_s e nel log)
'
' USO
'   ? POPOLA_LINK_FRONTEND_pf()                        ' una volta, poi controllare la tabella
'   ? VERIFICA_LINK_FRONTEND_pf()                      ' prova senza modifiche
'   ? INVIA_LINK_FRONTEND_pf()                         ' invio a tutti i front end attivi
'   ? INVIA_LINK_FRONTEND_pf("C:\...\GE_MSys_DF01_COMANDI.mdb")   ' un solo front end
'   Parametro bSilenzioso = True: nessun MsgBox (per automazione e test).
'
'==================================================================================================
' INDICE  (ogni etichetta si trova nel codice con Ctrl+F)
'==================================================================================================
' #01#        COSTANTI E VARIABILI DI MODULO
' #02#        INVIA_LINK_FRONTEND_pf
'   #02.01#     preparazione: log, controllo tabelle, conferma utente
'   #02.02#     apertura elenco assegnazioni attive
'   #02.03#     ciclo sulle assegnazioni
'     #02.03.01#  cambio front end: chiusura precedente, controlli, apertura
'     #02.03.02#  controlli sul link del catalogo
'     #02.03.03#  invio del link
'     #02.03.04#  registrazione esito
'   #02.04#     chiusura, statistiche, messaggio finale
'   #02.99#     gestione errori
' #03#        VERIFICA_LINK_FRONTEND_pf
' #04#        POPOLA_LINK_FRONTEND_pf
'   #04.01#     preparazione: log, nomi del catalogo
'   #04.02#     scansione dei database in RADICE_MSYS
'   #04.03#     lettura dei link di un database e inserimento assegnazioni
'   #04.04#     chiusura e messaggio finale
'   #04.99#     gestione errori
' #05#        INVIO DI UN SINGOLO LINK  (InviaUnLink)
'   #05.01#     ricerca della tabella nel front end
'   #05.02#     tabella locale: stop (R3)
'   #05.03#     cancellazione del link esistente (R2)
'   #05.04#     creazione del nuovo link
' #06#        CONTROLLI SU FILE E DATABASE
'   #06.01#     TabellaInSorgente   (con cache per database)
'   #06.02#     PercorsoInRadice
'   #06.03#     PercorsoEscluso
'   #06.04#     TabellaEsiste
' #07#        LOG
'   #07.01#     ApriLog
'   #07.02#     ScriviLog
'   #07.03#     ChiudiLog
' #08#        UTILITA
'   #08.01#     SqlTesto
'   #08.02#     AggiornaEsito
'   #08.03#     ElencaDatabase   (ricorsiva)
'==================================================================================================


'==================================================================================================
' #01# COSTANTI E VARIABILI DI MODULO
'==================================================================================================
Private Const NOME_MODULO As String = "Msys_TB99_LINK_ Mdl05_INVIA_LINK_FRONTEND"
Private Const RADICE_MSYS As String = "C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\"
Private Const TAB_CATALOGO As String = "Msys_TB99_LINK"
Private Const TAB_FRONTEND As String = "Msys_TB99_LINK_PATH_DB"
Private Const TAB_OGGETTI As String = "Msys_TB99_LINK_Object"
Private Const LOG_NOME As String = "TB99_Invio_Link_FrontEnd_Log.txt"

Private m_fso As Object             ' Scripting.FileSystemObject
Private m_log As Object             ' TextStream del log
Private m_logPath As String         ' percorso del file di log
Private m_pos As String             ' posizione corrente nel codice (per gli errori)
Private m_cacheSorg As Object       ' Dictionary: percorso DB sorgente -> Dictionary delle tabelle


'==================================================================================================
' #02# INVIA_LINK_FRONTEND_pf
'      Invia ai front end i link previsti in Msys_TB99_LINK_PATH_DB (righe con Scel_b = True).
'      sFrontEnd    = "" tutti i front end, altrimenti solo quello indicato (percorso completo)
'      bSimulazione = True: controlla e scrive gli esiti SIMULA_*, senza modificare i front end
'      bSilenzioso  = True: nessun MsgBox
'      Restituisce il riepilogo "TOT=..;CREATO=..;SOVRASCRITTO=..;..."
'==================================================================================================
Public Function INVIA_LINK_FRONTEND_pf(Optional ByVal sFrontEnd As String = "", _
                                       Optional ByVal bSimulazione As Boolean = False, _
                                       Optional ByVal bSilenzioso As Boolean = False) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim dbFE As DAO.Database
    Dim sSql As String, sFE As String, sFECorr As String, sName As String
    Dim sSorg As String, sVerifica As String, sEsito As String, sNota As String
    Dim nTot As Long
    Dim bFEValido As Boolean, sEsitoFE As String
    Dim stat As Object, k As Variant, sRiep As String

    On Error GoTo GestErrore

    '----------------------------------------------------------------------------------------------
    ' #02.01# preparazione: log, controllo tabelle, conferma utente
    '----------------------------------------------------------------------------------------------
    m_pos = "02.01"
    Set db = CurrentDb
    Set stat = CreateObject("Scripting.Dictionary")
    Set m_cacheSorg = CreateObject("Scripting.Dictionary")
    m_cacheSorg.CompareMode = 1
    ApriLog IIf(bSimulazione, "VERIFICA (simulazione) INVIO LINK AI FRONT END", "INVIO LINK AI FRONT END")
    ScriviLog "02.01", "INFO", "Filtro front end: " & IIf(Len(sFrontEnd) = 0, "(tutti)", sFrontEnd)

    If Not TabellaEsiste(db, TAB_FRONTEND) Or Not TabellaEsiste(db, TAB_CATALOGO) Then
        ScriviLog "02.01", "ERRORE", "Manca la tabella " & TAB_FRONTEND & " o " & TAB_CATALOGO
        ChiudiLog
        If Not bSilenzioso Then MsgBox "Manca la tabella " & TAB_FRONTEND & " o " & TAB_CATALOGO, vbCritical, NOME_MODULO
        INVIA_LINK_FRONTEND_pf = "ERRORE=TABELLE_MANCANTI"
        Exit Function
    End If

    If Not bSilenzioso And Not bSimulazione Then
        If MsgBox("Invio dei collegamenti del catalogo TB99 ai front end." & vbCrLf & vbCrLf & _
                  "I link esistenti vengono SOVRASCRITTI; le tabelle locali non vengono toccate." & vbCrLf & _
                  "Front end: " & IIf(Len(sFrontEnd) = 0, "tutti quelli attivi", sFrontEnd) & vbCrLf & vbCrLf & _
                  "Procedere?", vbYesNo + vbQuestion, NOME_MODULO) = vbNo Then
            ScriviLog "02.01", "INFO", "Operazione annullata dall'utente"
            ChiudiLog
            INVIA_LINK_FRONTEND_pf = "ANNULLATO"
            Exit Function
        End If
    End If

    '----------------------------------------------------------------------------------------------
    ' #02.02# apertura elenco assegnazioni attive
    '----------------------------------------------------------------------------------------------
    m_pos = "02.02"
    '// una riga per coppia tabella/destinazione: PATH_DB -> Object (ID_OBJ) -> catalogo (Name)
    sSql = "SELECT F.Database_s & F.[Database] AS FRONTEND_s, O.Name, C.Database_s & C.[Database] AS DB_SORG, " & _
           "C.Verifica_s AS VERIFICA, C.Name AS NOME_CAT " & _
           "FROM (" & TAB_FRONTEND & " AS F INNER JOIN " & TAB_OGGETTI & " AS O ON F.ID_OBJ = O.ID_OBJ) " & _
           "LEFT JOIN " & TAB_CATALOGO & " AS C ON O.Name = C.Name WHERE F.Scel_b = True"
    If Len(sFrontEnd) > 0 Then sSql = sSql & " AND F.Database_s & F.[Database] = " & SqlTesto(sFrontEnd)
    sSql = sSql & " GROUP BY F.Database_s & F.[Database], O.Name, C.Database_s & C.[Database], C.Verifica_s, C.Name" & _
                  " ORDER BY F.Database_s & F.[Database], O.Name"
    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    ScriviLog "02.02", "INFO", "Assegnazioni da elaborare aperte"

    '----------------------------------------------------------------------------------------------
    ' #02.03# ciclo sulle assegnazioni
    '----------------------------------------------------------------------------------------------
    sFECorr = vbNullString
    Do While Not rs.EOF
        nTot = nTot + 1
        sFE = Nz(rs!FRONTEND_s, "")
        sName = Nz(rs!Name, "")
        sSorg = Nz(rs!DB_SORG, "")
        sVerifica = Nz(rs!VERIFICA, "")
        sEsito = vbNullString: sNota = vbNullString

        '------------------------------------------------------------------------------------------
        ' #02.03.01# cambio front end: chiusura precedente, controlli, apertura
        '------------------------------------------------------------------------------------------
        m_pos = "02.03.01"
        If StrComp(sFE, sFECorr, vbTextCompare) <> 0 Then
            If Not dbFE Is Nothing Then dbFE.Close: Set dbFE = Nothing
            sFECorr = sFE
            bFEValido = False: sEsitoFE = vbNullString
            ScriviLog "02.03.01", "INFO", "=== FRONT END: " & sFE
            If Not m_fso.FileExists(sFE) Then
                sEsitoFE = "FE_MANCANTE"
            ElseIf Not PercorsoInRadice(sFE) Or StrComp(sFE, db.Name, vbTextCompare) = 0 Then
                sEsitoFE = "FE_NON_VALIDO"
            Else
                On Error Resume Next
                Set dbFE = DBEngine.OpenDatabase(sFE, False, bSimulazione)
                If Err.Number <> 0 Then
                    sEsitoFE = "FE_NON_APRIBILE"
                    ScriviLog "02.03.01", "ERRORE", "Apertura non riuscita (" & Err.Number & "): " & Err.Description
                    Err.Clear
                    Set dbFE = Nothing
                Else
                    bFEValido = True
                End If
                On Error GoTo GestErrore
            End If
            If Not bFEValido Then ScriviLog "02.03.01", "AVVISO", "Front end saltato: " & sEsitoFE
        End If

        '------------------------------------------------------------------------------------------
        ' #02.03.02# controlli sul link del catalogo
        '------------------------------------------------------------------------------------------
        m_pos = "02.03.02"
        If Not bFEValido Then
            sEsito = sEsitoFE
        ElseIf IsNull(rs!NOME_CAT) Then
            sEsito = "NON_A_CATALOGO"
        ElseIf sVerifica <> "OK" Or Not PercorsoInRadice(sSorg) Then
            sEsito = "CATALOGO_NON_OK": sNota = "Verifica_s=" & sVerifica & " Database=" & sSorg
        ElseIf StrComp(sFE, sSorg, vbTextCompare) = 0 Then
            sEsito = "SORGENTE"
        ElseIf Not TabellaInSorgente(sSorg, sName) Then
            sEsito = "SORGENTE_MANCANTE": sNota = sSorg
        Else
        '------------------------------------------------------------------------------------------
        ' #02.03.03# invio del link
        '------------------------------------------------------------------------------------------
            m_pos = "02.03.03"
            sEsito = InviaUnLink(dbFE, sName, sSorg, bSimulazione, sNota)
        End If

        '------------------------------------------------------------------------------------------
        ' #02.03.04# registrazione esito
        '------------------------------------------------------------------------------------------
        m_pos = "02.03.04"
        AggiornaEsito db, sFE, sName, sEsito, sNota
        stat(sEsito) = stat(sEsito) + 1
        ScriviLog "02.03.04", IIf(sEsito Like "*CREATO" Or sEsito Like "*SOVRASCRITTO", "OK", "AVVISO"), _
                  sName & " -> " & sEsito & IIf(Len(sNota) > 0, " (" & sNota & ")", "")
        rs.MoveNext
    Loop

    '----------------------------------------------------------------------------------------------
    ' #02.04# chiusura, statistiche, messaggio finale
    '----------------------------------------------------------------------------------------------
    m_pos = "02.04"
    If Not dbFE Is Nothing Then dbFE.Close: Set dbFE = Nothing
    rs.Close: Set rs = Nothing
    sRiep = "TOT=" & nTot
    For Each k In stat.Keys
        sRiep = sRiep & ";" & k & "=" & stat(k)
    Next k
    ScriviLog "02.04", "FINE", "Riepilogo: " & sRiep
    ChiudiLog
    If Not bSilenzioso Then
        MsgBox IIf(bSimulazione, "Verifica completata (nessuna modifica).", "Invio completato.") & vbCrLf & vbCrLf & _
               Replace(sRiep, ";", vbCrLf) & vbCrLf & vbCrLf & "Log: " & m_logPath, vbInformation, NOME_MODULO
    End If
    INVIA_LINK_FRONTEND_pf = sRiep
    Exit Function

    '----------------------------------------------------------------------------------------------
    ' #02.99# gestione errori
    '----------------------------------------------------------------------------------------------
GestErrore:
    ScriviLog m_pos, "ERRORE", "codice fermo a #" & m_pos & "# - errore " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not dbFE Is Nothing Then dbFE.Close
    If Not rs Is Nothing Then rs.Close
    ChiudiLog
    If Not bSilenzioso Then MsgBox "Errore " & Err.Number & " a #" & m_pos & "#: " & Err.Description, vbCritical, NOME_MODULO
    INVIA_LINK_FRONTEND_pf = "ERRORE=" & m_pos
End Function


'==================================================================================================
' #03# VERIFICA_LINK_FRONTEND_pf
'      Come INVIA_LINK_FRONTEND_pf ma in simulazione: apre i front end in sola lettura,
'      scrive gli esiti SIMULA_* e non modifica nulla.
'==================================================================================================
Public Function VERIFICA_LINK_FRONTEND_pf(Optional ByVal sFrontEnd As String = "", _
                                          Optional ByVal bSilenzioso As Boolean = False) As String
    VERIFICA_LINK_FRONTEND_pf = INVIA_LINK_FRONTEND_pf(sFrontEnd, True, bSilenzioso)
End Function


'==================================================================================================
' #04# POPOLA_LINK_FRONTEND_pf
'      Censisce i link previsti: legge ogni database in RADICE_MSYS e, per ogni link gia presente
'      il cui nome esiste nel catalogo, aggancia la destinazione a TUTTI gli oggetti della tabella
'      (Msys_TB99_LINK_PATH_DB.ID_OBJ). Gli oggetti di una tabella sono quelli che la usano e che
'      stanno nel suo database sorgente (es. le query di MSys_DF01_COMANDI in MSys_DF01.mdb).
'      Una tabella senza oggetti (per esempio un separatore) non riceve destinazioni.
'      Scel_b = True solo per i front-end (GE_* e MSYS_ATTIVA_GEST*); gli altri si attivano a mano.
'      Restituisce il numero di righe inserite.
'==================================================================================================
Public Function POPOLA_LINK_FRONTEND_pf(Optional ByVal bSilenzioso As Boolean = False) As Long
    Dim db As DAO.Database
    Dim dbX As DAO.Database
    Dim td As DAO.TableDef
    Dim rs As DAO.Recordset
    Dim catalogo As Object, esistenti As Object, elenco As Collection
    Dim f As Variant, sTipo As String, sChiave As String, nIns As Long, nDb As Long
    Dim rsO As DAO.Recordset

    On Error GoTo GestErrore

    '----------------------------------------------------------------------------------------------
    ' #04.01# preparazione: log, nomi del catalogo
    '----------------------------------------------------------------------------------------------
    m_pos = "04.01"
    Set db = CurrentDb
    ApriLog "POPOLA ASSEGNAZIONI LINK FRONT END"
    Set catalogo = CreateObject("Scripting.Dictionary"): catalogo.CompareMode = 1
    Set esistenti = CreateObject("Scripting.Dictionary"): esistenti.CompareMode = 1
    Set rs = db.OpenRecordset("SELECT Name FROM " & TAB_CATALOGO, dbOpenSnapshot)
    Do While Not rs.EOF
        catalogo(rs!Name & "") = True
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT Database_s & [Database] AS FRONTEND_s, ID_OBJ FROM " & TAB_FRONTEND, dbOpenSnapshot)
    Do While Not rs.EOF
        esistenti(rs!FRONTEND_s & "|" & rs!ID_OBJ) = True
        rs.MoveNext
    Loop
    rs.Close
    ScriviLog "04.01", "INFO", "Nomi a catalogo: " & catalogo.Count & " - assegnazioni gia presenti: " & esistenti.Count

    '----------------------------------------------------------------------------------------------
    ' #04.02# scansione dei database in RADICE_MSYS
    '----------------------------------------------------------------------------------------------
    m_pos = "04.02"
    Set elenco = New Collection
    ElencaDatabase m_fso.GetFolder(RADICE_MSYS), elenco
    ScriviLog "04.02", "INFO", "Database trovati: " & elenco.Count

    For Each f In elenco
        If StrComp(f, db.Name, vbTextCompare) <> 0 Then
        '------------------------------------------------------------------------------------------
        ' #04.03# lettura dei link di un database e inserimento assegnazioni
        '------------------------------------------------------------------------------------------
            m_pos = "04.03"
            On Error Resume Next
            Set dbX = DBEngine.OpenDatabase(f, False, True)
            If Err.Number <> 0 Then
                ScriviLog "04.03", "AVVISO", "Non apribile: " & f & " (" & Err.Description & ")"
                Err.Clear: Set dbX = Nothing
            End If
            On Error GoTo GestErrore
            If Not dbX Is Nothing Then
                nDb = nDb + 1
                sTipo = IIf(UCase$(m_fso.GetFileName(f)) Like "GE_*" Or UCase$(m_fso.GetFileName(f)) Like "MSYS_ATTIVA_GEST*", "FRONTEND", "BACKEND")
                For Each td In dbX.TableDefs
                    If Len(td.Connect) > 0 And catalogo.Exists(td.Name) Then
                        '// tutti gli oggetti della tabella (stanno nel suo database sorgente)
                        Set rsO = db.OpenRecordset("SELECT ID_OBJ FROM " & TAB_OGGETTI & " WHERE Name = " & SqlTesto(td.Name), dbOpenSnapshot)
                        If rsO.EOF Then ScriviLog "04.03", "AVVISO", td.Name & ": nessun oggetto a catalogo, destinazione non creata"
                        Do While Not rsO.EOF
                            sChiave = f & "|" & rsO!ID_OBJ
                            If Not esistenti.Exists(sChiave) Then
                                db.Execute "INSERT INTO " & TAB_FRONTEND & " (ID_OBJ, [Database], Database_s, Scel_b, Esito_s, DATAINS_d) VALUES (" & _
                                           rsO!ID_OBJ & ", " & SqlTesto(m_fso.GetFileName(f)) & ", " & _
                                           SqlTesto(m_fso.GetParentFolderName(f) & "\") & ", " & _
                                           IIf(sTipo = "FRONTEND", "True", "False") & ", 'DA_INVIARE', Now())", dbFailOnError
                                esistenti(sChiave) = True
                                nIns = nIns + 1
                            End If
                            rsO.MoveNext
                        Loop
                        rsO.Close: Set rsO = Nothing
                        ScriviLog "04.03", "OK", sTipo & " " & f & " <- " & td.Name
                    End If
                Next td
                dbX.Close: Set dbX = Nothing
            End If
        End If
    Next f

    '----------------------------------------------------------------------------------------------
    ' #04.04# chiusura e messaggio finale
    '----------------------------------------------------------------------------------------------
    m_pos = "04.04"
    ScriviLog "04.04", "FINE", "Database letti: " & nDb & " - assegnazioni inserite: " & nIns
    ChiudiLog
    If Not bSilenzioso Then MsgBox "Database letti: " & nDb & vbCrLf & "Assegnazioni inserite: " & nIns & vbCrLf & vbCrLf & _
                                   "Controllare la query Msys_TB99_LINK_Qry99_03_SELECT_PATH_DB.", vbInformation, NOME_MODULO
    POPOLA_LINK_FRONTEND_pf = nIns
    Exit Function

    '----------------------------------------------------------------------------------------------
    ' #04.99# gestione errori
    '----------------------------------------------------------------------------------------------
GestErrore:
    ScriviLog m_pos, "ERRORE", "codice fermo a #" & m_pos & "# - errore " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not dbX Is Nothing Then dbX.Close
    If Not rs Is Nothing Then rs.Close
    ChiudiLog
    If Not bSilenzioso Then MsgBox "Errore " & Err.Number & " a #" & m_pos & "#: " & Err.Description, vbCritical, NOME_MODULO
    POPOLA_LINK_FRONTEND_pf = -1
End Function


'==================================================================================================
' #05# INVIO DI UN SINGOLO LINK  (InviaUnLink)
'      Restituisce l'esito; in sNota l'eventuale dettaglio dell'errore.
'==================================================================================================
Private Function InviaUnLink(ByVal dbFE As DAO.Database, ByVal sName As String, ByVal sSorg As String, _
                             ByVal bSimulazione As Boolean, ByRef sNota As String) As String
    Dim td As DAO.TableDef
    Dim sAzione As String

    '----------------------------------------------------------------------------------------------
    ' #05.01# ricerca della tabella nel front end
    '----------------------------------------------------------------------------------------------
    m_pos = "05.01"
    On Error Resume Next
    Set td = dbFE.TableDefs(sName)
    Err.Clear
    On Error GoTo GestErrore

    '----------------------------------------------------------------------------------------------
    ' #05.02# tabella locale: stop (R3)
    '----------------------------------------------------------------------------------------------
    m_pos = "05.02"
    If td Is Nothing Then
        sAzione = "CREATO"
    ElseIf Len(td.Connect) = 0 Then
        InviaUnLink = "TABELLA_LOCALE"
        sNota = "tabella vera nel front end: non toccata"
        Exit Function
    Else
        sAzione = "SOVRASCRITTO"
        sNota = "prima: " & Mid$(td.Connect, 11)
    End If
    Set td = Nothing

    If bSimulazione Then
        InviaUnLink = "SIMULA_" & sAzione
        Exit Function
    End If

    '----------------------------------------------------------------------------------------------
    ' #05.03# cancellazione del link esistente (R2)
    '----------------------------------------------------------------------------------------------
    m_pos = "05.03"
    If sAzione = "SOVRASCRITTO" Then dbFE.TableDefs.Delete sName

    '----------------------------------------------------------------------------------------------
    ' #05.04# creazione del nuovo link
    '----------------------------------------------------------------------------------------------
    m_pos = "05.04"
    Set td = dbFE.CreateTableDef(sName)
    td.Connect = ";DATABASE=" & sSorg
    td.SourceTableName = sName
    dbFE.TableDefs.Append td
    dbFE.TableDefs.Refresh
    InviaUnLink = sAzione
    Exit Function

GestErrore:
    sNota = "#" & m_pos & "# errore " & Err.Number & ": " & Err.Description
    InviaUnLink = "ERRORE"
End Function


'==================================================================================================
' #06# CONTROLLI SU FILE E DATABASE
'==================================================================================================

'--------------------------------------------------------------------------------------------------
' #06.01# TabellaInSorgente: True se il database sorgente esiste e contiene la tabella locale.
'         Le tabelle di ogni sorgente si leggono una volta sola (cache m_cacheSorg).
'--------------------------------------------------------------------------------------------------
Private Function TabellaInSorgente(ByVal sSorg As String, ByVal sName As String) As Boolean
    Dim dbS As DAO.Database, td As DAO.TableDef, tabelle As Object

    If Not m_cacheSorg.Exists(sSorg) Then
        Set tabelle = CreateObject("Scripting.Dictionary"): tabelle.CompareMode = 1
        If m_fso.FileExists(sSorg) Then
            On Error Resume Next
            Set dbS = DBEngine.OpenDatabase(sSorg, False, True)
            If Err.Number = 0 Then
                For Each td In dbS.TableDefs
                    If Len(td.Connect) = 0 Then tabelle(td.Name) = True
                Next td
                dbS.Close
            Else
                ScriviLog "06.01", "AVVISO", "Sorgente non apribile: " & sSorg & " (" & Err.Description & ")"
            End If
            Err.Clear
            On Error GoTo 0
        End If
        Set m_cacheSorg(sSorg) = tabelle
    End If
    TabellaInSorgente = m_cacheSorg(sSorg).Exists(sName)
End Function

'--------------------------------------------------------------------------------------------------
' #06.02# PercorsoInRadice: True se il percorso sta dentro RADICE_MSYS
'--------------------------------------------------------------------------------------------------
Private Function PercorsoInRadice(ByVal sPath As String) As Boolean
    PercorsoInRadice = (StrComp(Left$(sPath, Len(RADICE_MSYS)), RADICE_MSYS, vbTextCompare) = 0)
End Function

'--------------------------------------------------------------------------------------------------
' #06.03# PercorsoEscluso: cartelle di salvataggio, copie e distribuzioni non si leggono
'--------------------------------------------------------------------------------------------------
Private Function PercorsoEscluso(ByVal sPath As String) As Boolean
    Dim s As String
    s = UCase$(sPath) & "\"
    PercorsoEscluso = (InStr(s, "\AA_SALVATAGGI\") > 0 Or InStr(s, "\OLD\") > 0 Or _
                       InStr(s, "\BACKUP") > 0 Or InStr(s, "\EXE\") > 0 Or InStr(s, "\.GIT\") > 0)
End Function

'--------------------------------------------------------------------------------------------------
' #06.04# TabellaEsiste: True se la tabella (locale o collegata) esiste nel database
'--------------------------------------------------------------------------------------------------
Private Function TabellaEsiste(ByVal db As DAO.Database, ByVal sName As String) As Boolean
    Dim td As DAO.TableDef
    On Error Resume Next
    Set td = db.TableDefs(sName)
    TabellaEsiste = (Err.Number = 0)
    Err.Clear
End Function


'==================================================================================================
' #07# LOG
'==================================================================================================

'--------------------------------------------------------------------------------------------------
' #07.01# ApriLog: crea <cartella del database>\LOGS\AAAAMMGG_HHMMSS_<LOG_NOME>
'--------------------------------------------------------------------------------------------------
Private Sub ApriLog(ByVal sTitolo As String)
    Dim sCartella As String
    On Error Resume Next
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    sCartella = CurrentProject.Path & "\LOGS\"
    If Not m_fso.FolderExists(sCartella) Then m_fso.CreateFolder sCartella
    m_logPath = sCartella & Format$(Now, "yyyymmdd_hhnnss") & "_" & LOG_NOME
    Set m_log = m_fso.OpenTextFile(m_logPath, 8, True)
    On Error GoTo 0
    ScriviLog "07.01", "INIZIO", sTitolo & " - database: " & CurrentDb.Name
End Sub

'--------------------------------------------------------------------------------------------------
' #07.02# ScriviLog: una riga uguale nel file e nella finestra Immediata
'--------------------------------------------------------------------------------------------------
Private Sub ScriviLog(ByVal sPos As String, ByVal sEsito As String, ByVal sMsg As String)
    Dim s As String
    s = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & NOME_MODULO & " | Pos: " & sPos & " | " & sEsito & " | " & sMsg
    Debug.Print s
    On Error Resume Next
    If Not m_log Is Nothing Then m_log.WriteLine s
End Sub

'--------------------------------------------------------------------------------------------------
' #07.03# ChiudiLog
'--------------------------------------------------------------------------------------------------
Private Sub ChiudiLog()
    On Error Resume Next
    If Not m_log Is Nothing Then m_log.Close
    Set m_log = Nothing
End Sub


'==================================================================================================
' #08# UTILITA
'==================================================================================================

'--------------------------------------------------------------------------------------------------
' #08.01# SqlTesto: stringa tra apici per SQL, con gli apici interni raddoppiati
'--------------------------------------------------------------------------------------------------
Private Function SqlTesto(ByVal s As String) As String
    SqlTesto = "'" & Replace(s, "'", "''") & "'"
End Function

'--------------------------------------------------------------------------------------------------
' #08.02# AggiornaEsito: scrive Esito_s, Nota_s e DATA_INVIO_d su tutte le righe PATH_DB della
'         coppia destinazione / tabella (gli oggetti della tabella con quella destinazione)
'--------------------------------------------------------------------------------------------------
Private Sub AggiornaEsito(ByVal db As DAO.Database, ByVal sFE As String, ByVal sName As String, ByVal sEsito As String, ByVal sNota As String)
    db.Execute "UPDATE " & TAB_FRONTEND & " SET Esito_s = " & SqlTesto(sEsito) & _
               ", Nota_s = " & SqlTesto(Left$(sNota, 255)) & ", DATA_INVIO_d = Now()" & _
               " WHERE Database_s & [Database] = " & SqlTesto(sFE) & _
               " AND ID_OBJ IN (SELECT ID_OBJ FROM " & TAB_OGGETTI & " WHERE Name = " & SqlTesto(sName) & ")", dbFailOnError
End Sub

'--------------------------------------------------------------------------------------------------
' #08.03# ElencaDatabase: aggiunge a elenco tutti i .mdb/.accdb della cartella e sottocartelle,
'         saltando i percorsi esclusi (#06.03#)
'--------------------------------------------------------------------------------------------------
Private Sub ElencaDatabase(ByVal oCartella As Object, ByRef elenco As Collection)
    Dim oFile As Object, oSub As Object, sExt As String
    If PercorsoEscluso(oCartella.Path) Then Exit Sub
    For Each oFile In oCartella.Files
        sExt = LCase$(m_fso.GetExtensionName(oFile.Name))
        If sExt = "mdb" Or sExt = "accdb" Then elenco.Add oFile.Path
    Next oFile
    For Each oSub In oCartella.SubFolders
        ElencaDatabase oSub, elenco
    Next oSub
End Sub
