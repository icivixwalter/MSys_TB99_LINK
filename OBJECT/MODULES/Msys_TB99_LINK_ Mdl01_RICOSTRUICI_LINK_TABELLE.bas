Option Compare Database
Option Explicit

'======================================================================
' A) MODULO
'======================================================================
' MODULO     : Msys_TB99_LINK_ Mdl01_RICOSTRUICI_LINK_TABELLE
' PATH       : c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_TB99_LINK\MDB\
' DB         : MSys_TB99_LINK.mdb
' PATH_LOG   : \LOGS
' PROCEDURA  : RICOSTRUICI_LINK_TABELLE_pf
' CODICE     : @SHADOW_SYSTEM_CATALOG
' AUTORE     : System
' DATA       : 2026
' VERSIONE   : 3.0 TB99 SELF-HEALING LINK ENGINE WITH LOGGING
' CODICE_EST : IL codice di estrazione di tutti gli oggetti è --->  Msys_TB99_LINK
'
' FAQ        : @TB99@LINK_(modulo per la ricostruzione delle Tabelle Link di tutti i subprogetti)
'            : @COME@RICOSTRUIRE@_(I COLLEGAMENTI TRA LE TABELLE, @RICOSTRUIRE@LINK)

'======================================================================
' B) INDICE DELLE FUNZIONI E PROCEDURE DEL MODULO
'======================================================================
'
'  02)_PROCEDURA_PRINCIPALE = RICOSTRUICI_LINK_TABELLE_pf
'  B.1) RICOSTRUICI_LINK_TABELLE_pf   - Funzione PRINCIPALE: Avvia la ricostruzione
'                                        automatica di tutti i link tabelle
'                                        basata sulla query TB99.
'                                        Genera log dettagliato e statistiche.
'
'  21)_CANCELLAZIONE_LINK_ESISTENTI
'  B.2) CancellaTuttiLinkEsistenti_PFunct - Funzione: Cancella TUTTI i link tabelle
'                                        esistenti nel database corrente
'                                        (NON cancella le tabelle fisiche).
'                                        Utile per reset completo prima di
'                                        una ricostruzione pulita.
'
'
'  B.3) WriteLog                      - Sub: Scrive messaggi formattati con
'                                        timestamp su file di log e sulla
'                                        finestra Immediate (Debug).
'                                        Utilizzata da tutte le procedure.
'
'  23)_VERIFICA_QUERY = QueryExists
'  B.4) QueryExists                   - Funzione: Verifica l'esistenza di una
'                                        query nel database corrente tramite
'                                        MSysObjects (Type=5).
'
'  24)_VERIFICA_LINK = LinkExistsInMSysObjects
'  B.5) LinkExistsInMSysObjects       - Funzione: Controlla se un link tabella
'                                        esiste già nel database corrente
'                                        (Type=6 in MSysObjects).
'
'  25)_VERIFICA_TABELLA = TableExistsInDatabase
'  B.6) TableExistsInDatabase         - Funzione: Verifica l'esistenza di una
'                                        tabella in un database esterno.
'                                        Apre il DB sorgente, controlla in
'                                        MSysObjects e lo chiude rilasciando
'                                        le risorse.
'
'  26)_ELIMINAZIONE_LINK = DeleteLink
'  B.7) DeleteLink                    - Sub: Elimina un link tabella specifico
'                                        dal database corrente tramite
'                                        TableDefs.Delete.
'
'  27)_CREAZIONE_LINK = CreateLink
'  B.8) CreateLink                    - Funzione: Crea un nuovo link tabella
'                                        verso un database esterno.
'                                        Imposta Connect string e
'                                        SourceTableName, quindi Appende
'                                        il TableDef e fa RefreshLink.

'======================================================================
' C) SCOPO DEL MODULO
'======================================================================
' Questo modulo NON si limita più a ricreare collegamenti.
'
' È un motore di controllo e auto-riparazione dei link di Access
' basato sulla query centrale:
'
'       Msys_TB99_LINK_Qry99_01_SELECT_Archivio
'
' La query TB99 rappresenta un repository logico dei collegamenti
' dell'intero sistema e viene utilizzata come fonte unica di verità.
'
' Il modulo genera un file di log dettagliato di ogni operazione
' per consentire audit, debug e tracciabilità delle riparazioni.

'======================================================================
' D) LOGICA OPERATIVA
'======================================================================
'
' Per ogni record della query TB99:
'
'  1) Legge il percorso completo del database (campo Database)
'  2) Legge il nome della tabella (campo Name)
'  3) Legge il tipo oggetto (campo Type, deve essere 6 = link)
'  4) Verifica che il path sia strutturalmente valido
'  5) Verifica esistenza file fisico del database di destinazione
'  6) APRE il database di destinazione
'  7) Verifica se la tabella esiste NEL DATABASE DI DESTINAZIONE
'  8) Se esiste, elimina il link corrente nel DB corrente
'  9) Ricostruisce il link dal DB di destinazione al DB corrente
' 10) Aggiorna lo stato nel campo Verifica_s con uno dei valori:
'
'       -> OK              = link ricostruito e operativo
'       -> MISSING         = file database non trovato
'       -> INVALID_PATH    = path strutturalmente errato
'       -> NOT_REGISTERED  = tabella non trovata nel DB sorgente O Type != 6
'       -> NOT_CHECKED     = non ancora verificato
'
' 11) Traccia ogni operazione su file di log e finestra Immediate
' 12) Rilascia le risorse del DB di destinazione prima di passare al record successivo

'======================================================================
' E) DESCRIZIONE GENERALE
'======================================================================
' Il modulo sincronizza i collegamenti tra:
'
'   - Database di destinazione (sorgente dati reale)
'   - Database corrente (dove vengono ricreati i link)
'   - Msys_TB99_LINK_Qry99_01_SELECT_Archivio (repository logico centrale)
'
' e mantiene coerenza tra filesystem, database sorgente e database corrente.

'======================================================================
' F) ARCHITETTURA CONCETTUALE
'======================================================================
'
'   Query TB99 = SINGLE SOURCE OF TRUTH (metadati link)
'   Database di destinazione = sorgente fisica delle tabelle
'   Database corrente = destinatario dei link
'   File di log = tracciamento operazioni
'   Modulo VBA = motore di verifica e ricostruzione link
'
'                     [Query TB99]
'                          |
'                          | (legge Name + Database)
'                          v
'                  [Database Destinazione]
'                          |
'                          | (apri -> verifica tabella -> chiudi)
'                          v
'                   [Ricostruzione Link]
'                          |
'                          | (elimina link esistente -> crea nuovo link)
'                          v
'                  [Database Corrente]
'                          |
'                          v
'                    [File di Log]

'======================================================================
' G) ISTRUZIONI PER L'UTILIZZO
'======================================================================
'
'  G.1) QUERY DI RIFERIMENTO
'    Deve esistere:
'       Msys_TB99_LINK_Qry99_01_SELECT_Archivio
'
'    Campi minimi richiesti:
'       - Name          (nome tabella da collegare)
'       - Database      (percorso completo DB di destinazione)
'       - Type          (deve essere 6 per i link table)
'       - Verifica_s    (campo aggiornabile per lo stato)
'
'  G.2) CONFIGURAZIONE LOG
'    Variabili modificabili in testata:
'       MODULE_PATH     = Percorso base del modulo
'       LOG_FOLDER      = Cartella per i file di log (creata automaticamente)
'       LOG_FILE_NAME   = Nome base del file di log
'
'  G.3) AVVIO PROCEDURA
'    - F5 nel modulo
'    - oppure Call RICOSTRUICI_LINK_TABELLE_pf
'
'  G.4) OUTPUT DEL SISTEMA
'    - Aggiornamento campo Verifica_s nella query TB99
'    - Auto-riparazione link mancanti/corrotti
'    - File di log dettagliato in LOG_FOLDER
'    - Logging in finestra Immediate
'    - Statistiche finali di coerenza sistema

'======================================================================
' H) NOTE ARCHITETTURALI IMPORTANTI
'======================================================================
'
' -> Verifica lato sorgente: apre il DB di destinazione e controlla
'    l'esistenza reale della tabella prima di ricreare il link
' -> Chiude immediatamente il DB di destinazione dopo la verifica
'    per rilasciare le risorse e passare al record successivo
' -> Elimina preventivamente i link corrotti prima della ricostruzione
' -> Usa la query TB99 come livello di astrazione e controllo
' -> Genera file di log con timestamp per ogni esecuzione
' -> Traccia ogni singola operazione con dettaglio dello stato
' -> Permette migrazione sicura dei percorsi database
' -> Evita corruzione strutturale dei link Access

'======================================================================
' I) MARKER LOGICI DI FLUSSO (ALLINEATO AL CODICE REALE)
'======================================================================
' NEL CODICE SOTTOETICHETTE NUMERATE ES:  01_01[SET_DB] EC...
'
'  01)_VARIABILI_PER_LOGGING        = Inizializzazione variabili / DB corrente
'  02)_CREAZIONE_LOG                = Creazione file di log con timestamp
'  03)_VERIFICA_ESISTENZA_QUERY     = Controllo esistenza query TB99
'  04)_CONFERMA_UTENTE              = Conferma operazione da parte dell'utente
'  05)_APERTURA_RECORDSET_QUERY     = Apertura query Msys_TB99_LINK_Qry99_01_SELECT_Archivio
'  06)_INIZIALIZZAZIONE_STATISTICHE = Reset contatori statistiche
'  07)_CICLO_PRINCIPALE             = Loop principale su tutti i record
'  08)_LETTURA_RECORD               = Lettura Name, Database, Type
'  09)_VALIDAZIONE_PATH             = Verifica Type=6 e path valido
'  10)_APERTURA_DB_DESTINAZIONE     = OpenDatabase() sul campo Database
'  11)_VERIFICA_TABELLA_SORGENTE    = Controllo esistenza tabella nel DB destinazione
'  12)_CHIUSURA_DB_DESTINAZIONE     = Close database e rilascio risorse
'  13)_ELIMINAZIONE_LINK_PRECEDENTE = DeleteLink() se esiste
'  14)_RICOSTRUZIONE_LINK           = CreateLink() dal DB destinazione a DB corrente
'  15)_AGGIORNAMENTO_RECORDSET      = Update campo Verifica_s
'  16)_AGGIORNAMENTO_STATISTICHE    = Incremento contatori per stato
'  17)_RECORD_SUCCESSIVO            = MoveNext
'  18)_STATISTICHE_FINALI           = Report stato sistema su Debug e Log
'  19)_CHIUSURA_RISORSE             = Cleanup DAO, FSO e file di log
'  20)_MESSAGGIO_FINALE             = MsgBox con esito finale operazione
'  21)_CANCELLAZIONE_LINK_ESISTENTI = Cancella TUTTI i link esistenti (non le tabelle fisiche)
'  22)_FUNZIONE_LOG                 = WriteLog - Scrive messaggi su file di log
'  23)_VERIFICA_QUERY               = QueryExists - Controlla esistenza query in MSysObjects
'  24)_VERIFICA_LINK                = LinkExistsInMSysObjects - Controlla esistenza link in MSysObjects
'  25)_VERIFICA_TABELLA             = TableExistsInDatabase - Verifica esistenza tabella in DB esterno
'  26)_ELIMINAZIONE_LINK            = DeleteLink - Elimina link tabella da TableDefs
'  27)_CREAZIONE_LINK               = CreateLink - Crea nuovo link tabella verso database esterno
'  28)_FUNZIONE_HELPER              = WriteLogMessage - Scrive messaggio con logger o Debug
'  99)_GESTIONE_ERRORI              = Error handler globale con logging



'======================================================================
' 01)_VARIABILI_PER_LOGGING = Inizializzazione variabili / DB corrente
'======================================================================
' NOTA: Le variabili MODULE_PATH, LOG_FOLDER e LOG_FILE_NAME
'======================================================================
Private Const MODULE_PATH As String = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_TB99_LINK\MDB\"
                                            
Private LOG_FOLDER As String
Private Const LOG_FILE_NAME As String = "TB99_Link_Rebuild_Log.txt"



'======================================================================
' 02)_PROCEDURA_PRINCIPALE = RICOSTRUICI_LINK_TABELLE_pf
'======================================================================
' AVVIO : Call RICOSTRUICI_LINK_TABELLE_pf()
'======================================================================
Public Function RICOSTRUICI_LINK_TABELLE_pf()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fso As Object
    Dim LogFile As Object
    Dim logFilePath As String

    Dim sSql As String
    Dim sName As String
    Dim sDatabasePath As String
    Dim sStatus As String
    Dim lType As Long

    Dim tableExistsInSource As Boolean
    Dim startTime As Date
    Dim endTime As Date

' 02_01[SET_DB] = Imposta database corrente
    Set db = CurrentDb
' 02_02[SET_FSO] = Crea oggetto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
' 02_03[TIMER_START] = Avvia contatore temporale
    startTime = Now

'======================================================================
' 02)_CREAZIONE_LOG = Creazione file di log con timestamp
'======================================================================
' 02_04[SET_LOG_FOLDER] = Imposta percorso cartella LOGS
    LOG_FOLDER = MODULE_PATH & "LOGS\"
    
' 02_05[CREATE_LOG_FOLDER] = Crea cartella LOG se non esiste
    If Not fso.FolderExists(LOG_FOLDER) Then
        On Error Resume Next
        fso.CreateFolder LOG_FOLDER
        If Err.Number <> 0 Then
            ' Fallback su C:\TEMP\
            LOG_FOLDER = "C:\TEMP\"
            fso.CreateFolder LOG_FOLDER
        End If
        On Error GoTo 0
    End If
    
' 02_06[BUILD_LOG_PATH] = Costruisce percorso completo file di log
    logFilePath = LOG_FOLDER & Format(startTime, "YYYYMMDD_HHMMSS") & "_" & LOG_FILE_NAME
    
' 02_07[CREATE_LOG_FILE] = Crea il file di log
    Set LogFile = fso.CreateTextFile(logFilePath, True)
    
' 02_08[WRITE_HEADER] = Scrive intestazione nel log
    Call WriteLog(LogFile, "========================================================================")
    Call WriteLog(LogFile, "INIZIO RICOSTRUZIONE LINK TB99")
    Call WriteLog(LogFile, "Data/Ora avvio: " & Format(startTime, "dd/mm/yyyy HH:MM:SS"))
    Call WriteLog(LogFile, "Database corrente: " & CurrentDb.Name)
    Call WriteLog(LogFile, "File di log: " & logFilePath)
    Call WriteLog(LogFile, "========================================================================")
    Call WriteLog(LogFile, "")

'======================================================================
' 03)_VERIFICA_ESISTENZA_QUERY = Controllo esistenza query TB99
'======================================================================
' 03_01[CHECK_QUERY] = Verifica esistenza query Msys_TB99_LINK_Qry99_01_SELECT_Archivio
    Call WriteLog(LogFile, "[VERIFICA_QUERY] Controllo esistenza query Msys_TB99_LINK_Qry99_01_SELECT_Archivio...")
    
    If Not QUERYExists("Msys_TB99_LINK_Qry99_01_SELECT_Archivio") Then
        Call WriteLog(LogFile, "[ERRORE] Query Msys_TB99_LINK_Qry99_01_SELECT_Archivio NON TROVATA!")
        Call WriteLog(LogFile, "[FINE] Operazione terminata con errore")
        LogFile.Close
        Set LogFile = Nothing
        MsgBox "Query Msys_TB99_LINK_Qry99_01_SELECT_Archivio non trovata", vbCritical
        Exit Function
    End If
    
    Call WriteLog(LogFile, "[VERIFICA_QUERY] Query trovata con successo.")
    Call WriteLog(LogFile, "")

'======================================================================
' 04)_CONFERMA_UTENTE = Conferma operazione da parte dell'utente
'======================================================================
' 04_01[ASK_CONFIRM] = Richiede conferma all'utente
    Call WriteLog(LogFile, "[UTENTE] Richiesta conferma operazione...")
    
    If MsgBox("Avvio ricostruzione link tabelle da TB99-", vbYesNo) = vbNo Then
        Call WriteLog(LogFile, "[UTENTE] Operazione ANNULLATA dall'utente.")
        Call WriteLog(LogFile, "[FINE] Log chiuso.")
        LogFile.Close
        Set LogFile = Nothing
        Exit Function
    End If
    
    Call WriteLog(LogFile, "[UTENTE] Operazione CONFERMATA.")
    Call WriteLog(LogFile, "")

'======================================================================
' 05)_APERTURA_RECORDSET_QUERY = Apertura query Msys_TB99_LINK_Qry99_01_SELECT_Archivio
'======================================================================
' 05_01[BUILD_SQL] = Costruisce query SQL
    sSql = "SELECT * FROM Msys_TB99_LINK_Qry99_01_SELECT_Archivio"
    
' ==================== INSERISCI QUI IL CONTROLLO ====================
    ' 05_01A[CHECK_RECORDS] = Verifica se la query contiene record
    Dim rsCount As DAO.Recordset
    Set rsCount = db.OpenRecordset("SELECT Count(*) AS Tot FROM Msys_TB99_LINK_Qry99_01_SELECT_Archivio", dbOpenSnapshot)
    
    If rsCount.Fields("Tot").Value = 0 Then
        rsCount.Close
        Set rsCount = Nothing
        Call WriteLog(LogFile, "[AVVISO] La query Msys_TB99_LINK_Qry99_01_SELECT_Archivio NON contiene record!")
        Call WriteLog(LogFile, "[FINE] Nessun link da ricostruire.")
        LogFile.Close
        Set LogFile = Nothing
        MsgBox "NESSUN LINK DA RICOSTRUIRE" & vbCrLf & vbCrLf & _
               "La query Msys_TB99_LINK_Qry99_01_SELECT_Archivio non contiene record.", vbInformation
        Exit Function
    End If
    
    rsCount.Close
    Set rsCount = Nothing
' =====================================================================

' 05_02[OPEN_RECORDSET] = Apre recordset
    Set rs = db.OpenRecordset(sSql, dbOpenDynaset)
    
    
    Call WriteLog(LogFile, "[APERTURA_QUERY] Recordset aperto. Record trovati: " & rs.RecordCount)
    Call WriteLog(LogFile, "")

'======================================================================
' 06)_INIZIALIZZAZIONE_STATISTICHE = Reset contatori statistiche
'======================================================================
' 06_01[RESET_COUNTERS] = Inizializza contatori a zero
    Dim nTot As Long: nTot = 0
    Dim nOk As Long: nOk = 0
    Dim nMissing As Long: nMissing = 0
    Dim nInvalidPath As Long: nInvalidPath = 0
    Dim nNotRegistered As Long: nNotRegistered = 0
    
    Call WriteLog(LogFile, "[STATISTICHE] Contatori inizializzati.")
    Call WriteLog(LogFile, "")

'======================================================================
' 07)_CICLO_PRINCIPALE = Loop principale su tutti i record
'======================================================================
    Call WriteLog(LogFile, "========================================================================")
    Call WriteLog(LogFile, "INIZIO CICLO ELABORAZIONE RECORD")
    Call WriteLog(LogFile, "========================================================================")
    Call WriteLog(LogFile, "")
    
' 07_01[LOOP_START] = Inizio ciclo Do While
    Do While Not rs.EOF
        nTot = nTot + 1
        
        Call WriteLog(LogFile, "----------------------------------------")
        Call WriteLog(LogFile, "RECORD N. " & nTot)
        Call WriteLog(LogFile, "----------------------------------------")

'======================================================================
' 08)_LETTURA_RECORD = Lettura Name, Database, Type
'======================================================================
' 08_01[READ_NAME] = Legge campo Name
        sName = Nz(rs!Name, "")
' 08_02[READ_DATABASE] = Legge campo Database
        sDatabasePath = Nz(rs!DB_COMPLETO, "")   '// 15/11/2026: percorso = Database_s (cartella) & Database (file)
' 08_03[READ_TYPE] = Legge campo Type
        lType = Nz(rs!Type, 0)

        Call WriteLog(LogFile, "[LETTURA_RECORD] Nome tabella: '" & sName & "'")
        Call WriteLog(LogFile, "[LETTURA_RECORD] Database path: '" & sDatabasePath & "'")
        Call WriteLog(LogFile, "[LETTURA_RECORD] Type: " & lType)
        
' 08_04[INIT_STATUS] = Inizializza stato
        sStatus = "NOT_CHECKED"

'======================================================================
' 09)_VALIDAZIONE_PATH = Verifica Type=6 e path valido
'======================================================================
        Call WriteLog(LogFile, "[VALIDAZIONE] Avvio validazione...")
        
' 09_01[CHECK_TYPE] = Verifica che sia un link table (Type = 6)
        If lType <> 6 Then
            sStatus = "NOT_REGISTERED"
            Call WriteLog(LogFile, "[VALIDAZIONE] Type = " & lType & " (diverso da 6) -> NON è un link table")
            Call WriteLog(LogFile, "[VALIDAZIONE] Stato assegnato: " & sStatus)
            
' 09_02[CHECK_PATH_FORMAT] = Verifica path strutturalmente valido
        ElseIf Len(Trim(sDatabasePath)) = 0 Or InStr(1, sDatabasePath, ":\") = 0 Then
            sStatus = "INVALID_PATH"
            Call WriteLog(LogFile, "[VALIDAZIONE] Path strutturalmente INVALIDO")
            Call WriteLog(LogFile, "[VALIDAZIONE] Stato assegnato: " & sStatus)
            
' 09_03[CHECK_FILE_EXISTS] = Verifica esistenza file fisico del database
        ElseIf Not fso.FileExists(sDatabasePath) Then
            sStatus = "MISSING"
            Call WriteLog(LogFile, "[VALIDAZIONE] File database NON TROVATO sul filesystem")
            Call WriteLog(LogFile, "[VALIDAZIONE] Path cercato: '" & sDatabasePath & "'")
            Call WriteLog(LogFile, "[VALIDAZIONE] Stato assegnato: " & sStatus)
            
        Else
            Call WriteLog(LogFile, "[VALIDAZIONE] Path valido. File database ESISTE.")
            Call WriteLog(LogFile, "[VALIDAZIONE] Inizio verifica lato SORGENTE...")
            
'======================================================================
' 10)_APERTURA_DB_DESTINAZIONE = OpenDatabase() sul campo Database
' 11)_VERIFICA_TABELLA_SORGENTE = Controllo esistenza tabella nel DB destinazione
' 12)_CHIUSURA_DB_DESTINAZIONE = Close database e rilascio risorse
'======================================================================
            
' 10_01[OPEN_EXTERNAL_DB] = Apre il database di destinazione
' 11_01[CHECK_TABLE_EXISTS] = Verifica esistenza tabella
' 12_01[CLOSE_EXTERNAL_DB] = Chiude database e rilascia risorse
            tableExistsInSource = TableExistsInDatabase(sDatabasePath, sName, LogFile)
            
' 11_02[EVALUATE_RESULT] = Valuta risultato verifica
            If tableExistsInSource Then
                Call WriteLog(LogFile, "[VERIFICA_TABELLA] Tabella '" & sName & "' TROVATA nel database di destinazione.")
                
'======================================================================
' 13)_ELIMINAZIONE_LINK_PRECEDENTE = DeleteLink() se esiste
'======================================================================
                
' 13_01[CHECK_LINK_EXISTS] = Controlla se link esiste già
                If LinkExistsInMSysObjects(sName) Then
                    Call WriteLog(LogFile, "[ELIMINAZIONE_LINK] Link esistente TROVATO in MSysObjects.")
                    Call WriteLog(LogFile, "[ELIMINAZIONE_LINK] Eliminazione in corso...")
' 13_02[DELETE_LINK] = Elimina link esistente
                    DeleteLink sName, LogFile
                Else
                    Call WriteLog(LogFile, "[ELIMINAZIONE_LINK] Link esistente NON trovato in MSysObjects.")
                End If
                
'======================================================================
' 14)_RICOSTRUZIONE_LINK = CreateLink() dal DB destinazione a DB corrente
'======================================================================
                
                Call WriteLog(LogFile, "[RICOSTRUZIONE_LINK] Creazione nuovo link...")
                Call WriteLog(LogFile, "[RICOSTRUZIONE_LINK] Tabella: '" & sName & "'")
                Call WriteLog(LogFile, "[RICOSTRUZIONE_LINK] Database sorgente: '" & sDatabasePath & "'")
                
' 14_01[CREATE_NEW_LINK] = Crea nuovo link
                If CreateLink(sName, sDatabasePath, LogFile) Then
                    sStatus = "OK"
                    Call WriteLog(LogFile, "[RICOSTRUZIONE_LINK] Link CREATO con successo!")
                    Call WriteLog(LogFile, "[RICOSTRUZIONE_LINK] Stato assegnato: " & sStatus)
                Else
                    sStatus = "NOT_REGISTERED"
                    Call WriteLog(LogFile, "[RICOSTRUZIONE_LINK] ERRORE: Impossibile creare il link!")
                    Call WriteLog(LogFile, "[RICOSTRUZIONE_LINK] Stato assegnato: " & sStatus)
                End If
            Else
                sStatus = "NOT_REGISTERED"
                Call WriteLog(LogFile, "[VERIFICA_TABELLA] Tabella '" & sName & "' NON trovata nel database di destinazione.")
                Call WriteLog(LogFile, "[VERIFICA_TABELLA] Stato assegnato: " & sStatus)
            End If
            
        End If
        
'======================================================================
' 15)_AGGIORNAMENTO_RECORDSET = Update campo Verifica_s
'======================================================================
        Call WriteLog(LogFile, "[AGGIORNAMENTO] Salvataggio stato nel campo Verifica_s...")
        
' 15_01[CHECK_UPDATABLE] = Verifica se recordset è aggiornabile
        If rs.Updatable Then
' 15_02[EDIT_RECORD] = Entra in modalità modifica
            rs.Edit
' 15_03[UPDATE_STATUS] = Aggiorna campo Verifica_s
            rs!Verifica_s = sStatus

' 15_04[UPDATE_SUBPROJECT] = Aggiorna campo COD_SUB_PROGETTO_s SOLO SE VUOTO (2026-11-15: non sovrascrive i codici del catalogo)
            If Len(Nz(rs!COD_SUB_PROGETTO_s, "")) > 0 Then
                Call WriteLog(LogFile, "[AGGIORNAMENTO] COD_SUB_PROGETTO_s gia censito: '" & rs!COD_SUB_PROGETTO_s & "' (non modificato)")
            ElseIf Len(sName) >= 9 Then
                rs!COD_SUB_PROGETTO_s = Left(sName, 9)
                Call WriteLog(LogFile, "[AGGIORNAMENTO] COD_SUB_PROGETTO_s impostato a: '" & Left(sName, 9) & "'")
            Else
                rs!COD_SUB_PROGETTO_s = sName
                Call WriteLog(LogFile, "[AGGIORNAMENTO] COD_SUB_PROGETTO_s impostato a (nome completo <9 char): '" & sName & "'")
            End If
            
' 15_05[SAVE_RECORD] = Salva modifiche
            rs.Update
            Call WriteLog(LogFile, "[AGGIORNAMENTO] Recordset aggiornato con stato: '" & sStatus & "' nel campo Verifica_s")
        Else
            Call WriteLog(LogFile, "[AGGIORNAMENTO] WARNING: Recordset NON aggiornabile per " & sName)
        End If

'======================================================================
' 16)_AGGIORNAMENTO_STATISTICHE = Incremento contatori per stato
'======================================================================
' 16_01[INCREMENT_OK] = Incrementa contatore OK
' 16_02[INCREMENT_MISSING] = Incrementa contatore MISSING
' 16_03[INCREMENT_INVALID] = Incrementa contatore INVALID_PATH
' 16_04[INCREMENT_NOTREG] = Incrementa contatore NOT_REGISTERED
        Select Case sStatus
            Case "OK": nOk = nOk + 1
            Case "MISSING": nMissing = nMissing + 1
            Case "INVALID_PATH": nInvalidPath = nInvalidPath + 1
            Case "NOT_REGISTERED": nNotRegistered = nNotRegistered + 1
        End Select
        
        Call WriteLog(LogFile, "[STATISTICHE] Aggiornamento contatori:")
        Call WriteLog(LogFile, "            OK: " & nOk)
        Call WriteLog(LogFile, "            MISSING: " & nMissing)
        Call WriteLog(LogFile, "            INVALID_PATH: " & nInvalidPath)
        Call WriteLog(LogFile, "            NOT_REGISTERED: " & nNotRegistered)
        Call WriteLog(LogFile, "            Totale elaborati: " & nTot)
        
        Call WriteLog(LogFile, "----------------------------------------")
        Call WriteLog(LogFile, "")
        
'======================================================================
' 17)_RECORD_SUCCESSIVO = MoveNext
'======================================================================
' 17_01[NEXT_RECORD] = Passa al record successivo
        rs.MoveNext

    Loop
' 07_02[LOOP_END] = Fine ciclo Do While

'======================================================================
' 18)_STATISTICHE_FINALI = Report stato sistema su Debug e Log
'======================================================================
' 18_01[TIMER_END] = Rileva ora fine elaborazione
    endTime = Now
    
' 18_02[WRITE_FINAL_STATS] = Scrive statistiche finali su log
    Call WriteLog(LogFile, "========================================================================")
    Call WriteLog(LogFile, "STATISTICHE FINALI")
    Call WriteLog(LogFile, "========================================================================")
    Call WriteLog(LogFile, "Data/Ora fine: " & Format(endTime, "dd/mm/yyyy HH:MM:SS"))
    Call WriteLog(LogFile, "Durata totale: " & Format(endTime - startTime, "hh:mm:ss"))
    Call WriteLog(LogFile, "")
    Call WriteLog(LogFile, "TOTALE RECORD ELABORATI: " & nTot)
    Call WriteLog(LogFile, "OK (link ricostruiti): " & nOk)
    Call WriteLog(LogFile, "MISSING (DB non trovato): " & nMissing)
    Call WriteLog(LogFile, "INVALID_PATH (path errato): " & nInvalidPath)
    Call WriteLog(LogFile, "NOT_REGISTERED (tabella inesistente): " & nNotRegistered)
    Call WriteLog(LogFile, "========================================================================")
    
' 18_03[DEBUG_PRINT] = Stampa statistiche su finestra Immediate
    Debug.Print "========== STATISTICHE FINALI =========="
    Debug.Print "TOTALE RECORD: " & nTot
    Debug.Print "OK (link ricostruiti): " & nOk
    Debug.Print "MISSING (DB non trovato): " & nMissing
    Debug.Print "INVALID_PATH (path errato): " & nInvalidPath
    Debug.Print "NOT_REGISTERED (tabella inesistente): " & nNotRegistered
    Debug.Print "========================================"

'======================================================================
' 19)_CHIUSURA_RISORSE = Cleanup DAO, FSO e file di log
'======================================================================
' 19_01[CLOSE_RECORDSET] = Chiude recordset
    rs.Close
' 19_02[RELEASE_RS] = Rilascia recordset
    Set rs = Nothing
' 19_03[RELEASE_DB] = Rilascia database
    Set db = Nothing
' 19_04[RELEASE_FSO] = Rilascia FileSystemObject
    Set fso = Nothing
    
    Call WriteLog(LogFile, "[CHIUSURA] File di log salvato in: " & logFilePath)
' 19_05[CLOSE_LOG] = Chiude file di log
    LogFile.Close
' 19_06[RELEASE_LOG] = Rilascia oggetto log
    Set LogFile = Nothing

'======================================================================
' 20)_MESSAGGIO_FINALE = MsgBox con esito finale operazione
'======================================================================
' 20_01[SHOW_RESULT] = Mostra messaggio riepilogativo all'utente
    MsgBox "Ricostruzione link completata!" & vbCrLf & vbCrLf & _
           "Totale elaborati: " & nTot & vbCrLf & _
           "Link OK: " & nOk & vbCrLf & _
           "DB non trovato (MISSING): " & nMissing & vbCrLf & _
           "Path invalido (INVALID_PATH): " & nInvalidPath & vbCrLf & _
           "Tabella non censita (NOT_REGISTERED): " & nNotRegistered & vbCrLf & vbCrLf & _
           "Log salvato in: " & logFilePath, vbInformation

    Exit Function

'======================================================================
' 99)_GESTIONE_ERRORI = Error handler globale con logging
'======================================================================
ErrorHandler:
' 99_01[BUILD_ERROR_MSG] = Costruisce messaggio errore
    Dim errorMsg As String
    errorMsg = "Errore: " & Err.Description & vbCrLf & _
               "Numero: " & Err.Number & vbCrLf & _
               "In tabella: " & sName
    
' 99_02[LOG_ERROR] = Scrive errore nel log
    If Not LogFile Is Nothing Then
        Call WriteLog(LogFile, "[ERRORE_GLOBALE] " & errorMsg)
        Call WriteLog(LogFile, "[FINE] Operazione terminata con errore")
        LogFile.Close
    End If
    
' 99_03[SHOW_ERROR] = Mostra errore all'utente
    MsgBox errorMsg, vbCritical

End Function


'======================================================================
' 21)_CANCELLAZIONE_LINK_ESISTENTI = Cancella TUTTI i link esistenti
'======================================================================
' AVVIO : Call CancellaTuttiLinkEsistenti_PFunct()
'
'COSA FA LA FUNZIONE
'La funzione CancellaTuttiLinkEsistenti() esegue queste operazioni:
'
'Azione Descrizione
'- Crea log Genera un file di log con timestamp in \LOGS\
'- Verifica query   Controlla che esista Msys_TB99_LINK_Qry99_01_SELECT_Archivio
'- Conferma utente  Chiede conferma (già integrata)
'- Legge tutti i link   Apre il recordset della query TB99
'- Cancella uno per uno Per ogni link (Type=6), lo elimina da TableDefs
'- Statistiche  Mostra quanti link cancellati e eventuali errori
'- Messaggio finale Riepilogo dell'operazione



'+-------------------------------------------------------------------+
'¦ 21) CANCELLAZIONE_LINK_ESISTENTI                                  ¦
'+-------------------------------------------------------------------¦
'¦ 21_01[DECLARATIONS]      - Dichiarazione variabili                ¦
'¦ 21_02[LOGGER_INIT]       - Inizializzazione Logger                ¦
'¦ 21_03[CHECK_EXISTING_LINKS] - Verifica link esistenti             ¦
'¦ 21_04[CREATE_FSO]        - Crea cartella LOGS                     ¦
'¦ 21_05[CREATE_LOGGER]     - Crea istanza logger                    ¦
'¦ 21_06[CHECK_LOGGER]      - Verifica logger                        ¦
'¦ 21_07[WRITE_HEADER]      - Scrive intestazione log                ¦
'¦ 21_08[VERIFY_QUERY]      - Verifica query TB99                    ¦
'¦ 21_09[ASK_CONFIRM]       - Richiede conferma utente               ¦
'¦ 21_10[OPEN_RECORDSET]    - Apertura recordset per cancellazione   ¦
'¦ 21_11[LOOP_START]        - Loop cancellazione                     ¦
'¦ 21_12[CHECK_IF_LINK]     - Verifica Type=6                        ¦
'¦ 21_13[CHECK_LINK_EXISTS] - Verifica in MSysObjects                ¦
'¦ 21_14[DELETE_LINK]       - Elimina il link                        ¦
'¦ 21_15[NEXT_RECORD]       - Passa al record successivo             ¦
'¦ 21_16[STATISTICHE_FINALI] - Report statistiche                    ¦
'¦ 21_17[CLOSE_RESOURCES]   - Chiusura risorse                       ¦
'¦ 21_18[CLOSE_LOGGER]      - Chiusura logger                        ¦
'¦ 21_19[MESSAGGIO_FINALE]  - MsgBox esito finale                    ¦
'¦ 99)[GESTIONE_ERRORI]     - Error handler                          ¦
'+-------------------------------------------------------------------+
'
'+-------------------------------------------------------------------+
'¦ FUNZIONI DI SUPPORTO                                              ¦
'+-------------------------------------------------------------------¦
'¦ 22) WriteLog              - Scrive su file log e Debug            ¦
'¦ 23) QueryExists           - Verifica esistenza query              ¦
'¦ 24) LinkExistsInMSysObjects - Verifica link in MSysObjects        ¦
'¦ 25) TableExistsInDatabase - Verifica tabella in DB esterno        ¦
'¦ 26) DeleteLink            - Elimina link da TableDefs             ¦
'¦ 27) CreateLink            - Crea nuovo link                       ¦
'¦ 28) WriteLogMessage       - Wrapper per logging                   ¦
'+-------------------------------------------------------------------+

' Pre-controllo (riga 21_03)
'queryName = "MSysObjects_Qry01_06_SELECT_TABELLE_LINK"
' - Verifica esistenza link nel DB

' Verifica tracciamento (riga 21_08)
'"Msys_TB99_LINK_Qry99_01_SELECT_Archivio"
' - Controlla se esiste la query di archivio

' Recordset principale (riga 21_10)
'sSQL = "SELECT * FROM " & queryName
' - Usa MSysObjects_Qry01_06_SELECT_TABELLE_LINK


'======================================================================
Public Function CancellaTuttiLinkEsistenti_PFunct() As Boolean
    
    '======================================================================
    ' 21_01[DECLARATIONS] = Dichiarazione variabili
    '======================================================================
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsCheck As DAO.Recordset
    Dim sSql As String
    Dim sName As String
    Dim queryName As String
    Dim nTot As Long
    Dim nCancellati As Long
    Dim nErrori As Long
    Dim nLinkEsistenti As Long
    Dim startTime As Date
    Dim endTime As Date
    
    '======================================================================
    ' 21_02[LOGGER_INIT] = Inizializzazione Logger
    '======================================================================
    Dim logger As Object
    Dim logFilePath As String
    Dim logFileName As String
    Dim fso As Object
    Dim dbPath As String
    Dim dbFolder As String
    Dim logEnabled As Boolean
    
    On Error GoTo ErrorHandler
    
    '======================================================================
    ' 21_03[CHECK_EXISTING_LINKS] = Verifica se esistono link nel progetto corrente
    '======================================================================
    queryName = "MSysObjects_Qry01_06_SELECT_TABELLE_LINK"
    
    ' Verifica se la query esiste
    If Not QUERYExists(queryName) Then
        MsgBox "Query " & queryName & " non trovata!", vbCritical
        CancellaTuttiLinkEsistenti_PFunct = False
        Exit Function
    End If
    
    Set db = CurrentDb
    
    ' Apre il recordset per verificare se esistono link
    Set rsCheck = db.OpenRecordset("SELECT * FROM " & queryName, dbOpenSnapshot)
    
    ' SE IL RECORDSET È VUOTO - ESCE CON MESSAGGIO
    If rsCheck.EOF Then
        rsCheck.Close
        Set rsCheck = Nothing
        Set db = Nothing
        MsgBox "NESSUN LINK DA CANCELLARE" & vbCrLf & vbCrLf & _
               "Il progetto corrente non contiene tabelle link (Type=6).", vbInformation
        CancellaTuttiLinkEsistenti_PFunct = True
        Exit Function
    End If
    
    ' Se arriviamo qui, ci sono link da cancellare
    nLinkEsistenti = rsCheck.RecordCount
    
    ' Chiudi il recordset di verifica
    rsCheck.Close
    Set rsCheck = Nothing
    
    '======================================================================
    ' 21_04[CREATE_FSO] = Crea oggetto FileSystemObject per il percorso
    '======================================================================
    Set fso = CreateObject("Scripting.FileSystemObject")
    dbPath = CurrentDb.Name
    dbFolder = fso.GetParentFolderName(dbPath)
    logFileName = "DELETE_LINKS_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"
    logFilePath = dbFolder & "\LOGS\" & logFileName
    
    ' Crea la cartella LOGS se non esiste
    If Not fso.FolderExists(dbFolder & "\LOGS\") Then
        On Error Resume Next
        fso.CreateFolder dbFolder & "\LOGS\"
        On Error GoTo 0
    End If
    
    '======================================================================
    ' 21_05[CREATE_LOGGER] = Istanza della classe log unica MSys_TB12_Cls01_LOG
    '                       file: LOGS\DELETE_LINKS_<data_ora>_INFO_<giorno>.txt (+ _ERRORI_ se ci sono errori)
    '======================================================================
    Set logger = New MSys_TB12_Cls01_LOG
    logger.CartellaLog = dbFolder & "\LOGS"
    logger.NomeBase = Replace(logFileName, ".txt", "")
    logger.origine = "Mdl01_RICOSTRUICI_LINK_TABELLE"
    logger.ScriviAVideo = False          ' il Debug.Print lo fa gia WriteLogMessage
    logFilePath = logger.PercorsoFile("INFO")

    '======================================================================
    ' 21_06[CHECK_LOGGER] = Logger pronto
    '======================================================================
    logEnabled = True
    
    '======================================================================
    ' 21_07[WRITE_HEADER] = Scrive intestazione nel log
    '======================================================================
    startTime = Now
    
    Call WriteLogMessage(logger, logEnabled, "========================================================================", 1)
    Call WriteLogMessage(logger, logEnabled, "INIZIO CANCELLAZIONE TUTTI I LINK ESISTENTI", 1)
    Call WriteLogMessage(logger, logEnabled, "Data/Ora avvio: " & Format(startTime, "dd/mm/yyyy HH:MM:SS"), 1)
    Call WriteLogMessage(logger, logEnabled, "Database corrente: " & CurrentDb.Name, 1)
    Call WriteLogMessage(logger, logEnabled, "File di log: " & logFilePath, 1)
    Call WriteLogMessage(logger, logEnabled, "========================================================================", 1)
    Call WriteLogMessage(logger, logEnabled, "", 1)
    
    '======================================================================
    ' 21_08[VERIFY_QUERY] = Verifica esistenza query TB99 (solo per verifica)
    '======================================================================
    Call WriteLogMessage(logger, logEnabled, "[VERIFICA_QUERY] Controllo esistenza query Msys_TB99_LINK_Qry99_01_SELECT_Archivio...", 1)
    
    If Not QUERYExists("Msys_TB99_LINK_Qry99_01_SELECT_Archivio") Then
        Call WriteLogMessage(logger, logEnabled, "[AVVISO] Query Msys_TB99_LINK_Qry99_01_SELECT_Archivio NON TROVATA!", 2)
        Call WriteLogMessage(logger, logEnabled, "[AVVISO] Procedo con la cancellazione senza tracciamento.", 2)
    Else
        Call WriteLogMessage(logger, logEnabled, "[VERIFICA_QUERY] Query trovata con successo.", 1)
    End If
    Call WriteLogMessage(logger, logEnabled, "", 1)
    
    '======================================================================
    ' 21_09[ASK_CONFIRM] = Richiede conferma all'utente
    '======================================================================
    Call WriteLogMessage(logger, logEnabled, "[UTENTE] Richiesta conferma cancellazione " & nLinkEsistenti & " link...", 1)
    
    If MsgBox("ATTENZIONE! Verranno cancellati " & nLinkEsistenti & " link tabelle presenti nel progetto corrente." & vbCrLf & vbCrLf & _
              "Questa operazione NON cancella le tabelle fisiche, solo i collegamenti." & vbCrLf & vbCrLf & _
              "Procedere-", vbYesNo + vbExclamation) = vbNo Then
        Call WriteLogMessage(logger, logEnabled, "[UTENTE] Operazione ANNULLATA dall'utente.", 1)
        Call WriteLogMessage(logger, logEnabled, "[FINE] Log chiuso.", 1)
        
        Set logger = Nothing
        Set fso = Nothing
        CancellaTuttiLinkEsistenti_PFunct = False
        Exit Function
    End If
    
    Call WriteLogMessage(logger, logEnabled, "[UTENTE] Operazione CONFERMATA.", 1)
    Call WriteLogMessage(logger, logEnabled, "", 1)
    
    '======================================================================
    ' 21_10[OPEN_RECORDSET] = Apertura recordset per cancellazione
    '======================================================================
    sSql = "SELECT * FROM " & queryName
    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    
    Call WriteLogMessage(logger, logEnabled, "[APERTURA_QUERY] Recordset aperto. Record trovati: " & rs.RecordCount, 1)
    Call WriteLogMessage(logger, logEnabled, "", 1)
    Call WriteLogMessage(logger, logEnabled, "========================================================================", 1)
    Call WriteLogMessage(logger, logEnabled, "INIZIO CANCELLAZIONE LINK", 1)
    Call WriteLogMessage(logger, logEnabled, "========================================================================", 1)
    Call WriteLogMessage(logger, logEnabled, "", 1)
    
    '======================================================================
    ' 21_11[LOOP_START] = Loop su tutti i record della query
    '======================================================================
    nTot = 0
    nCancellati = 0
    nErrori = 0
    
    Do While Not rs.EOF
        nTot = nTot + 1
        sName = Nz(rs!Name, "")
        
        ' Salta nomi vuoti
        If Len(sName) = 0 Then
            Call WriteLogMessage(logger, logEnabled, "[AVVISO] RECORD N. " & nTot & " - Nome vuoto, saltato.", 2)
            rs.MoveNext
            GoTo NextRecord
        End If
        
        Call WriteLogMessage(logger, logEnabled, "----------------------------------------", 1)
        Call WriteLogMessage(logger, logEnabled, "RECORD N. " & nTot & " - Tabella: '" & sName & "'", 1)
        Call WriteLogMessage(logger, logEnabled, "----------------------------------------", 1)
        
        '======================================================================
        ' 21_12[CHECK_IF_LINK] = Verifica se è effettivamente un link (Type=6)
        '======================================================================
        If Nz(rs!Type, 0) = 6 Then
            
            '======================================================================
            ' 21_13[CHECK_LINK_EXISTS] = Verifica se il link esiste in MSysObjects
            '======================================================================
            If LinkExistsInMSysObjects(sName) Then
                Call WriteLogMessage(logger, logEnabled, "[CANCELLAZIONE] Link trovato, eliminazione in corso...", 2)
                
                '======================================================================
                ' 21_14[DELETE_LINK] = Elimina il link
                '======================================================================
                On Error Resume Next
                db.TableDefs.Delete sName
                If Err.Number = 0 Then
                    nCancellati = nCancellati + 1
                    Call WriteLogMessage(logger, logEnabled, "[CANCELLAZIONE] Link '" & sName & "' eliminato con SUCCESSO.", 1)
                Else
                    nErrori = nErrori + 1
                    Call WriteLogMessage(logger, logEnabled, "[CANCELLAZIONE] ERRORE eliminazione link '" & sName & "': " & Err.Description, 3)
                End If
                On Error GoTo 0
            Else
                Call WriteLogMessage(logger, logEnabled, "[CANCELLAZIONE] Link '" & sName & "' NON trovato in MSysObjects (già assente).", 2)
            End If
        Else
            Call WriteLogMessage(logger, logEnabled, "[CANCELLAZIONE] Type = " & Nz(rs!Type, 0) & " -> NON è un link table, saltato.", 2)
        End If
        
        Call WriteLogMessage(logger, logEnabled, "----------------------------------------", 1)
        Call WriteLogMessage(logger, logEnabled, "", 1)
        
NextRecord:
        '======================================================================
        ' 21_15[NEXT_RECORD] = Passa al record successivo
        '======================================================================
        rs.MoveNext
    Loop
    
    '======================================================================
    ' 21_16[STATISTICHE_FINALI] = Report statistiche cancellazione
    '======================================================================
    endTime = Now
    
    Call WriteLogMessage(logger, logEnabled, "========================================================================", 1)
    Call WriteLogMessage(logger, logEnabled, "STATISTICHE FINALI CANCELLAZIONE", 1)
    Call WriteLogMessage(logger, logEnabled, "========================================================================", 1)
    Call WriteLogMessage(logger, logEnabled, "Data/Ora fine: " & Format(endTime, "dd/mm/yyyy HH:MM:SS"), 1)
    Call WriteLogMessage(logger, logEnabled, "Durata totale: " & Format(endTime - startTime, "hh:mm:ss"), 1)
    Call WriteLogMessage(logger, logEnabled, "", 1)
    Call WriteLogMessage(logger, logEnabled, "LINK ESISTENTI PRIMA DELLA CANCELLAZIONE: " & nLinkEsistenti, 1)
    Call WriteLogMessage(logger, logEnabled, "TOTALE RECORD ESAMINATI: " & nTot, 1)
    Call WriteLogMessage(logger, logEnabled, "LINK CANCELLATI CON SUCCESSO: " & nCancellati, 1)
    Call WriteLogMessage(logger, logEnabled, "ERRORI DURANTE CANCELLAZIONE: " & nErrori, 1)
    Call WriteLogMessage(logger, logEnabled, "========================================================================", 1)
    
    '======================================================================
    ' 21_17[CLOSE_RESOURCES] = Chiusura risorse
    '======================================================================
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set fso = Nothing
    
    '======================================================================
    ' 21_18[CLOSE_LOGGER] = Chiusura logger
    '======================================================================
    Call WriteLogMessage(logger, logEnabled, "[CHIUSURA] File di log salvato in: " & logFilePath, 1)
    
    Set logger = Nothing
    
    '======================================================================
    ' 21_19[MESSAGGIO_FINALE] = MsgBox con esito finale
    '======================================================================
    MsgBox "Cancellazione link completata!" & vbCrLf & vbCrLf & _
           "Link esistenti prima della cancellazione: " & nLinkEsistenti & vbCrLf & _
           "Link cancellati con successo: " & nCancellati & vbCrLf & _
           "Errori: " & nErrori & vbCrLf & vbCrLf & _
           "Log salvato in: " & logFilePath, vbInformation
    
    CancellaTuttiLinkEsistenti_PFunct = True
    Exit Function
    
    '======================================================================
    ' 99)_GESTIONE_ERRORI = Error handler globale con logging
    '======================================================================
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "Errore in CancellaTuttiLinkEsistenti_PFunct: " & Err.Description & vbCrLf & _
               "Numero: " & Err.Number & vbCrLf & _
               "In: " & Err.Source
    
    ' Log dell'errore
    Call WriteLogMessage(logger, logEnabled, "[ERRORE_GLOBALE] " & errorMsg, 3)
    Call WriteLogMessage(logger, logEnabled, "[FINE] Operazione terminata con errore", 1)
    
    ' Cleanup
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not rsCheck Is Nothing Then
        rsCheck.Close
        Set rsCheck = Nothing
    End If
    Set db = Nothing
    Set fso = Nothing
    
    Set logger = Nothing
    
    MsgBox errorMsg, vbCritical
    CancellaTuttiLinkEsistenti_PFunct = False
End Function


'======================================================================
' 22)_FUNZIONE_LOG = WriteLog
'======================================================================
' Scrive messaggi su file di log e finestra Debug
'======================================================================
Private Sub WriteLog(ByVal LogFile As Object, ByVal Message As String)
' 22_01[WRITE_TO_FILE] = Scrive su file di log
    On Error Resume Next
    LogFile.WriteLine Now & " - " & Message
    On Error GoTo 0
    
' 22_02[WRITE_TO_DEBUG] = Scrive su finestra Immediate
    Debug.Print Now & " - " & Message
End Sub


'======================================================================
' 23)_VERIFICA_QUERY = QueryExists
'======================================================================
' Controlla esistenza query in MSysObjects (Type=5)
'======================================================================
Public Function QUERYExists(ByVal queryName As String) As Boolean
    Dim rs As DAO.Recordset
    
' 23_01[OPEN_MSYSOBJECTS] = Cerca query in MSysObjects (Type=5)
    On Error Resume Next
    Set rs = CurrentDb.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=5 AND Name='" & queryName & "'")
    On Error GoTo 0
    
' 23_02[CHECK_RECORD] = Verifica se query esiste
    If Not rs Is Nothing Then
        QUERYExists = Not rs.EOF
' 23_03[CLOSE_RS] = Chiude recordset
        rs.Close
        Set rs = Nothing
    Else
        QUERYExists = False
    End If
End Function


'======================================================================
' 24)_VERIFICA_LINK = LinkExistsInMSysObjects
'======================================================================
' Controlla esistenza link in MSysObjects (Type=6)
'======================================================================
Private Function LinkExistsInMSysObjects(ByVal tableName As String) As Boolean
    Dim rs As DAO.Recordset
    
' 24_01[OPEN_MSYSOBJECTS] = Cerca link in MSysObjects (Type=6)
    On Error Resume Next
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE Type=6 AND Name='" & tableName & "'")
    On Error GoTo 0
    
' 24_02[CHECK_RECORD] = Verifica se link esiste
    If Not rs Is Nothing Then
        LinkExistsInMSysObjects = Not rs.EOF
' 24_03[CLOSE_RS] = Chiude recordset
        rs.Close
        Set rs = Nothing
    Else
        LinkExistsInMSysObjects = False
    End If
End Function


'======================================================================
' 25)_VERIFICA_TABELLA = TableExistsInDatabase
'======================================================================
' Verifica esistenza tabella in database esterno
'======================================================================
Private Function TableExistsInDatabase(ByVal databasePath As String, ByVal tableName As String, ByVal LogFile As Object) As Boolean
    Dim dbExternal As DAO.Database
    Dim rs As DAO.Recordset
    
    Call WriteLog(LogFile, "[APERTURA_DB] Tentativo apertura database: '" & databasePath & "'")
    
' 25_01[OPEN_EXTERNAL_DB] = Tenta apertura database esterno
    On Error Resume Next
    Set dbExternal = OpenDatabase(databasePath)
    
    If Err.Number <> 0 Then
        Call WriteLog(LogFile, "[APERTURA_DB] ERRORE: Impossibile aprire il database!")
        Call WriteLog(LogFile, "[APERTURA_DB] Errore n. " & Err.Number & ": " & Err.Description)
        TableExistsInDatabase = False
        Exit Function
    End If
    On Error GoTo 0
    
    Call WriteLog(LogFile, "[APERTURA_DB] Database aperto con successo.")
    Call WriteLog(LogFile, "[VERIFICA_TABELLA] Ricerca tabella '" & tableName & "' in MSysObjects...")
    
' 25_02[CHECK_TABLE_IN_EXTERNAL] = Cerca tabella nel DB esterno (Type=1,4,6)
    Set rs = dbExternal.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE (Type=1 OR Type=4 OR Type=6) AND Name='" & tableName & "'")
    
    TableExistsInDatabase = Not rs.EOF
    
    If Not TableExistsInDatabase Then
        Call WriteLog(LogFile, "[VERIFICA_TABELLA] Tabella '" & tableName & "' NON TROVATA")
    Else
        Call WriteLog(LogFile, "[VERIFICA_TABELLA] Tabella '" & tableName & "' TROVATA")
    End If
    
' 25_03[CLOSE_EXTERNAL_RS] = Chiude recordset esterno
    rs.Close
' 25_04[CLOSE_EXTERNAL_DB] = Chiude database esterno
    dbExternal.Close
' 25_05[RELEASE_EXTERNAL_OBJECTS] = Rilascia oggetti
    Set rs = Nothing
    Set dbExternal = Nothing
    Call WriteLog(LogFile, "[CHIUSURA_DB] Database chiuso e risorse rilasciate.")
    
End Function


'======================================================================
' 26)_ELIMINAZIONE_LINK = DeleteLink
'======================================================================
' Elimina link tabella da TableDefs
'======================================================================
Private Sub DeleteLink(ByVal tableName As String, ByVal LogFile As Object)
    Dim db As DAO.Database
    
' 26_01[SET_CURRENT_DB] = Imposta database corrente
    Set db = CurrentDb
    
    Call WriteLog(LogFile, "[DELETE_LINK] Eliminazione link '" & tableName & "' da TableDefs...")
    
' 26_02[DELETE_TABLEDEF] = Elimina TableDef dal database
    On Error Resume Next
    db.TableDefs.Delete tableName
    If Err.Number = 0 Then
        Call WriteLog(LogFile, "[DELETE_LINK] Link '" & tableName & "' eliminato con successo.")
    Else
        Call WriteLog(LogFile, "[DELETE_LINK] Link non trovato o errore: " & Err.Description)
    End If
    On Error GoTo 0
    
' 26_03[RELEASE_DB] = Rilascia oggetto database
    Set db = Nothing
End Sub


'======================================================================
' 27)_CREAZIONE_LINK = CreateLink
'======================================================================
' Crea nuovo link tabella verso database esterno
'======================================================================
Private Function CreateLink(ByVal tableName As String, ByVal sourceDatabasePath As String, ByVal LogFile As Object) As Boolean
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    
' 27_01[SET_CURRENT_DB] = Imposta database corrente
    Set db = CurrentDb
    
    Call WriteLog(LogFile, "[CREATE_LINK] Creazione TableDef per '" & tableName & "'...")
    
    On Error GoTo CreateLinkError
    
' 27_02[CREATE_TABLEDEF] = Crea nuovo TableDef per il link
    Set td = db.CreateTableDef(tableName)
    Call WriteLog(LogFile, "[CREATE_LINK] TableDef creato.")
    
' 27_03[SET_CONNECT_STRING] = Imposta stringa di connessione
    td.Connect = ";DATABASE=" & sourceDatabasePath
' 27_04[SET_SOURCE_TABLE] = Imposta nome tabella sorgente
    td.SourceTableName = tableName
    Call WriteLog(LogFile, "[CREATE_LINK] Connect string impostata")
    Call WriteLog(LogFile, "[CREATE_LINK] SourceTableName impostato: '" & tableName & "'")
    
' 27_05[APPEND_TABLEDEF] = Aggiunge TableDef al database corrente
    Call WriteLog(LogFile, "[CREATE_LINK] Appendo TableDef al database...")
    db.TableDefs.Append td
    Call WriteLog(LogFile, "[CREATE_LINK] TableDef appeso con successo.")
    
' 27_06[REFRESH_LINK] = Aggiorna il link
    Call WriteLog(LogFile, "[CREATE_LINK] Refresh del link...")
    td.RefreshLink
    Call WriteLog(LogFile, "[CREATE_LINK] Link aggiornato con successo.")
    
    CreateLink = True
    Call WriteLog(LogFile, "[CREATE_LINK] Link creato correttamente per '" & tableName & "'")
    
' 27_07[RELEASE_OBJECTS] = Rilascia oggetti
    Set td = Nothing
    Set db = Nothing
    Exit Function
    
CreateLinkError:
' 27_99[ERROR_HANDLER] = Gestione errore creazione link
    Call WriteLog(LogFile, "[CREATE_LINK] ERRORE: " & Err.Description)
    CreateLink = False
    Set td = Nothing
    Set db = Nothing
End Function


'======================================================================
' 28)_FUNZIONE_HELPER = WriteLogMessage
'======================================================================
' Scrive un messaggio usando il logger o Debug.Print come fallback
'======================================================================

Private Sub WriteLogMessage(ByVal logger As Object, ByVal logEnabled As Boolean, ByVal Message As String, ByVal level As Integer)
    ' livelli usati nel modulo: 1 = messaggio normale, 2 = avviso, 3 = errore
    ' errore -> file ERRORI + copia nel file INFO (classe MSys_TB12_Cls01_LOG, metodo Scrivi #03.01#)

    ' Scrive sempre su Debug
    Debug.Print Message

    If logEnabled And Not logger Is Nothing Then
        On Error Resume Next
        Select Case level
            Case 2:    logger.Scrivi "CancellaTuttiLinkEsistenti_PFunct", "-", "AVVISO", Message
            Case 3:    logger.Scrivi "CancellaTuttiLinkEsistenti_PFunct", "-", "ERRORE", Message
            Case Else: logger.Scrivi "CancellaTuttiLinkEsistenti_PFunct", "-", "INFO", Message
        End Select
        On Error GoTo 0
    End If

End Sub