Attribute VB_Name = "Msys_TB99_LINK_ Mdl04_IMPORTA_NOMI_OGGETTI"
' =============================================================================
' FUNZIONE: ImportaOggettiDaDatabaseEsterno
' SCOPO:
'   1 - Legge i progetti dalla query Msys_TB99_LINK_Qry99_10_GROUP_COD_SUB_PROGETTO_s
'   2 - Per ogni progetto apre il database esterno
'   3 - Interroga MSysObjects per recuperare QUERY, FORM, REPORT, MODULI, CLASSI
'   4 - Salva i nomi degli oggetti nella tabella detail Msys_TB99_LINK_Object
' =============================================================================

' 01_INIZIO_FUNZIONE                    =    Attiva error handling
' 02_DICHIARAZIONE_VARIABILI            =    Dim variabili
' 03_CONFIGURAZIONE_INIZIALE            =    Database corrente, timer
' 04_PULISCI_TABELLA_DETAIL             =    Cancella i vecchi dati dalla tabella detail
' 05_RECUPERO_PROGETTI                  =    Recupera i progetti dalla query (solo Verifica_s = 'OK')
' 06_PREPARAZIONE_CONTATORI             =    Setup progresso e contatori
' 07_INIZIO_CICLO_PROGETTI              =    Start loop DO WHILE sui progetti
'    08_LEGGI_DATI_PROGETTO             =    Legge COD_SUB_PROGETTO_s e Database
'    09_APRI_DATABASE_ESTERNO           =    Apre il database esterno (con gestione errori)
'    10_VERIFICA_APERTURA               =    Se errore, salta il progetto
'    11_QUERY_SU_MSYSOBJECTS            =    Interroga MSysObjects nel db esterno
'       11a_FILTRO_TYPE                 =    Filtra Type: 5 (QUERY), -32768 (FORM), -32764 (REPORT), -32766 (MODULI), -32761 (CLASSI)
'       11b_ESCLUDI_TMP                 =    Esclude oggetti temporanei (~TMP, ~tmp, ~tmb)
'       11c_ESCLUDI_TABELLE             =    Esclude Type = 1 (TABELLE)
'    12_APRI_RECORDSET_OGGETTI          =    Apre recordset con gli oggetti trovati
'    13_INSERISCI_NEL_DETAIL            =    Cicla e inserisce i record nella tabella detail
'       13a_AddNew                      =    Aggiunge nuovo record
'       13b_Set_COD_SUB_PROGETTO_s      =    Popola campo COD_SUB_PROGETTO_s
'       13c_Set_Type                    =    Popola campo Type (tipo oggetto)
'       13d_Set_Name                    =    Popola campo Name (nome oggetto)
'       13e_Update                      =    Salva il record
'    14_AGGIORNA_CONTATORE_OGGETTI      =    Incrementa contatore lngRecordsAdded
'    15_CHIUDI_RECORDSET_OGGETTI        =    Close e Set Nothing
'    16_CHIUDI_DATABASE_ESTERNO         =    Close e Set Nothing (con gestione errori)
'    17_AGGIORNA_BARRA_PROGRESSO        =    Update meter e gestione errore 7952
'    18_PROSSIMO_PROGETTO               =    MoveNext
' 19_FINE_CICLO_PROGETTI                =    End Loop
' 20_CALCOLO_STATISTICHE                =    Timer, calcolo tempo impiegato
' 21_FORMAZIONE_REPORT                  =    Crea stringa con statistiche finali
' 22_SCRITTURA_SU_LOG                   =    Scrive report su file LOG
' 23_PULISCI_BARRA                      =    SysCmd clear status e remove meter
' 24_MOSTRA_MESSAGGIO                   =    MsgBox con report statistiche
' 25_CLEANUP_NORMALE                    =    Close recordset e Set Nothing
' 26_USCITA_NORMALE                     =    Exit Sub
' 27_GESTIONE_ERRORI                    =    MsgBox errore con dettagli e LOG
' 28_FINE_FUNZIONE                      =    End Sub

'// ========================================================================================================== //
'// ImportaOggettiDaDatabaseEsterno               *** INIZIO ***
'// ========================================================================================================== //

' =============================================================================
' FUNZIONE: ImportaOggettiDaDatabaseEsterno
' SCOPO:
'   1 - Legge i progetti dalla query Msys_TB99_LINK_Qry99_10_GROUP_COD_SUB_PROGETTO_s
'   2 - Per ogni progetto apre il database esterno
'   3 - Interroga MSysObjects per recuperare QUERY, FORM, REPORT, MODULI, CLASSI
'   4 - Salva i nomi degli oggetti nella tabella detail Msys_TB99_LINK_Object
' =============================================================================

Public Sub ImportaOggettiDaDatabaseEsterno()
    
    ' 01_INIZIO_FUNZIONE
    On Error GoTo lbl_GestioneErrori
    
    ' 02_DICHIARAZIONE_VARIABILI
    Dim db As DAO.Database
    Dim dbEsterno As DAO.Database
    Dim rsProgetti As DAO.Recordset
    Dim rsOggetti As DAO.Recordset
    Dim rsDestinazione As DAO.Recordset
    Dim strSQL As String
    Dim strDatabasePath As String
    Dim strCOD_SUB_PROGETTO_s As String
    Dim intCounter As Integer
    Dim intTotal As Integer
    Dim lngRecordsAdded As Long
    Dim dblStartTime As Double
    Dim dblEndTime As Double
    Dim strStatistics As String
    Dim strLogPath As String
    Dim intFileNum As Integer
    Dim strLogLine As String
    Dim strErrMsg As String
    
    ' 03_CONFIGURAZIONE_INIZIALE
    Set db = CurrentDb()
    dblStartTime = Timer
    
    ' Creazione cartella LOGS se non esiste
    strLogPath = CurrentProject.Path & "\LOGS\"
    On Error Resume Next
    MkDir strLogPath
    On Error GoTo lbl_GestioneErrori
    
    ' Apertura file di log con timestamp
    strLogPath = strLogPath & "IMPORTA_NOMI_LOG_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    intFileNum = FreeFile
    Open strLogPath For Output As #intFileNum
    
    ' Intestazione LOG
    Print #intFileNum, "=================================================="
    Print #intFileNum, "LOG IMPORTAZIONE OGGETTI DA DATABASE ESTERNI"
    Print #intFileNum, "Data/Ora: " & Now()
    Print #intFileNum, "=================================================="
    Print #intFileNum, ""
    
    ' Debug.Print
    Debug.Print "=================================================="
    Debug.Print "LOG IMPORTAZIONE OGGETTI DA DATABASE ESTERNI"
    Debug.Print "Data/Ora: " & Now()
    Debug.Print "=================================================="
    
    ' 04_PULISCI_TABELLA_DETAIL
    strLogLine = "04_PULISCI_TABELLA_DETAIL - Cancellazione vecchi dati da Msys_TB99_LINK_Object"
    Print #intFileNum, strLogLine
    Debug.Print strLogLine
    
    On Error Resume Next
    db.Execute "DELETE FROM Msys_TB99_LINK_Object WHERE Name Is Null", dbFailOnError   '// 15/11/2026: toglie solo gli oggetti importati da qui, non quelli legati a una tabella
    If Err.Number <> 0 Then
        strLogLine = "    ATTENZIONE: Errore cancellazione tabella - " & Err.Number & " - " & Err.Description
        Print #intFileNum, strLogLine
        Debug.Print strLogLine
        Err.Clear
    End If
    On Error GoTo lbl_GestioneErrori
    
    ' 05_RECUPERO_PROGETTI
    strLogLine = "05_RECUPERO_PROGETTI - Lettura progetti dalla query"
    Print #intFileNum, strLogLine
    Debug.Print strLogLine
    
    ' Verifica esistenza query
    On Error Resume Next
    Set rsProgetti = db.OpenRecordset("Msys_TB99_LINK_Qry99_10_GROUP_COD_SUB_PROGETTO_s", dbOpenSnapshot)
    If Err.Number <> 0 Then
        strErrMsg = "ERRORE: Query Msys_TB99_LINK_Qry99_10_GROUP_COD_SUB_PROGETTO_s non trovata!" & vbCrLf & _
                    "Errore: " & Err.Number & " - " & Err.Description
        Print #intFileNum, strErrMsg
        Debug.Print strErrMsg
        MsgBox strErrMsg, vbCritical, "Errore Importazione"
        GoTo lbl_CleanupNormale
    End If
    On Error GoTo lbl_GestioneErrori
    
    ' Filtra solo Verifica_s = 'OK'
    rsProgetti.Filter = "Verifica_s = 'OK'"
    Set rsProgetti = rsProgetti.OpenRecordset()
    
    If rsProgetti.EOF Then
        strLogLine = "NESSUN PROGETTO TROVATO (con Verifica_s = 'OK')! Operazione interrotta."
        Print #intFileNum, strLogLine
        Debug.Print strLogLine
        MsgBox "Nessun progetto trovato con Verifica_s = 'OK'!", vbExclamation, "Importa Oggetti"
        GoTo lbl_CleanupNormale
    End If
    
    ' 06_PREPARAZIONE_CONTATORI
    rsProgetti.MoveLast
    intTotal = rsProgetti.RecordCount
    rsProgetti.MoveFirst
    intCounter = 0
    lngRecordsAdded = 0
    
    strLogLine = "06_PREPARAZIONE_CONTATORI - Progetti trovati: " & intTotal
    Print #intFileNum, strLogLine
    Debug.Print strLogLine
    Print #intFileNum, ""
    
    ' Inizializza il meter
    On Error Resume Next
    SysCmd acSysCmdRemoveMeter
    SysCmd acSysCmdInitMeter, "Importazione oggetti da database esterni...", intTotal
    On Error GoTo lbl_GestioneErrori
    
    ' 07_INIZIO_CICLO_PROGETTI
    Do While Not rsProgetti.EOF
        
        ' 08_LEGGI_DATI_PROGETTO
        strCOD_SUB_PROGETTO_s = Nz(rsProgetti!COD_SUB_PROGETTO_s, "")
        strDatabasePath = Nz(rsProgetti!DB_COMPLETO, "")
        
        intCounter = intCounter + 1
        
        strLogLine = "[" & intCounter & "/" & intTotal & "] Progetto: " & strCOD_SUB_PROGETTO_s
        Print #intFileNum, strLogLine
        Debug.Print strLogLine
        
        strLogLine = "    Database: " & strDatabasePath
        Print #intFileNum, strLogLine
        Debug.Print strLogLine
        
        SysCmd acSysCmdSetStatus, "Elaborazione (" & intCounter & "/" & intTotal & "): " & strCOD_SUB_PROGETTO_s
        
        ' Verifica se il file esiste
        If Dir(strDatabasePath) = "" Then
            strLogLine = "    ERRORE: File database non trovato!"
            Print #intFileNum, strLogLine
            Debug.Print strLogLine
            GoTo lbl_SkipProgetto
        End If
        
        ' 09_APRI_DATABASE_ESTERNO
        On Error Resume Next
        Set dbEsterno = DBEngine.OpenDatabase(strDatabasePath)
        
        ' 10_VERIFICA_APERTURA
        If Err.Number <> 0 Then
            strLogLine = "    ERRORE: Impossibile aprire il database! Errore: " & Err.Number & " - " & Err.Description
            Print #intFileNum, strLogLine
            Debug.Print strLogLine
            Err.Clear
            GoTo lbl_SkipProgetto
        End If
        On Error GoTo lbl_GestioneErrori
        
        strLogLine = "    Database aperto con successo"
        Print #intFileNum, strLogLine
        Debug.Print strLogLine
        
        ' 11_QUERY_SU_MSYSOBJECTS
        ' NOTA: Usa dbEsterno per la connessione al database esterno
        On Error Resume Next
        
        ' Costruisci la query
        strSQL = "SELECT " & _
                 "'" & strCOD_SUB_PROGETTO_s & "' AS COD_SUB_PROGETTO_s, " & _
                 "Type, " & _
                 "Name " & _
                 "FROM MSysObjects " & _
                 "WHERE Type IN (5, -32768, -32764, -32766, -32761) " & _
                 "AND Name IS NOT NULL " & _
                 "AND Left(Name, 4) <> '~TMP' " & _
                 "AND Left(Name, 4) <> '~tmb' " & _
                 "AND Left(Name, 4) <> '~tmp' " & _
                 "ORDER BY Name"
        
        ' 12_APRI_RECORDSET_OGGETTI - Usando dbEsterno
        Set rsOggetti = dbEsterno.OpenRecordset(strSQL, dbOpenSnapshot)
        
        If Err.Number <> 0 Then
            strLogLine = "    ERRORE nella query MSysObjects! Errore: " & Err.Number & " - " & Err.Description
            Print #intFileNum, strLogLine
            Debug.Print strLogLine
            Err.Clear
            GoTo lbl_CloseExternal
        End If
        On Error GoTo lbl_GestioneErrori
        
      ' 13_INSERISCI_NEL_DETAIL - Versione con Database e Tipo Oggetto
            If Not rsOggetti.EOF Then
                Dim intOggettiTrovati As Integer
                Dim intInseriti As Integer
                Dim strInsertSQL As String
                Dim strNomeOggetto As String
                Dim strTipoOggetto As String
                Dim vEsclusioni As Variant
                Dim i As Integer
                Dim blnEscludi As Boolean
                
                ' Lista dei prefissi da escludere
                vEsclusioni = Array("~sq_c", "~sq_f", "~sq_h", "~sq_t", "~sq_v", "~TMP", "~tmp", "~tmb")
                
                intOggettiTrovati = 0
                intInseriti = 0
                
                Do While Not rsOggetti.EOF
                    intOggettiTrovati = intOggettiTrovati + 1
                    strNomeOggetto = Nz(rsOggetti!Name, "")
                    blnEscludi = False
                    
                    ' Determina il TIPO_OGGETTO_s in base al Type
                    Select Case rsOggetti!Type
                        Case 5
                            strTipoOggetto = "QUERY"
                        Case -32768
                            strTipoOggetto = "FORMS"
                        Case -32764
                            strTipoOggetto = "REPORTS"
                        Case -32766
                            strTipoOggetto = "MACRO"
                        Case -32761
                            strTipoOggetto = "MODULES"
                        Case Else
                            strTipoOggetto = ""
                    End Select
                    
                    ' Filtra solo i Type validi (non vuoti)
                    If strTipoOggetto <> "" Then
                        
                        ' Controlla se il nome inizia con uno dei prefissi da escludere
                        For i = LBound(vEsclusioni) To UBound(vEsclusioni)
                            If Left(strNomeOggetto, Len(vEsclusioni(i))) = vEsclusioni(i) Then
                                blnEscludi = True
                                Exit For
                            End If
                        Next i
                        
                        If Not blnEscludi Then
                       ' INSERT con Database e TIPO_OGGETTO_s
                            strInsertSQL = "INSERT INTO Msys_TB99_LINK_Object ([Database], Database_s, TIPO_OGGETTO_s, NOME_OGGETTO_s) VALUES (" & _
                               "'" & Replace(Mid(strDatabasePath, InStrRev(strDatabasePath, "\") + 1), "'", "''") & "', " & _
                               "'" & Replace(Left(strDatabasePath, InStrRev(strDatabasePath, "\")), "'", "''") & "', " & _
                               "'" & strTipoOggetto & "', " & _
                               "'" & Replace(strNomeOggetto, "'", "''") & "')"
                                            
                            db.Execute strInsertSQL, dbFailOnError
                            intInseriti = intInseriti + 1
                            lngRecordsAdded = lngRecordsAdded + 1
                        End If
                    End If
                   
                    
                    rsOggetti.MoveNext
                Loop
                
                strLogLine = "    Oggetti trovati: " & intOggettiTrovati & " - Importati: " & intInseriti
                Print #intFileNum, strLogLine
                Debug.Print strLogLine
                
            Else
                strLogLine = "    Nessun oggetto trovato"
                Print #intFileNum, strLogLine
                Debug.Print strLogLine
            End If

        
        
      
        ' 14_CHIUDI_RECORDSET_OGGETTI
lbl_CloseExternal:
        On Error Resume Next
        If Not rsOggetti Is Nothing Then
            rsOggetti.Close
            Set rsOggetti = Nothing
        End If
        
        ' 15_CHIUDI_DATABASE_ESTERNO
        If Not dbEsterno Is Nothing Then
            dbEsterno.Close
            Set dbEsterno = Nothing
        End If
        On Error GoTo lbl_GestioneErrori
        
lbl_SkipProgetto:
        ' 16_AGGIORNA_BARRA_PROGRESSO
        On Error Resume Next
        SysCmd acSysCmdUpdateMeter, intCounter
        If Err.Number = 7952 Then
            Err.Clear
            SysCmd acSysCmdRemoveMeter
            SysCmd acSysCmdInitMeter, "Importazione oggetti...", intTotal
            SysCmd acSysCmdUpdateMeter, intCounter
        End If
        On Error GoTo lbl_GestioneErrori
        
        Print #intFileNum, ""
        
        ' 17_PROSSIMO_PROGETTO
        rsProgetti.MoveNext
        
    Loop
    
    ' 18_FINE_CICLO_PROGETTI
    
    ' 19_CALCOLO_STATISTICHE
    dblEndTime = Timer
    Dim dblElapsedSeconds As Double
    dblElapsedSeconds = dblEndTime - dblStartTime
    
    ' 20_FORMAZIONE_REPORT
    strStatistics = "========== STATISTICHE IMPORTAZIONE OGGETTI ==========" & vbCrLf & vbCrLf & _
                    "Importazione completata con successo!" & vbCrLf & vbCrLf & _
                    "DATI GENERALI:" & vbCrLf & _
                    "   - Progetti elaborati: " & intTotal & vbCrLf & _
                    "   - Oggetti importati: " & Format(lngRecordsAdded, "#,##0") & vbCrLf & vbCrLf & _
                    "TIPI OGGETTI IMPORTATI:" & vbCrLf & _
                    "   - Type 5: QUERY" & vbCrLf & _
                    "   - Type -32768: FORM" & vbCrLf & _
                    "   - Type -32764: REPORT" & vbCrLf & _
                    "   - Type -32766: MODULI" & vbCrLf & _
                    "   - Type -32761: CLASSI" & vbCrLf & vbCrLf & _
                    "TEMPI:" & vbCrLf & _
                    "   - Tempo impiegato: " & Format(dblElapsedSeconds, "0.00") & " secondi" & vbCrLf & _
                    "   - Velocita media: " & Format(lngRecordsAdded / IIf(dblElapsedSeconds = 0, 1, dblElapsedSeconds), "0") & " record/sec" & vbCrLf & vbCrLf & _
                    "DESTINAZIONE:" & vbCrLf & _
                    "   - Tabella: Msys_TB99_LINK_Object" & vbCrLf & _
                    "   - Operazione riuscita: SI" & vbCrLf & vbCrLf & _
                    "LOG FILE:" & vbCrLf & _
                    "   - Percorso: " & strLogPath
    
    ' 21_SCRITTURA_SU_LOG
    Print #intFileNum, ""
    Print #intFileNum, "=================================================="
    Print #intFileNum, "STATISTICHE FINALI"
    Print #intFileNum, "=================================================="
    Print #intFileNum, ""
    Print #intFileNum, "Progetti elaborati: " & intTotal
    Print #intFileNum, "Oggetti importati: " & Format(lngRecordsAdded, "#,##0")
    Print #intFileNum, "Tempo impiegato: " & Format(dblElapsedSeconds, "0.00") & " secondi"
    Print #intFileNum, "Velocita media: " & Format(lngRecordsAdded / IIf(dblElapsedSeconds = 0, 1, dblElapsedSeconds), "0") & " record/sec"
    Print #intFileNum, ""
    Print #intFileNum, "=================================================="
    Print #intFileNum, "FINE LOG"
    Print #intFileNum, "=================================================="
    
    ' Debug.Print finali
    Debug.Print ""
    Debug.Print "=================================================="
    Debug.Print "STATISTICHE FINALI"
    Debug.Print "=================================================="
    Debug.Print "Progetti elaborati: " & intTotal
    Debug.Print "Oggetti importati: " & Format(lngRecordsAdded, "#,##0")
    Debug.Print "Tempo impiegato: " & Format(dblElapsedSeconds, "0.00") & " secondi"
    Debug.Print "LOG salvato in: " & strLogPath
    Debug.Print "=================================================="
    
    ' 22_PULISCI_BARRA
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    SysCmd acSysCmdRemoveMeter
    On Error GoTo lbl_GestioneErrori
    
    ' 23_MOSTRA_MESSAGGIO
    MsgBox strStatistics, vbInformation, "Report Importazione Oggetti - " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
    
    ' 24_CLEANUP_NORMALE
lbl_CleanupNormale:
    On Error Resume Next
    If Not rsProgetti Is Nothing Then
        rsProgetti.Close
        Set rsProgetti = Nothing
    End If
    If Not rsOggetti Is Nothing Then
        rsOggetti.Close
        Set rsOggetti = Nothing
    End If
    If Not dbEsterno Is Nothing Then
        dbEsterno.Close
        Set dbEsterno = Nothing
    End If
    Set db = Nothing
    SysCmd acSysCmdClearStatus
    SysCmd acSysCmdRemoveMeter
    
    ' Chiusura file LOG
    If intFileNum > 0 Then
        Close #intFileNum
    End If
    
    ' 25_USCITA_NORMALE
    Exit Sub
    
    ' 26_GESTIONE_ERRORI
lbl_GestioneErrori:
    Dim strErrDetails As String
    strErrDetails = "Errore: " & Err.Number & " - " & Err.Description & vbCrLf & _
                    "Progetto: " & strCOD_SUB_PROGETTO_s & vbCrLf & _
                    "Database: " & strDatabasePath & vbCrLf & _
                    "Progetti elaborati: " & intCounter
    
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    SysCmd acSysCmdRemoveMeter
    
    ' Scrive errore su LOG
    If intFileNum > 0 Then
        Print #intFileNum, ""
        Print #intFileNum, "=================================================="
        Print #intFileNum, "ERRORE! Operazione interrotta"
        Print #intFileNum, "=================================================="
        Print #intFileNum, strErrDetails
        Close #intFileNum
    End If
    
    ' Debug.Print errore
    Debug.Print ""
    Debug.Print "=================================================="
    Debug.Print "ERRORE! Operazione interrotta"
    Debug.Print "=================================================="
    Debug.Print strErrDetails
    Debug.Print "=================================================="
    
    MsgBox "ERRORE durante l'importazione!" & vbCrLf & vbCrLf & _
           strErrDetails & vbCrLf & vbCrLf & _
           "LOG salvato in: " & strLogPath, _
           vbCritical, "Errore Importazione"
    
    Resume lbl_CleanupNormale
    
    ' 27_FINE_FUNZIONE
End Sub

'// ========================================================================================================== //
'// NOTE:
'//   - La tabella master (Msys_TB99_LINK) contiene già le TABELLE collegate
'//   - La tabella detail (Msys_TB99_LINK_Object) salva QUERY, FORM, REPORT, MODULI, CLASSI
'//   - Type 1 (TABELLE) viene escluso per evitare duplicati
'//   - I LOG vengono salvati nella cartella LOGS con formato IMPORTA_NOMI_LOG_YYYYMMDD_HHNNSS.txt
'//   - Tutte le operazioni vengono tracciate in DEBUG (finestra immediata VBA)
'// ========================================================================================================== //









'// ========================================================================================================== //
'// ImportaOggettiDaDatabaseEsterno               *** FINE ***
'// ========================================================================================================== //

'// NOTE:
'//   - La tabella master (Msys_TB99_LINK) contiene già le TABELLE collegate
'//   - La tabella detail (Msys_TB99_LINK_Object) salva QUERY, FORM, REPORT, MODULI, CLASSI
'//   - Type 1 (TABELLE) viene escluso per evitare duplicati
'//   - I database esterni devono essere accessibili e non protetti da password
'//   - La routine cancella prima i vecchi dati dalla tabella detail
'//   - La barra di avanzamento mostra il progresso sui progetti (non sui singoli oggetti)
'//   - MSysObjects è una tabella di sistema, richiede permessi di lettura
'//   - I LOG vengono salvati nella cartella LOGS con formato IMPORTA_NOMI_LOG_YYYYMMDD_HHNNSS.txt
'//   - Tutte le operazioni vengono tracciate in DEBUG (finestra immediata VBA)
'// ========================================================================================================== //