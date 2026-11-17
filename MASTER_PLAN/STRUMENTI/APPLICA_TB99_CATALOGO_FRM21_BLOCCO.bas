


'//=====================================================================================================================//
'//     @TB99@CATALOGO   IMPORTAZIONE DEGLI OGGETTI DAL CATALOGO TB99                         *** INIZIO ***
'//     2026.11.17 - Al posto delle collection scritte nel codice, gli oggetti si leggono dalle tabelle fisiche:
'//
'//         Msys_TB99_LINK          = 1 riga per tabella collegata: Name + database in cui si trova (Database_s & Database)
'//         Msys_TB99_LINK_Object   = 1 riga per oggetto (QUERY, FORMS, REPORTS, MACRO, MODULES) che usa la tabella
'//                                   e che sta nello stesso database della tabella (legame sul campo Name)
'//
'//     ATTIVAZIONE : automatica, quando il database corrente contiene le tabelle locali Msys_TB99_LINK e
'//                   Msys_TB99_LINK_Object (cioe in MSys_TB99_LINK.mdb). Negli altri database la form funziona
'//                   come prima, con le collection.
'//
'//     SCELTA      : la combinata 01 elenca i database sorgente del catalogo (mini-progetto, file, n. tabelle).
'//                   Gli oggetti del progetto scelto sono:
'//                     - TABELLE  = le righe di Msys_TB99_LINK di quel database        (diventano COLLEGAMENTI)
'//                     - OGGETTI  = le righe di Msys_TB99_LINK_Object legate a quelle tabelle e nello stesso database
'//
'//     ORDINE      : 1) CANCELLA nel database corrente   link -> form -> report -> macro -> query -> moduli
'//                   2) IMPORTA dal database sorgente    link -> query -> form -> report -> macro -> moduli
'//
'//     SICUREZZA   : - le tabelle si COLLEGANO, non si importano mai; una tabella LOCALE con lo stesso nome non viene toccata
'//                   - non si toccano: questa form, la progress bar, le form UTILITA_Msys_Frm21_*, gli oggetti Msys_TB99_*
'//                   - un oggetto aperto non viene cancellato ne sostituito
'//                   - il database sorgente non puo essere il database corrente
'//
'//     USO DA CODICE (senza messaggi):
'//         s = Forms("UTILITA_Msys_Frm21_IMPORTA_OGGETTI").TB99_Esegui_pf("CANCELLA+IMPORTA", "TUTTI", "C:\...\MSys_DF01.mdb", True)
'//         par_Azione_s = "CANCELLA" | "IMPORTA" | "CANCELLA+IMPORTA"
'//         par_Tipo_s   = "TUTTI" | "TABLES" | "QUERY" | "FORMS" | "REPORTS" | "MACRO" | "MODULES"
'//=====================================================================================================================//


'//MODALITA CATALOGO: True se nel database corrente esistono, locali, le due tabelle del catalogo
Public Function TB99_Attivo_b() As Boolean
    Dim dbC As DAO.Database
    Dim tdfL As DAO.TableDef
    Dim tdfO As DAO.TableDef

    On Error GoTo TB99_Attivo_b_Err
    TB99_Attivo_b = False

    Set dbC = CurrentDb
    Set tdfL = dbC.TableDefs("Msys_TB99_LINK")
    Set tdfO = dbC.TableDefs("Msys_TB99_LINK_Object")
    If tdfL.Connect = "" And tdfO.Connect = "" Then TB99_Attivo_b = True

TB99_Attivo_b_Err:
    Set tdfL = Nothing
    Set tdfO = Nothing
    Set dbC = Nothing
End Function


'//SQL DELLA COMBINATA 01: un database sorgente per riga
'//  colonna 0 (nascosta, valore) = percorso completo   colonna 1 = mini-progetto   colonna 2 = file   colonna 3 = n. tabelle
Private Function TB99_SqlCmb01_pf() As String
    TB99_SqlCmb01_pf = "SELECT L.Database_s & L.Database AS DB_COMPLETO, L.COD_SUB_PROGETTO_s AS PROGETTO, " & _
                       "L.Database AS DB, Count(L.Name) AS N_TAB " & _
                       "FROM Msys_TB99_LINK AS L " & _
                       "WHERE L.Database Is Not Null AND L.Database_s Is Not Null " & _
                       "GROUP BY L.Database_s, L.Database, L.COD_SUB_PROGETTO_s " & _
                       "ORDER BY L.COD_SUB_PROGETTO_s, L.Database;"
End Function


'//FORM_LOAD IN MODALITA CATALOGO: imposta le combinate e NON importa nulla all'apertura
Private Sub TB99_Form_Load_Sub()

    On Error Resume Next

    Call DisabilitaButton_Sub
    Call Disabilito_sourceDBPath_Sub
    Call DisabilitaButton_CMB_02_Sub

    Me.ctr_AVVISO_ABILITATO_TXT = True
    Me.CurrentProject_txt = Application.CurrentProject.Name
    Me.LBL_TIT01_TXT.Caption = "IMPORTA OGGETTI DAL CATALOGO TB99 (Msys_TB99_LINK + Msys_TB99_LINK_Object) " & _
                               "NEL DATABASE: " & Application.CurrentProject.Name
    Me.Cmb_01_Lbl.Caption = "DATABASE SORGENTE"
    Me.Cmb_02_Lbl.Caption = "OGGETTI"

    '//COMBINATA 01 = i database del catalogo
    Err.Clear
    Me.Cmb_01_txt.RowSourceType = "Table/Query"
    If Err.Number <> 0 Then
        Err.Clear
        Me.Cmb_01_txt.RowSourceType = "Tabella/query"
    End If
    Me.Cmb_01_txt.ColumnCount = 4
    Me.Cmb_01_txt.BoundColumn = 1
    Me.Cmb_01_txt.ColumnWidths = "0;1700;3000;500"
    Me.Cmb_01_txt.ListWidth = 5400
    Me.Cmb_01_txt.ListRows = 24
    Me.Cmb_01_txt.RowSource = TB99_SqlCmb01_pf()
    Me.Cmb_01_txt.Value = Null

    '//COMBINATA 02 = una sola voce
    Me.Cmb_02_txt.RowSourceType = "Value List"
    Me.Cmb_02_txt.RowSource = "OGGETTI DAL CATALOGO TB99;"
    Me.Cmb_02_txt.Value = Null

    Me.sourceDBPath_s_Txt = ""
    Me.TXT_OBJECT_ESTERNI = "SCEGLI NELLA COMBINATA 'DATABASE SORGENTE' IL DATABASE DA CUI IMPORTARE GLI OGGETTI." & vbCrLf & _
                            "Gli oggetti vengono letti dalle tabelle Msys_TB99_LINK e Msys_TB99_LINK_Object."

    Call Imposta_Colore_Sub

End Sub


'//COMBINATA 01 DOPO AGGIORNAMENTO IN MODALITA CATALOGO
Private Sub TB99_Cmb01_AfterUpdate_Sub()

    Dim path_s As String

    path_s = Nz(Me.Cmb_01_txt.Value, "")

    Call DisabilitaButton_Sub

    If path_s = "" Then
        Me.sourceDBPath_s_Txt = ""
        Me.TXT_OBJECT_ESTERNI = ""
        Call Disabilito_sourceDBPath_Sub
        Call DisabilitaButton_CMB_02_Sub
        Exit Sub
    End If

    Me.sourceDBPath_s_Txt = path_s
    Me.CODICE_PROGETTO_TXT = Nz(Me.Cmb_01_txt.Column(1), "")
    Me.LBL_TIT01_TXT.Caption = "CATALOGO TB99 | progetto ---> " & Nz(Me.Cmb_01_txt.Column(1), "") & _
                               " | database: " & path_s

    Call Abilito_sourceDBPath_Sub
    Call AbilitaButton_CMB_02_Sub
    Me.Cmb_02_txt.Value = "OGGETTI DAL CATALOGO TB99"
    Call TB99_AbilitaComandi_Sub

    '//elenco degli oggetti che verranno cancellati e importati
    Me.TXT_OBJECT_ESTERNI = TB99_Riepilogo_pf(path_s)

End Sub


'//COMBINATA 02 DOPO AGGIORNAMENTO IN MODALITA CATALOGO
Private Sub TB99_Cmb02_AfterUpdate_Sub()
    If Nz(Me.Cmb_02_txt.Value, "") = "OGGETTI DAL CATALOGO TB99" And Nz(Me.sourceDBPath_s_Txt, "") <> "" Then
        Call TB99_AbilitaComandi_Sub
    Else
        Call DisabilitaButton_Sub
    End If
End Sub


'//ABILITA TUTTI I COMANDI IMPORTA / CANCELLA
Private Sub TB99_AbilitaComandi_Sub()
    Call AbilitaButton_TABLE_Sub
    Call AbilitaButton_QUERY_Sub
    Call AbilitaButton_FORM_Sub
    Call AbilitaButton_REPORT_Sub
    Call AbilitaButton_MACRO_Sub
    Call AbilitaButton_MODULI_Sub
    Call AbilitaButton_IMPORTA_TUTTO_Sub
    Call AbilitaButton_DELETE_TUTTO_Sub
End Sub


'//OGGETTI DA NON TOCCARE MAI: questa form, la progress bar, le UTILITA_Msys_Frm21_*, gli oggetti del catalogo Msys_TB99_*
Private Function TB99_Protetto_b(par_Nome_s As String) As Boolean
    Dim n As String
    Dim nomeParent_s As String

    n = LCase(Trim(par_Nome_s))
    TB99_Protetto_b = False

    If n = "" Then TB99_Protetto_b = True: Exit Function
    If n = LCase(Me.Name) Then TB99_Protetto_b = True: Exit Function
    If Left(n, 19) = "utilita_msys_frm21_" Then TB99_Protetto_b = True: Exit Function
    If Left(n, 10) = "msys_tb99_" Then TB99_Protetto_b = True: Exit Function
    If Left(n, 4) = "~tmp" Or Left(n, 3) = "~sq" Then TB99_Protetto_b = True: Exit Function

    '//la maschera principale che contiene questa form come sottomaschera
    On Error Resume Next
    nomeParent_s = ""
    nomeParent_s = LCase(Me.Parent.Name)
    On Error GoTo 0
    If nomeParent_s <> "" And n = nomeParent_s Then TB99_Protetto_b = True
End Function


'//LA COLLECTION DEGLI OGGETTI LETTA DAL CATALOGO (sostituisce le collection scritte nel codice)
'//  par_Tipo_s = "TABLES" | "QUERY" | "FORMS" | "REPORTS" | "MACRO" | "MODULES"
'//  par_SourcePath_s = percorso completo del database sorgente (vuoto = casella sourceDBPath_s_Txt)
Public Function TB99_Collection_pf(par_Tipo_s As String, Optional par_SourcePath_s As String = "") As Collection

    Dim col As Collection
    Dim dbC As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql_s As String
    Dim path_s As String
    Dim nome_s As String

    Set col = New Collection
    Set TB99_Collection_pf = col

    path_s = par_SourcePath_s
    If path_s = "" Then path_s = Nz(Me.sourceDBPath_s_Txt, "")
    If path_s = "" Then Exit Function
    path_s = Replace(path_s, "'", "''")

    If UCase(par_Tipo_s) = "TABLES" Then
        sql_s = "SELECT DISTINCT L.Name FROM Msys_TB99_LINK AS L " & _
                "WHERE L.Database_s & L.Database = '" & path_s & "' " & _
                "ORDER BY L.Name;"
    Else
        sql_s = "SELECT DISTINCT O.NOME_OGGETTO_s FROM Msys_TB99_LINK AS L " & _
                "INNER JOIN Msys_TB99_LINK_Object AS O ON L.Name = O.Name " & _
                "WHERE L.Database_s & L.Database = '" & path_s & "' " & _
                "AND O.Database_s & O.Database = '" & path_s & "' " & _
                "AND O.TIPO_OGGETTO_s = '" & Replace(UCase(par_Tipo_s), "'", "''") & "' " & _
                "ORDER BY O.NOME_OGGETTO_s;"
    End If

    Set dbC = CurrentDb
    Set rs = dbC.OpenRecordset(sql_s, dbOpenSnapshot)
    Do While Not rs.EOF
        nome_s = Nz(rs.Fields(0).Value, "")
        If Not TB99_Protetto_b(nome_s) Then
            On Error Resume Next
            col.Add nome_s, LCase(nome_s)          '//la chiave evita i doppioni
            On Error GoTo 0
        End If
        rs.MoveNext
    Loop
    rs.Close

    Set rs = Nothing
    Set dbC = Nothing
    Set TB99_Collection_pf = col

End Function


'//RIEPILOGO TESTUALE DEGLI OGGETTI DEL PROGETTO SCELTO
Private Function TB99_Riepilogo_pf(par_SourcePath_s As String) As String
    Dim tipi_v As Variant
    Dim t As Variant
    Dim col As Collection
    Dim v As Variant
    Dim s As String

    tipi_v = Array("TABLES", "QUERY", "FORMS", "REPORTS", "MACRO", "MODULES")

    s = "DATABASE SORGENTE: " & par_SourcePath_s & vbCrLf
    s = s & "DATABASE CORRENTE: " & CurrentProject.FullName & vbCrLf
    s = s & "IMPORTA TUTTO = prima CANCELLA nel database corrente, poi IMPORTA (le tabelle diventano collegamenti)" & vbCrLf
    s = s & String(100, "-") & vbCrLf

    For Each t In tipi_v
        Set col = TB99_Collection_pf(CStr(t), par_SourcePath_s)
        s = s & TB99_NomeTipo_pf(CStr(t)) & " (" & col.Count & ")" & vbCrLf
        For Each v In col
            s = s & "      " & v & vbCrLf
        Next v
    Next t

    TB99_Riepilogo_pf = s
End Function


'//DESCRIZIONE DEL TIPO
Private Function TB99_NomeTipo_pf(par_Tipo_s As String) As String
    Select Case UCase(par_Tipo_s)
        Case "TABLES":  TB99_NomeTipo_pf = "TABELLE COLLEGATE"
        Case "QUERY":   TB99_NomeTipo_pf = "QUERY"
        Case "FORMS":   TB99_NomeTipo_pf = "FORM"
        Case "REPORTS": TB99_NomeTipo_pf = "REPORT"
        Case "MACRO":   TB99_NomeTipo_pf = "MACRO"
        Case "MODULES": TB99_NomeTipo_pf = "MODULI"
        Case Else:      TB99_NomeTipo_pf = par_Tipo_s
    End Select
End Function


'//TIPO DEL CATALOGO --> COSTANTE ACCESS (AcObjectType)
Private Function TB99_AcTipo_pf(par_Tipo_s As String) As Integer
    Select Case UCase(par_Tipo_s)
        Case "TABLES":  TB99_AcTipo_pf = acTable
        Case "QUERY":   TB99_AcTipo_pf = acQuery
        Case "FORMS":   TB99_AcTipo_pf = acForm
        Case "REPORTS": TB99_AcTipo_pf = acReport
        Case "MACRO":   TB99_AcTipo_pf = acMacro
        Case "MODULES": TB99_AcTipo_pf = acModule
        Case Else:      TB99_AcTipo_pf = -1
    End Select
End Function


'//ESISTE L'OGGETTO NEL DATABASE? (tabelle con TableDefs, gli altri con MSysObjects)
Private Function TB99_Esiste_b(par_db As DAO.Database, par_Tipo_s As String, par_Nome_s As String) As Boolean
    Dim rs As DAO.Recordset
    Dim tdfE As DAO.TableDef
    Dim tipoSys_s As String

    TB99_Esiste_b = False
    On Error GoTo TB99_Esiste_b_Err

    If UCase(par_Tipo_s) = "TABLES" Then
        Set tdfE = par_db.TableDefs(par_Nome_s)
        TB99_Esiste_b = True
        Set tdfE = Nothing
        Exit Function
    End If

    Select Case UCase(par_Tipo_s)
        Case "QUERY":   tipoSys_s = "5"
        Case "FORMS":   tipoSys_s = "-32768"
        Case "REPORTS": tipoSys_s = "-32764"
        Case "MACRO":   tipoSys_s = "-32766"
        Case "MODULES": tipoSys_s = "-32761"
        Case Else:      Exit Function
    End Select

    Set rs = par_db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = " & tipoSys_s & _
                                  " AND Name = '" & Replace(par_Nome_s, "'", "''") & "'", dbOpenSnapshot)
    TB99_Esiste_b = Not rs.EOF
    rs.Close
    Set rs = Nothing
    Exit Function

TB99_Esiste_b_Err:
    TB99_Esiste_b = False
End Function


'//CANCELLA UN OGGETTO NEL DATABASE CORRENTE - restituisce l'esito
Private Function TB99_Cancella_pf(par_Tipo_s As String, par_Nome_s As String, ByRef par_Fatto_b As Boolean) As String

    Dim dbC As DAO.Database
    Dim tdfC As DAO.TableDef

    par_Fatto_b = False
    On Error GoTo TB99_Cancella_pf_Err

    Set dbC = CurrentDb

    If Not TB99_Esiste_b(dbC, par_Tipo_s, par_Nome_s) Then
        TB99_Cancella_pf = "NON PRESENTE"
        GoTo TB99_Cancella_pf_Exit
    End If

    If SysCmd(acSysCmdGetObjectState, TB99_AcTipo_pf(par_Tipo_s), par_Nome_s) <> 0 Then
        TB99_Cancella_pf = "APERTO: NON TOCCATO"
        GoTo TB99_Cancella_pf_Exit
    End If

    If UCase(par_Tipo_s) = "TABLES" Then
        Set tdfC = dbC.TableDefs(par_Nome_s)
        If tdfC.Connect = "" Then
            TB99_Cancella_pf = "TABELLA LOCALE: NON TOCCATA"
            GoTo TB99_Cancella_pf_Exit
        End If
        Set tdfC = Nothing
        dbC.TableDefs.Delete par_Nome_s
        TB99_Cancella_pf = "LINK CANCELLATO"
    Else
        DoCmd.DeleteObject TB99_AcTipo_pf(par_Tipo_s), par_Nome_s
        TB99_Cancella_pf = "CANCELLATO"
    End If
    par_Fatto_b = True

TB99_Cancella_pf_Exit:
    Set tdfC = Nothing
    Set dbC = Nothing
    Exit Function

TB99_Cancella_pf_Err:
    TB99_Cancella_pf = "ERRORE " & Err.Number & ": " & Err.Description
    Resume TB99_Cancella_pf_Exit
End Function


'//IMPORTA UN OGGETTO DAL DATABASE SORGENTE (le tabelle vengono COLLEGATE) - restituisce l'esito
Private Function TB99_Importa_pf(par_dbSrc As DAO.Database, par_SourcePath_s As String, par_Tipo_s As String, _
                                 par_Nome_s As String, ByRef par_Fatto_b As Boolean) As String

    Dim dbC As DAO.Database
    Dim tdfS As DAO.TableDef
    Dim tdfN As DAO.TableDef
    Dim connect_s As String
    Dim sorgente_s As String
    Dim fatto_b As Boolean
    Dim esito_s As String

    par_Fatto_b = False
    On Error GoTo TB99_Importa_pf_Err

    If Not TB99_Esiste_b(par_dbSrc, par_Tipo_s, par_Nome_s) Then
        TB99_Importa_pf = "NON TROVATO NEL DATABASE SORGENTE"
        Exit Function
    End If

    Set dbC = CurrentDb

    '//se esiste gia nel database corrente lo tolgo (stessa regola della cancellazione)
    If TB99_Esiste_b(dbC, par_Tipo_s, par_Nome_s) Then
        esito_s = TB99_Cancella_pf(par_Tipo_s, par_Nome_s, fatto_b)
        If Not fatto_b Then
            TB99_Importa_pf = "NON IMPORTATO (" & esito_s & ")"
            GoTo TB99_Importa_pf_Exit
        End If
    End If

    If UCase(par_Tipo_s) = "TABLES" Then
        '//COLLEGAMENTO: se nel sorgente la tabella e' a sua volta un link, punto direttamente alla tabella fisica
        Set tdfS = par_dbSrc.TableDefs(par_Nome_s)
        If tdfS.Connect <> "" Then
            connect_s = tdfS.Connect
            sorgente_s = tdfS.SourceTableName
        Else
            connect_s = ";DATABASE=" & par_SourcePath_s
            sorgente_s = par_Nome_s
        End If
        Set tdfS = Nothing

        Set tdfN = dbC.CreateTableDef(par_Nome_s)
        tdfN.Connect = connect_s
        tdfN.SourceTableName = sorgente_s
        dbC.TableDefs.Append tdfN
        TB99_Importa_pf = IIf(fatto_b, "RICOLLEGATA  ", "COLLEGATA  ") & connect_s
    Else
        DoCmd.TransferDatabase acImport, "Microsoft Access", par_SourcePath_s, TB99_AcTipo_pf(par_Tipo_s), par_Nome_s, par_Nome_s
        TB99_Importa_pf = IIf(fatto_b, "SOSTITUITO", "IMPORTATO")
    End If
    par_Fatto_b = True

TB99_Importa_pf_Exit:
    Set tdfN = Nothing
    Set tdfS = Nothing
    Set dbC = Nothing
    Exit Function

TB99_Importa_pf_Err:
    TB99_Importa_pf = "ERRORE " & Err.Number & ": " & Err.Description
    Resume TB99_Importa_pf_Exit
End Function


'//PROCEDURA PRINCIPALE: CANCELLA e/o IMPORTA gli oggetti del progetto scelto
'//  par_Azione_s     = "CANCELLA" | "IMPORTA" | "CANCELLA+IMPORTA"
'//  par_Tipo_s       = "TUTTI" oppure un tipo ("TABLES", "QUERY", "FORMS", "REPORTS", "MACRO", "MODULES")
'//  par_SourcePath_s = percorso completo del database sorgente (vuoto = casella sourceDBPath_s_Txt)
'//  par_Silenzioso_b = True = nessun messaggio (uso da codice); il risultato e' il testo del log
Public Function TB99_Esegui_pf(par_Azione_s As String, par_Tipo_s As String, _
                               Optional par_SourcePath_s As String = "", _
                               Optional par_Silenzioso_b As Boolean = False) As String

    Dim path_s As String
    Dim azione_s As String
    Dim ordineCanc_v As Variant
    Dim ordineImp_v As Variant
    Dim t As Variant
    Dim v As Variant
    Dim col As Collection
    Dim dbSrc As DAO.Database
    Dim fatto_b As Boolean
    Dim esito_s As String
    Dim log_s As String
    Dim fileLog_s As String
    Dim cartella_s As String
    Dim nFile As Integer
    Dim nCanc As Long, nImp As Long, nScarti As Long, nTot As Long
    Dim conferma_s As String
    Dim cnt_s As String

    On Error GoTo TB99_Esegui_pf_Err

    azione_s = UCase(Trim(par_Azione_s))
    path_s = par_SourcePath_s
    If path_s = "" Then path_s = Nz(Me.sourceDBPath_s_Txt, "")

    '//CONTROLLI
    '//...........................................................................//
    If azione_s <> "CANCELLA" And azione_s <> "IMPORTA" And azione_s <> "CANCELLA+IMPORTA" Then
        TB99_Esegui_pf = "AZIONE NON VALIDA: " & par_Azione_s
        GoTo TB99_Esegui_pf_Msg
    End If
    If path_s = "" Then
        TB99_Esegui_pf = "SCEGLI PRIMA IL DATABASE SORGENTE NELLA COMBINATA 01."
        GoTo TB99_Esegui_pf_Msg
    End If
    If Dir(path_s) = "" Then
        TB99_Esegui_pf = "IL DATABASE SORGENTE NON ESISTE: " & path_s
        GoTo TB99_Esegui_pf_Msg
    End If
    If StrComp(path_s, CurrentProject.FullName, vbTextCompare) = 0 Then
        TB99_Esegui_pf = "IL DATABASE SORGENTE E' IL DATABASE CORRENTE: OPERAZIONE ANNULLATA."
        GoTo TB99_Esegui_pf_Msg
    End If
    If LCase(CurrentProject.Name) = "msys_attiva_gest_oggetti.mdb" Then
        TB99_Esegui_pf = "SIAMO NEL PROGETTO ORIGINALE MSYS_ATTIVA_GEST_OGGETTI.mdb: OPERAZIONE ANNULLATA."
        GoTo TB99_Esegui_pf_Msg
    End If

    '//ORDINE DEI TIPI
    '//...........................................................................//
    If UCase(par_Tipo_s) = "TUTTI" Then
        ordineCanc_v = Array("TABLES", "FORMS", "REPORTS", "MACRO", "QUERY", "MODULES")
        ordineImp_v = Array("TABLES", "QUERY", "FORMS", "REPORTS", "MACRO", "MODULES")
    Else
        If TB99_AcTipo_pf(par_Tipo_s) = -1 Then
            TB99_Esegui_pf = "TIPO NON VALIDO: " & par_Tipo_s
            GoTo TB99_Esegui_pf_Msg
        End If
        ordineCanc_v = Array(UCase(par_Tipo_s))
        ordineImp_v = Array(UCase(par_Tipo_s))
    End If

    '//CONFERMA
    '//...........................................................................//
    nTot = 0
    cnt_s = ""
    For Each t In ordineImp_v
        Set col = TB99_Collection_pf(CStr(t), path_s)
        cnt_s = cnt_s & "   " & TB99_NomeTipo_pf(CStr(t)) & ": " & col.Count & vbCrLf
        nTot = nTot + col.Count
    Next t

    If nTot = 0 Then
        TB99_Esegui_pf = "NEL CATALOGO NON CI SONO OGGETTI PER: " & path_s
        GoTo TB99_Esegui_pf_Msg
    End If

    If Not par_Silenzioso_b Then
        conferma_s = "AZIONE: " & azione_s & vbCrLf & _
                     "DA:  " & path_s & vbCrLf & _
                     "IN:  " & CurrentProject.FullName & vbCrLf & vbCrLf & cnt_s & vbCrLf
        If InStr(azione_s, "CANCELLA") > 0 Then
            conferma_s = conferma_s & "Gli oggetti con questi nomi nel database corrente verranno CANCELLATI" & _
                         IIf(InStr(azione_s, "IMPORTA") > 0, " e poi reimportati.", ".") & vbCrLf
        End If
        conferma_s = conferma_s & "Continuo?"
        If MsgBox(conferma_s, vbYesNo + vbQuestion + vbDefaultButton2, "CATALOGO TB99") <> vbYes Then
            TB99_Esegui_pf = "OPERAZIONE ANNULLATA DALL'UTENTE."
            Exit Function
        End If
    End If

    DoCmd.Hourglass True

    log_s = "TB99 - " & azione_s & " - " & par_Tipo_s & " - " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    log_s = log_s & "DATABASE SORGENTE: " & path_s & vbCrLf
    log_s = log_s & "DATABASE CORRENTE: " & CurrentProject.FullName & vbCrLf
    log_s = log_s & String(100, "=") & vbCrLf

    '//1) CANCELLAZIONE
    '//...........................................................................//
    If InStr(azione_s, "CANCELLA") > 0 Then
        log_s = log_s & vbCrLf & "1) CANCELLAZIONE NEL DATABASE CORRENTE" & vbCrLf
        For Each t In ordineCanc_v
            Set col = TB99_Collection_pf(CStr(t), path_s)
            log_s = log_s & "   " & TB99_NomeTipo_pf(CStr(t)) & " (" & col.Count & ")" & vbCrLf
            For Each v In col
                SysCmd acSysCmdSetStatus, "TB99 CANCELLA " & TB99_NomeTipo_pf(CStr(t)) & ": " & v
                esito_s = TB99_Cancella_pf(CStr(t), CStr(v), fatto_b)
                If fatto_b Then
                    nCanc = nCanc + 1
                ElseIf esito_s <> "NON PRESENTE" Then
                    nScarti = nScarti + 1
                End If
                log_s = log_s & "      " & Left(v & Space(70), 70) & " " & esito_s & vbCrLf
            Next v
        Next t
    End If

    '//2) IMPORTAZIONE
    '//...........................................................................//
    If InStr(azione_s, "IMPORTA") > 0 Then
        log_s = log_s & vbCrLf & "2) IMPORTAZIONE DAL DATABASE SORGENTE" & vbCrLf
        Set dbSrc = DBEngine.Workspaces(0).OpenDatabase(path_s, False, True)
        For Each t In ordineImp_v
            Set col = TB99_Collection_pf(CStr(t), path_s)
            log_s = log_s & "   " & TB99_NomeTipo_pf(CStr(t)) & " (" & col.Count & ")" & vbCrLf
            For Each v In col
                SysCmd acSysCmdSetStatus, "TB99 IMPORTA " & TB99_NomeTipo_pf(CStr(t)) & ": " & v
                esito_s = TB99_Importa_pf(dbSrc, path_s, CStr(t), CStr(v), fatto_b)
                If fatto_b Then
                    nImp = nImp + 1
                Else
                    nScarti = nScarti + 1
                End If
                log_s = log_s & "      " & Left(v & Space(70), 70) & " " & esito_s & vbCrLf
            Next v
        Next t
        dbSrc.Close
        Set dbSrc = Nothing
    End If

    log_s = log_s & vbCrLf & String(100, "=") & vbCrLf
    log_s = log_s & "CANCELLATI: " & nCanc & "   IMPORTATI/COLLEGATI: " & nImp & "   NON ESEGUITI: " & nScarti & vbCrLf

    '//FILE DI LOG
    '//...........................................................................//
    cartella_s = CurrentProject.Path & "\LOGS\"
    If Dir(cartella_s, vbDirectory) = "" Then cartella_s = CurrentProject.Path & "\"
    fileLog_s = cartella_s & "TB99_IMPORTA_OGGETTI_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    nFile = FreeFile
    Open fileLog_s For Output As #nFile
    Print #nFile, log_s
    Close #nFile
    log_s = log_s & "LOG: " & fileLog_s & vbCrLf

    SysCmd acSysCmdClearStatus
    DoCmd.Hourglass False
    Application.RefreshDatabaseWindow

    On Error Resume Next
    Me.TXT_OBJECT_ESTERNI = log_s
    On Error GoTo 0

    If Not par_Silenzioso_b Then
        MsgBox "OPERAZIONE TERMINATA" & vbCrLf & vbCrLf & _
               "Cancellati: " & nCanc & vbCrLf & _
               "Importati / collegati: " & nImp & vbCrLf & _
               "Non eseguiti: " & nScarti & vbCrLf & vbCrLf & _
               "Dettaglio nella casella di testo e nel log:" & vbCrLf & fileLog_s, _
               IIf(nScarti = 0, vbInformation, vbExclamation), "CATALOGO TB99"
    End If

    TB99_Esegui_pf = log_s
    Exit Function

TB99_Esegui_pf_Msg:
    If Not par_Silenzioso_b Then MsgBox TB99_Esegui_pf, vbExclamation, "CATALOGO TB99"
    Exit Function

TB99_Esegui_pf_Err:
    '//salvo l'errore prima di On Error Resume Next, che lo azzera
    esito_s = "ERRORE " & Err.Number & ": " & Err.Description
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    DoCmd.Hourglass False
    If Not dbSrc Is Nothing Then dbSrc.Close
    Set dbSrc = Nothing
    TB99_Esegui_pf = log_s & vbCrLf & esito_s
    Me.TXT_OBJECT_ESTERNI = TB99_Esegui_pf
    If Not par_Silenzioso_b Then MsgBox esito_s, vbCritical, "CATALOGO TB99"
End Function

'//=====================================================================================================================//
'//     @TB99@CATALOGO   IMPORTAZIONE DEGLI OGGETTI DAL CATALOGO TB99                         *** FINE ***
'//=====================================================================================================================//
