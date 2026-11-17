param([Parameter(Mandatory=$true)][string]$DbPath, [Parameter(Mandatory=$true)][string]$Blocco)
# Inserisce nella classe della form UTILITA_Msys_Frm21_IMPORTA_OGGETTI il blocco TB99 e gli agganci
$ErrorActionPreference = "Stop"
$NomeForm = "UTILITA_Msys_Frm21_IMPORTA_OGGETTI"
$testo = [IO.File]::ReadAllText($Blocco, [Text.Encoding]::UTF8)

# procedura -> righe da inserire subito dopo la dichiarazione
$ganci = [ordered]@{
  "Form_Load"               = "    If TB99_Attivo_b() Then Call TB99_Form_Load_Sub: Exit Sub          '//@TB99@CATALOGO"
  "Cmb_01_TXT_AfterUpdate"  = "    If TB99_Attivo_b() Then Call TB99_Cmb01_AfterUpdate_Sub: Exit Sub  '//@TB99@CATALOGO"
  "Cmb_02_txt_AfterUpdate"  = "    If TB99_Attivo_b() Then Call TB99_Cmb02_AfterUpdate_Sub: Exit Sub  '//@TB99@CATALOGO"
  "CollectionForm_PFunct"   = "    If TB99_Attivo_b() Then Set CollectionForm_PFunct = TB99_Collection_pf(""FORMS""): Exit Function     '//@TB99@CATALOGO"
  "CollectionQuery_PFunct"  = "    If TB99_Attivo_b() Then Set CollectionQuery_PFunct = TB99_Collection_pf(""QUERY""): Exit Function    '//@TB99@CATALOGO"
  "CollectionTable_PFunct"  = "    If TB99_Attivo_b() Then Set CollectionTable_PFunct = TB99_Collection_pf(""TABLES""): Exit Function   '//@TB99@CATALOGO"
  "CollectionMacro_PFunct"  = "    If TB99_Attivo_b() Then Set CollectionMacro_PFunct = TB99_Collection_pf(""MACRO""): Exit Function    '//@TB99@CATALOGO"
  "CollectionModuli_PFunct" = "    If TB99_Attivo_b() Then Set CollectionModuli_PFunct = TB99_Collection_pf(""MODULES""): Exit Function '//@TB99@CATALOGO"
  "CollectionReport_PFunct" = "    If TB99_Attivo_b() Then Set CollectionReport_PFunct = TB99_Collection_pf(""REPORTS""): Exit Function '//@TB99@CATALOGO"
  "Cmd_Importa_TABELLE_Click" = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""IMPORTA"", ""TABLES""): Exit Sub    '//@TB99@CATALOGO"
  "Cmd_Delete_TABLE_Click"    = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""CANCELLA"", ""TABLES""): Exit Sub   '//@TB99@CATALOGO"
  "Cmd_IMPORTA_QUERY_Click"   = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""IMPORTA"", ""QUERY""): Exit Sub     '//@TB99@CATALOGO"
  "Cmd_Delete_QUERY_Click"    = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""CANCELLA"", ""QUERY""): Exit Sub    '//@TB99@CATALOGO"
  "Cmd_Importa_FORM_Click"    = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""IMPORTA"", ""FORMS""): Exit Sub     '//@TB99@CATALOGO"
  "Cmd_Delete_FORM_Click"     = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""CANCELLA"", ""FORMS""): Exit Sub    '//@TB99@CATALOGO"
  "Cmd_Importa_REPORT_Click"  = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""IMPORTA"", ""REPORTS""): Exit Sub   '//@TB99@CATALOGO"
  "Cmd_Delete_REPORT_Click"   = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""CANCELLA"", ""REPORTS""): Exit Sub  '//@TB99@CATALOGO"
  "Cmd_Importa_MACRO_Click"   = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""IMPORTA"", ""MACRO""): Exit Sub     '//@TB99@CATALOGO"
  "Cmd_Delete_MACRO_Click"    = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""CANCELLA"", ""MACRO""): Exit Sub    '//@TB99@CATALOGO"
  "Cmd_IMPORTA_MODULI_Click"  = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""IMPORTA"", ""MODULES""): Exit Sub   '//@TB99@CATALOGO"
  "Cmd_Delete_MODULI_Click"   = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""CANCELLA"", ""MODULES""): Exit Sub  '//@TB99@CATALOGO"
  "Cmd_IMPORTA_TUTTO_Click"   = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""CANCELLA+IMPORTA"", ""TUTTI""): Exit Sub  '//@TB99@CATALOGO prima cancella poi importa"
  "Cmd_DELETE_TUTTO_Click"    = "    If TB99_Attivo_b() Then Call TB99_Esegui_pf(""CANCELLA"", ""TUTTI""): Exit Sub   '//@TB99@CATALOGO"
}

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 3
$acc.OpenCurrentDatabase($DbPath)
try {
  $acc.DoCmd.OpenForm($NomeForm, 1)   # acDesign
  $mod = $acc.Forms.Item($NomeForm).Module
  $tutto = $mod.Lines(1, $mod.CountOfLines)
  if ($tutto -match "TB99_Attivo_b") { throw "Il blocco TB99 e' gia presente: nessuna modifica." }
  "Righe prima: " + $mod.CountOfLines
  foreach ($p in $ganci.Keys) {
    $riga = $mod.ProcBodyLine($p, 0)
    $decl = $mod.Lines($riga, 1)
    if ($decl -notmatch [regex]::Escape($p)) { throw "Dichiarazione inattesa per $p : $decl" }
    $mod.InsertLines($riga + 1, $ganci[$p])
    "  agganciato {0,-26} riga {1}" -f $p, ($riga + 1)
  }
  $mod.InsertLines($mod.CountOfLines + 1, $testo)
  "Righe dopo: " + $mod.CountOfLines
  $acc.DoCmd.Close(2, $NomeForm, 1)   # acForm, acSaveYes
  "Form salvata."
} finally {
  $acc.CloseCurrentDatabase(); $acc.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($acc)
}
