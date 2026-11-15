# ==================================================================================================
#  RICOSTRUISCI_TB99_OGGETTI_1_ESPORTA.ps1
#  SCOPO  : fase 1 di 2 - per ogni database in PROGETTI_MDB\MSYS\ esporta in testo query (SQL),
#           maschere, report, macro e moduli, in una cartella di lavoro.
#           La fase 2 (RICOSTRUISCI_TB99_OGGETTI_2_CARICA.py + .ps1) trova le tabelle usate
#           e riempie Msys_TB99_LINK_Object.
#  USO    : powershell -File RICOSTRUISCI_TB99_OGGETTI_1_ESPORTA.ps1 -Lavoro <cartella>
#  NOTE   : ogni database viene copiato in una cartella temporanea ed esportato dalla copia, con
#           un processo Access separato e un tempo massimo; i database originali non vengono aperti
#           da Access. Alla fine nessun processo Access avviato dallo script resta aperto.
#  INDICE : #01# parametri   #02# elenco database   #03# esportazione   #04# chiusura
# ==================================================================================================
param([Parameter(Mandatory=$true)][string]$Lavoro, [int]$TempoMax = 240)

# #01# parametri
$RADICE = 'C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\'
$ESCLUSI = '\\(AA_SALVATAGGI|OLD|EXE|\.git)\\|\\BACKUP|\\MSys_TB99_LINK\\'
New-Item -ItemType Directory -Force $Lavoro | Out-Null

# #02# elenco database
$dbs = Get-ChildItem $RADICE -Recurse -Include *.mdb, *.accdb -ErrorAction SilentlyContinue | ? { $_.FullName -notmatch $ESCLUSI }
"database da esportare: $($dbs.Count)"

# #03# esportazione, un database alla volta
$tmp = Join-Path $env:TEMP 'tb99_ricostruisci'
foreach ($f in $dbs) {
  $chiave = ($f.FullName.Substring($RADICE.Length) -replace '[\\:]', '#')
  $out = Join-Path $Lavoro $chiave; Remove-Item -Recurse -Force $out -ErrorAction SilentlyContinue; New-Item -ItemType Directory $out | Out-Null
  [IO.File]::WriteAllText((Join-Path $out '_database.txt'), $f.FullName, [Text.Encoding]::UTF8)
  New-Item -ItemType Directory -Force $tmp | Out-Null; $copia = Join-Path $tmp ('copia' + $f.Extension); Copy-Item $f.FullName $copia -Force
  $prima = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | % Id)
  $job = Start-Job -ArgumentList $copia, $out -ScriptBlock { param($db, $out)
    $idx = New-Object System.Collections.Generic.List[string]
    # query con DAO (piu veloce e affidabile)
    $e = New-Object -ComObject DAO.DBEngine.120; $d = $e.OpenDatabase($db, $false, $true); $i = 0
    foreach ($q in $d.QueryDefs) { if ($q.Name -like '~*') { continue }; $i++; $fn = "Q_$i.txt"
      [IO.File]::WriteAllText((Join-Path $out $fn), $q.SQL, [Text.Encoding]::UTF8); $idx.Add("$fn`t5`t$($q.Name)") }
    $d.Close()
    # maschere, report, macro, moduli con Access
    $a = New-Object -ComObject Access.Application; $a.AutomationSecurity = 3; $a.OpenCurrentDatabase($db)
    $tipi = @{ 2 = -32768; 3 = -32764; 4 = -32766; 5 = -32761 }
    foreach ($pair in @(@(2, $a.CurrentProject.AllForms), @(3, $a.CurrentProject.AllReports), @(4, $a.CurrentProject.AllMacros), @(5, $a.CurrentProject.AllModules))) {
      foreach ($o in $pair[1]) { if ($o.Name -like '~*') { continue }; $i++; $fn = "O_$i.txt"
        try { $a.SaveAsText($pair[0], $o.Name, (Join-Path $out $fn)); $idx.Add("$fn`t$($tipi[$pair[0]])`t$($o.Name)") } catch {} } }
    $a.CloseCurrentDatabase(); $a.Quit()
    [IO.File]::WriteAllLines((Join-Path $out '_indice.tsv'), $idx, [Text.Encoding]::UTF8); $idx.Count }
  if (Wait-Job $job -Timeout $TempoMax) { $n = Receive-Job $job; "OK      $n oggetti  $($f.FullName.Substring($RADICE.Length))" }
  else { "TIMEOUT          $($f.FullName.Substring($RADICE.Length))" }
  Remove-Job $job -Force
  @(Get-Process MSACCESS -ErrorAction SilentlyContinue | % Id) | ? { $prima -notcontains $_ } | % { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }
  Remove-Item $copia -Force -ErrorAction SilentlyContinue
}

# #04# chiusura
Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
"fine esportazione - processi Access ancora aperti: " + @(Get-Process MSACCESS -ErrorAction SilentlyContinue).Count
