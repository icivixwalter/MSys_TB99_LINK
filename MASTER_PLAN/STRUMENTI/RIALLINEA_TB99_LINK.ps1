# ==================================================================================================
#  RIALLINEA_TB99_LINK.ps1                                         versione 2.0 del 15/11/2026
#  SCOPO  : controlla il catalogo Msys_TB99_LINK e ricostruisce i link fisici dentro MSys_TB99_LINK.mdb
#           - ogni tabella deve stare in un database dentro PROGETTI_MDB\MSYS\ (regola del progetto)
#           - Verifica_s = OK / MISSING / NOT_REGISTERED
#           - i record fuori da MSYS\ vengono tolti, insieme ai loro oggetti e alle loro destinazioni
#  STRUTTURA (dal 15/11/2026):
#           Msys_TB99_LINK          Name, Type, Connect, Database (solo file), Database_s (solo cartella)
#           Msys_TB99_LINK_Object   oggetti che usano la tabella (legame con LINK: Name)
#           Msys_TB99_LINK_PATH_DB  destinazioni dell oggetto (legame con Object: ID_OBJ) - toccata solo per i record tolti
#  USO    : powershell -File RIALLINEA_TB99_LINK.ps1      (con MSys_TB99_LINK.mdb chiuso)
#  INDICE : #01# impostazioni  #02# verifica catalogo  #03# fuori da MSYS\  #04# link fisici  #05# riepilogo
# ==================================================================================================
$ErrorActionPreference = 'Stop'

# #01# impostazioni
$MDB    = 'C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_TB99_LINK\MDB\MSys_TB99_LINK.mdb'
$RADICE = 'C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\'
$e  = New-Object -ComObject DAO.DBEngine.120
$db = $e.OpenDatabase($MDB)
function X($sql) { $db.Execute($sql, 128) }
function T([string]$v) { "'" + $v.Replace("'", "''") + "'" }
$cache = @{}
function Tabelle($path) {
  if (-not $cache.ContainsKey($path)) { $s = @{}
    if (Test-Path $path) { $x = $e.OpenDatabase($path, $false, $true); foreach ($t in $x.TableDefs) { if (-not $t.Connect) { $s[$t.Name] = 1 } }; $x.Close() }
    $cache[$path] = $s }
  return $cache[$path]
}

# #02# verifica di ogni tabella del catalogo
$righe = @(); $r = $db.OpenRecordset('SELECT ID_LINK, Name, Database_s, [Database] FROM Msys_TB99_LINK', 4)
while (-not $r.EOF) { $righe += ,@([string]$r.Fields(0).Value, [string]$r.Fields(1).Value, ([string]$r.Fields(2).Value + [string]$r.Fields(3).Value)); $r.MoveNext() }; $r.Close()
$stat = @{}; $fuori = @()
foreach ($x in $righe) {
  $full = $x[2]
  if (-not $full.StartsWith($RADICE, [StringComparison]::OrdinalIgnoreCase)) { $fuori += $x[1]; continue }
  if (-not (Test-Path $full)) { $st = 'MISSING' } elseif ((Tabelle $full).ContainsKey($x[1])) { $st = 'OK' } else { $st = 'NOT_REGISTERED' }
  X "UPDATE Msys_TB99_LINK SET Verifica_s = '$st', Connect = $(T (';DATABASE=' + $full)), DATAAGG_d = Now() WHERE ID_LINK = $($x[0])"
  $stat[$st] = 1 + [int]$stat[$st]
}

# #03# record fuori da MSYS\: prima oggetti e destinazioni (integrita referenziale), poi la tabella
foreach ($n in $fuori) {
  X "DELETE FROM Msys_TB99_LINK_PATH_DB WHERE ID_OBJ IN (SELECT ID_OBJ FROM Msys_TB99_LINK_Object WHERE Name = $(T $n))"
  X "DELETE FROM Msys_TB99_LINK_Object WHERE Name = $(T $n)"
  X "DELETE FROM Msys_TB99_LINK WHERE Name = $(T $n)"
  try { $db.TableDefs.Delete($n) } catch {}
  "TOLTO (fuori da MSYS\): $n"
}

# #04# link fisici dentro MSys_TB99_LINK.mdb ricostruiti dal catalogo (solo stato OK)
$ok = 0; $err = 0
$r = $db.OpenRecordset("SELECT Name, Database_s & [Database] AS FULLDB FROM Msys_TB99_LINK WHERE Verifica_s = 'OK'", 4)
while (-not $r.EOF) {
  $n = [string]$r.Fields(0).Value; $p = [string]$r.Fields(1).Value
  try {
    $td = $null; foreach ($t in $db.TableDefs) { if ($t.Name -eq $n) { $td = $t; break } }
    if ($td -ne $null -and -not $td.Connect) { throw 'esiste una tabella LOCALE con lo stesso nome' }
    if ($td -ne $null) { $db.TableDefs.Delete($n) }
    $td = $db.CreateTableDef($n); $td.Connect = ';DATABASE=' + $p; $td.SourceTableName = $n; $db.TableDefs.Append($td)
    $c = $db.OpenRecordset("SELECT COUNT(*) FROM [$n]", 4); $null = $c.Fields(0).Value; $c.Close(); $ok++
  } catch { $err++; "ERRORE link $n : $($_.Exception.Message)" }
  $r.MoveNext()
}
$r.Close()
foreach ($t in @($db.TableDefs)) { if ($t.Connect -and $t.Connect -notlike (';DATABASE=' + $RADICE + '*')) { "TOLTO link orfano: $($t.Name)"; $db.TableDefs.Delete($t.Name) } }

# #05# riepilogo
"Verifica: " + (($stat.GetEnumerator() | % { "$($_.Key)=$($_.Value)" }) -join ', ') + "  - link ricostruiti OK=$ok ERRORI=$err"
$db.Close()
