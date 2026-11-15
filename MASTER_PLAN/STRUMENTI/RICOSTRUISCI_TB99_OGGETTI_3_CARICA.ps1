# ==================================================================================================
#  RICOSTRUISCI_TB99_OGGETTI_3_CARICA.ps1
#  SCOPO  : fase 3 di 3 - scrive l'elenco del catalogo, lancia l'analisi (fase 2) e ricarica
#           Msys_TB99_LINK_Object con gli oggetti trovati.
#  USO    : powershell -File RICOSTRUISCI_TB99_OGGETTI_3_CARICA.ps1 -Lavoro <cartella della fase 1>
#           da eseguire con MSys_TB99_LINK.mdb chiuso
#  INDICE : #01# catalogo   #02# analisi   #03# caricamento   #04# riepilogo
# ==================================================================================================
param([Parameter(Mandatory=$true)][string]$Lavoro)
$ErrorActionPreference = 'Stop'
$MDB = 'C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_TB99_LINK\MDB\MSys_TB99_LINK.mdb'
$e = New-Object -ComObject DAO.DBEngine.120

# #01# catalogo: una riga per tabella, Name <TAB> database sorgente completo
$db = $e.OpenDatabase($MDB, $false, $true); $nomi = @()
$r = $db.OpenRecordset('SELECT Name, Database_s & [Database] FROM Msys_TB99_LINK', 4); while (-not $r.EOF) { $nomi += ([string]$r.Fields(0).Value + "`t" + [string]$r.Fields(1).Value); $r.MoveNext() }; $db.Close()
$cat = Join-Path $Lavoro '_catalogo.txt'; $tsv = Join-Path $Lavoro '_oggetti.tsv'
[IO.File]::WriteAllLines($cat, $nomi, (New-Object Text.UTF8Encoding $false))

# #02# analisi (fase 2)
python -I (Join-Path $PSScriptRoot 'RICOSTRUISCI_TB99_OGGETTI_2_ANALIZZA.py') $Lavoro $cat $tsv
if ($LASTEXITCODE -ne 0) { throw 'analisi non riuscita' }

# #03# caricamento
#      catena: Msys_TB99_LINK -> Msys_TB99_LINK_Object (Name) -> Msys_TB99_LINK_PATH_DB (ID_OBJ)
#      le destinazioni vengono salvate con la chiave dell'oggetto (tabella + nome + database),
#      poi si ricarica Object e si riagganciano le destinazioni ai nuovi ID_OBJ
function S([string]$v) { if ($v -eq '') { 'Null' } else { "'" + $v.Replace("'", "''") + "'" } }
$db = $e.OpenDatabase($MDB)
foreach ($t in @($db.TableDefs)) { if ($t.Name -eq 'tmp_TB99_PATH_SALVA') { $db.TableDefs.Delete($t.Name) } }
$db.Execute('SELECT P.*, O.Name AS K_NAME, O.NOME_OGGETTO_s AS K_NOME, O.[Database] AS K_FILE, O.Database_s AS K_DIR INTO tmp_TB99_PATH_SALVA FROM Msys_TB99_LINK_PATH_DB AS P INNER JOIN Msys_TB99_LINK_Object AS O ON P.ID_OBJ = O.ID_OBJ', 128)
$salvate = $db.OpenRecordset('SELECT COUNT(*) FROM tmp_TB99_PATH_SALVA', 4).Fields(0).Value
$db.Execute('DELETE FROM Msys_TB99_LINK_PATH_DB', 128)
$db.Execute('DELETE FROM Msys_TB99_LINK_Object WHERE Name Is Not Null', 128)   # le righe senza tabella (da Mdl04) restano
$n = 0
foreach ($l in [IO.File]::ReadAllLines($tsv, [Text.Encoding]::UTF8)) {
  if (-not $l) { continue }; $c = $l.Split("`t")   # Name | NOME | TIPO | Type | Database | Database_s | sub
  $db.Execute("INSERT INTO Msys_TB99_LINK_Object (Name, NOME_OGGETTO_s, TIPO_OGGETTO_s, [Database], Database_s, Scel_b, DATAINS_d, DATAAGG_d) VALUES ($(S $c[0]), $(S $c[1]), $(S $c[2]), $(S $c[4]), $(S $c[5]), False, Now(), Now())", 128)
  $n++
}
$db.Execute('INSERT INTO Msys_TB99_LINK_PATH_DB (ID_OBJ, [Database], Database_s, Scel_b, Esito_s, Nota_s, DATAINS_d, DATA_INVIO_d) SELECT O.ID_OBJ, S.[Database], S.Database_s, S.Scel_b, S.Esito_s, S.Nota_s, S.DATAINS_d, S.DATA_INVIO_d FROM tmp_TB99_PATH_SALVA AS S INNER JOIN Msys_TB99_LINK_Object AS O ON (S.K_NAME = O.Name) AND (S.K_NOME = O.NOME_OGGETTO_s) AND (S.K_FILE = O.[Database]) AND (S.K_DIR = O.Database_s)', 128)
$riprese = $db.OpenRecordset('SELECT COUNT(*) FROM Msys_TB99_LINK_PATH_DB', 4).Fields(0).Value
"destinazioni salvate: $salvate - riagganciate: $riprese" + $(if ($salvate -ne $riprese) { '  (le altre appartenevano a oggetti che non esistono piu: restano in tmp_TB99_PATH_SALVA)' } else { '' })
if ($salvate -eq $riprese) { $db.TableDefs.Delete('tmp_TB99_PATH_SALVA') }

# #04# riepilogo
"righe caricate in Msys_TB99_LINK_Object: $n"
$r = $db.OpenRecordset('SELECT TIPO_OGGETTO_s, COUNT(*) FROM Msys_TB99_LINK_Object GROUP BY TIPO_OGGETTO_s', 4)
while (-not $r.EOF) { '  {0,-8} {1}' -f $r.Fields(0).Value, $r.Fields(1).Value; $r.MoveNext() }
"tabelle del catalogo con almeno un oggetto: " + $db.OpenRecordset('SELECT COUNT(*) FROM (SELECT DISTINCT Name FROM Msys_TB99_LINK_Object)', 4).Fields(0).Value
$db.Close()
