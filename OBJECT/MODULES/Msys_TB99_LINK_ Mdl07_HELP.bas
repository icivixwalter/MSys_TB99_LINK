Option Compare Database
Option Explicit
'==================================================================================================
' Msys_TB99_LINK_ Mdl07_HELP  -  apertura degli help HTML degli oggetti del progetto
'   cartelle : <progetto>\OBJECT\HTML\<TIPO>\<NomeOggetto>\<NomeOggetto>_HELP_<nn>_<CHIAVE>.html
'              TIPO = FORM | QUERY | REPORT | MODULE | TABLE | MACRO
'              il database sta in <progetto>\MDB, quindi OBJECT e' la cartella sorella di MDB
'   modifica : si scrive il file .md accanto all'html, poi OBJECT\HTML\GENERA_HELP_HTML.bat
'              rigenera le pagine (regole in OBJECT\HTML\LEGGIMI_HELP.md)
'   esempio  : HELP_Apri_pf "FORM", Me.Name, "01_COLLEGAMENTI"
'   INDICE
'     #01# HELP_Radice_pf    cartella OBJECT\HTML del progetto
'     #02# HELP_Cartella_pf  cartella help di un oggetto
'     #03# HELP_File_pf      percorso completo di una pagina
'     #04# HELP_Apri_pf      apre la pagina (se manca, l'indice 00 dell'oggetto)
'==================================================================================================

Public Function HELP_Radice_pf() As String
    '#01# <progetto>\OBJECT\HTML\  (il progetto e' la cartella sopra quella del database)
    Dim p As String
    p = CurrentProject.Path
    HELP_Radice_pf = Left$(p, InStrRev(p, "\")) & "OBJECT\HTML\"
End Function

Public Function HELP_Cartella_pf(ByVal sTipo As String, ByVal sOggetto As String) As String
    '#02# <progetto>\OBJECT\HTML\<TIPO>\<NomeOggetto>\
    HELP_Cartella_pf = HELP_Radice_pf() & UCase$(sTipo) & "\" & sOggetto & "\"
End Function

Public Function HELP_File_pf(ByVal sTipo As String, ByVal sOggetto As String, Optional ByVal sPagina As String = "00_INDICE") As String
    '#03# <cartella oggetto>\<NomeOggetto>_HELP_<nn>_<CHIAVE>.html
    HELP_File_pf = HELP_Cartella_pf(sTipo, sOggetto) & sOggetto & "_HELP_" & sPagina & ".html"
End Function

Public Function HELP_Apri_pf(ByVal sTipo As String, ByVal sOggetto As String, Optional ByVal sPagina As String = "00_INDICE") As Boolean
    '#04# apre la pagina; se manca prova l'indice 00 dell'oggetto; se manca anche quello avvisa
    Dim f As String
    On Error GoTo ERRORE
    f = HELP_File_pf(sTipo, sOggetto, sPagina)
    If Len(Dir$(f)) = 0 Then f = HELP_File_pf(sTipo, sOggetto)
    If Len(Dir$(f)) = 0 Then
        MsgBox "Help non trovato:" & vbCrLf & HELP_File_pf(sTipo, sOggetto, sPagina) & vbCrLf & vbCrLf & _
               "Se esiste il file .md, generare le pagine con OBJECT\HTML\GENERA_HELP_HTML.bat", vbExclamation, "Help - " & sOggetto
        Exit Function
    End If
    Application.FollowHyperlink f
    HELP_Apri_pf = True
    Exit Function
ERRORE:
    MsgBox "Impossibile aprire l'help:" & vbCrLf & f & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Help - " & sOggetto
End Function