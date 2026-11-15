# -*- coding: utf-8 -*-
# ==================================================================================================
#  RICOSTRUISCI_TB99_OGGETTI_2_ANALIZZA.py
#  SCOPO  : fase 2 di 3 - legge i testi esportati dalla fase 1 e trova, per ogni oggetto (query,
#           maschera, report, macro, modulo), quali tabelle del catalogo Msys_TB99_LINK usa,
#           direttamente o attraverso le query e le sottomaschere dello stesso database.
#           REGOLA: si tengono solo gli oggetti che stanno NELLO STESSO DATABASE della tabella
#           fisica (il mini-progetto originale), per esempio le query di MSys_DF01_COMANDI
#           che si trovano in MSys_DF01.mdb.
#  USO    : python RICOSTRUISCI_TB99_OGGETTI_2_ANALIZZA.py <cartella lavoro> <catalogo.txt> <uscita.tsv>
#           catalogo.txt = una riga per tabella: Name <TAB> percorso completo del database sorgente
#  USCITA : TSV  Name | NOME_OGGETTO_s | TIPO_OGGETTO_s | Type | Database | Database_s | COD_SUB_PROGETTO_s
#  INDICE : #01# impostazioni   #02# lettura testi   #03# tabelle usate   #04# dipendenze indirette
#           #05# scrittura
#  NOTA   : 15/11/2026 - la propagazione segue anche le maschere usate come sottomaschera
# ==================================================================================================
import os, re, sys

# --- #01# impostazioni ----------------------------------------------------------------------------
RADICE = 'C:\\CASA\\LINGUAGGI\\ACCESS\\PROGETTI_MDB\\MSYS\\'
TIPO = {5: 'QUERY', -32768: 'FORMS', -32764: 'REPORTS', -32766: 'MACRO', -32761: 'MODULES'}

def leggi(p):
    b = open(p, 'rb').read()
    if b[:2] == b'\xff\xfe': return b.decode('utf-16')
    if b[:3] == b'\xef\xbb\xbf': return b[3:].decode('utf-8')
    try: return b.decode('utf-8')
    except UnicodeDecodeError: return b.decode('cp1252')

def main(lavoro, catalogo, uscita):
    # tabelle del catalogo, esclusi i separatori di gruppo
    sorgente = {}
    for riga in open(catalogo, encoding='utf-8-sig'):
        riga = riga.rstrip('\r\n')
        if not riga: continue
        nome, _, db_tab = riga.partition('\t')
        sorgente[nome] = os.path.normcase(db_tab.strip())
    nomi = [n for n in sorgente if '{' not in n and '}' not in n]
    rx = {n: re.compile(r'(?<![\w])' + re.escape(n) + r'(?![\w])', re.I) for n in nomi}
    righe = set()
    for cart in sorted(os.listdir(lavoro)):
        d = os.path.join(lavoro, cart)
        if not os.path.isfile(os.path.join(d, '_indice.tsv')): continue
        db = open(os.path.join(d, '_database.txt'), encoding='utf-8-sig').read().strip()
        rel = db[len(RADICE):] if db.upper().startswith(RADICE.upper()) else db
        sub = rel.split('\\')[0]
        db_file = os.path.basename(db); db_path = os.path.dirname(db) + '\\'
        # --- #02# lettura testi -----------------------------------------------------------------
        ogg = []
        for l in open(os.path.join(d, '_indice.tsv'), encoding='utf-8-sig').read().splitlines():
            if not l: continue
            fn, tipo, nome = l.split('\t'); tipo = int(tipo)
            if tipo not in TIPO: continue
            ogg.append((nome, tipo, leggi(os.path.join(d, fn))))
        # --- #03# tabelle usate direttamente ----------------------------------------------------
        usa = {}
        for nome, tipo, t in ogg:
            usa[(nome, tipo)] = {n for n in nomi if rx[n].search(t)}
        # --- #04# dipendenze indirette attraverso query e maschere (sottomaschere) dello stesso db -
        query = {nome: (nome, tipo) for nome, tipo, _ in ogg if tipo in (5, -32768)}
        rxq = {q: re.compile(r'(?<![\w])' + re.escape(q) + r'(?![\w])', re.I) for q in query}
        richiama = {}
        for nome, tipo, t in ogg:
            richiama[(nome, tipo)] = {q for q in query if q != nome and rxq[q].search(t)}
        cambiato = True
        while cambiato:
            cambiato = False
            for k, qs in richiama.items():
                for q in qs:
                    nuove = usa[query[q]] - usa[k]
                    if nuove: usa[k] |= nuove; cambiato = True
        for (nome, tipo), tab in usa.items():
            for n in tab:
                if sorgente[n] and sorgente[n] != os.path.normcase(db): continue   # solo nel database della tabella
                righe.add((n, nome, TIPO[tipo], str(tipo), db_file, db_path, sub))
    # --- #05# scrittura ---------------------------------------------------------------------------
    with open(uscita, 'w', encoding='utf-8') as f:
        for r in sorted(righe): f.write('\t'.join(r) + '\n')
    print('righe oggetto-tabella:', len(righe))

if __name__ == '__main__':
    main(*sys.argv[1:4])
