# -*- coding: utf-8 -*-
"""
GENERA_HELP  -  trasforma gli help .md in pagine .html (solo libreria standard di Python)

  Struttura:   OBJECT\\HTML\\<TIPO>\\<NomeOggetto>\\<NomeOggetto>_HELP_<nn>_<CHIAVE>.md
               TIPO = FORM | QUERY | REPORT | MODULE | TABLE | MACRO
  Per ogni .md nasce accanto il .html con lo stesso nome; in ogni cartella viene copiato HELP_stile.css.
  In OBJECT\\HTML nasce INDICE_HELP.html con tutti gli oggetti che hanno un help.
  Si modifica SOLO il .md: l'html viene riscritto ad ogni esecuzione.

  Formato del .md (vedi OBJECT\\HTML\\LEGGIMI_HELP.md):
    intestazione  ---  titolo / menu / descrizione  (+ guida / sottotitolo / versione nella pagina 00)  ---
    ## titolo      ### sottotitolo
    - elenco       1. elenco numerato      | tabella | con | righe |   (cella "{r}testo" rossa, "{v}testo" verde)
    **grassetto**  *corsivo*  `codice`  [testo](altra_pagina.md)  [[Pulsante]]  [[Pulsante|simula]]
    > [!NOTA] / > [!ATTENZIONE] / > [!OK]   riquadro (le righe seguenti iniziano con >)
    {{SCHEDE}}     schede con tutte le altre pagine (di solito nella pagina 00)
    righe che iniziano con <   html scritto a mano, copiato cosi' com'e'

  Indice:  #01# impostazioni  #02# lettura .md  #03# testo in linea  #04# blocchi
           #05# pagina html   #06# cartella di un oggetto   #07# indice generale   #08# avvio
"""
import os
import re
import shutil
import sys
import datetime
from html import escape

# ------------------------------------------------------------------------------------------------
# #01# impostazioni
# ------------------------------------------------------------------------------------------------
RADICE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))     # ...\OBJECT\HTML
STRUMENTI = os.path.join(RADICE, "_STRUMENTI")
CSS_NOME = "HELP_stile.css"
TIPI = ["FORM", "QUERY", "REPORT", "MODULE", "TABLE", "MACRO"]
NOMI_TIPO = {"FORM": "Form", "QUERY": "Query", "REPORT": "Report", "MODULE": "Moduli",
             "TABLE": "Tabelle", "MACRO": "Macro"}
RE_FILE = re.compile(r"^(?P<ogg>.+)_HELP_(?P<nn>\d{2})_(?P<chiave>[A-Za-z0-9_]+)\.md$")
OGGI = datetime.date.today().strftime("%d/%m/%Y")


# ------------------------------------------------------------------------------------------------
# #02# lettura .md: intestazione (--- chiave: valore ---) e corpo
# ------------------------------------------------------------------------------------------------
def leggi_md(percorso):
    with open(percorso, encoding="utf-8-sig") as f:
        testo = f.read().replace("\r\n", "\n")
    meta = {}
    if testo.startswith("---\n"):
        fine = testo.find("\n---", 4)
        if fine > 0:
            for riga in testo[4:fine].split("\n"):
                if ":" in riga:
                    k, v = riga.split(":", 1)
                    meta[k.strip().lower()] = v.strip()
            testo = testo[fine + 4:].lstrip("\n")
    return meta, testo


# ------------------------------------------------------------------------------------------------
# #03# testo in linea
# ------------------------------------------------------------------------------------------------
def link_html(url):
    if re.match(r"^[a-z]+:", url, re.I) or url.startswith("#"):
        return url
    return re.sub(r"\.md(#|$)", r".html\1", url)


def in_linea(t):
    pezzi = re.split(r"(`[^`]+`)", t)
    out = []
    for p in pezzi:
        if p.startswith("`") and p.endswith("`") and len(p) > 1:
            out.append("<code>" + escape(p[1:-1], quote=False) + "</code>")
            continue
        # pulsanti  [[testo]]  [[testo|classe]]
        p = re.sub(r"\[\[([^\]|]+)\|([a-z]+)\]\]", r'<span class="pulsante \2">\1</span>', p)
        p = re.sub(r"\[\[([^\]]+)\]\]", r'<span class="pulsante">\1</span>', p)
        # link  [testo](url)
        p = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", lambda m: '<a href="%s">%s</a>' % (link_html(m.group(2)), m.group(1)), p)
        # grassetto e corsivo
        p = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", p)
        p = re.sub(r"(?<![\*\w])\*(?![\s\*])(.+?)(?<![\s\*])\*(?![\*\w])", r"<i>\1</i>", p)
        out.append(p)
    return "".join(out)


# ------------------------------------------------------------------------------------------------
# #04# blocchi
# ------------------------------------------------------------------------------------------------
RE_UL = re.compile(r"^[-*] +(.*)$")
RE_OL = re.compile(r"^\d+[.)] +(.*)$")
RE_TAB_SEP = re.compile(r"^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$")


def celle(riga):
    riga = riga.strip()
    if riga.startswith("|"):
        riga = riga[1:]
    if riga.endswith("|"):
        riga = riga[:-1]
    return [c.strip() for c in riga.split("|")]


def cella(tag, testo):
    m = re.match(r"^\{([a-z]+)\}(.*)$", testo)
    if m:
        return '<%s class="%s">%s</%s>' % (tag, m.group(1), in_linea(m.group(2).strip()), tag)
    return "<%s>%s</%s>" % (tag, in_linea(testo), tag)


def blocchi(righe, schede_html=""):
    out = []
    par = []
    i = 0

    def chiudi_par():
        if par:
            out.append("<p>" + in_linea("\n".join(par)) + "</p>")
            par.clear()

    while i < len(righe):
        r = righe[i]
        s = r.strip()
        if not s:
            chiudi_par(); i += 1; continue
        if s == "{{SCHEDE}}":
            chiudi_par(); out.append(schede_html); i += 1; continue
        m = re.match(r"^(#{2,4}) +(.*)$", s)
        if m:
            chiudi_par()
            liv = len(m.group(1))
            out.append("<h%d>%s</h%d>" % (liv, in_linea(m.group(2)), liv)); i += 1; continue
        if s.startswith(">"):
            chiudi_par()
            dentro = []
            while i < len(righe) and righe[i].strip().startswith(">"):
                dentro.append(re.sub(r"^\s*> ?", "", righe[i])); i += 1
            classe = "box"
            if dentro:
                m = re.match(r"^\[!(\w+)\]\s*(.*)$", dentro[0].strip())
                if m:
                    tipo = m.group(1).upper()
                    classe += {"ATTENZIONE": " attenzione", "OK": " ok"}.get(tipo, "")
                    dentro[0] = m.group(2)
            out.append('<div class="%s">%s</div>' % (classe, blocchi(dentro)))
            continue
        if s.startswith("|") and i + 1 < len(righe) and RE_TAB_SEP.match(righe[i + 1].strip()):
            chiudi_par()
            h = "<table><tr>" + "".join(cella("th", c) for c in celle(s)) + "</tr>"
            i += 2
            while i < len(righe) and righe[i].strip().startswith("|"):
                h += "\n<tr>" + "".join(cella("td", c) for c in celle(righe[i])) + "</tr>"
                i += 1
            out.append(h + "\n</table>"); continue
        if RE_UL.match(s) or RE_OL.match(s):
            chiudi_par()
            num = bool(RE_OL.match(s))
            rx = RE_OL if num else RE_UL
            voci = []
            while i < len(righe):
                t = righe[i]
                m = rx.match(t.strip()) if not t.startswith(" ") else None
                if m:
                    voci.append(m.group(1))
                elif t.startswith(" ") and t.strip() and voci:
                    voci[-1] += "\n" + t.strip()
                else:
                    break
                i += 1
            tag = "ol" if num else "ul"
            out.append("<%s>\n%s\n</%s>" % (tag, "\n".join("<li>" + in_linea(v) + "</li>" for v in voci), tag))
            continue
        if s.startswith("<"):
            chiudi_par()
            while i < len(righe) and righe[i].strip():
                out.append(righe[i]); i += 1
            continue
        par.append(s); i += 1
    chiudi_par()
    return "\n".join(out)


# ------------------------------------------------------------------------------------------------
# #05# pagina html
# ------------------------------------------------------------------------------------------------
MODELLO = """<!doctype html>
<html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<!-- PAGINA GENERATA da {md} con OBJECT\\HTML\\GENERA_HELP_HTML.bat: modificare il file .md, non questo -->
<title>{titolo} - {guida}</title><link rel="stylesheet" href="{css}"></head>
<body><header><h1>{guida}</h1><p>{sottotitolo}</p></header>
<nav>{nav}</nav>
<main>
<h1 class="pagina">{titolo}</h1>
{corpo}
</main>
<footer>{guida} - versione {versione} - generata il {oggi} da {md} - <a href="{indice}">indice di tutti gli help</a></footer></body></html>
"""


# ------------------------------------------------------------------------------------------------
# #06# cartella di un oggetto: tutte le sue pagine
# ------------------------------------------------------------------------------------------------
def genera_oggetto(tipo, cartella):
    oggetto = os.path.basename(cartella)
    pagine = []
    for nome in sorted(os.listdir(cartella)):
        m = RE_FILE.match(nome)
        if not m:
            continue
        if m.group("ogg") != oggetto:
            print("  ATTENZIONE: %s non inizia con il nome della cartella (%s): saltato" % (nome, oggetto))
            continue
        meta, corpo = leggi_md(os.path.join(cartella, nome))
        pagine.append({"md": nome, "html": nome[:-3] + ".html", "nn": m.group("nn"),
                       "chiave": m.group("chiave"), "meta": meta, "corpo": corpo})
    if not pagine:
        return None
    capo = pagine[0]["meta"]
    guida = capo.get("guida", "Guida di " + oggetto)
    sottotitolo = capo.get("sottotitolo", "%s <b>%s</b>" % (NOMI_TIPO.get(tipo, tipo), oggetto))
    versione = capo.get("versione", "-")
    shutil.copyfile(os.path.join(STRUMENTI, CSS_NOME), os.path.join(cartella, CSS_NOME))
    for p in pagine:
        p["titolo"] = p["meta"].get("titolo", p["chiave"].replace("_", " ").title())
        p["menu"] = p["meta"].get("menu", p["titolo"])
    schede = '<div class="schede">\n' + "\n".join(
        '<a href="%s"><b>%s</b><br>%s</a>' % (p["html"], in_linea(p["menu"]), in_linea(p["meta"].get("descrizione", "")))
        for p in pagine[1:]) + "\n</div>"
    for p in pagine:
        nav = "".join('<a href="%s"%s>%s</a>' % (q["html"], ' class="qui"' if q is p else "", in_linea(q["menu"]))
                      for q in pagine)
        html = MODELLO.format(md=p["md"], titolo=escape(p["titolo"]), guida=guida, sottotitolo=sottotitolo,
                              css=CSS_NOME, nav=nav, corpo=blocchi(p["corpo"].split("\n"), schede),
                              versione=versione, oggi=OGGI, indice="../../INDICE_HELP.html")
        with open(os.path.join(cartella, p["html"]), "w", encoding="utf-8", newline="\r\n") as f:
            f.write(html)
    print("  %-7s %-55s %2d pagine" % (tipo, oggetto, len(pagine)))
    return {"tipo": tipo, "oggetto": oggetto, "guida": guida, "prima": pagine[0]["html"],
            "pagine": pagine, "descr": capo.get("descrizione", "")}


# ------------------------------------------------------------------------------------------------
# #07# indice generale OBJECT\HTML\INDICE_HELP.html
# ------------------------------------------------------------------------------------------------
def genera_indice(oggetti):
    shutil.copyfile(os.path.join(STRUMENTI, CSS_NOME), os.path.join(RADICE, CSS_NOME))
    progetto = os.path.basename(os.path.dirname(os.path.dirname(RADICE)))
    corpo = ['<div class="box">Un help per ogni oggetto, diviso per tipo. Ogni pagina si modifica nel suo file '
             '<code>.md</code>; poi <code>OBJECT\\HTML\\GENERA_HELP_HTML.bat</code> rigenera le pagine html.</div>']
    for tipo in TIPI:
        elenco = [o for o in oggetti if o["tipo"] == tipo]
        if not elenco:
            continue
        corpo.append("<h2>%s</h2>" % NOMI_TIPO[tipo])
        corpo.append('<table><tr><th>Oggetto</th><th>Guida</th><th>Pagine</th></tr>')
        for o in elenco:
            link = "%s/%s/%s" % (tipo, o["oggetto"], o["prima"])
            pag = ", ".join('<a href="%s/%s/%s">%s</a>' % (tipo, o["oggetto"], p["html"], in_linea(p["menu"]))
                            for p in o["pagine"])
            corpo.append('<tr><td><a href="%s"><code>%s</code></a></td><td>%s</td><td>%s</td></tr>'
                         % (link, o["oggetto"], o["guida"], pag))
        corpo.append("</table>")
    html = MODELLO.format(md="(indice generale)", titolo="Indice degli help", guida="Help del progetto " + progetto,
                          sottotitolo="cartella OBJECT\\HTML - %d oggetti con help" % len(oggetti), css=CSS_NOME,
                          nav="", corpo="\n".join(corpo), versione="-", oggi=OGGI, indice="INDICE_HELP.html")
    with open(os.path.join(RADICE, "INDICE_HELP.html"), "w", encoding="utf-8", newline="\r\n") as f:
        f.write(html)


# ------------------------------------------------------------------------------------------------
# #08# avvio
# ------------------------------------------------------------------------------------------------
def main():
    print("GENERA_HELP - radice: " + RADICE)
    oggetti = []
    for tipo in TIPI:
        base = os.path.join(RADICE, tipo)
        if not os.path.isdir(base):
            continue
        for nome in sorted(os.listdir(base)):
            cartella = os.path.join(base, nome)
            if os.path.isdir(cartella):
                o = genera_oggetto(tipo, cartella)
                if o:
                    oggetti.append(o)
    genera_indice(oggetti)
    print("Fatto: %d oggetti, indice in %s" % (len(oggetti), os.path.join(RADICE, "INDICE_HELP.html")))


if __name__ == "__main__":
    sys.exit(main())
