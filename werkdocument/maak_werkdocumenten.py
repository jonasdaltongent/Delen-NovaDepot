#!/usr/bin/env python3
"""
maak_werkdocumenten.py
Genereert het materiaal voor les 08 "Delen met de juiste rechten" (NovaDepot, ORLO).

  WP3_Delen-met-de-juiste-rechten.docx   het werkdocument dat de leerling INLEVERT
  Leveringsplanning.docx                 het oefendocument van de duo-test

Beide gaan als "Een kopie maken voor elke leerling" in de opdracht (soort "kopie" in classroom.json): elke
leerling krijgt een eigen Google-document en deelt de Leveringsplanning met de buur. Geen zip in deze les.

Gebruik:  python3 maak_werkdocumenten.py
Vereist:  python-docx

LET OP bij aanpassen:
 - De Leveringsplanning bestaat uit VIER LOSSE REGELS (gewone alinea's, geen tabel): in stap 4 wist de buur de
   regel van 10.30 uur, en dat moet kunnen met selecteren en Delete.
 - De leveringen komen uit les 05 en 06: Bakkerij Korstjes (7.30 uur, poort 3), Papierhandel Vellekens (8.15 uur),
   Koffiebranderij De Bonenbaas (10.30 uur) en Speelgoed Tolletje (13.00 uur, 3 pallets, poort 3), op dinsdag
   13 oktober 2026. De poorten en pallets van Vellekens en De Bonenbaas zijn voor deze les verzonnen.
 - De personages zijn die van de hele reeks: Tom Peeters (magazijn), Sara Claes (preventiedienst), Lotte Maes
   (personeelsdienst). Alle namen en bedrijven zijn fictief. Geen e-mailadressen, geen namen van leerlingen.
 - Rij 1 van deel 1 en deel 3 staat als voorbeeld ingevuld.
"""
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
UIT = os.path.join(HERE, "WP3_Delen-met-de-juiste-rechten.docx")
PLANNING = os.path.join(HERE, "Leveringsplanning.docx")

NAVY = RGBColor(0x2A, 0x39, 0x73)
GREY = RGBColor(0x4D, 0x55, 0x73)
GRIJS = RGBColor(0x80, 0x80, 0x80)

# De vier leveringen van dinsdag 13 oktober 2026: uur, leverancier, poort, pallets
LEVERINGEN = [
    ("7.30 uur", "Bakkerij Korstjes", "poort 3", "2 pallets"),
    ("8.15 uur", "Papierhandel Vellekens", "poort 1", "1 pallet"),
    ("10.30 uur", "Koffiebranderij De Bonenbaas", "poort 2", "1 pallet"),
    ("13.00 uur", "Speelgoed Tolletje", "poort 3", "3 pallets"),
]


# ---------- hulpfuncties (zoals in les 06 en 07) ----------
def lettertype(run, naam="Arial"):
    run.font.name = naam
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), naam)


def basis_document(titel):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(11)
    rpr = st.element.get_or_add_rPr()
    rpr.rFonts.set(qn("w:eastAsia"), "Arial")
    taal = OxmlElement("w:lang")
    taal.set(qn("w:val"), "nl-BE")
    rpr.append(taal)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.1
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Cm(2)
    doc.core_properties.title = titel
    doc.core_properties.author = "Toegepaste Informatica"
    return doc


def tekst(doc, s, vet=False, klein=False, cursief=False, na=6, grootte=11, kleur=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(na)
    r = p.add_run(s)
    lettertype(r)
    r.bold = vet
    r.italic = cursief
    r.font.size = Pt(9.5 if klein else grootte)
    if klein:
        r.font.color.rgb = GREY
    if kleur is not None:
        r.font.color.rgb = kleur
    return p


def kop(doc, s, grootte=13, voor=14):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(voor)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(s)
    lettertype(r)
    r.bold = True
    r.font.size = Pt(grootte)
    r.font.color.rgb = NAVY
    return p


def _voeg_in_voor(ouder, el, opvolgers):
    """Voeg el in vóór het eerste bestaande element uit opvolgers (OOXML-volgorde), anders achteraan."""
    for tag in opvolgers:
        ref = ouder.find(qn(tag))
        if ref is not None:
            ref.addprevious(el)
            return el
    ouder.append(el)
    return el


def vaste_tabel(t, breedtes_cm, randkleur="8C8C8C"):
    """Vaste kolombreedtes en dunne randen, zodat Google Documenten de tabel toont zoals bedoeld."""
    tbl = t._tbl
    tblpr = tbl.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = _voeg_in_voor(tblpr, OxmlElement("w:tblW"),
                             ("w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd",
                              "w:tblLayout", "w:tblCellMar", "w:tblLook"))
    tblw.set(qn("w:w"), str(int(round(sum(breedtes_cm) * 567))))
    tblw.set(qn("w:type"), "dxa")
    randen = OxmlElement("w:tblBorders")
    for kant in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + kant)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), randkleur)
        randen.append(el)
    _voeg_in_voor(tblpr, randen, ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"))
    indeling = OxmlElement("w:tblLayout")
    indeling.set(qn("w:type"), "fixed")
    _voeg_in_voor(tblpr, indeling, ("w:tblCellMar", "w:tblLook"))
    for kolom, b in zip(tbl.tblGrid.findall(qn("w:gridCol")), breedtes_cm):
        kolom.set(qn("w:w"), str(int(round(b * 567))))
    for rij in t.rows:
        for cel, b in zip(rij.cells, breedtes_cm):
            cel.width = Cm(b)
    return t


def schaduw(cel, kleur="E4E8F6"):
    tcpr = cel._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), kleur)
    tcpr.append(shd)


def cel_tekst(cel, s, vet=False, grootte=10.5, cursief=False):
    cel.text = ""
    p = cel.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(s)
    lettertype(r)
    r.bold = vet
    r.italic = cursief
    r.font.size = Pt(grootte)


def antwoordlijnen(doc, aantal=2):
    for _ in range(aantal):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("_" * 78)
        lettertype(r)
        r.font.color.rgb = GRIJS


def tabel(doc, koppen, breedtes, rijen, hoogte=0.9):
    """Een tabel met kopregel; rij 1 staat al ingevuld (cursief) als voorbeeld; kolom 1 vet."""
    t = doc.add_table(rows=len(rijen) + 1, cols=len(koppen))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    vaste_tabel(t, breedtes)
    for i, k in enumerate(koppen):
        cel_tekst(t.rows[0].cells[i], k, vet=True)
        schaduw(t.rows[0].cells[i])
    for i, rij in enumerate(rijen, start=1):
        for j, w in enumerate(rij):
            if w:
                cel_tekst(t.rows[i].cells[j], w, cursief=(i == 1 and j > 0), grootte=9.5, vet=(j == 0))
        t.rows[i].height = Cm(hoogte)
    return t


# ---------- 1. Het werkdocument ----------
def werkdocument():
    doc = basis_document("WP3 Delen met de juiste rechten - werkdocument")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("WERKDOCUMENT — DELEN MET DE JUISTE RECHTEN")
    lettertype(r)
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = NAVY
    tekst(doc, "Les 08 — Mijn digitale werkplek 3  ·  Toegepaste Informatica  ·  NovaDepot", klein=True, na=8)

    t = doc.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    vaste_tabel(t, [7.0, 4.5, 5.5])
    for i, label in enumerate(["Naam:", "Klas:", "Datum:"]):
        cel_tekst(t.rows[0].cells[i], label + " ", vet=True, grootte=11)

    kop(doc, "Zo werk je", 12, voor=10)
    for s in [
        "1.  Op de lespagina lees je wat je moet doen, stap voor stap.",
        "2.  Je test het delen met je buur, in je Leveringsplanning.",
        "3.  In dit werkdocument schrijf je op wat je merkte.",
        "4.  Klik op Inleveren: Classroom levert dit document en je Leveringsplanning samen in.",
    ]:
        tekst(doc, s, na=1)

    # ---- deel 1 ----
    kop(doc, "Deel 1 — Wat kan je met elke rol? (stap 2, 3 en 4)")
    tekst(doc, "Test het in de planning van je buur. Vul na elke test in: ja of nee. Rij 1 is een voorbeeld.", klein=True)
    tabel(doc, ["Wat kan je?", "Als Kijker", "Als Reageerder", "Als Bewerker"], [6.0, 3.6, 3.7, 3.7], [
        ("De planning lezen", "ja", "ja", "ja"),
        ("Een reactie schrijven", "", "", ""),
        ("Een regel zelf veranderen", "", "", ""),
    ])
    tekst(doc, "Zelf veranderen: de tekst verandert meteen. Een voorstel (suggestie) telt niet mee.", klein=True, na=2)

    # ---- deel 2 ----
    kop(doc, "Deel 2 — Een regel terug (stap 4)")
    tekst(doc, "Vraag 1. Welke regel wiste je buur in jouw planning?", vet=True, na=2)
    antwoordlijnen(doc, 1)
    tekst(doc, "Vraag 2. Hoe kreeg je de regel terug? Schrijf het in je eigen woorden.", vet=True, na=2)
    antwoordlijnen(doc, 3)

    # ---- deel 3 ----
    kop(doc, "Deel 3 — Wie krijgt welke rol? (stap 5 en 6)")
    tekst(doc, "Kies Kijker, Reageerder of Bewerker. Geef zo weinig rechten als nodig. Rij 1 is een voorbeeld.", klein=True)
    tabel(doc, ["Wie?", "Wat moet die persoon kunnen?", "Welke rol?"], [4.8, 8.2, 4.0], [
        ("Sara Claes (preventiedienst)", "de planning lezen, maar niets veranderen", "Kijker"),
        ("Lotte Maes (personeelsdienst)", "de planning lezen en een opmerking schrijven", ""),
        ("Een collega in het magazijn", "de planning lezen en aanpassen", ""),
        ("Je leraar, voor je map NovaDepot", "je map bekijken", ""),
    ], hoogte=1.0)
    tekst(doc, "")
    tekst(doc, "Waarom kies je bij Algemene toegang Beperkt, en niet Iedereen met de link?", vet=True, na=2)
    antwoordlijnen(doc, 2)

    # ---- zelfcontrole ----
    kop(doc, "Zelfcontrole — aankruisen vóór je inlevert")
    for s in [
        "In mijn Leveringsplanning staat de regel van 10.30 uur er weer in.",
        "Mijn buur heeft geen toegang meer tot mijn Leveringsplanning.",
        "Mijn leraar kan mijn map NovaDepot bekijken, als Kijker.",
        "Bij Algemene toegang staat Beperkt.",
        "Deel 1, 2 en 3 zijn ingevuld.",
    ]:
        tekst(doc, "☐  " + s, na=1)

    # ---- extra ----
    kop(doc, "Extra — niet verplicht")
    tekst(doc, "Alleen als je al ingeleverd hebt. Klik in de opdracht op Inleveren ongedaan maken. "
               "Lever daarna opnieuw in.", klein=True)
    tekst(doc, "a) Geef de herstelde versie van je planning een naam, bijvoorbeeld Planning in orde. "
               "Welke naam gaf je?", vet=True, na=2)
    antwoordlijnen(doc, 1)
    tekst(doc, "b) Beantwoord de reactie van je buur in je planning. Wat antwoordde je?", vet=True, na=2)
    antwoordlijnen(doc, 1)

    doc.save(UIT)


# ---------- 2. De Leveringsplanning ----------
def leveringsplanning():
    doc = basis_document("Leveringsplanning")
    tekst(doc, "LEVERINGSPLANNING", vet=True, grootte=24, kleur=NAVY, na=2)
    tekst(doc, "Dinsdag 13 oktober 2026", vet=True, grootte=16, na=0)
    tekst(doc, "NovaDepot  ·  magazijn  ·  opgemaakt door Tom Peeters  ·  fictief bedrijf", klein=True, na=16)
    for uur, wie, poort, pallets in LEVERINGEN:
        # elke levering is één gewone regel: selecteren en Delete volstaat om ze te wissen
        tekst(doc, f"{uur}  ·  {wie}  ·  {poort}  ·  {pallets}", grootte=13, na=10)
    tekst(doc, "")
    tekst(doc, "Vragen of opmerkingen? Schrijf ze als reactie in dit document.", klein=True)
    doc.save(PLANNING)


if __name__ == "__main__":
    werkdocument()
    leveringsplanning()
    print("Klaar.")
    print("  werkdocument om in te leveren : " + os.path.basename(UIT))
    print("  oefendocument (duo-test)      : " + os.path.basename(PLANNING))
