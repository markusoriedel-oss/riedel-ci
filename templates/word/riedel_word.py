"""Hilfsfunktionen, um Word-Dokumente auf Basis der RIEDEL-Vorlagen mit python-docx zu füllen.

Beispiel:
    import riedel_word as rw
    doc = rw.aus_vorlage("RIEDEL Marktwerteinschätzung.dotx")
    doc.add_paragraph("Marktwerteinschätzung", style="Title")
    rw.kennzahlen(doc, [("Wertindikation", "7,1 bis 8,1 Mio. €"), ("Mittelwert", "7,55 Mio. €")])
    doc.add_paragraph("1. Grundstück", style="Heading 1")
    rw.tabelle(doc, [["Schritt", "Grundlage", "Wert"], ["…", "…", "…"], ["Ergebnis", "", "…"]],
               [42, 88, 40], ergebnis=True, rechts=(2,))
    doc.save("Ergebnis.docx")
"""
import io
import zipfile

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Mm, Pt

TEXTBREITE = 167.1        # Satzspiegel der Marktwerteinschätzung in mm (20 mm links bis Logokante 187,1 mm)
LINIE = "CCCCCC"
LINIE_STARK = "000000"


def twips(mm):
    return int(round(mm / 25.4 * 1440))


def aus_vorlage(dotx):
    """Öffnet eine .dotx als neues Dokument und entfernt den Musterinhalt (Kopf, Fuß und Formate bleiben)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(dotx) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            daten = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                daten = daten.replace(b"wordprocessingml.template.main+xml", b"wordprocessingml.document.main+xml")
            zout.writestr(item, daten)
    buf.seek(0)
    doc = Document(buf)
    body = doc.element.body
    for el in list(body):
        if not el.tag.endswith("sectPr"):
            body.remove(el)
    return doc


def fusszeile_betreff(doc, betreff):
    """Titel in der Fußzeile der Folgeseiten setzen (Dokumenteigenschaft und angezeigtes Feldergebnis)."""
    doc.core_properties.subject = betreff
    doc.core_properties.title = betreff
    for p in doc.sections[0].footer.paragraphs:
        for r in p.runs:
            if r.text == "Titel des Dokuments":
                r.text = betreff


def nb(text):
    """Geschützte Leerzeichen vor Einheiten."""
    for a in (" %", " €", " m²", " Mio.", " m ", " Minuten"):
        text = text.replace(a, " " + a[1:])
    return text.replace("§ 34", "§ 34")


def zelle_raender(zelle, links=None, rechts=None):
    tcpr = zelle._tc.get_or_add_tcPr()
    mar = tcpr.find(qn("w:tcMar"))
    if mar is None:
        mar = OxmlElement("w:tcMar")
        tcpr.append(mar)
    for seite, wert in (("left", links), ("right", rechts)):
        if wert is not None:
            el = OxmlElement(f"w:{seite}")
            el.set(qn("w:w"), str(wert))
            el.set(qn("w:type"), "dxa")
            mar.append(el)


def tabellen_look(tabelle, kopf=True, ergebnis=False):
    tblpr = tabelle._tbl.tblPr
    for alt in tblpr.findall(qn("w:tblLook")):
        tblpr.remove(alt)
    look = OxmlElement("w:tblLook")
    look.set(qn("w:val"), "0000")
    look.set(qn("w:firstRow"), "1" if kopf else "0")
    look.set(qn("w:lastRow"), "1" if ergebnis else "0")
    look.set(qn("w:firstColumn"), "0")
    look.set(qn("w:lastColumn"), "0")
    look.set(qn("w:noHBand"), "1")
    look.set(qn("w:noVBand"), "1")
    tblpr.append(look)


def tabelle(doc, zeilen, breiten_mm, kopf=True, ergebnis=False, rechts=()):
    """Tabelle im RIEDEL-Format. Erste und letzte Spalte schließen bündig mit dem Satzspiegel ab."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    summe = sum(breiten_mm)
    breiten_mm = [b * TEXTBREITE / summe for b in breiten_mm]
    t = doc.add_table(rows=len(zeilen), cols=len(breiten_mm))
    t.style = doc.styles["RIEDEL Tabelle"]
    t.autofit = False
    tabellen_look(t, kopf, ergebnis)
    tblpr = t._tbl.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    tblw.set(qn("w:w"), str(twips(TEXTBREITE)))
    tblw.set(qn("w:type"), "dxa")
    tblpr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    for gc, b in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), breiten_mm):
        gc.set(qn("w:w"), str(twips(b)))
    for i, zeile in enumerate(zeilen):
        row = t.rows[i]
        trpr = row._tr.get_or_add_trPr()
        trpr.append(OxmlElement("w:cantSplit"))
        for j, wert in enumerate(zeile):
            c = row.cells[j]
            c.width = Mm(breiten_mm[j])
            if j == 0:
                zelle_raender(c, links=0)
            if j == len(zeile) - 1:
                zelle_raender(c, rechts=0)
            p = c.paragraphs[0]
            if kopf and i == 0:
                p.style = "RIEDEL Tabellenkopf"
            elif ergebnis and i == len(zeilen) - 1:
                p.style = "RIEDEL Tabellenergebnis"
            else:
                p.style = "RIEDEL Tabellentext"
            p.text = wert
            if j in rechts:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # Ergebniszeile: Bezeichnung über leere Folgezellen ziehen, damit sie nicht umbricht
    if ergebnis:
        letzte = zeilen[-1]
        k = 1
        while k < len(letzte) - 1 and not letzte[k]:
            k += 1
        if k > 1:
            row = t.rows[-1]
            zelle = row.cells[0].merge(row.cells[k - 1])
            for extra in zelle.paragraphs[1:]:
                extra._p.getparent().remove(extra._p)
    return t


def kennzahlen(doc, werte):
    """Kennzahlenfeld: gleich starke schwarze Linie oben und unten, keine Farbfläche."""
    n = len(werte)
    t = doc.add_table(rows=1, cols=n)
    t.autofit = False
    tblpr = t._tbl.tblPr
    tblpr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}><w:top w:val="single" w:sz="8" w:space="0" w:color="{LINIE_STARK}"/>'
        f'<w:left w:val="nil"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="{LINIE_STARK}"/>'
        f'<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'))
    tblpr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblpr.append(parse_xml(
        f'<w:tblCellMar {nsdecls("w")}><w:top w:w="142" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
        f'<w:bottom w:w="170" w:type="dxa"/><w:right w:w="170" w:type="dxa"/></w:tblCellMar>'))
    tblw = tblpr.find(qn("w:tblW"))
    tblw.set(qn("w:w"), str(twips(TEXTBREITE)))
    tblw.set(qn("w:type"), "dxa")
    b = TEXTBREITE / n
    for gc in t._tbl.tblGrid.findall(qn("w:gridCol")):
        gc.set(qn("w:w"), str(twips(b)))
    for j, (bez, wert) in enumerate(werte):
        c = t.rows[0].cells[j]
        c.width = Mm(b)
        p = c.paragraphs[0]
        p.style = "RIEDEL Kennzahl Bezeichnung"
        p.add_run(bez)
        c.add_paragraph(wert, style="RIEDEL Kennzahl Wert")
    return t


def abstand(doc, pt=8):
    p = doc.add_paragraph(style="RIEDEL Quelle")
    p.paragraph_format.line_spacing = Pt(pt)
    p.paragraph_format.space_after = Pt(0)
    return p


def abbildung(doc, pfad, unterschrift, breite_mm=TEXTBREITE):
    """Grafik über die volle Satzbreite mit Abbildungsunterschrift (Formatvorlage Caption)."""
    p = doc.add_paragraph(style="RIEDEL Quelle")
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(pfad), width=Mm(breite_mm))
    return doc.add_paragraph(unterschrift, style="Caption")
