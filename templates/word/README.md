# Word-Vorlagen RIEDEL Immobilien GmbH (CI 2026)

| Datei | Verwendung |
|-------|-----------|
| `RIEDEL Briefpapier.dotx` | Briefe. Briefbogen nach der Satzdatei des Marketings (Martin et Karczinski, 2026): Logo, Falzmarken, Anschriftfeld nach Maß, Betreff mit Datum, Fußzeile mit Firmenangaben auf Seite 1, Folgeseiten nur mit Logo. |
| `RIEDEL Marktwerteinschätzung.dotx` | Marktwerteinschätzungen, Bewertungen und sachliche Berichte. Gleiches Logo und gleiche Fußzeile auf Seite 1, breiter Satzspiegel bis zur Logokante, ohne Falzmarken und Anschriftfeld. Folgeseiten mit Titel und „Seite x von y“. |
| `riedel_word.py` | Hilfsfunktionen für python-docx: Dokument aus Vorlage öffnen, Tabelle und Kennzahlenfeld im RIEDEL-Format, Abbildung mit Unterschrift, Fußzeilentitel. |
| `riedel_diagramm.py` | Diagramme im RIEDEL CI (matplotlib, PNG 300 dpi): Bandbreitendiagramm, Säulendiagramm. |

## Gestaltungsregeln

- Schrift: Suisse Intl Light 9,5 pt für Fließtext. In der Marktwerteinschätzung Titel und Kapitel in Cardinal Photo. Beide Schriften sind lizenziert und müssen installiert sein.
- **Text immer schwarz.** Beige nur im Logo.
- Tabellen: nur waagerechte Linien, zwischen den Zeilen hellgrau `#CCCCCC`, unter dem Kopf und um die Ergebniszeile schwarz. Keine Farbflächen. Kopf in Suisse Intl Medium 9,5 pt.
- Kennzahlenfeld: gleich starke schwarze Linie oben und unten.
- Kapitel im Format „1. Grundstück und Baurecht“, Unterkapitel „4.1 …“, ohne Einzug.
- Abbildungen: Dunkelblau `#00102A`, Beige `#A0927E`, Orange `#EB5D49` nur zur Hervorhebung, Linien `#CCCCCC`.
- Tabellen und Grafiken nur dort, wo sie das Verständnis wesentlich unterstützen, sonst Fließtext.
- Geschützte Leerzeichen vor Einheiten (`nb()` in `riedel_word.py`), keine Gedankenstriche.

## Formatvorlagen

**Briefpapier:** Standard, RIEDEL Absenderzeile, RIEDEL Anschrift, RIEDEL Betreff, RIEDEL Datum, RIEDEL Grußformel, RIEDEL Fußzeile.

**Marktwerteinschätzung:** Standard, Title, Subtitle, Heading 1, Heading 2, RIEDEL Tabelle (Tabellenformat), RIEDEL Tabellenkopf, RIEDEL Tabellentext, RIEDEL Tabellenergebnis, RIEDEL Kennzahl Bezeichnung, RIEDEL Kennzahl Wert, Caption, RIEDEL Quelle, RIEDEL Hinweis Überschrift, RIEDEL Hinweis, RIEDEL Unterschrift.

## Verwendung mit python-docx

```python
import riedel_word as rw
doc = rw.aus_vorlage("RIEDEL Marktwerteinschätzung.dotx")
rw.fusszeile_betreff(doc, "Marktwerteinschätzung Musterstraße 1, 80000 München")
doc.add_paragraph("Marktwerteinschätzung", style="Title")
doc.add_paragraph("Musterstraße 1, 80000 München. Einfamilienhaus mit 180 m² Wohnfläche.", style="Subtitle")
rw.kennzahlen(doc, [("Wertindikation", "1,9 bis 2,1 Mio. €"), ("Mittelwert", "2,0 Mio. €")])
doc.add_paragraph("1. Grundstück", style="Heading 1")
doc.save("Marktwerteinschätzung Musterstraße 1.docx")
```

Die .dotx-Dateien lassen sich in Word auch direkt per Doppelklick als neues Dokument öffnen.
