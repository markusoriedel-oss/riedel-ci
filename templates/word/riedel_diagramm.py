"""Diagramme im RIEDEL CI für Word-Dokumente (PNG, 300 dpi).

Aufbau nach dem RIEDEL PowerPoint-Master: Balken in Dunkelblau und Beige, Orange nur zur Hervorhebung, feine waagerechte Hilfslinien,
Werte direkt an den Balken, Legende und Quelle klein darunter. Farben und Schrift aus dem neuen CI.
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.patches
import matplotlib.ticker
import matplotlib.pyplot as plt
from matplotlib import font_manager

FONTS = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
for f in ("SuisseIntl-Light.otf", "SuisseIntl-Book.ttf", "SuisseIntl-Medium.otf"):
    p = os.path.join(FONTS, f)
    if os.path.exists(p):
        font_manager.fontManager.addfont(p)

# Farben für Abbildungen aus riedel-basis.css (riedel-builders): Dunkelblau, Beige, Orange zur Hervorhebung
TEXT = "#000000"
DUNKELBLAU = "#00102A"
BEIGE = "#A0927E"
ORANGE = "#EB5D49"
LINIE = "#CCCCCC"
FLAECHE = "#EFEFEF"     # neutrale, sehr helle Hinterlegung
GRAU = "#000000"

MM = 1 / 25.4


def _stil():
    plt.rcParams.update({
        "font.family": "Suisse Intl",
        "font.weight": 300,
        "font.size": 7.5,
        "text.color": TEXT,
        "axes.labelcolor": GRAU,
        "xtick.color": GRAU,
        "ytick.color": TEXT,
        "axes.edgecolor": LINIE,
        "svg.fonttype": "none",
    })


def zahl(x, stellen=2):
    return f"{x:,.{stellen}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def wertbandbreite(pfad, zeilen, band, mittel, achse, breite_mm=167, einheit="Mio. €"):
    """Waagerechtes Bandbreitendiagramm.

    zeilen: Liste (Bezeichnung, von, bis); bei von == bis wird ein Punktwert gezeichnet.
    band:   (von, bis) der Wertindikation, als Fläche hinterlegt.
    mittel: Mittelwert als Linie.
    achse:  (min, max, schritt) der x-Achse.
    """
    _stil()
    hoehe_mm = 19 + 8.5 * len(zeilen)
    fig, ax = plt.subplots(figsize=(breite_mm * MM, hoehe_mm * MM), dpi=300)
    fig.subplots_adjust(left=0.30, right=0.985, top=1 - 7 / hoehe_mm, bottom=14 / hoehe_mm)

    ax.axvspan(band[0], band[1], color=FLAECHE, zorder=0, lw=0)
    ax.axvline(mittel, color=ORANGE, lw=0.9, zorder=2)
    ax.text(mittel, len(zeilen) - 0.35, f"Mittelwert {zahl(mittel)} {einheit}", ha="center", va="bottom",
            fontsize=6.5, color=ORANGE)
    # Legende unter der Achse: Fläche = Wertindikation
    ax.add_patch(matplotlib.patches.Rectangle((0.0, -0.40), 0.018, 0.07, transform=ax.transAxes,
                                              color=FLAECHE, clip_on=False, lw=0))
    ax.text(0.026, -0.365, f"Wertindikation {zahl(band[0], 1)} bis {zahl(band[1], 1)} {einheit}",
            transform=ax.transAxes, ha="left", va="center", fontsize=6.5, color=GRAU)

    for i, (name, von, bis) in enumerate(reversed(zeilen)):
        y = i
        if bis > von:
            ax.barh(y, bis - von, left=von, height=0.32, color=BEIGE, zorder=3, lw=0)
            ax.text(von - 0.04, y, zahl(von, 1), ha="right", va="center", fontsize=7)
            ax.text(bis + 0.04, y, zahl(bis, 1), ha="left", va="center", fontsize=7)
        else:
            ax.plot([von], [y], "o", ms=5.5, color=DUNKELBLAU, zorder=4, mec="white", mew=0.8)
            ax.text(von, y + 0.28, zahl(von), ha="center", va="bottom", fontsize=7)

    ax.set_yticks(range(len(zeilen)))
    ax.set_yticklabels([z[0] for z in reversed(zeilen)])
    lo, hi, step = achse
    ticks = [lo + k * step for k in range(int(round((hi - lo) / step)) + 1)]
    ax.set_xticks(ticks)
    ax.set_xticklabels([zahl(t, 1) for t in ticks])
    ax.set_xlim(lo, hi)
    ax.set_ylim(-0.6, len(zeilen) - 0.4)
    ax.grid(axis="x", color=LINIE, lw=0.4, zorder=1)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(LINIE)
    ax.tick_params(axis="y", length=0, pad=6)
    ax.tick_params(axis="x", length=0, pad=3, labelsize=6.5)
    ax.text(1.0, -0.365, einheit, transform=ax.transAxes, ha="right", va="center", fontsize=6.5, color=GRAU)
    fig.savefig(pfad, dpi=300, transparent=False, facecolor="white")
    plt.close(fig)
    return pfad


def balken(pfad, kategorien, werte, einheit, breite_mm=167, hoehe_mm=55, wertformat=lambda v: zahl(v, 0)):
    """Säulendiagramm mit einer Reihe, Werte über den Säulen."""
    _stil()
    fig, ax = plt.subplots(figsize=(breite_mm * MM, hoehe_mm * MM), dpi=300)
    fig.subplots_adjust(left=0.07, right=0.99, top=0.86, bottom=0.14)
    farben = [DUNKELBLAU] * (len(werte) - 1) + [ORANGE]
    ax.bar(kategorien, werte, width=0.45, color=farben, zorder=3, lw=0)
    for x, v in zip(kategorien, werte):
        ax.text(x, v, wertformat(v), ha="center", va="bottom", fontsize=7, color=TEXT)
    ax.grid(axis="y", color=LINIE, lw=0.4, zorder=0)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(LINIE)
    ax.tick_params(length=0, labelsize=6.5)
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: zahl(v, 0)))
    ax.text(0, 1.04, einheit, transform=ax.transAxes, ha="right", va="bottom", fontsize=6.5, color=GRAU)
    fig.savefig(pfad, dpi=300, facecolor="white")
    plt.close(fig)
    return pfad
