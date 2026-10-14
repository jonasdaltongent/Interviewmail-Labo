#!/usr/bin/env python3
"""
maak_stapjes.py — de beelden bij de stapjes op de lespagina (assets/stapjes/), uit het werkdocument zelf

Gebruik:  python3 maak_stapjes.py <leeg.pdf> <met-x.pdf>     (vereist PyMuPDF en Pillow)

  leeg.pdf    Interviewmail zoals Google Documenten het toont: upload Interviewmail.docx in Drive (het wordt een
              Google-document) en kies Bestand › Downloaden › PDF-document.
  met-x.pdf   hetzelfde, met een x in het eerste vak van deel 4 (voor s6-kruis.png). Op 14-10-2026 is dat
              *Test Interviewmail (Claude)* in de demomap van Jonas' Drive.

Waarom zo (14-10-2026, Stap voor stap / Kort, zoals MW les 3): dezelfde letters en vakken die de leerling
ziet, op ware grootte (Google Documenten op 100 %), met één rood kader (#C5221F, witte rand erbuiten).
Verandert het werkdocument, draai dan dit script opnieuw.
"""
import os
import sys

import fitz
from PIL import Image, ImageDraw

HIER = os.path.dirname(os.path.abspath(__file__))
UIT = os.path.join(HIER, "..", "assets", "stapjes")
SCHAAL = 2 * 96 / 72          # eerst op dubbele grootte tekenen, dan verkleinen: scherpe letters
ROOD = (197, 34, 31)
VAK = (0xF3, 0xF5, 0xFC)      # HEX_VAK in huisstijl.py
KADER = (0xEE, 0xF0, 0xF9)    # HEX_KADER
KOP = (0xE4, 0xE8, 0xF6)      # HEX_KOP: kopregel van een tabel
VENSTERKOP = (0xF2, 0xF6, 0xFC)


def kleur(fill):
    return tuple(round(c * 255) for c in fill) if fill else None


def vlakken(page, rgb):
    """De gevulde rechthoeken met deze kleur, van boven naar onder en van links naar rechts."""
    uit = []
    for d in page.get_drawings():
        k = kleur(d.get("fill"))
        if k and all(abs(a - b) <= 2 for a, b in zip(k, rgb)):
            uit.append(fitz.Rect(d["rect"]))
    uit.sort(key=lambda r: (round(r.y0), r.x0))
    return uit


def tekst(page, s, nr=0):
    treffers = page.search_for(s)
    if len(treffers) <= nr:
        raise SystemExit(f"'{s}' niet gevonden op bladzijde {page.number + 1}")
    return treffers[nr]


def bladzijde_met(doc, s):
    for p in doc:
        if p.search_for(s):
            return p
    raise SystemExit(f"'{s}' staat in geen enkele bladzijde")


def bewaar(page, naam, uitsnede, kaders=(), boven=None, onder=None):
    """Snijdt uitsnede (in pdf-punten) uit, tekent de rode kaders en bewaart op ware grootte.
    boven/onder: wat boven of onder die hoogte valt, wordt wit (geen halve regels van de buurrij)."""
    uitsnede = fitz.Rect(uitsnede) & page.rect
    pix = page.get_pixmap(matrix=fitz.Matrix(SCHAAL, SCHAAL), clip=uitsnede, alpha=False)
    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    d = ImageDraw.Draw(im)
    if boven is not None:
        d.rectangle([0, 0, im.width, (boven - uitsnede.y0) * SCHAAL], fill="white")
    if onder is not None:
        d.rectangle([0, (onder - uitsnede.y0) * SCHAAL, im.width, im.height], fill="white")
    for k in kaders:
        k = fitz.Rect(k)
        box = [(k.x0 - uitsnede.x0) * SCHAAL - 6, (k.y0 - uitsnede.y0) * SCHAAL - 6,
               (k.x1 - uitsnede.x0) * SCHAAL + 6, (k.y1 - uitsnede.y0) * SCHAAL + 6]
        wit = [box[0] - 4, box[1] - 4, box[2] + 4, box[3] + 4]
        d.rounded_rectangle(wit, radius=16, outline=(255, 255, 255), width=4)
        d.rounded_rectangle(box, radius=12, outline=ROOD, width=5)
    im = im.resize((im.width // 2, im.height // 2), Image.LANCZOS)
    im.save(os.path.join(UIT, naam + ".png"), optimize=True)
    print("gemaakt:", naam + ".png", im.size)


def rand(r, x=8, y=8):
    return fitz.Rect(r.x0 - x, r.y0 - y, r.x1 + x, r.y1 + y)


def band(r, links, rechts):
    """Dezelfde hoogte als r, over de volle breedte van links tot rechts."""
    return fitz.Rect(links, r.y0, rechts, r.y1)


def main(leeg, met_x):
    os.makedirs(UIT, exist_ok=True)
    doc = fitz.open(leeg)

    # ---- bladzijde 1: Je onderzoek en deel 1
    p = bladzijde_met(doc, "Je onderzoek")
    kader = vlakken(p, KADER)[0]
    k_in, k_uit = kader.x0 + 2, kader.x1 - 2
    interview, begel = tekst(p, "Het interview:"), tekst(p, "ines.coppens@onderzoekslabo.example")
    vraagt, wanneer = tekst(p, "Wat je vraagt:"), tekst(p, "Tegen wanneer je een antwoord hoopt:")
    bewaar(p, "s1-je-onderzoek", rand(kader, 10, 5),
           [fitz.Rect(k_in, interview.y0 - 1, k_uit, begel.y1 + 1)], boven=kader.y0 - 1)
    bewaar(p, "s1-vraag", rand(kader, 10, 5),
           [fitz.Rect(k_in, vraagt.y0 - 1, k_uit, wanneer.y1 + 1)], boven=kader.y0 - 1)
    bewaar(p, "s4-adres", rand(kader, 10, 5), [tekst(p, "hanne.declercq@delantaarn.example")],
           boven=kader.y0 - 1)

    kop1 = tekst(p, "Deel 1")
    koppen = [k for k in vlakken(p, KOP) if k.y0 > kop1.y1]
    links, rechts = min(k.x0 for k in koppen), max(k.x1 for k in koppen)
    kopregel = fitz.Rect(links, min(k.y0 for k in koppen), rechts, max(k.y1 for k in koppen))
    rijen = [v for v in vlakken(p, VAK) if v.y0 > kopregel.y1]       # rij 2 tot 6, de laatste kolom
    rij1 = fitz.Rect(links, kopregel.y1, rechts, rijen[0].y0)
    bewaar(p, "s2-deel1", rand(fitz.Rect(links, kop1.y0, rechts, rijen[0].y1)), [rij1], onder=rijen[0].y1 + 1)
    for nr, v in enumerate(rijen, start=2):
        rij = band(v, links, rechts)
        bewaar(p, f"s2-rij{nr}", rand(rij, 8, 8), [v], boven=rij.y0 - 1, onder=rij.y1 + 1)

    # ---- bladzijde 2: deel 2 en vraag 1
    p = bladzijde_met(doc, "Aan, Cc of Bcc?")
    kop2, vraag1 = tekst(p, "Deel 2"), tekst(p, "Vraag 1.")
    vakken = vlakken(p, VAK)
    sit = [v for v in vakken if kop2.y1 < v.y0 < vraag1.y0]          # situatie a, b, c, d
    antw1 = [v for v in vakken if v.y0 > vraag1.y1][0]
    links, rechts = antw1.x0, antw1.x1
    a = fitz.Rect(sit[0].x0, sit[0].y0 + 3, sit[0].x1, sit[0].y1)      # 3 punten lager: de kopregel blijft leesbaar
    bewaar(p, "s3-situatie-a", rand(fitz.Rect(links, kop2.y0, rechts, sit[-1].y1)), [a])
    bewaar(p, "s3-situaties-bcd", rand(fitz.Rect(links, sit[1].y0, rechts, sit[-1].y1)), [sit[1] | sit[3]],
           boven=sit[1].y0 - 1, onder=sit[-1].y1 + 1)
    bewaar(p, "s3-vraag1", rand(fitz.Rect(links, vraag1.y0, rechts, antw1.y1), 8, 10),
           [fitz.Rect(links, vraag1.y0 - 2, rechts, antw1.y1)])

    # ---- bladzijde 3: het venster Nieuw bericht
    p = bladzijde_met(doc, "Nieuw bericht")
    kop = vlakken(p, VENSTERKOP)[0]
    links, rechts = kop.x0, kop.x1
    velden = [v for v in vlakken(p, VAK) if v.y0 >= kop.y1 - 1]     # Aan, Cc, Bcc, Onderwerp, dan het bericht
    namen = ["Aan", "Cc", "Bcc", "Onderwerp", "aanspreking", "inleiding", "midden", "slot", "slotgroet",
             "handtekening"]
    if len(velden) != len(namen):
        raise SystemExit(f"verwacht {len(namen)} vakken in het venster, gevonden: {len(velden)}")
    rijen = {n: band(v, links, rechts) for n, v in zip(namen, velden)}
    boven = fitz.Rect(links, kop.y0, rechts, rijen["aanspreking"].y1)
    bewaar(p, "s4-venster", rand(boven), onder=boven.y1 + 1)
    for n in ("Aan", "Cc", "Onderwerp", "aanspreking"):
        bewaar(p, "s4-" + n.lower(), rand(boven), [rijen[n]], onder=boven.y1 + 1)
    bericht = fitz.Rect(links, rijen["aanspreking"].y0, rechts, rijen["handtekening"].y1)
    for n in ("inleiding", "midden", "slot"):
        bewaar(p, "s5-" + n, rand(bericht), [rijen[n]], boven=bericht.y0 - 1)
    bewaar(p, "s5-slotgroet", rand(bericht), [rijen["slotgroet"] | rijen["handtekening"]], boven=bericht.y0 - 1)

    # ---- bladzijde 4: deel 4 en deel 5
    p = bladzijde_met(doc, "Feedback van je buur")
    kop4, kop5 = tekst(p, "Deel 4"), tekst(p, "Deel 5")
    vakken = vlakken(p, VAK)
    top, tip, aangepast = (tekst(p, "Dit is goed"), tekst(p, "Dit kan beter"), tekst(p, "Wat heb ik aangepast?"))
    vinkjes = [v for v in vakken if kop4.y1 < v.y0 < top.y0]
    antw_top = [v for v in vakken if top.y1 < v.y0 < tip.y0][0]
    antw_tip = [v for v in vakken if tip.y1 < v.y0 < aangepast.y0][0]
    antw_aangepast = [v for v in vakken if aangepast.y1 < v.y0 < kop5.y0][0]
    antw5 = [v for v in vakken if v.y0 > kop5.y1][0]
    links, rechts = antw5.x0, antw5.x1
    bewaar(p, "s6-deel4", rand(fitz.Rect(links, kop4.y0, rechts, vinkjes[2].y1)), onder=vinkjes[2].y1 + 1)
    blok = rand(fitz.Rect(links, top.y0, rechts, antw_aangepast.y1), 8, 10)
    bewaar(p, "s6-top", blok, [fitz.Rect(links, top.y0 - 2, rechts, antw_top.y1)])
    bewaar(p, "s6-tip", blok, [fitz.Rect(links, tip.y0 - 2, rechts, antw_tip.y1)])
    bewaar(p, "s6-aangepast", blok, [fitz.Rect(links, aangepast.y0 - 2, rechts, antw_aangepast.y1)])
    vraag5 = tekst(p, "Welke twee dingen")
    bewaar(p, "s7-deel5", rand(fitz.Rect(links, kop5.y0, rechts, antw5.y1), 8, 10),
           [fitz.Rect(links, vraag5.y0 - 2, rechts, antw5.y1)])

    # ---- de x in het eerste vak van deel 4
    p = bladzijde_met(fitz.open(met_x), "Feedback van je buur")
    kop4, top = tekst(p, "Deel 4"), tekst(p, "Dit is goed")
    vinkjes = [v for v in vlakken(p, VAK) if kop4.y1 < v.y0 < top.y0]
    bewaar(p, "s6-kruis", rand(fitz.Rect(links, vinkjes[0].y0, rechts, vinkjes[2].y1), 8, 8), [vinkjes[0]],
           boven=vinkjes[0].y0 - 1, onder=vinkjes[2].y1 + 1)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])
