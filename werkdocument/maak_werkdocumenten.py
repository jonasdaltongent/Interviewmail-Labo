#!/usr/bin/env python3
"""
maak_werkdocumenten.py
Genereert het werkdocument voor les 3 "De formele e-mail" (Onderzoekslabo Jongeren & Welzijn, klas 3MWWE):

  Interviewmail.docx   het werkdocument dat de leerling INLEVERT (wordt in Classroom een Google-document)

Gebruik:  python3 maak_werkdocumenten.py     (vereist python-docx; huisstijl.py en go-dalton-gent.png
          staan ernaast, uit _sjablonen/lespakket/werkdocument/)

Waarom geen zip en geen Gmail: de leerlingen hebben met hun schoolaccount geen Gmail (Jonas, 09-10-2026 en
10-10-2026: "Werkdocument zoals 3MWb"), en per les gebruiken ze één werkplek naast Classroom en de lespagina
(_afspraken/didactiek.md). De e-mail komt daarom in dit document, in een venster dat eruitziet als "Nieuw
bericht" in Gmail. Model: Stagemail van MW les 3.

12-10-2026: in de nieuwe huisstijl, zoals Stagemail (Jonas: "Pas maar toe op alle lessen van deze week"):
geen kopregel, geen *Zo werk je*, geen uitleg die op de lespagina staat, geen knop Verzenden. De e-mails van
Senne en Noor staan op de lespagina (stap 2) en als citaten in de tabel van deel 1: niet nog eens volledig
hier. Deel 4 is de criteriatabel van lesvoorbereiding § 22, met de punten.

LET OP bij aanpassen:
 - Deel 1, rij 1 (het onderwerp) is het uitgewerkte voorbeeld: dat doet de leraar voor (dia 6). Laat het ingevuld.
 - De adressen eindigen op .example: dat domein bestaat nooit echt (RFC 2606). Geen echte organisatie.
 - Alle namen en gegevens zijn fictief (dat staat op de lespagina, niet in het werkdocument). Noors e-mail
   gaat bewust over een ander onderzoek en aan iemand anders dan de e-mail van de leerling: de leerling neemt
   de opbouw over, niet de zinnen.
 - De punten in deel 4 komen uit lesvoorbereiding.md § 22. Verander je ze, doe het op beide plaatsen.
"""
import os
from huisstijl import *

HERE = os.path.dirname(os.path.abspath(__file__))

JONGERENWERKER = ("mevrouw Hanne De Clercq", "hanne.declercq@delantaarn.example")
BEGELEIDER = ("mevrouw Ines Coppens", "ines.coppens@onderzoekslabo.example")


def interviewmail():
    doc = nieuw_document()
    kopblok(doc, "INTERVIEWMAIL", punten=20)

    kader(doc, "Je onderzoek", [
        "Je onderzoeksgroep in het Onderzoekslabo Jongeren & Welzijn onderzoekt schermgebruik en slaap bij "
        "jongeren (het onderzoeksplan uit les 2). Jullie willen een interview met een jongerenwerker.",
        ("Het interview:", f"{JONGERENWERKER[0]}, jongerenwerker bij Jongerenhuis De Lantaarn  —  "
                           f"{JONGERENWERKER[1]}"),
        ("Je onderzoeksbegeleider:", f"{BEGELEIDER[0]}  —  {BEGELEIDER[1]}"),
        ("Wat je vraagt:", "een interview van ongeveer 20 minuten  ·  twee mogelijke momenten: woensdag "
                           "21 oktober om 14 uur of donderdag 22 oktober om 16 uur  ·  in het jongerenhuis of online."),
        ("Wat er met haar antwoorden gebeurt:", "ze worden gebruikt voor een schoolonderzoek en haar naam komt "
                                                "niet in het verslag."),
        ("Tegen wanneer je een antwoord hoopt:", "maandag 19 oktober."),
    ])

    # ---- Deel 1: de citaten staan er al; rij 1 is ook verklaard (de demo)
    deel(doc, 1, "Goed of zwak?")
    tabel(doc, ["Onderdeel", "Senne", "Noor", "Waarom werkt het bij Noor beter?"], [
        ("1  onderwerp", "vraagje", "Vraag om een interview over groepschats – Onderzoekslabo",
         "Bij vraagje weet de ontvanger niet waarover het gaat. Bij Noor zie je waarover (een interview over "
         "groepschats) en van wie (Onderzoekslabo) nog voor je de e-mail opent."),
        ("2  aanspreking", "hallo,", "Beste meneer Willems,", ""),
        ("3  inleiding", "ik zit in het 3de jaar en we moeten een interview doen voor een taak over slaap",
         "Mijn naam is Noor Van Hecke en ik zit in 3MWWE … namens het Onderzoekslabo Jongeren & Welzijn …", ""),
        ("4  midden: de vraag", "kunt ge morgen na school?",
         "ongeveer 20 minuten … woensdag 28 oktober om 15 uur of donderdag 29 oktober om 13 uur, bij het CLB of "
         "online … Uw naam komt niet in ons verslag", ""),
        ("5  slot", "(geen slot)", "vóór dinsdag 20 oktober … Alvast bedankt voor uw tijd.", ""),
        ("6  slotgroet en handtekening", "thx!! 🙂  Senne",
         "Met vriendelijke groeten, Noor Van Hecke, 3MWWE, Onderzoekslabo Jongeren & Welzijn, GO! Dalton Gent", ""),
    ], [2.6, 3.0, 4.4, BREEDTE - 2.6 - 3.0 - 4.4], invul=(3,), rijhoogte=1.2, voorbeeld_rij=1,
        vet=(0,), klein=(0, 1, 2, 3))

    # ---- Deel 2
    deel(doc, 2, "Aan, Cc of Bcc?")
    tabel(doc, ["Situatie", "Aan, Cc of Bcc?"], [
        ("a.  Je mailt de jongerenwerker die je wil interviewen.", ""),
        ("b.  Je onderzoeksbegeleider wil je e-mail lezen, maar moet zelf niets doen.", ""),
        ("c.  Je nodigt 24 leerlingen van het 4de jaar uit voor je enquête. Ze mogen elkaars e-mailadres "
         "niet zien.", ""),
        ("d.  Je twee groepsgenoten moeten weten wat je vroeg.", ""),
    ], [12.6, BREEDTE - 12.6], invul=(1,), rijhoogte=1.0)
    vraag(doc, "Vraag 1. Waarom zet je de 24 leerlingen niet in Cc?")
    antwoordvak(doc, 1.8)

    # ---- Deel 3 (een eigen bladzijde: het venster breekt niet)
    deel(doc, 3, "Mijn interviewaanvraag", nieuwe_pagina=True)
    mailvenster(doc, [("aanspreking", 1.0),
                      ("inleiding:\nwie ben je, namens wie, waarom mail je?", 2.4),
                      ("midden:\nwat vraag je?", 3.4),
                      ("slot:\ntegen wanneer, bedanken", 1.8),
                      ("slotgroet", 1.0),
                      ("handtekening:\nnaam, klas, labo, school", 2.4)])

    # ---- Deel 4: de buur leest na = de criteria van de beoordeling (lesvoorbereiding § 22)
    deel(doc, 4, "Feedback van je buur", nieuwe_pagina=True)
    controletabel(doc, [
        ("Bij Aan staat de jongerenwerker, bij Cc de begeleider. Bcc is leeg.", 3),
        ("Het onderwerp zegt waarover de e-mail gaat en van wie hij komt.", 2),
        ("De aanspreking is *Beste mevrouw De Clercq,* (of *Geachte*), met een komma. De slotgroet is *Met "
         "vriendelijke groeten,* met daaronder naam, klas, Onderzoekslabo en school.", 3),
        ("In de inleiding staat wie ze is, namens wie ze mailt en waarom.", 3),
        ("In het midden staat wat ze vraagt: hoe lang, wanneer, waar, en wat er met de antwoorden gebeurt.", 4),
        ("In het slot staat tegen wanneer ze een antwoord hoopt, en ze bedankt.", 2),
        ("Overal staat *u* en *uw*, nergens *je* of *jij*. Geen emoji's of chattaal. Korte alinea's.", 2),
        ("Deel 1 en 2 zijn ingevuld.", 1),
    ])
    vraag(doc, "Dit is goed (een top van je buur):")
    antwoordvak(doc, 1.2)
    vraag(doc, "Dit kan beter (een tip van je buur):")
    antwoordvak(doc, 1.2)
    vraag(doc, "Wat heb ik aangepast?")
    antwoordvak(doc, 1.2)

    # ---- Deel 5: eerst het codewoord van de zelftest (15-10-2026, Jonas: "Codewoord idee is heel goed!"). De
    # lespagina toont het pas als de drie vragen van stap 7 juist zijn; zo zie je wie de zelftest echt maakte.
    deel(doc, 5, "Codewoord en slotvraag")
    vraag(doc, "Het codewoord van de zelftest (stap 7 op de lespagina):")
    antwoordvak(doc, 0.9)
    vraag(doc, "Welke twee dingen moet mevrouw De Clercq na het lezen van je e-mail zeker weten?")
    antwoordvak(doc, 1.8)

    doc.save(os.path.join(HERE, "Interviewmail.docx"))


if __name__ == "__main__":
    interviewmail()
    print("Gemaakt: Interviewmail.docx")
