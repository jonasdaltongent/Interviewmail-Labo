# Handleiding voor de leraar — Les 3: De formele e-mail

**Vak:** Toegepaste Informatica
**Doelgroep:** klas 3MWWE — 2de graad Maatschappij- en welzijnswetenschappen, doorstroomfinaliteit
**Lesduur:** 1 × 50 minuten: 10 minuten instructie + 40 minuten keuzewerktijd
**Context:** het Onderzoekslabo Jongeren & Welzijn (fictief) vraagt een interview aan bij een jongerenwerker van Jongerenhuis De Lantaarn (fictief)
**Lokaal:** 18 — Windows 11, Google Workspace in Chrome
**Lesdag:** donderdag 15 oktober 2026, 8ste lesuur (rooster van 25-09-2026)
**Kernleerplandoel:** `BV2_04.01` — digitaal communiceren (toepassen; summatief via het portfolio)
**Opdracht in Classroom:** *Les 3 — De formele e-mail*, deadline vrijdag 16 oktober 2026, 20.00 uur

<!-- Beginvragen (stap 1 van CONTEXT.md), beantwoord door Jonas op 10-10-2026:
     klas 3MWWE · les 3 van het lessenplan (module 1), do 15-10, 8ste uur · les 2 is gegeven (do 01-10,
     verder in de studie op 08-10; afgeleid uit _index/lessen.md en het logboek): lesstart haalt les 2 op ·
     lokaal 18, Windows 11 · werkdocument zoals 3MWb (geen Gmail: geen toegang met het schoolaccount), met
     een venster dat "Nieuw bericht" nabootst · AI-hulp (Gemini Notebook): aan · mascotte: de poses
     "laptop" (dia 7) en "letop" (dia 6) bestaan nog niet; hun plaats blijft klaar, en de hoofdsessie zet
     ze erin zodra Jonas zegt dat ze er zijn (verwacht in de nacht van 10 op 11-10). Erin gezet op 11-10-2026. -->

## 1. Inhoud van het pakket

```text
W07 - Les 03 - MWW - De formele e-mail/
├── index.html              de lespagina: startpagina in beelden · route · één stap · checklist
├── presentatie.html        10 dia's; dia 1 is de vaste startdia "Zo start je"
├── css/style.css           lespagina (uit het sjabloon, met .nodig, .vast, .mail, .ccb, .bouwplan, de mascotte, Stap voor stap / Kort)
├── css/slides.css          dia's (uit het sjabloon, met .startroute, .met-mascotte en .kloklijst)
├── js/script.js, js/slides.js
├── assets/                 logo's van het Onderzoekslabo en GO! Dalton Gent, lettertypes (OFL), mascotte Dalton, beelden/ (Nano Banana),
│                           stapjes/ (27 uitsneden van het werkdocument)
├── werkdocument/
│   ├── Interviewmail.docx          het werkdocument dat de leerling INLEVERT
│   ├── maak_werkdocumenten.py      maakt het opnieuw (python-docx)
│   ├── maak_stapjes.py             maakt de beelden bij de stapjes opnieuw, uit de pdf van Google Documenten
│   └── huisstijl.py, go-dalton-gent.png   de huisstijl van het werkdocument (uit _sjablonen/lespakket/werkdocument/)
├── lesvoorbereiding.md     de lesvoorbereiding (28 punten)
├── dalton-lesfiche.html    de Dalton-lesfiche in de kleurcode: openen, Kopieer de fiche, plakken
├── lesdoelen.json          de doelen voor je jaaroverzicht
├── classroom.json          de opdracht voor de Classroom-koppeling
├── ai-bron.txt             de volledige tekst van de lespagina: de bron van de AI-hulp
└── README.md               deze handleiding
```

**Geen zip en geen Gmail.** De leerlingen hebben geen Gmail met hun schoolaccount ("Werkdocument zoals
3MWb"), en alles gebeurt in één werkplek naast Classroom en de lespagina (`_afspraken/didactiek.md`). De
e-mail komt in het werkdocument *Interviewmail*, in een venster dat eruitziet als *Nieuw bericht*; er wordt
niets verstuurd.

**Adressen.** In het werkdocument en op de lespagina staan verzonnen adressen die eindigen op `.example`, een
domein dat nooit echt bestaat (RFC 2606). Zo staat er geen echte organisatie of persoon in het pakket.

**De twee voorbeeldmails.** *Senne* (zwak: onderwerp *vraagje*, *hallo,*, *kunt ge morgen na school?*, *thx!! 🙂*,
begeleider in Aan) en *Noor* (goed: aan een CLB-medewerker, over groepsdruk in groepschats, met andere momenten
dan de leerling). Noors e-mail gaat bewust over een ander onderzoek aan een andere persoon: de leerling neemt de
opbouw over, niet de zinnen. Ze staan op de lespagina (stap 2) en in deel 1 van het werkdocument.

### 1b. Hoe de lespagina werkt

- **Startpagina**: *Wat heb je nodig?* in drie grote beelden in twee kolommen (de opdracht, de lespagina, het
  werkdocument), *Na deze les kan je*, en *Vast?* als genummerde rij: *Toon stapjes*, je buur, de AI-hulp, je
  leraar. Bijgewerkt op 14-10-2026 volgens `_afspraken/lespagina.md` (*Groot genoeg*, de kortere hulpvolgorde).
- **Stap voor stap / Kort** (14-10-2026, Jonas: "Ik merk net dat de les van 3MWWE nog niet voldoet aan stap voor
  stap en kort"; zoals MW les 3): elke handeling is één vette regel; open je ze, dan zie je de kleine stapjes,
  elk met één beeld (`assets/stapjes/`, of een tekening van de opdracht of de tabbladen). Bovenaan *Wat moet je
  doen?* zet de knop *Stap voor stap* alles open en *Kort* alles dicht; een leerling kan ook één regel openen. De
  les begint in *Stap voor stap* (`data-modus-start="stap"`); de keuze wordt per les onthouden. *Hulp nodig?* is
  weg: wat erin stond, staat nu als grijze regel onder het stapje waar het misloopt (de vragen bij rij 2 tot 6
  in stap 2, de zinsbeginnen in stap 5), en het tipvak bevat de meest waarschijnlijke fout.
- **Stap 7** toont onder de handelingen het filmpje **Inleveren** (`assets/video/inleveren.mp4`, 10 seconden, 15-10-2026,
  Jonas: "tot ze het systeem snappen"), ook in *Kort* in beeld; het vervangt de tekening *Hier lever je in*. Het stapje
  *Klik in het venster Je werk inleveren? nog eens op Inleveren* kwam erbij (het venster is nagekeken met *demoleerling3*).
  Bij 3MWb was *Waar moet ik inleveren?* de vraag die het vaakst kwam (12-10-2026).
- **Dia 7** is de controle (*Waar zet je je begeleider?*, iedereen antwoordt met vingers), **dia 8** *Iedereen doet
  stap 1* (14-10-2026, `_afspraken/presentatie.md`).
- **Dia 9** toont kloktijden uit de planning (`_tools/zet_lesmomenten.py`), met stap 1 als *samen ✓*: buiten de les
  die van do 15-10, 8ste uur (tot 15.24); tijdens de les rekent ze vanaf nu tot de bel. **T** rekent opnieuw.
- **Route** links, **één stap** in het midden, **checklist** rechts (20 taken). Op een smal scherm staat alles
  onder elkaar.
- De **titel van de pagina** begint met *Lespagina · Les 3*: zo heet de link in de opdracht en het tabblad in
  Chrome, naast het tabblad *Interviewmail*.
- **AI-hulp**: aan (`data-ai-hulp="aan"` op `<body>`), want de Gemini Notebook hangt aan de opdracht
  (`_afspraken/ai-hulp.md`).
- **Mascotte Dalton**: op de startpagina, bij *Hulp nodig?*, in de laatste stap en in de eindmelding; op de
  dia's 2, 3, 6 (*letop*) en 9 (*laptop*); niet op dia 7 en 8. De poses van dia 6 en 7 staan erin sinds 11-10-2026 (`1320dcf`).
- `?leraar` achter het adres toont de screenshot-plaatsen (deze les heeft er geen: de beelden zijn tekeningen).

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages ✅ *gebeurd op 10-10-2026*

Repository [`jonasdaltongent/Interviewmail-Labo`](https://github.com/jonasdaltongent/Interviewmail-Labo), Pages op branch `main`, map `/ (root)`. De lespagina:
<https://jonasdaltongent.github.io/Interviewmail-Labo/>, de dia's: <https://jonasdaltongent.github.io/Interviewmail-Labo/presentatie.html>. Dat adres staat ook op dia 7, in
`classroom.json` en in `lesdoelen.json`.

### Stap 2 — De opdracht in Classroom, met de koppeling ✅ *concept op 10-10-2026, ingepland op 12-10-2026*

Als **concept** in [3MWWE - Informatica](https://classroom.google.com/c/MjUzNjQyOTk0NDFa), **ingepland**: Classroom
publiceert de opdracht zelf op donderdag 15 oktober om 14.05, een half uur vóór het 8ste uur
(`_afspraken/publiceren.md`). Op **Toewijzen** klikken hoeft niet meer. Teruggelezen via de API op 14-10-2026:
`DRAFT`, ingepland, drie bijlagen; het Google-document *Interviewmail* is de versie van 12-10-2026 (zelfde
inhoud als `werkdocument/Interviewmail.docx`).

| | Opdracht: **Les 3 — De formele e-mail** |
|---|---|
| **Onderwerp** | Module 1 — Digitaal organiseren en communiceren |
| **Bijlage 1** | de link naar de lespagina |
| **Bijlage 2** | `Interviewmail` (Google-document) — **Een kopie maken voor elke leerling** |
| **Bijlage 3** | de Gemini Notebook *AI-hulp · Les 3 · De formele e-mail* (*Bovenaan de lesgroep plaatsen* uit) |
| **Punten** | 20 |
| **Deadline** | vrijdag 16 oktober 2026, 20.00 uur |

```bash
python3 "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/_tools/zet_opdracht_klaar.py" "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/2026-2027/W07 - Les 03 - MWW - De formele e-mail" --proef
```

Instructietekst (het script vult de deadline in; er is geen e-mailadres nodig):

```text
Open de lespagina en werk stap 1 tot 7 af.
Je schrijft alles in je werkdocument Interviewmail. Je hebt niets anders nodig.

Klik op Inleveren, ten laatste vrijdag 16 oktober 2026 om 20.00 uur.
```

### Stap 3 — Met de hand, als de koppeling niet werkt

1. Upload `werkdocument/Interviewmail.docx` **in Drive zelf** (**Nieuw** › **Bestand uploaden**, met *Uploads
   converteren* aan).
2. Voeg het in de opdracht toe met **Bijvoegen** › **Drive**, met **Een kopie maken voor elke leerling**
   (alleen vóór je de opdracht post).
3. Voeg de lespagina toe met **Link**.

### De AI-hulp

✅ **Aan het concept gehangen (10-10-2026).** Gemini Notebook *AI-hulp · Les 3 · De formele e-mail* in jouw profiel
(dalton.gent), met als bron `https://jonasdaltongent.github.io/Interviewmail-Labo/ai-bron.txt` (de versie met de
herstellingen van 10-10-2026), **Chat instellen** › **Aangepast** met de vaste instructie uit `_afspraken/ai-hulp.md`
› **Korter** (teruggelezen). In het concept bijgevoegd met **Bestaande selecteren**, *Bovenaan de lesgroep
plaatsen* uitgezet; teruggelezen via de API: drie bijlagen, nog altijd `DRAFT`.

**Nieuwe bron op 14-10-2026** (Jonas: "Ook de ai hulp"): `…/ai-bron.txt?v=2026-10-14`, de tekst van commit
`a863096`, teruggelezen in de notebook (geen *naam, klas en datum* meer, *of haar naam in je verslag komt*, stap 7
met het tabblad Google Classroom). **De oude bron** (`…/ai-bron.txt` zonder `?v=`, de tekst van 10-10-2026, met
*of het anoniem blijft*) **staat er nog**: verwijderen is definitief, dus dat doe je zelf: in de notebook bij
**Sources** naast die bron **⋮** › **Bron verwijderen** › **Verwijderen**. Tot dan antwoordt de AI-hulp uit
allebei.

Een eerste notebook met dezelfde naam staat per ongeluk in het testaccount *demoleerling3* (de lesbouwer werkte
eerst in het verkeerde Chrome-profiel). Die hangt nergens aan; wissen kan je zelf als je wil.

### De demo klaarzetten (1 minuut)

Op dia 5 toon je waar de leerling in *Interviewmail* typt. Gebruik daarvoor je eigen kopie: upload
`werkdocument/Interviewmail.docx` in Drive (met *Uploads converteren* aan) en laat ze met niemand gedeeld. Open
ze en de lespagina in twee tabbladen.

### Afvinklijst vóór de les

- [x] De lespagina staat online en het adres op dia 7 klopt (nagekeken op 14-10-2026).
- [x] De opdracht staat klaar in `3MWWE - Informatica`, met de lespagina, de AI-hulp en *Interviewmail* als
      **kopie per leerling**, en is ingepland voor do 15-10 om 14.05 (nagekeken op 14-10-2026).
- [ ] Het wachtwoord van de computers staat op het bord (niet op de dia's: die zijn openbaar).
- [ ] Je eigen kopie van *Interviewmail* en de lespagina staan open in twee tabbladen voor de demo.
- [ ] De Dalton-lesfiche staat in je planner (open `dalton-lesfiche.html`, klik op **Kopieer de fiche**, plak).
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities, `F` is volledig scherm.

### 2b. Nagelezen klikpaden

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Naar de opdracht (dia 1) | classroom.google.com › **Inloggen** › de lesgroep › **Schoolwerk** › de opdracht › **Instructies bekijken** | [Classroom 6020285](https://support.google.com/edu/classroom/answer/6020285?hl=nl&co=GENIE.Platform%3DDesktop) (09-10-2026) |
| Inleveren | **Inleveren** | idem |
| Aan, Cc, Bcc (uitleg) | de velden **Aan**, **Cc**, **Bcc**; Cc: "geen actie hoeven te ondernemen"; Bcc: adres verborgen | [Gmail 2819488](https://support.google.com/mail/answer/2819488?hl=nl&co=GENIE.Platform%3DDesktop) (09-10-2026) |

Er staan geen knopnamen van Gmail of Documenten in de stappen: de leerlingen typen alleen in een tabel van hun
werkdocument. De e-mailregels komen uit `_vakinhoud/digitale-communicatie.md` (Team Taaladvies).

## 3. Het verloop van de les

| Fase | Tijd | Dia | Wat |
|---|---|---|---|
| Start | 0–2' | 1 | aanmelden met de vaste startdia, tot de lespagina open is |
| Lesstart | 2–4' | 2 | waarom heette het bestand `interview_respondent-02`? (vingers, dan →) |
| Ik doe | 4–9'30" | 3–6 | lesdoel · het eindproduct in acht delen · Aan, Cc of Bcc · rij 1 van deel 1 (het onderwerp) |
| Controle | 1' | 7 | *Waar zet je je begeleider?* — iedereen antwoordt met vingers; wie fout heeft, komt bij jou |
| Samen | 3' | 8 | iedereen doet stap 1 terwijl jij de schermen ziet; duim op als ze klaar zijn |
| Keuzewerktijd | tot het einde | 9 | stap 2–7, *Zo werk je verder* blijft staan, met **kloktijden** (T rekent opnieuw) |
| Afsluiten | laatste minuut | 10 | de exitvraag (= deel 5); wie niet klaar is, levert toch in |

**Eerste rondgang** (stap 2): is iedereen aan rij 2 van deel 1 begonnen? Wie nog geen twee tabbladen open heeft
(de lespagina en *Interviewmail*), eerst dat. **Tweede rondgang** (stap 4–5): Aan en Cc juist overgetypt, Bcc leeg, *u* in plaats van
*je*? Vraag: *"Lees je midden luidop: weet mevrouw De Clercq nu wanneer en hoe lang?"*

## 4. Verbetersleutel

- **Deel 1**, rij 2 tot 6 (de citaten staan er al; de leerling verklaart in de laatste kolom): aanspreking: *hallo* is te los en zonder naam; *Beste meneer Willems,* is beleefd en
  persoonlijk · inleiding: Senne zegt niet wie hij is en namens wie; Noor wel · midden: Senne vraagt iets
  onhaalbaars (*morgen*) zonder duur of plaats; Noor zegt hoe lang, wanneer, waar en wat er met de antwoorden
  gebeurt · slot: Senne geeft geen termijn en bedankt niet echt; Noor vraagt een antwoord vóór een datum en
  bedankt vooraf · slotgroet: *thx* en een emoji, alleen een voornaam; Noor: *Met vriendelijke groeten,* en een
  volledige handtekening. (Rij 1, het onderwerp, is al verklaard.)
- **Deel 2**: a **Aan** · b **Cc** · c **Bcc** · d **Cc**. Vraag 1: een e-mailadres is een persoonsgegeven; in Cc
  ziet iedereen de adressen van de anderen; met Bcc niet (zoals `respondent-02` in les 2).
- **Deel 3**, de e-mail: zie de criteria in `lesvoorbereiding.md` §22 (20 punten).
- **Deel 5**: wat je vraagt (een interview van 20 minuten op een van de twee momenten) en tegen wanneer je een
  antwoord nodig hebt (vóór maandag 19 oktober).

### Essentiële fouten — geef hier altijd feedback op

- Geen onderwerp.
- Geen concrete vraag (wat en wanneer).
- De begeleider in *Aan* of de jongerenwerker in *Cc*.
- *Hallo* of chattaal.
- Geen naam onder de e-mail.
- In deel 2 de respondenten in *Cc*.

## 5. Schermafbeeldingen (optioneel)

De beelden bij de stapjes (`assets/stapjes/`, 27 stuks) zijn uitsneden van het werkdocument zoals Google
Documenten het toont, op ware grootte, met één rood kader (#C5221F, witte rand erbuiten), zoals in MW les 3.
`werkdocument/maak_stapjes.py` maakt ze uit de pdf die Google Documenten van *Interviewmail* maakt, en uit een
tweede pdf met een x in deel 4 (*Test Interviewmail (Claude)*, in de demomap van je Drive); uitleg bovenaan het
script. Verandert het werkdocument, draai het dan opnieuw. De opdracht in Classroom en de tabbladen van Chrome
zijn tekeningen in HTML en CSS (`.mock-opdracht`, `.teken-tabs`, het kader is `.wijs`), net als de beelden op de
startpagina, in stap 2 en op de dia's.

## 6. Het materiaal opnieuw maken

```bash
cd "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/2026-2027/W07 - Les 03 - MWW - De formele e-mail/werkdocument" && python3 maak_werkdocumenten.py
```

## 7. Leerplandoelen in je jaaroverzicht

`lesdoelen.json` (klasgroep `3 MWW`): `BV2_04.01` kern · `BV2_02.08`, `BV2_02.08.02`, `BV2_02.13` ondersteunend ·
`BV2_04.04` voorbereidend. De pre-push hook schrijft ze in `Leerplandoelen 2026-2027.xlsx`.

## 8. Wat nog moet blijken in de klas

- **Het filmpje Inleveren** (15-10-2026, in de laatste stap, in elke les tot de leerlingen het systeem kennen): leveren ze
  in zonder te vragen? Kijken ze naar het filmpje, en klikken ze in het venster *Je werk inleveren?* nog eens op *Inleveren*?
- **Stap voor stap / Kort** (14-10-2026, voor het eerst in 3MWWE): kiezen leerlingen zelf *Kort*, of laten ze alles
  open? Klappen ze in *Kort* een regel open als ze vastzitten? Wordt stap 2 te lang, met de e-mails van Senne en
  Noor onder vijf stapjes?
- Komen ze vlotter op gang na de controledia en *Iedereen doet stap 1*?

- **Het nieuwe werkdocument** (12-10-2026): geen naam/klas/datum en *Zo werk je* meer, de volledige e-mails van
  Senne en Noor alleen nog op de lespagina, antwoorden in lichtblauwe vakken, deel 4 met de punten van § 22.
  Vergelijken ze vlot met de citaten in de tabel? Zet de buur in deel 4 een x in het vak?
- Helpen de vaste startdia en de startpagina in beelden om zonder hulp tot bij de opdracht, de lespagina en het
  werkdocument te geraken, ook nu de klas het werkdocument voor het eerst zo gebruikt?
- Vinden de leerlingen **Inleveren** zonder te vragen, met de tekening in stap 7?
- Lukt het verklaren in deel 1 in 8 minuten (vijf rijen, alleen de laatste kolom)? De
  minimumroute slaat deel 1 over.
- Neemt de leerling de opbouw van Noors e-mail over zonder haar zinnen te kopiëren?
- Typen de leerlingen netjes in de vakken van het venster, of schuiven de tabellen in Google Documenten? Op
  14-10-2026 als pdf uit Google Documenten nagekeken: vier bladzijden, de tabellen en het venster staan goed.
- Klopt de tijd van stap 5 (10 minuten) voor inleiding, midden en slot?
- Geven de leerlingen in stap 6 bruikbare feedback met de checklist (top en tip), of vinken ze alles aan?
- Kunnen de leerlingen van 3MWWE de AI-hulp openen, helpt hij, en klikken ze op de bronnummers?
- **De eerste beelden van Nano Banana** (14-10-2026, Jonas: "Laad maar op, we zullen het morgen eens testen!"):
  mevrouw De Clercq als casusbeeld in stap 4 en op dia 8 (in de plaats van de mascotte), met *Gemaakt met AI*, en
  Jongerenhuis De Lantaarn klein op dia 4. Maakt het de ontvanger concreter (formeel schrijven aan iemand die je niet
  kent)? Reageren de leerlingen op het label *Gemaakt met AI*? De bron van de AI-hulp is niet aangepast (alleen het
  label kwam erbij).
- Komt de leerling vanzelf op *u* en *uw*? Wat doen ze met de datum voor het antwoord (maandag 19 oktober)?
