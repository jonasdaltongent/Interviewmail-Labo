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
     ze erin zodra Jonas zegt dat ze er zijn (verwacht in de nacht van 10 op 11-10). -->

## 1. Inhoud van het pakket

```text
W07 - Les 03 - MWW - De formele e-mail/
├── index.html              de lespagina: startpagina in beelden · route · één stap · checklist
├── presentatie.html        8 dia's; dia 1 is de vaste startdia "Zo start je"
├── css/style.css           lespagina (uit het sjabloon, met .nodig, .vast, .mail, .ccb, .bouwplan, de mascotte)
├── css/slides.css          dia's (uit het sjabloon, met .startroute en .met-mascotte)
├── js/script.js, js/slides.js
├── assets/                 logo's van het Onderzoekslabo en GO! Dalton Gent, lettertypes (OFL), mascotte Dalton
├── werkdocument/
│   ├── Interviewmail.docx          het werkdocument dat de leerling INLEVERT
│   └── maak_werkdocumenten.py      maakt het opnieuw (python-docx)
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

- **Startpagina**: *Wat heb je nodig?* in drie beelden (de opdracht, de lespagina, het werkdocument), *Na deze
  les kan je*, en *Vast?* als genummerde rij.
- **Route** links, **één stap** in het midden, **checklist** rechts (22 taken). Op een smal scherm staat alles
  onder elkaar.
- De **titel van de pagina** begint met *Lespagina · Les 3*: zo heet de link in de opdracht en het tabblad in
  Chrome, naast het tabblad *Interviewmail*.
- **AI-hulp**: aan (`data-ai-hulp="aan"` op `<body>`), want de Gemini Notebook hangt aan de opdracht
  (`_afspraken/ai-hulp.md`).
- **Mascotte Dalton**: op de startpagina, bij *Hulp nodig?*, in de laatste stap en in de eindmelding; op de
  dia's 2, 3 en 8. Zie *Mascotte: wat nog volgt* hieronder.
- `?leraar` achter het adres toont de screenshot-plaatsen (deze les heeft er geen: de beelden zijn tekeningen).

### Mascotte: wat nog volgt

De poses `dalton_letop.webp` (dia 6, *Fout 1 doen we samen*) en `dalton_laptop.webp` (dia 7, *Zo werk je
verder*) bestaan nog niet. Dia 6 en 7 hebben al `class="slide met-mascotte"` (de tekst blijft links, 300 px
rechts is vrij; gemeten op 1280 × 720: niets valt buiten het beeld). In `presentatie.html` staat op de plaats van
de afbeelding een commentaar `<!-- MASCOTTE VOLGT: <img … src = "assets/mascotte/dalton_….webp" …> -->`. Zodra de
`.webp` in `assets/mascotte/` staat: haal de commentaartekens weg en verwijder de spaties rond `=` in `src`.
(Die spaties staan er zodat `controleer_les.py` het commentaar niet als ontbrekend bestand telt.)

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages ✅ *gebeurd op 10-10-2026*

Repository [`jonasdaltongent/Interviewmail-Labo`](https://github.com/jonasdaltongent/Interviewmail-Labo), Pages op branch `main`, map `/ (root)`. De lespagina:
<https://jonasdaltongent.github.io/Interviewmail-Labo/>, de dia's: <https://jonasdaltongent.github.io/Interviewmail-Labo/presentatie.html>. Dat adres staat ook op dia 7, in
`classroom.json` en in `lesdoelen.json`.

### Stap 2 — De opdracht in Classroom, met de koppeling ✅ *concept op 10-10-2026*

Als **concept** in [3MWWE - Informatica](KLASLINK). Toewijzen doe je zelf met **Toewijzen**.

| | Opdracht: **Les 3 — De formele e-mail** |
|---|---|
| **Onderwerp** | Module 1 — Digitaal organiseren en communiceren |
| **Bijlage 1** | de link naar de lespagina |
| **Bijlage 2** | `Interviewmail` (Google-document) — **Een kopie maken voor elke leerling** |
| **Bijlage 3** | de Gemini Notebook *AI-hulp · Les 3 · De formele e-mail* (zie hieronder) |
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

AIHULPSTATUS

### De demo klaarzetten (1 minuut)

Op dia 5 toon je waar de leerling in *Interviewmail* typt. Gebruik daarvoor je eigen kopie: upload
`werkdocument/Interviewmail.docx` in Drive (met *Uploads converteren* aan) en laat ze met niemand gedeeld. Open
ze en de lespagina in twee tabbladen.

### Afvinklijst vóór de les

- [ ] De lespagina staat online en het adres op dia 7 klopt.
- [ ] De opdracht staat klaar in `3MWWE - Informatica`, met de lespagina, de AI-hulp en *Interviewmail* als
      **kopie per leerling**, en is toegewezen.
- [ ] Het wachtwoord van de computers staat op het bord (niet op de dia's: die zijn openbaar).
- [ ] Je eigen kopie van *Interviewmail* en de lespagina staan open in twee tabbladen voor de demo.
- [ ] De Dalton-lesfiche staat in je planner (open `dalton-lesfiche.html`, klik op **Kopieer de fiche**, plak).
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities, `F` is volledig scherm.
- [ ] Zijn `dalton_letop.webp` en `dalton_laptop.webp` er al? Zet ze dan op dia 6 en 7 (zie hierboven).

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
| Keuzewerktijd | 9'30"–50' | 7 | stap 1–7, *Zo werk je verder* blijft staan |
| Afsluiten | laatste minuut | 8 | de exitvraag (= deel 5); wie niet klaar is, levert toch in |

**Eerste rondgang** (stap 1–2): heeft iedereen de lespagina én *Interviewmail* open, in twee tabbladen, en is rij
2 van deel 1 begonnen? **Tweede rondgang** (stap 4–5): Aan en Cc juist overgetypt, Bcc leeg, *u* in plaats van
*je*? Vraag: *"Lees je midden luidop: weet mevrouw De Clercq nu wanneer en hoe lang?"*

## 4. Verbetersleutel

- **Deel 1**, rij 2 tot 6: aanspreking: *hallo* is te los en zonder naam; *Beste meneer Willems,* is beleefd en
  persoonlijk · inleiding: Senne zegt niet wie hij is en namens wie; Noor wel · midden: Senne vraagt iets
  onhaalbaars (*morgen*) zonder duur of plaats; Noor zegt hoe lang, wanneer, waar en wat er met de antwoorden
  gebeurt · slot: Senne geeft geen termijn en bedankt niet echt; Noor vraagt een antwoord vóór een datum en
  bedankt vooraf · slotgroet: *thx* en een emoji, alleen een voornaam; Noor: *Met vriendelijke groeten,* en een
  volledige handtekening. (Rij 1, het onderwerp, staat ingevuld.)
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

Niet nodig: de beelden op de startpagina, in stap 2 en op de dia's zijn tekeningen in HTML en CSS.

## 6. Het materiaal opnieuw maken

```bash
cd "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/2026-2027/W07 - Les 03 - MWW - De formele e-mail/werkdocument" && python3 maak_werkdocumenten.py
```

## 7. Leerplandoelen in je jaaroverzicht

`lesdoelen.json` (klasgroep `3 MWW`): `BV2_04.01` kern · `BV2_02.08`, `BV2_02.08.02`, `BV2_02.13` ondersteunend ·
`BV2_04.04` voorbereidend. De pre-push hook schrijft ze in `Leerplandoelen 2026-2027.xlsx`.

## 8. Wat nog moet blijken in de klas

- Helpen de vaste startdia en de startpagina in beelden om zonder hulp tot bij de opdracht, de lespagina en het
  werkdocument te geraken, ook nu de klas het werkdocument voor het eerst zo gebruikt?
- Lukt het vergelijken in deel 1 in 8 minuten (vijf rijen met vier vakken), of is dat te veel typwerk? De
  minimumroute slaat deel 1 over.
- Neemt de leerling de opbouw van Noors e-mail over zonder haar zinnen te kopiëren?
- Typen de leerlingen netjes in de vakken van het venster, of schuiven de tabellen in Google Documenten (het
  midden-vak is hoog gemaakt)? Dit is nog niet in Google Documenten bekeken.
- Klopt de tijd van stap 5 (10 minuten) voor inleiding, midden en slot?
- Geven de leerlingen in stap 6 bruikbare feedback met de checklist (top en tip), of vinken ze alles aan?
- Kunnen de leerlingen van 3MWWE de AI-hulp openen, helpt hij, en klikken ze op de bronnummers?
- Komt de leerling vanzelf op *u* en *uw*? Wat doen ze met de datum voor het antwoord (maandag 19 oktober)?
- De dia's 6 en 7 passen met de mascotte erbij? (Gemeten zonder: de tekst blijft links; kijk na zodra de
  poses er zijn.)
