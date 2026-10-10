# Handleiding voor de leraar — Les 08: Delen met de juiste rechten

**Vak:** Toegepaste Informatica
**Doelgroep:** 3ORLO, 4ORLOa en 4ORLOb — 2de graad Organisatie en logistiek, arbeidsmarktgerichte finaliteit
**Lesduur:** 1 × 50 minuten: 10 minuten instructie + 40 minuten keuzewerktijd
**Context:** NovaDepot, fictief logistiek bedrijf en groothandel — module 2 van het jaarplan
(`Leerplannen/Lessenplan_TOINFO_ORLO_arbeidsmarkt_2026-2027.md`)
**Toestel:** **lokaal 18, Windows 11** — Google Workspace in Chrome
**Lesdag:** uur 2 van week 7: 3ORLO wo 14 oktober (2de uur) · 4ORLOb wo 14 oktober (3de uur) · 4ORLOa do 15 oktober (7de uur)
**Kernleerplandoel:** `BV2_04.03` (digitale inhouden beheren: delen en beschermen) — toepassen
**Deadline:** vrijdag 16 oktober 2026, 20.00 uur · 20 punten

<!-- Beginvragen (stap 1 van CONTEXT.md), beantwoord door Jonas op 09-10-2026:
     klassen 3ORLO, 4ORLOa, 4ORLOb · les 08 van W07 · toestel lokaal 18, Windows 11 ·
     lesstart = ophalen uit de vorige les · delen tussen leerlingaccounts mag · werkdocument: ja. -->

De leerlingen testen in duo wat een **kijker**, een **reageerder** en een **bewerker** kan: elk deelt de eigen
**Leveringsplanning** met de buur. Als bewerker wist de buur de regel van 10.30 uur; de leerling zet de planning terug
met de **versiegeschiedenis**. Daarna halen ze de toegang van de buur weg en delen ze hun map **NovaDepot** met jou,
als kijker.

---

## 1. Inhoud van het pakket

```text
W07 - Les 08 - ORLO - Delen met de juiste rechten/
├── index.html                   # de lespagina: route · één stap · checklist
├── presentatie.html             # 8 dia's voor de lesstart en "Ik doe"
├── css/style.css, css/slides.css
├── js/script.js, js/slides.js
├── assets/
│   ├── novadepot-logo.svg/.png, novadepot-icon.svg, dalton-gent-logo.png
│   ├── fonts/                   # Atkinson Hyperlegible + Montserrat (OFL, zelf gehost)
│   ├── mascotte/                # Dalton de bever, de poses uit het sjabloon (.webp)
│   └── screenshots/             # zes schermafbeeldingen (09-10-2026) + LEESMIJ.md
├── werkdocument/
│   ├── WP3_Delen-met-de-juiste-rechten.docx   # het werkdocument (kopie per leerling)
│   ├── Leveringsplanning.docx                 # het oefendocument van de duo-test (kopie per leerling)
│   └── maak_werkdocumenten.py                 # maakt beide opnieuw (python-docx)
├── lesvoorbereiding.md
├── dalton-lesfiche.html         # openen, Kopieer de fiche, plakken in je planner
├── classroom.json               # de opdracht voor _tools/zet_opdracht_klaar.py
├── lesdoelen.json               # leerplandoelen voor je jaaroverzicht
└── README.md                    # deze handleiding
```

### 1b. Hoe de lespagina werkt

Dezelfde opbouw als les 06 en 07: links de **route** (zes stappen in twee groepen: *Testen met je buur* · *Afronden*),
midden **één stap** met vijf vaste blokken, rechts de **checklist** met 19 taken.

- Stap 2, 3 en 4 volgen elkaar op met **dezelfde buur en hetzelfde document**: eerst kijker, dan reageerder, dan
  bewerker. Na elke test vullen ze één kolom van deel 1 in.
- Elke nieuwe handeling heeft een **schermafbeelding met nummers** (delen, algemene toegang, reactie, versie terugzetten,
  map delen).
- De **zelftest** in stap 6 heeft drie vragen: de kleinste rol, een versie terugzetten, *Beperkt*.
- Windows 11 (lokaal 18). `localStorage` bewaart alleen de vinkjes en de laatste stap (voorvoegsel
  `novadepot_delen_v1_`), met een wisknop.

---

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages ✅ *gebeurd op 09-10-2026*

Repository [`jonasdaltongent/Delen-NovaDepot`](https://github.com/jonasdaltongent/Delen-NovaDepot),
Pages op branch `main`, map `/ (root)`. De lespagina staat op
[jonasdaltongent.github.io/Delen-NovaDepot](https://jonasdaltongent.github.io/Delen-NovaDepot/), de dia's op
[…/presentatie.html](https://jonasdaltongent.github.io/Delen-NovaDepot/presentatie.html); dat adres staat ook
op dia 7, in `classroom.json` en in `lesdoelen.json`. Live bestanden nagekeken: gelijk aan de lokale.

### Stap 2 — De opdracht in Classroom, met de koppeling ✅ *concept op 09-10-2026*

Als **concept** klaargezet in [3ORLO](https://classroom.google.com/c/MTYyNjY1NjU1OTEx), [4ORLOa](https://classroom.google.com/c/MjUzMTgzNDM2NjZa) en [4ORLOb](https://classroom.google.com/c/MjUzMTc2OTkwNzNa) (teruggelezen: onderwerp, deadline, 20 punten, werkdocument en
Leveringsplanning `STUDENT_COPY`). Toewijzen doe je zelf met **Toewijzen**.

| | Opdracht: **Werkplek 3 — Delen met de juiste rechten** |
|---|---|
| **Onderwerp** | Digitale competenties (bestaat al) |
| **Bijlage 1** | de link naar de lespagina |
| **Bijlage 2** | `WP3_Delen-met-de-juiste-rechten` (Google-document) — **Een kopie maken voor elke leerling** |
| **Bijlage 3** | `Leveringsplanning` (Google-document) — **Een kopie maken voor elke leerling** |
| **Punten** | 20 |
| **Deadline** | vrijdag 16 oktober 2026, 20.00 uur |

Proef zonder Google (09-10-2026: in orde):

```bash
python3 "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/_tools/zet_opdracht_klaar.py" "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/2026-2027/W07 - Les 08 - ORLO - Delen met de juiste rechten" --proef
```

Daarna `--maak` (concept); de AI publiceert niet in Classroom. Instructietekst (het script vult je e-mailadres en de
deadline in; in `classroom.json` staat alleen `{email}`):

```text
1. Open de lespagina (link), je werkdocument WP3_Delen-met-de-juiste-rechten en je Leveringsplanning.
2. Je werkt samen met je buur: elk deelt de eigen Leveringsplanning.
3. Het e-mailadres van je leraar, voor stap 5: <je e-mailadres>
4. Klik op Inleveren, ten laatste vrijdag 16 oktober 2026 om 20.00 uur. Je werkdocument en je planning gaan samen mee.
```

### Stap 3 — Met de hand, als de koppeling niet werkt

1. Upload beide `.docx`-bestanden **in Drive zelf** (**Nieuw** › **Bestand uploaden**, met *Uploads converteren* aan).
2. Voeg ze in de opdracht toe met **Bijvoegen** › **Drive**, allebei met **Een kopie maken voor elke leerling** (alleen
   vóór je de opdracht post).
3. Voeg de lespagina toe met **Link** en zet je e-mailadres in de instructie.

### De demo klaarzetten (1 minuut)

Je doet op dia 5 en 6 voor in een eigen Leveringsplanning. In je map *Demo voor schermafbeeldingen (Claude)* staat er
al een (een Google-document, met vier regels, met niemand gedeeld). Op dia 5 klik je op **Annuleren** in plaats van
**Verzenden** als je met niemand wil delen. Op dia 6 wis je de regel van 10.30 uur live en zet je ze terug.

### Afvinklijst vóór de les

- [ ] De lespagina staat online en het adres op dia 7 klopt.
- [ ] De opdracht staat in de drie klassen, met drie bijlagen, allebei de documenten als **kopie per leerling**, en je
      e-mailadres in de instructie.
- [ ] Met een **leerlingaccount** getest: de kopie van de planning delen met een ander leerlingaccount lukt (zie §8).
- [ ] Je Leveringsplanning voor de demo staat klaar.
- [ ] De Dalton-lesfiche staat in je planner (open `dalton-lesfiche.html`, klik op **Kopieer de fiche**, plak).
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities, `F` is volledig scherm.

### 2b. Nagelezen klikpaden

Op **jouw scherm** (Chrome, profiel Jonas, 09-10-2026: een demodocument en de demomap, zonder te delen) en in de
Nederlandse help (nagelezen 09-10-2026):

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Een document delen | rechtsboven **Delen** › naam › rol (**Kijker** · **Reageerder** · **Bewerker**) › **Verzenden** | jouw scherm (09-10-2026); de help schrijft *Sturen of Delen* ([Drive 2494822](https://support.google.com/drive/answer/2494822?hl=nl)) |
| Een map delen | rechtermuisknop op de map › **Delen** › **Delen** (of **Link kopiëren**) | jouw scherm (09-10-2026); de help: *Meer acties › Delen › Delen* |
| Algemene toegang | onder **Algemene toegang** de pijl-omlaag › **Beperkt** › **Klaar** | jouw scherm · [Drive 2494893](https://support.google.com/drive/answer/2494893?hl=nl) |
| Rol veranderen, toegang verwijderen | **Delen** › rechts van de naam de pijl-omlaag › een andere rol of **Toegang verwijderen** › **Opslaan** | [Drive 2494893](https://support.google.com/drive/answer/2494893?hl=nl) (niet op jouw scherm: daarvoor moet het document met iemand gedeeld zijn) |
| Gedeelde bestanden | in Drive links **Gedeeld met mij** | [Drive 2375057](https://support.google.com/drive/answer/2375057?hl=nl) · jouw scherm |
| Reactie | tekst selecteren › **Reactie toevoegen** (in de werkbalk of naast de tekst) › typen › **Invoegen** | jouw scherm (de knop) · [Docs 65129](https://support.google.com/docs/answer/65129?hl=nl) |
| Reactie beantwoorden | klik op de reactie › **Reageren** › typen › **Reageren** | [Docs 65129](https://support.google.com/docs/answer/65129?hl=nl) |
| Vorige versie | rechtsboven het klokje (**Laatste bewerking**) › een versie › **Deze versie herstellen** › **Herstellen** | jouw scherm (09-10-2026) · [Docs 190843](https://support.google.com/docs/answer/190843?hl=nl) |
| Versie een naam geven (extra) | klokje › naast de versie **Meer** (drie puntjes) › **Deze versie een naam geven** | [Docs 190843](https://support.google.com/docs/answer/190843?hl=nl) |
| Inleveren | **Inleveren** · **Inleveren ongedaan maken** | [Classroom 6020285](https://support.google.com/edu/classroom/answer/6020285?hl=nl) |

Klik je op de **tijd** van een versie, dan opent een vakje om ze een naam te geven (dat bleek bij het opnemen): **Esc**
sluit het. Dat staat bij *Hulp nodig?* in stap 4.

---

## 3. Het verloop van de les

| Fase | Tijd | Wat |
|---|---|---|
| **Instructie** | **10'** | Dia 1 terwijl de computers opstarten. Dia 2 lesstart (3'): *welke naam hoort bij de facturen?* Dia 3–6 (5'30"): lesdoel, drie rollen voorspellen, demo delen, demo een regel terug. Dia 7 *Zo werk je verder* (30"). |
| **Keuzewerktijd** | **40'** | Dia 7 blijft staan; de duo's werken stap 1 tot 6 af, inleveren inbegrepen. Dia 8 in de laatste minuut. |

**Eerste rondgang:** heeft iedereen een buur (of een groepje van drie)? Delen ze met het juiste schoolaccount, en
staat *Algemene toegang* op *Beperkt*?

**Tweede rondgang, rond stap 4 en 5:** lukt het terugzetten? *"Hoe weet je welke versie de goede is?"* Is de buur
verwijderd, en is de map met jou gedeeld als kijker?

**Je krijgt een e-mail per leerling** die de map NovaDepot met je deelt. Zo zie je meteen wie stap 5 haalde, en in les
10 (*Opgeruimd?*) open je de mappen via *Gedeeld met mij*.

Zeg aan het einde mondeling dat wie niet klaar is, toch inlevert, en dat eerst de toegang van de buur weg moet.

---

## 4. Verbetersleutel

### Deel 1 — Wat kan je met elke rol?

| Wat kan je? | Als Kijker | Als Reageerder | Als Bewerker |
|---|---|---|---|
| De planning lezen | ja | ja | ja (voorbeeld, al ingevuld) |
| Een reactie schrijven | nee | ja | ja |
| Een regel zelf veranderen | nee | nee (alleen een voorstel) | ja |

### Deel 2 — Een regel terug

1. De regel van **10.30 uur**: Koffiebranderij De Bonenbaas, poort 2, 1 pallet.
2. Het klokje rechtsboven (*Laatste bewerking*) › een eerdere versie waarin de regel er nog staat › *Deze versie
   herstellen* › *Herstellen*. In eigen woorden is goed, zolang de versiegeschiedenis erin zit.

### Deel 3 — Wie krijgt welke rol?

| Wie? | Rol |
|---|---|
| Sara Claes (preventiedienst) | Kijker (voorbeeld, al ingevuld) |
| Lotte Maes (personeelsdienst) | Reageerder |
| Een collega in het magazijn | Bewerker |
| Je leraar, voor je map NovaDepot | Kijker |

**Waarom Beperkt?** Dan kunnen alleen de mensen die je toevoegt het openen. Met *Iedereen met de link* kan iedereen die
de link krijgt het openen, ook buiten de school.

### Extra

- a) Een naam voor de herstelde versie, bijvoorbeeld *Planning in orde*.
- b) Een antwoord op de reactie van de buur.

### Essentiële fouten — geef hier altijd feedback op

- *Iedereen met de link* gekozen.
- De buur heeft nog toegang tot de planning (stap 5 niet gedaan).
- Iemand **Bewerker** gemaakt waar **Kijker** volstaat, vooral de leraar bij de map.
- De regel opnieuw getypt in plaats van de versie te herstellen: de planning klopt, maar het doel is gemist.

---

## 5. Schermafbeeldingen

Alle zes zijn er (09-10-2026), uit Chrome (profiel Jonas): een demodocument dat met niemand gedeeld is, en de demomap.
Het adres in het veld is vervangen door *naam van je buur*, jouw naam door *jij (eigenaar)*; Mac-sneltoetsen zijn
afgedekt. De lijst staat in `assets/screenshots/LEESMIJ.md`. Er is geen filmpje in deze les.

---

## 6. Het materiaal opnieuw maken

```bash
python3 "werkdocument/maak_werkdocumenten.py"
```

Maakt het werkdocument en de Leveringsplanning opnieuw. Vereist `python-docx`. De planning bestaat uit vier losse regels
(geen tabel): zo kan de buur één regel selecteren en wissen.

---

## 7. Leerplandoelen in je jaaroverzicht

De repository krijgt dezelfde `pre-push` hook als de andere lessen (`_tools/update_leerdoelen.py` met
`lesdoelen.json`). Het blad **ORLO** kent `BV2_04.03`, `BV2_04.04` en `BV2_05.01` (nagekeken op 09-10-2026).
`BV2_05.01` heeft officieel het niveau *Creëren*; deze les is daar maar een begin van (ondersteunend).

---

## 8. Wat nog moet blijken in de klas

1. **Delen tussen leerlingaccounts.** Dat mag (jouw antwoord, 09-10-2026). Nog niet getest: verschijnen de namen van
   klasgenoten in de lijst bij *Delen*? (Anders typen ze het schoolmailadres; dat staat bij *Hulp nodig?*.)
2. **De kopie uit Classroom delen.** Vóór het inleveren is de leerling eigenaar van de eigen kopie (bij het inleveren gaat
   het eigendom naar jou). Test met een leerlingaccount dat delen lukt.
3. **De knop Opslaan** na het veranderen van een rol komt uit de help, niet van jouw scherm. Bij *Hulp nodig?* in stap 3
   staat: zie je *Klaar*, klik dan op *Klaar*.
4. **De versiegeschiedenis** groepeert bewerkingen die kort na elkaar gebeuren. De versie van vóór het wissen is er
   altijd (minstens de versie toen de kopie gemaakt werd); de leerling kiest de versie waarin de regel er nog staat.
5. **Les 07 gemist?** Zonder map NovaDepot kan stap 5 niet. De kant-en-klare map hangt aan de opdracht van les 07
   (`NovaDepot.zip`).
6. **4ORLOa** heeft les 08 op donderdag (7de uur): tot de deadline op vrijdag blijft er weinig tijd.
7. **Nieuw sinds 09-10-2026, voor alle klassen**: de vaste startdia *Zo start je* (dia 1) en de startpagina met
   *Wat heb je nodig?* in beelden en *Vast?* als rij. Geraken de leerlingen zo zonder hulp tot bij de opdracht,
   de lespagina en het werkdocument? (`_afspraken/presentatie.md`, `_afspraken/lespagina.md`)
