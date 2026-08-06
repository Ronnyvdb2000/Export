# VLIF4-automatisering — technische en gebruikshandleiding

## 1. Wat dit systeem doet

Elke **maandag, woensdag en vrijdag om 17:00 UTC** controleert het systeem
automatisch of er een nieuwe versie van de VLIF4-codelijst gepubliceerd is
op de Vlaamse overheidswebsite. Zo ja: de PDF wordt gedownload, omgezet
naar een Excel-werkboek, gevalideerd tegen een AI-model, en het resultaat
wordt gemaild — allemaal zonder dat jij iets moet opstarten.

**Je hoeft normaal gezien nooit iets handmatig te doen.** Dit document
beschrijft wat er automatisch gebeurt, en hoe je het kan testen of
bijsturen indien nodig.

---

## 2. Architectuur — de twee workflows

Het systeem bestaat uit twee GitHub Actions workflows die na elkaar draaien:

```
┌─────────────────────────┐      commit + push       ┌──────────────────────────┐
│  vlif4_weekly_check.yml │  ────── naar input/ ────► │   vlif4-pipeline.yml     │
│  (detectie + download)  │                            │  (build + valideer+mail)│
└─────────────────────────┘                            └──────────────────────────┘
   draait op cron-schema                                  draait automatisch bij
   (ma/wo/vr 17:00 UTC)                                    elke push op input/*.pdf
```

### 2.1 `vlif4_weekly_check.yml` — detectie en download

- **Trigger:** cron-schema (`0 17 * * 1,3,5`) + handmatig (`workflow_dispatch`)
- **Doet:**
  1. Haalt de VLIF-webpagina op
     (`lv.vlaanderen.be/subsidies/vlif-steun/...`)
  2. Zoekt de link met tekst "VLIF4 Codelijst blok ..." en leest de
     versie eruit (bv. "03-2026")
  3. Vergelijkt die versie met `data/last_version.txt`
  4. **Geen wijziging** → workflow stopt, er gebeurt verder niets
  5. **Wel een nieuwe versie** → downloadt de PDF naar `input/`, werkt
     `data/last_version.txt` bij, en **commit + pusht** dit naar de
     `main`-branch

Dit script staat in `scripts/check_new_pdf.py`.

### 2.2 `vlif4-pipeline.yml` — build, validatie, mail

- **Trigger:** automatisch bij elke push die een bestand onder
  `input/**/*.pdf` wijzigt (dus ook de commit van hierboven) + handmatig
  (`workflow_dispatch`, met optioneel een ander e-mailadres als input)
- **Doet (drie jobs na elkaar):**
  1. **`build`** — `src/layout_engine.py` + `src/build_excel.py` zetten
     elke PDF in `input/` om naar een Excel-werkboek in `output/`
  2. **`validate`** — `src/validatie_output.py` controleert elk
     Excel-bestand tegen de bron-PDF met een lokaal Phi-3.5-taalmodel
     (gedownload en gecached), en schrijft een validatierapport
  3. **`notify`** — `send_email.py` verstuurt een mail via SendGrid met
     de Excel-bestanden én de validatierapporten als bijlage

### 2.3 `convert.yml` — enkel handmatig testen (optioneel)

- **Trigger:** enkel `workflow_dispatch` (sinds de aanpassing geen
  automatische push-trigger meer, om dubbele builds te vermijden)
- **Doet:** enkel stap 1 (build), zonder validatie of mail — handig om
  snel te testen of de PDF-conversie zelf werkt, zonder de volledige
  (tragere) validatie-flow te doorlopen.

---

## 3. Waarom je geen "naar welk e-mailadres?"-vraag meer ziet

`notify_email` is een **workflow_dispatch-input** — zo'n veld verschijnt
alleen als je een workflow *handmatig* start via de Actions-tab in
GitHub. Nu `vlif4-pipeline.yml` automatisch draait via een **push**-trigger,
is er geen mens die een formulier invult, dus wordt automatisch de
fallback gebruikt: de **`NOTIFY_EMAIL`-secret**. Die staat al correct
ingesteld op repo-niveau — er verandert dus niets aan wie de mail
ontvangt, alleen hoe dat e-mailadres wordt doorgegeven.

Wil je toch ooit naar een ander adres mailen? Start `vlif4-pipeline.yml`
dan handmatig via "Run workflow" — dan verschijnt het invulveld wél
opnieuw.

---

## 4. Benodigde GitHub Secrets (al ingesteld)

| Secret | Gebruikt door | Doel |
|---|---|---|
| `NOTIFY_EMAIL` | `vlif4-pipeline.yml` → `send_email.py` | ontvanger van de resultaatmail |
| `SENDGRID_API_KEY` | `send_email.py` | authenticatie bij SendGrid |
| `SENDGRID_FROM` | `send_email.py` | geverifieerd afzenderadres |

Geen actie nodig — deze staan al op repo-niveau (Settings → Secrets and
variables → Actions → Repository secrets).

---

## 5. Bestandsoverzicht

| Pad | Rol |
|---|---|
| `.github/workflows/vlif4_weekly_check.yml` | detectie + download (cron) |
| `.github/workflows/vlif4-pipeline.yml` | build + validatie + mail (push-getriggerd) |
| `.github/workflows/convert.yml` | handmatige build-only test |
| `scripts/check_new_pdf.py` | scraper/versievergelijker |
| `data/last_version.txt` | laatst verwerkte blok-versie (bv. "03-2026") |
| `input/*.pdf` | brondocument(en) |
| `src/layout_engine.py` | PDF → ruwe Excel |
| `src/build_excel.py` | ruwe Excel → tabblad-werkboek |
| `src/validatie_output.py` | AI-validatie van het eindresultaat |
| `send_email.py` | verstuurt resultaatmail via SendGrid |

---

## 6. Gebruik — normale werking

**Niets.** Het systeem draait zelfstandig drie keer per week. Je merkt er
alleen iets van als:
- er een nieuwe codelijst-versie gepubliceerd wordt → je krijgt een mail
  met het resultaat, positief of met validatie-issues gemarkeerd
- er geen nieuwe versie is → er gebeurt niets, geen mail, geen melding

---

## 7. Handmatig testen

### 7.1 Enkel de detectie testen (geen download/mail)
Actions-tab → **VLIF4 wekelijkse check** → **Run workflow**.
Als de online versie gelijk is aan `data/last_version.txt`, gebeurt er
verder niets (verwacht gedrag).

### 7.2 De volledige flow forceren (detectie → build → validatie → mail)
1. Pas `data/last_version.txt` tijdelijk aan naar een andere waarde dan
   de huidige (bv. leegmaken of `00-0000`), en commit dat.
2. Trigger `vlif4_weekly_check.yml` handmatig (zie 7.1).
3. Het script ziet nu een "nieuwe" versie, downloadt de PDF opnieuw,
   commit en pusht.
4. Die push triggert automatisch `vlif4-pipeline.yml` — dit is de
   volledige end-to-end test, inclusief de mail.

### 7.3 Enkel de build testen (snel, zonder validatie/mail)
Actions-tab → **Convert PDF to Excel (PyMuPDF engine)** → **Run workflow**.

---

## 8. Troubleshooting

| Probleem | Mogelijke oorzaak |
|---|---|
| Geen mail ontvangen na een nieuwe versie | Controleer of `vlif4-pipeline.yml` effectief getriggerd is (Actions-tab); check de `notify`-job logs voor SendGrid-foutmeldingen |
| `vlif4_weekly_check.yml` vindt geen link | De structuur van de VLIF-pagina is gewijzigd — `LINK_TEXT_PATTERN` in `check_new_pdf.py` moet dan aangepast worden |
| Build faalt met "geen PDF's in input/" | `input/` is leeg — normaal enkel een probleem bij handmatig testen na verwijderen van bestanden |
| Validatie duurt lang / faalt | Phi-model download/cache-probleem — check de `validate`-job logs |

---

## 9. Wijzigingsgeschiedenis van deze automatisering

- Detectie + download toegevoegd (`vlif4_weekly_check.yml`,
  `scripts/check_new_pdf.py`)
- `vlif4-pipeline.yml` uitgebreid met een `push`-trigger op
  `input/**/*.pdf`, zodat build → validatie → mail automatisch volgt
  op een nieuwe download
- `convert.yml` teruggebracht naar enkel `workflow_dispatch`, om
  dubbele builds te vermijden nu `vlif4-pipeline.yml` ook op push draait
- Cron vastgezet op 17:00 UTC, ma/wo/vr (geen zomer-/wintertijd-correctie)
