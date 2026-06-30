# AfsprakenProgramma — Functionality Overview

Excel/VBA application for generating medical orders ("afspraken") on a NICU/PICU,
integrating with the MetaVision clinical information system. This document is a
functional inventory derived from parsing the full VBA source tree in `src/`
(~40k lines: ~50 standard modules, 26 classes, 25 user forms, 67 worksheet
code-behind modules, plus the SQL database schema in `src/sql/`).

It is organized to match the README's feature list ("Kenmerken" 1–15), followed
by the architectural, data, administration, and developer functionality that the
README leaves implicit.

> **Sources.** The functional inventory is derived primarily from the VBA source
> in `src/`. Details marked *(docs)* come from the Word documents in `docs/`
> (user manual, interface specifications, beheer/governance documents, and release
> notes); some of those describe MetaVision-side behaviour or later releases and
> are flagged where they may go beyond the parsed code snapshot.

## Architecture at a glance

- **All business logic lives in the `Mod*` standard modules.** The 67 worksheet
  code-behind files (`src/document/*.doccls`) are thin shells: they only toggle
  the window chrome on `Worksheet_Activate`, auto-hide lookup sheets on
  `Worksheet_Deactivate`, and host the workbook open/close lifecycle
  (`WbkAfspraken`).
- **Prescription state is held in named Excel ranges**, read/written through a
  safe layer (`ModRange.GetRangeValue` / `SetRangeValue`) that logs missing names
  instead of crashing.
- **Sheet naming convention:** `sht` + scope (`Glob` global / `Neo` neonatology /
  `Ped` pediatrics / `Pat` patient / `Div` divider) + type (`Gui` user interface /
  `Ber` calculation / `Tbl` lookup table / `Prt` print / `Data` / `Text`).
- **Code conventions** (`src/module/ModConventions.bas`): modules `Mod*`, classes
  `Class*`, forms `Form*`; cross-module calls are prefixed `ModName.SubName`;
  variables carry a type prefix (`str`, `int`, `bln`, `obj`, …); constants are
  `CONST_*`; named ranges start with `_` for patient data.

---

## Clinical functionality (README Kenmerken 1–15)

### 1. Patiëntgegevens (patient data)

*Code: `ModPatient`, `ClassPatientDetails`, `FormPatient`, `FormPatLijst`.*

- Enter/edit demographics (hospital number, names, admission & birth dates,
  weight, length, gender, gestational weeks/days, birth weight) via `FormPatient`
  with **live field validation**:
  - weight 0.4–200 kg, length 30–250 cm
  - gestation 21–49 weeks / 0–6 days, birth weight 400–9998 g
  - birth date ≤ today and ≤ admission date, admission date > 2006
  - actual weight ≥ birth weight / 1000
- **MetaVision sync:** auto-fill demographics from MetaVision by hospital number /
  current patient ID; automatically pick up MetaVision's selected patient and
  logged-in clinician.
- **Derived clinical values:** chronological age, gestational age, postmenstrual
  age, prematurity-corrected age, and body surface area (Du Bois formula).
- **Patient list picker** (`FormPatLijst`) backed by the live MetaVision/database
  list, filtered to admitted patients of the current department, sortable A–Z, with
  a standard/template-patient view.
- **Open / save a patient** against the application database, with version history
  (see *Data layer & persistence*).
- **Clear scopes:** PICU-only, NICU-only, patient-only, or everything.
- **Launch context** *(docs)*: the program is normally opened *from* the MetaVision
  Voorschrijven form (direct opening needs an admin password); it loads MetaVision's
  selected patient, authorizes by MetaVision login, and recognizes the department.
  Demographics sync keys on the patient number, falling back to bed location when
  the number is empty; lab sync pulls the last 24 hours.
- **Required-field rule** *(docs)*: all fields are mandatory except gestation and
  birth weight — but for patients younger than 28 days those two also become
  mandatory.

### 2. Totalen (totals)

*Code: calculation sheets `shtPedBerTot`, `shtNeoBerInfB`; `ModPedEntTPN`,
`ModNeoInfB`, `ModMetaVision`.*

- Fluid and ingredient totals are computed on dedicated calculation sheets and
  rolled up from feeds, additives, TPN components, IV-line flushes and base fluids.
- **MetaVision lab link** (`MetaVision_SyncLab`): populates the lab panel with
  recent blood/lab values and liver/kidney-function values.
- **Automatic eGFR calculation** (Schwartz for < 50 kg, MDRD otherwise) and an
  **Acute Kidney Injury alert** (creatinine rise > 26.5 µmol/L or > 1.5×, or
  diuresis < 0.5 ml/kg/h), surfaced as Dutch warning text.

### 3. Continue infusen (continuous IV medication)

*Code: `modPedContIV` (PICU), `ModNeoInfB` (NICU), `ClassNeoMedCont`, `FormMedIV`,
`FormPedMedIVPickList`, `FormNeoMedIVPickList`.*

- **Quick-pick** of continuous drips from a department formulary (max 15 PICU /
  10 NICU) with add/remove reconciliation; plus 5 free-text non-standard meds
  (PICU lines 16–20).
- **Dose → drip-rate / quantity calculation** from patient weight × solution
  volume × strength ÷ conversion factor, with clinically sensible rounding.
- **Automatic standard-concentration and standard-volume selection** per drug
  (default solvent NaCl 0.9%).
- **Dosing-limit display:** min / max / absolute-max dose and an advice string,
  sourced from configuration.
- **Concentration-limit enforcement:** quantity auto-clamped to the drug's
  min/max concentration and rounded to whole / 0.1 / 0.01 mL.
- **Solvent/diluent checking:** warns and resets when the chosen solvent
  (none / NaCl / glucose) conflicts with the drug's advised or mandatory solvent;
  drug-specific volume rules (epidural 24 mL, doxapram 12 mL, lidocaine cap 48 mL).
- **Special cases:** epinephrine (no auto-dose), epidural (weight-based quantity),
  doxapram (max-concentration based). Roll-up into totals.
- **Dose-colour semantics** *(docs)*: yellow = lower/upper limit exceeded
  (warning); red = absolute upper limit exceeded (invalid — must never be
  prescribed); blue strength/volume = non-standard solution. A red order **blocks
  printing** of the NICU worksheet and the pharmacy preparation letters.
- **Amount snapped to a multiple of the strength** (whole / 0.1 / 0.01 mL by
  magnitude) plus min/max concentration clamping *(docs; e.g. dopamine 40 mg/ml)*.
- **Epidural weight-banded dilution** *(docs)*: 0–6 kg → 2 ml/kg up to 24 ml,
  start 1 ml/h; 7–24 kg → 2 ml/kg up to 48 ml, start 2 ml/h; 25–48 kg → 1 ml/kg up
  to 48 ml, start 4 ml/h; > 48 kg undiluted, start 4 ml/h. Auto-rounding of the
  epidural volume can misfire at multi-decimal weights (1.25 kg → 1.3 ml); correct
  by editing the amount or rounding the weight.
- **NICU side-lines ("zijlijnen")** carry glucose, NaCl, sodium bicarbonate, or
  albumin, alongside an arterial line and free-text medication *(docs)*.
- **PICU non-standard meds (slots 16–20)** compute the dose in **both**
  Eenheid/kg/uur and Eenheid/kg/min *(docs)*.

### 4. Discontinue medicatie (intermittent medication)

*Code: `ModMedDisc`, `ModFormularium`, `ClassMedDisc`, `ClassDose`,
`ClassDoseRule`, `ClassFormularium`, `ClassSubstance`, `ClassSolution`,
`FormMedDisc`, `FormMedDiscPickList`, `FormPRN`, `ModWeb`.*

- Up to **30 intermittent orders**, each editable, clearable, sortable, with a
  free-text remark.
- Based on the **G-Standaard** and **Kinderformularium**; supports GPK drug
  numbers, manual entry, and **indication- and route-driven prescribing** from
  the lists valid for the selected drug.
- **Automatic dose-rule selection by demographics** (corrected age, gestational
  age, postmenstrual age, weight, gender, indication, route), pre-filling
  norm / min / max / absolute-max / max-per-dose, on a per-kg, per-m², or
  per-dose basis.
- **Two-way dose calculation** (norm dose × frequency ⇄ per-administration
  "keer" dose), capped at max-per-dose and rounded to product divisibility;
  combination-product splitting into substances with per-substance dosing text.
- Standard **frequency table**; dissolving/administration setup (department-
  specific solution selection, volume ⇄ max-concentration, minimum infusion time),
  day-dose and concentration.
- **PRN / as-needed** orders (`FormPRN`) with mandatory instruction text.
- **Medication safety surveillance ("medicatie bewaking"):** required-field
  validation, mandatory absolute-max dose for patients > 50 kg, **online dose
  retrieval from GenForm / G-Standaard** (REST/JSON via `ModWeb`), Tall-Man
  lettering, and **deep links to** the Kinderformularium (by ATC code),
  G-Standaard/GenForm, Farmacotherapeutisch Kompas, and the Parenteralia handbook.
- **HIX import:** paste an active medication list and fuzzy-match it to formulary
  drugs (reporting unmatched / overflow items).
- **Pharmacy email:** render the prescription / validation copy to PDF and email
  it to the hospital pharmacy.

**Business rules and definitions** *(docs: Discontinue_Medicatie, release notes)*:

- **Formulary scoping (double filter):** the drug list is the pharmacy assortment
  **filtered by presence in the Kinderformularium** (pediatric-relevant drugs
  only); per-drug indications map 1:1 to Kinderformularium indications, creating an
  explicit prescription ↔ indication ↔ dose link.
- **Quick-pick is MetaVision-history-scoped:** the generic quick-pick only lists
  generics previously prescribed in MetaVision; the multi-select method auto-removes
  duplicates (use single-line entry to keep duplicates).
- **"Deelbaarheid" (divisibility):** the smallest indivisible unit of a product
  (suppository strength, half-tablet, mL step). The keer-dose is always a multiple
  of it, and the **deelbaarheid unit sets the unit** for keer-dose, calculated
  dose, and the advice/min/max/abs-max doses. It is tunable to better approximate
  the advice dose (e.g. set to 1 mg, or 0.1 mL to dose in millilitres).
- **±10 % tolerance rule:** when no min/max dose is configured, the calculated dose
  must lie within 10 % of the advice dose (relative orange warning only).
- **Solution volume formula:** `Berekend Volume = Keer Dosering / Max Conc`.
  Common antibiotics ship with pre-filled solvent / Max Conc / infusion time
  (e.g. gentamicine 10 mg/ml).
- **Six dose-control checks:** (1) wrong frequency (only when the frequency list is
  restricted); (2) > 10 % off advice dose; (3) exceeds min/max/abs-max (red);
  (4) no solvent though one was specified; (5) Max Conc exceeded; (6) infusion
  time below minimum.
- **Per-dose vs cumulative toggle** and frequency-driven dosing for non-daily
  frequencies (per 36 h, per 2 days) — added in response to incident MIP 18-41384
  so doses match the Kinderformularium exactly. Dose calculation can be **switched
  off per order** (e.g. ointments).
- **Blue-marked drug** = a generic/form combination not yet known in MetaVision,
  needing a one-time action by the on-duty functional admin before planning.
- **Non-assortment / study drugs** can be entered manually (generic and indication
  become free-text). Parenteralia solutions are split PICU vs NICU.

**Combination-preparation policy** *(docs: Voorstel ... combinatiepreparaten)* —
adopted after MIP incidents from inconsistent practice:

- **Non-parenteral forms:** prescribe and monitor in **stuks / mL / doses (not mg)**.
- **Parenteral forms:** dose and monitor on the **sum in (milli)grams**, with
  products renamed to make the sum explicit (e.g. "Imipenem + cilastatine 1000 mg
  (500+500)", "Piperacilline + tazobactam 4500 mg (4000+500)").

**Medicatie bewaking (medication surveillance)** *(docs: Medicatie bewaking,
Implementatie ...)*:

- **Unit-aware, proactive engine:** the system computes resultant units itself
  (e.g. 10 ml × 3 mg/ml = 30 mg), converts dose units (mg ↔ mcg) and frequency
  expressions (2×/3 days ⇄ 1×/36 h), and shows limits **live at prescribing time**,
  per body weight **or per body surface (m²)**.
- **Four definable limits:** min and max per kg/m² per time, max per time, max per
  keer, with a selector for per-kg / per-m² / no-correction and per-dose / per-time.
- **Webservice architecture:** the GenForm request carries birth date, weight,
  length, gestation, GPK, route, and indication and returns JSON dose rules; backed
  by three components — **ZIndex.TypeProvider** (GPK → generics + ATC),
  **FormularyParser** (reads the Kinderformularium), and **GenPresCheck** (combines
  them into the dose rule). Rationale: standard G-Standaard surveillance is only
  limitedly usable for children and unsuitable for neonates. A retrieved rule set
  also narrows the allowed frequency list; route is mandatory for a search.

### 5. Lijnen en pacemaker (lines and pacemaker)

*Code: `ModPedLijnPM`, `FormPedLijnenPickList`.*

- Quick-pick of up to 6 intravascular lines, with flushes that feed into fluid
  totals; line comments.
- Pacemaker: copy standard settings into the active prescription; clear pacemaker
  data.

### 6. Voeding en TPN (nutrition and TPN)

*Code: `ModPedEntTPN`, `FormSpecialeVoeding`, `FormPedEntPickList`,
`FormNeoEntPickList`.*

- Feed/additive **quick-pick** (PICU: 1 feed + 3 additives; NICU: 1 feed +
  4 breast-milk + 4 formula additives) with auto-default frequency/volume.
- **Special-nutrition configuration:** 9 nutrient values (kcal, protein,
  carbohydrate, fat, Na, K, Ca, phosphate, Mg) rolled into nutritional totals.
- **TPN:** automatic selection of the correct standardized amino-acid composition
  **by weight** (Samenstelling B/C/D/E, NICU Mix, Nutriflex), with a safety rule
  blocking electrolyte additions to NICU Mix.
- **Automatic 3-day TPN build-up:** per-day, per-weight calculation of TPN volume,
  lipid dose, glucose concentration & base-fluid volume, electrolytes (NaCl, KCl,
  Ca-gluconate, MgCl, phosphate), trace elements (Peditrace) and vitamins
  (Vitintra, Soluvit); the chosen day is highlighted on the printout.
- Per-component manual mL entry; 24-hour volume → hourly infusion-rate conversion;
  roll-up into totals.

**Business rules** *(docs: Voeding_en_TPN_NICU/PICU, release notes)*:

- **Total fluid intake is a derived value** (computed from total oral feeding
  volume and weight), not a raw input.
- **"Extra" exclusion:** feeding (and the arterial line, continuous medication, and
  side-lines) marked "extra" is excluded from the totals.
- **TPN rest-volume (glucose base) formula:** `Totale Vocht Intake × Gewicht −
  Σ(orale voeding) − Σ(lijnen) − Σ(medicatie) − Σ(overige TPN)`; a negative result
  is flagged red.
- **PICU TPN component model:** five definable infusions — SST1 (protein *or*
  electrolyte), SST2 (electrolyte), CalcGluc + MgCl, KNaP, and Lipiden.
- **Minimum solvent-volume floor:** the system auto-sets a minimum glucose solvent
  volume so concentrations don't fall below their minimums; the pump rate cannot be
  lowered past it (reduce electrolytes/protein first).
- **Non-glucose solvent blocks protein:** choosing a non-glucose solvent zeroes the
  protein composition; the alternative is to start an SST2.
- **Two selectable lipid compositions** (incl. SMOF); weight-band boundary values
  resolve to the higher band; entering 0 ml auto-removes a composition; a protein
  composition without a chosen TPN day raises a warning.

### 7. Lab aanvragen (lab requests)

*Code: `ModPedLab`, `ModNeoLab`.*

- One-click toggling of standard test panels at fixed sampling rounds (admission,
  14:00, 19:00, 24:00, day-1; ~31–32 tests per round); lab comments. NICU labs are
  simpler (clear + comment).

### 8. Afspraken en controles (appointments and controls)

*Code: `ModPedAfspr`, `ModNeoAfspr`, `FormTekstInvoer`, `FormOpmerking`.*

- Quick-clear plus structured free-text entry: other appointments, compensation
  fluid (Ped), wound-culture location and other (Neo).

### 9. Infuusbrief Neonatologie (NICU infusion chart)

*Code: `ModNeoInfB`, sheets `shtNeoGuiInfB` / `shtNeoBerInfB` / `shtNeoBerAdvies`,
`FormCopy1700`.*

- Integrated overview of enteral feeding + continuous medication + IV lines + TPN.
- **Age/gestation-based increasing fluid intake** (advice formulas on
  `shtNeoBerAdvies`, driven by weight and gestation), with manual override per
  kg/day.
- **Phototherapy / glucose fluid correction** (intake auto-adjusted when TPN
  glucose changes).
- **Dual "Actueel" / "17:00" versions** with selective copy-forward / copy-back of
  feeding, continuous medication and TPN (`FormCopy1700`, showing only the
  differences).
- Enteral vs parenteral totals; arterial-line handling; one-click standard TPN +
  lipid composition.
- **Copy 17:00 → Actueel** defaults to only "TPN overnemen" checked; on taking over
  the TPN, rest-volume differences are compensated in the fluid intake so the TPN
  composition stays identical to the 17:00 version. Pharmacy printing is blocked
  from the Actueel version (must be the 17:00 version) *(docs)*.

### 10. Infuusbrief voor elektroliet-oplossingen en TPN

*Code: `ModPedPrint`, `ModNeoInfB`.* — Generation of the TPN / electrolyte
infusion chart (see §14, printing).

### 11. Werkbrieven Neonatologie (NICU work-sheets)

*Code: `ModNeoPrint`.* — `PrintNeoWerkBrief` / `SaveNeoWerkBrief` produce the NICU
nursing worksheet, validated for the 17:00 version and for valid continuous meds
and TPN.

### 12. Bereidingsvoorschriften apotheek (pharmacy preparation / VTGM)

*Code: `ModNeoPrint`.* — `PrintApotheekWerkBrief` produces one pharmacy
preparation letter **per continuous-medication item** (loops the 10 infusion-chart
med slots); `SendApotheekWerkBrief` emails the worksheet + all VTGM PDFs to the
pharmacy (Cc neonatology), gated on physician login, 17:00 version, and valid meds.

### 13. Acute medicatie / APLS (emergency drugs)

*Code: sheets `shtPedGuiAcuut` / `shtNeoGuiAcuut`; `ModPedPrint`, `ModNeoPrint`.* —
Acute-care ("Acuut Blad") sheets with weight-based emergency drug and intervention
calculations; printable.

### 14. Uitprint (printing)

*Code: `ModPedPrint`, `ModNeoPrint`, `FormPrintAfspraken`, `ModSheet`.*

- Print / preview / PDF of: continuous medication, discontinuous medication, the
  acute sheet, the **weight-correct TPN chart** (5 banded sheets: 2–6, 7–15,
  16–30, 31–50, > 50 kg), the NICU worksheet, and pharmacy letters — with the bed
  number stamped into the page header.

### 15. Koppeling met MetaVision (MetaVision integration)

*Code: `ModMetaVision`, `ModDatabase`, `ModString.ConcatenateKeyValue`.*

The integration is **two-way but asymmetric**, using two different mechanisms:

- **Inbound (pull, live SQL):** read patient demographics and department lists,
  lab signals, the logged-in user, and the **medication-order catalog**
  (`MetaVision_GetMedicatieOpdrachten`) directly from MetaVision's SQL Server
  (connection from a `secret` credentials file plus registry keys). This path is
  read-only against MetaVision (SELECT only).
- **Outbound (push, file handoff — NOT a SQL write-back):** generated orders are
  exported to a key-value data/"Tekst" file that **MetaVision imports**. The app
  does not INSERT into MetaVision's database; the write path is the shared file the
  MetaVision side reads. *(docs: Interface specification — see dedicated section
  below)*

> **Correction to an earlier caveat:** a code-only reading suggests "MetaVision is
> read-only / orders are not written back." That is only half right: there is no
> *direct SQL* write to MetaVision, but the round-trip is completed via the
> file-based export above plus MetaVision-side validation/planning forms (see
> "MetaVision-side workflow"). So README point 15 ("verwerking in MetaVision") is
> realized as **designed file export + active MetaVision-side import**, not as a
> database write from the VBA.

---

## Interface specification (AfsprakenProgramma ↔ MetaVision)

*Source: docs `Interface_AfsprakenProgramma_MetaVision_versie_1_0`,
`Interface_Definities_..._V5` (the two are near-identical design docs).*

- **Mechanism:** the app writes all generated orders to an Excel data/"Tekst"
  workbook named after the patient's MetaVision bed; MetaVision imports that file.
- **Serialization:** key-value (column 1 = key, column 2 = value). Hierarchical
  keys are `Group.Element` (e.g. `Patient.Nummer`); nested key-value pairs are
  separated by `||` with `^^` between key and value; a deeper sublevel uses `##`
  between pairs and `::` between key and value.
- **Exported data dictionary (21 groups):** `Afspraak` (timestamp), `Patient`,
  `Lab`, `MedDisc` (shared NICU + PICU); department-split groups `PedMedCont`,
  `PedEnt`, `PedTPN`, `PedAccess`, `PedPM`, `PedLab`, `PedAfspr`, `PedIntake`,
  `NeoEnt`, `NeoArtLijn`, `NeoMedCont`, `NeoZijLijn`, `NeoTPN`, `NeoIntake`,
  `NeoLab`, `NeoAfspr`.
- **Field-level payload (selected):** each `MedDisc` order carries Indic, Generiek,
  Vorm, Sterkte, Freq, Hoev, Route, Dose, OplHoev, OplKeuze, Tijd, PRN, Opm, **ATC**,
  **GPK**, and **Etiket** (official G-Standaard label text). Continuous-med groups
  carry Medicament/Sterkte/Volume/Oplossing/Stand/Dosering/Totaal/Advies plus
  `Extra` (does the volume count toward fluid balance) and `Tijd` (pump run-time).
  TPN groups serialize the full ingredient list with Stand and TotDag; Lab and
  Intake groups export computed per-kg/day totals.
- **Status:** these are interface *specifications*, but the export path is at least
  partially live — a release note records "imports powders correctly into
  MetaVision," and the order text the spec relies on is actually produced by the
  code (stored in the `PrescriptionText` table).

## MetaVision-side workflow: Valideren, Tekenen & Plannen

*Source: docs `Afspraken_Voorschrijven`, `Afspraken_Plannen`, user manual. These
forms live in MetaVision (not in the `src/` VBA tree) and complete the round-trip.*

- **Voorschrijven / Valideren / Tekenen:** the "Medische Afspraken Voorschrijven"
  form is where imported orders are validated and electronically signed. Signing
  requires three fields — **Supervisor, Voorschrijver, Besproken met**; unsigned
  orders show in red. A "Veranderingen" tab contrasts `== NIEUWE AFSPRAKEN ==` vs
  `== VERVALLEN AFSPRAKEN ==`. One-off ("Eenmalige") orders are a separate tab.
- **Medication-history rule:** every change stops the current prescription and
  starts a new one; an order dropping out of the actual list automatically gets a
  stop date/time; active prescriptions have no stop date. Full change history is
  kept (who/what/when).
- **Plannen:** signing creates a "Taak Medische Afspraak Wijziging" notifying the
  nurse. The "Medische Afspraken Plannen" form reconciles not-yet-planned vs
  no-longer-current orders (lapsed planned orders must be cancelled first).
  Administration times must match the order frequency (mismatch warns); PRN orders
  are planned via "Zo nodig." A clean Plannen form is the guarantee that everything
  ordered was processed; a Planning Log records all planning actions.

---

## Application shell & environment

*Code: `ModApplication`, `ModSetting`, `WbkAfspraken`.*

- Kiosk-style startup/shutdown: hide gridlines / headings / tabs / formula bar,
  set the window title to patient + bed, protect UI sheets, very-hide calculation
  sheets; restore Excel and quit on close.
- **Development-mode toggle** (unprotect/reveal sheets for editing).
- Environment detection (Development / Training / Acceptation / Production, and
  PICU vs NICU) from the workbook path and registry.
- **Test-vs-production database toggle** and **logging on/off** toggle.

## Data layer & persistence

*Code: `ModDatabase`, `ModRange`, `src/sql/GenerateDB.sql`.*

- Application SQL database (`UMCU_WKZ_AP*`), designed **append-only with version
  IDs**. 11 tables: `Patient`, `Prescriber`, `PrescriptionData`,
  `PrescriptionText`, `ConfigMedCont`, `ConfigMedDisc`, `ConfigMedDiscDose`,
  `ConfigMedDiscSolution`, `ConfigMedTallMan`, `ConfigParEnt`, `Log` — accessed via
  ~40 versioned table-valued functions and insert/update stored procedures.
- **Per-patient prescription version history:** save with a newer-version conflict
  warning; open the latest *or* a specific historical version; reusable
  "standard / template patients" (`standaard_NNN`, max 999).
- Per-workstation configuration resolved by computer name
  (`GetRegistryForComputerName`), falling back to the Windows registry.
- **Audit logging** to the database (`InsertLog`) and to log files.
- Safe named-range data layer with bulk rename, name registry export, and
  patient-data refresh tooling.
- **Formularium lazy-load performance redesign** *(docs: Performance verbetering,
  v0.60.51-beta)*: loading the formulary was made many times faster by splitting
  the discontinue-med query — `GetConfigMedDiscLatest` now returns only the
  drug-header columns (plus the Tall-Man join), and dose rules are fetched lazily
  per drug via the parameterized TVF
  `GetConfigMedDiscDoseLatestForGenericShape(@Generic, @Shape)`.

## Administration / Beheer

*Code: `ModAdmin`, `FormAdminNeoMedCont`, `FormAdminParent`, `FormColorPicker`,
`FormFontPicker`.* (The README's "Beheer" section is currently empty; this is what
the code provides.)

- Role-gated ribbon groups (Beheerders / Apotheek).
- **Editors for the medication knowledge base:**
  - NICU/PICU continuous-medication formulary (`FormAdminNeoMedCont`), with
    min/max consistency validation.
  - Parenteralia composition (`FormAdminParent`).
  - Discontinue-medication formulary.
  - Each supports **DB version selection, external xlsx import/export, and
    save-to-database-or-file**.
- Open log files and refresh the MetaVision order catalog.
- **Appearance / theming:** per-department color and font configuration
  (`FormColorPicker` / `FormFontPicker`) and rule-based conditional formatting
  (info / warning / error highlighting) on prescription sheets.

**Deployment & provisioning** *(docs: Applicatie_Beheer_Manual, Beheerdocument)*:

- Distributed as a GitHub release ZIP into one of three recognized folders
  (`AfsprakenProgramma_Productie` / `_Training` / `_Test`), unpacking to
  `AfsprakenProgramma.xlsm` + a `db` config folder, with the runnable copy under an
  `App` subfolder and a sibling `Data` folder.
- Requires the `secret` MetaVision credentials file and read access to registry
  `HKCU\SOFTWARE\UMCU\MV\`.
- Configuration tables are version-managed by printing/saving them as PDF.

## Developer tooling

*Code: `ModUtils`, `ModRange`, `FormNaamGeven`, `FormLog`.*

- **Source-control export** (`ExportForSourceControl`): writes all VBA code, sheet
  formulas, and defined names to the `src/` tree — this is how this repository is
  produced.
- Range-naming tools (`FormNaamGeven`, bulk rename, patient-data refresh),
  clipboard helpers, shell/git execution, and performance toggling.
- In-app log viewer (`FormLog`).

## Governance & change management

*Source: docs `Beheerdocument`, `Besluit Stuurgroep`, `Gebruikers groep`. Not code
functionality, but the operational regime the application runs under.*

- **Three-tier environment promotion:** Test (development + initial test) →
  Training (training + acceptance/FAT) → Productie; every change flows
  Test → Acceptance → Production.
- **Change classification:** *Fix* (no regression risk, no CAB), *Minor* (warrants
  regression tests — e.g. adding/removing continuous IV meds — no CAB), *Major*
  (new/changed functionality — requires CAB review). Post go-live (28 Nov 2017) the
  system is frozen except for legal/regulatory changes and bug fixing; the Steering
  Group acts as Change Advisory Board.
- **Functional ownership matrix:** named owners test and sign off each release
  (PICU, NICU, Pharmacy/KFS for Neo continuous-IV preparation & dosing math, and
  DIT/app-management for config, colours, datafiles, and the MetaVision catalog);
  pharmacy owns the Neo continuous-med + parenteralia tables, clinical divisions
  own the rest.
- **Validation intent:** the system is expected to run automated prescribing
  scenarios that **permute every prescribing variable** (continuous and
  discontinuous) for clinical validation.
- **Regulatory framing:** built to meet IGJ/GMP requirements; error dialogs are
  auto-emailed to the developer; full logging of who viewed which patient and what
  they changed.

## Supporting libraries & tests

- **VBA-Web v4.1.3** (`WebHelpers`, `WebClient`, `WebRequest`, `WebResponse`, and
  the authenticator classes) and **VBA-JSON v2.2.3** (`JsonConverter`) — MIT
  open-source, used for the GenForm REST calls.
- In-house utility libraries: arrays (merge sort), collections, locale-aware
  string/number parsing, dates (1904 empty-date convention), Excel-function
  wrappers, hash-code builder, branded message boxes, a modeless progress dialog,
  file I/O, Word-document reading, and CDO/SMTP email.
- Test suites runnable from the ribbon: `ModNeoInfB_Tests`, `ModMedDisc_Tests`,
  `ModTests`, with assertions in `ModAssert`.

---

## Notes & caveats

- MetaVision integration: there is **no direct SQL write-back** to MetaVision, but
  README point 15 is realized as a **file-based order export** that MetaVision
  imports, plus MetaVision-side validation/planning forms. See §15 and the
  "Interface specification" and "MetaVision-side workflow" sections.
- Login: `FormPassword` and `ClassUser` provide the UI and data model, but **no
  credential verification** exists in code; user identity comes from MetaVision /
  registry. Admin actions are gated by a hardcoded password (`"hla"` in
  `ModConst`).
- The authors' own comments flag some dead/duplicate code (e.g. a duplicate
  `btnPedMedContExport` case in the ribbon dispatcher in `ModRibbon`).

---

## Sheet inventory (feature surface)

| Scope | GUI (`Gui`) | Calculation (`Ber`) | Lookup (`Tbl`) / Data | Print (`Prt`) |
|---|---|---|---|---|
| **Global** | Front page; MedDisc list | Conv, Lab, Norm, MedDisc, MedDiscMail, Sql | Ent, ParEnt, MedDisc, MedOpdr, Names, Settings, Temp | MedDisc |
| **Neonatology** | Acuut, Afspr, InfB, Lab | Advies, Afspr, InfB, Lab | Ent, Lijst, MedIV; DataInfB | Afspr, Apoth, MedDisc, Werkbr |
| **Pediatrics** | Acuut, Afspr, EntTPN, Lab, LijnPM, MedIV | Afspr, Ent, IVenPM, Lab, MedIV, TPN, Tot | Afspr, Ent, IV, MedIV, ParEnt, Gewicht, Lengte | Afspr, MedDisc, TPN×5 (2–6/7–15/16–30/31–50/>50 kg) |
| **Patient** | — | — | Data, Details, Text | — |
