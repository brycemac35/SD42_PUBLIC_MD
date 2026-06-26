# SPC Time-Study Report — Reusable Template & Method

**Document type:** `REF-` (analytical reference)
**Naming:** `REF-<PLATFORM>-###_<Keywords>` (e.g. `REF-GEN-002_MacBook_M1_Imaging_Analysis`)
**Style:** SD42 house style — navy `1F4E79`, Arial, ISO 9001:2015 structured numbering
**Worked example throughout:** REF-GEN-002 (MacBook M1 imaging + JAMF enrollment combined time study)

This file is both a **fillable template** (placeholders in `«guillemets»`) and a **worked example** (the REF-GEN-002 values shown alongside). Copy it, replace the placeholders with your new device group's data, and follow the build instructions at the end to produce the DOCX.

---

## 0. How to use this template

1. Duplicate this file and rename for your device group.
2. Fill the **Data Input** block (Section 9) with your measurements.
3. Decide the structural choices flagged with ☑ (target window, subgroup size, fail definition, special-cause removal).
4. Run the build script (Section 10) to generate the SD42-styled DOCX with embedded charts.
5. Visually QA the PDF proof, then publish.

> **Principle:** confirm structural decisions *before* building (target window, subgrouping, which points are special causes). Document special-cause removals explicitly — never silently drop points.

---

## 1. Title Page

| Field | Placeholder | Example (REF-GEN-002) |
|---|---|---|
| Title | «Platform» Deployment Time Analysis | MacBook M1 Deployment Time Analysis |
| Subtitle | «Stage A» + «Stage B» — Combined Time Study | MacBook M1 Imaging + JAMF Enrollment — Combined Time Study |
| Doc number | REF-«PLATFORM»-### | REF-GEN-002 |
| Rev | «X.0» | 3.0 |
| Author | «Name» | Bryce M |
| Date | «YYYY-MM-DD» | 2026-06-25 |
| Purpose | 2–3 sentences: what's measured, method, audience | Combined two-stage prep time via SPC + MIL-STD-105E sampling |

**Optional title-page throughput box** (include when planning capacity):

| Measure | Lab | On-site |
|---|---|---|
| Cycle / batch size | «n» devices | «cart size» (staggered) |
| Per-device total | «mean» min | «×1.5–2» min |
| Devices/day («h» h, theoretical) | «calc» | «calc» |
| Devices/day (whole cycles) | «floor» («cycles») | «floor» («cycles») |

Add caveats as callouts (e.g. on-site bandwidth degradation, concurrency/staggering).

---

## 2. Executive Summary

Lead with the **headline number** and the **single most important finding**, not a neutral recap.

- Opening: what was measured and how (1 short paragraph).
- **Bold** headline result sentence.
- Root-cause callout (red) if a special cause exists.
- Secondary failure/rate callout (amber) if relevant.
- Closing: capability before/after in plain terms.
- **Headline metrics table** (before vs after).

> Example headline: *"Once the wifi/activation special cause is removed, a prepared device takes ≈73 min end to end (≈34 imaging + ≈40 enrollment)."*

---

## 3. Scope & Method

- **3.1 The process** — describe the stages. If multi-stage sequential, state that total = sum of stages.
- **3.2 Sampling & subgrouping** — n per subgroup, number of subgroups, sampling basis.
- **3.3 How to read the results** — control limits (voice of process) vs spec limits (voice of customer); capability benchmark 1.33; the fail definition.

---

## 4. Combined / Primary Analysis (before vs after)

If the special cause exists, present **before vs after** side by side.

- Embedded X̄ & R chart figure (before | after).
- **4.1 What the charts show** — bullets on R-chart control, σ reduction, residual signals.
- Special-cause removal callout (documented, not arbitrary).
- **4.2 Capability table** — Cp / Cpk / Pp / Ppk, before & after, vs benchmark.

---

## 5. Per-Stage Sections (repeat per stage)

For each stage (imaging, enrollment, …):

- Summary metrics table (before/after if applicable).
- Embedded X̄ & R chart.
- Sampling basis if the stage was independently sampled (see Section 7).
- Control status statement (in control / out of control + why).

---

## 6. Distribution & Capability

- Histogram(s) vs spec limits (before/after).
- Note bimodality / multiple populations if present (a key tell for special vs common cause).
- Capability index table with plain-language reading.

---

## 7. Sampling Basis — MIL-STD-105E / ANSI-ASQ Z1.4

Use when you need a defensible **sample size** from a finite population (lot). Used here as **sample-size rationale only** (not lot accept/reject) unless you explicitly want acceptance sampling.

**Step 1 — lot size → code letter (General Inspection Level II, the default):**

| Lot size | Code | n (Level II) |
|---|---|---|
| 2–8 | A/B | 2–3 |
| 9–15 | C | 5 |
| 16–25 | D | 8 |
| 26–50 | E | 13 |
| 51–90 | F | 20 |
| 91–150 | G | 32 |
| 151–280 | **G** | **32** |
| 281–500 | H | 50 |
| 501–1200 | J | 80 |

> ☑ **Inspection level:** Level I (less, smaller n), **Level II (default)**, or Level III (tighter, larger n).
> Example: lot 200 → Level II → code **G** → **n = 32**.

**Step 2 — accept/reject (only if doing acceptance sampling):** choose an AQL, read Ac/Re from the single-sampling table. Example: n=32, AQL 2.5% → Ac=2, Re=3. *Omit this step if using MIL-STD purely for sample size.*

---

## 8. Methodology — SPC Formulas

**X̄ & R control charts** (subgroup size n, k subgroups):

```
X̄ⱼ = mean of subgroup j
R̄ⱼ = max(j) − min(j)
X̳ (grand mean) = mean of X̄ⱼ
R̄ = mean of Rⱼ

X̄ chart:  UCL = X̳ + A₂·R̄ ,  CL = X̳ ,  LCL = X̳ − A₂·R̄
R  chart:  UCL = D₄·R̄ ,      CL = R̄ ,  LCL = D₃·R̄

within-process σ = R̄ / d₂
```

**Capability** (spec limits LSL, USL; mean μ):

```
Cp  = (USL − LSL) / (6·σ_within)
Cpk = min[ (USL − μ)/(3·σ_within) , (μ − LSL)/(3·σ_within) ]
Pp  = (USL − LSL) / (6·σ_overall)     ← σ_overall = sample stdev of all readings
Ppk = min[ (USL − μ)/(3·σ_overall) , (μ − LSL)/(3·σ_overall) ]
```

Benchmark: **≥ 1.33** capable; **< 1.00** routinely out of spec. Cp/Cpk use *within* σ; Pp/Ppk use *overall* σ. A gap between them signals batch-to-batch shift.

**Control-chart constants (by subgroup size n):**

| n | A₂ | D₃ | D₄ | d₂ |
|---|---|---|---|---|
| 2 | 1.880 | 0 | 3.267 | 1.128 |
| 3 | 1.023 | 0 | 2.574 | 1.693 |
| 4 | 0.729 | 0 | 2.282 | 2.059 |
| 5 | 0.577 | 0 | 2.114 | 2.326 |
| 6 | 0.483 | 0 | 2.004 | 2.534 |
| 7 | 0.419 | 0.076 | 1.924 | 2.704 |
| 8 | **0.373** | **0.136** | **1.864** | **2.847** |
| 9 | 0.337 | 0.184 | 1.816 | 2.970 |
| 10 | 0.308 | 0.223 | 1.777 | 3.078 |

**Fail definition options** (when "fail = slow"):

- **Option A** — fixed threshold borrowed from a prior study (only if same process / shared time budget).
- **Option B** — this dataset's own mean + 1σ_overall (self-referential; flags ~top 16%).
- **Option C** — mean + 1σ_within (excludes batch drift; tighter; correct when batches drift).

> ☑ Pick A/B/C. Example used **B/C** (they coincided because the JAMF process was in control: σ_overall ≈ σ_within).

**Special-cause removal:** only remove points with a *known, assignable* cause. Document the cause, recompute limits on the remaining data, and show before/after.

---

## 9. Combined Time Model (multi-stage)

**Per-device total** (sequential, independent stages):

```
μ_total = μ_stageA + μ_stageB + …
σ_total = √(σ_A² + σ_B² + …)        ← root-sum-square, NOT simple sum
```

**Per-lot wall-clock** (S devices per tech per cycle, rolling/staggered starts → cycle ≈ mean of the batch):

```
cycles = ⌈ lot_size / S ⌉
lot_time (per tech) ≈ cycles × μ_total
```

**Daily throughput** (W hours available):

```
theoretical devices/day = (W·60 / μ_total) × S       (fractional)
practical  devices/day  = ⌊ W·60 / (μ_total) ⌋_cycles × S   (whole cycles only)
```

> Example: μ_total ≈ 73 min, S=8, W=6h → ~39/day theoretical, **32/day** (4 whole cycles). Lot 200 → 25 cycles ≈ **31 h/tech**.
> **On-site:** degrade the *network-bound* stage only (enrollment ×1.5–2). Larger concurrent batches (full cart 20–25) share bandwidth → stagger to avoid compounding.

---

## 10. Findings & Recommendations

- **10.1 Key findings** — bold bullets, most important first.
- **10.2 Likely causes to investigate** — if special cause unresolved.
- **10.3 Recommended actions** — numbered table: # / Action / Owner / Priority.
- Closing note callout: state units assumed, special causes excluded, any management-supplied figures.

---

## 11. DATA INPUT — paste new measurements here

Fill this block for the new device group; everything else flows from it.

```
DOC NUMBER:        REF-«PLATFORM»-###
DEVICE GROUP:      «e.g. iPad 9th gen / Dell 3420»
STAGE A NAME:      «e.g. Imaging»          SOURCE: «e.g. SanDisk SDDDC4 USB»
STAGE B NAME:      «e.g. JAMF enrollment»  SOURCE: «e.g. JAMF cloud»

TARGET WINDOW (min):   LSL = «20»   USL = «50»     ☑ confirm
SUBGROUP SIZE n:       «8»                          ☑ confirm
FAIL DEFINITION:       «B: mean + 1σ_overall»       ☑ A/B/C

STAGE A DATA  (one subgroup per row, n values each):
  SG1: «v, v, v, v, v, v, v, v»
  SG2: «…»
  …

STAGE B DATA:
  SG1: «…»
  …

SPECIAL CAUSES (subgroups to remove + reason):
  «e.g. SG1, SG5 — wifi reconnect failure stalls activation»

OPERATIONAL FAILURE RATE (if supplied):  «e.g. 1/4»
SAMPLING (if used):  lot = «200»  level = «II»  → code «G» → n = «32»

CAPACITY PLANNING (optional):
  devices per cycle S = «8»   working hours W = «6»
  on-site enrollment multiplier = «1.5–2×»   cart size = «20–25»
```

**Worked example (REF-GEN-002), for reference:**

```
TARGET WINDOW:  imaging 20–50 ; combined 40–90
SUBGROUP SIZE:  8
FAIL DEF:       B (≈ C, process in control)

JAMF STAGE B DATA (n=32, 4 subgroups):
  J1: 22.32, 35.64, 35.72, 40.72, 49.04, 49.14, 54.34, 56.64
  J2: 33.07, 34.14, 35.28, 35.28, 35.43, 35.47, 43.69, 43.69
  J3: 20.00, 30.36, 32.17, 39.03, 42.47, 44.14, 45.80, 54.49
  J4: 31.25, 33.85, 39.13, 40.91, 41.61, 43.67, 46.78, 48.58
  → mean 39.8, σ 8.55, in control; fail threshold 48.3 → 6/32 slow

SPECIAL CAUSES (imaging + combined): SG1, SG5 — wifi reconnect → activation stall
OPERATIONAL FAILURE RATE: 1/4
SAMPLING: lot 200 → Level II → code G → n=32
CAPACITY: S=8, W=6h → 32/day lab (4 cycles), 24/day on-site (3 cycles); lot 200 ≈ 31 h/tech
```

---

## 12. BUILD INSTRUCTIONS

### 12.1 Environment

```bash
# DOCX generation
npm install docx                       # Node docx library
# Charts
pip install matplotlib openpyxl numpy --break-system-packages
# Fonts: Liberation Sans at /usr/share/fonts/truetype/liberation/ (Arial substitute)
```

### 12.2 SD42 style constants (apply in the docx build script)

```
NAVY      = 1F4E79   (headers, header band, table headers)
NAVY_LIGHT= D6E4F0   (header-right cell)
GREY_ROW  = F2F2F2   (zebra striping)
BORDER    = CCCCCC   (table grid)
Font      = Arial, body size 22 half-pt (11 pt)
Callouts  = red C00000/FDECEA (danger) ; amber BF8F00/FFF2CC (note) ; left bar
Header band = navy table, title + subtitle left, Doc/Author/Rev/Date right
Footer    = "School District No. 42 — Information Technology Services" + Page X of Y
```

### 12.3 Chart rendering (matplotlib)

```
- Agg backend; Liberation Sans; 170 DPI; navy #1F4E79 lines, red #C00000 limits,
  green #1E7B34 for "after" series, grey #595959 centre line.
- X̄ & R as stacked subplots (or 2×2 for before/after). Red markers on OOC points.
- Histogram vs LSL/USL dashed red, mean grey. Save PNG, embed via docx ImageRun.
```

### 12.4 Generate, validate, QA

```bash
node build.js                                              # writes the .docx
python3 /mnt/skills/public/docx/scripts/office/validate.py  <file>.docx
# Visual QA:
soffice --headless --convert-to pdf <file>.docx
pdftoppm -jpeg -r 90 <file>.pdf page                       # inspect page-*.jpg
```

### 12.5 Publish

- Output to `/mnt/user-data/outputs/`, present for download.
- SD42 style template (source of truth):
  `https://raw.githubusercontent.com/brycemac35/SD42_PUBLIC_MD/main/SD42_DOCX_Style_Template.md`
- Publish to GitHub manually; verify via `raw.githubusercontent.com`.

---

## 13. Checklist before publishing

- [ ] Target window (LSL/USL) confirmed for this device group
- [ ] Subgroup size set; correct A₂/D₃/D₄/d₂ constants used
- [ ] Special causes identified, cause documented, before/after shown
- [ ] Fail definition (A/B/C) chosen and stated
- [ ] Capability indices computed on correct σ (within vs overall)
- [ ] Sampling basis documented if a sample was drawn from a lot
- [ ] Combined σ via root-sum-square (not simple sum)
- [ ] Throughput figures show theoretical AND whole-cycle
- [ ] On-site caveats noted if applicable (bandwidth, concurrency, staggering)
- [ ] DOCX validated + PDF visually proofed
- [ ] Closing note states units, exclusions, supplied figures
