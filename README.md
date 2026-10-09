# Mapping Mental Health Risk: Designing the Sample for a Clinical Research Study

Data Sources: Census, DCHB and NFHS data.

**Files**

| File | Purpose |
|---|---|
| `mhrs_compute.R` | The calculation. **It is never edited.** |
| `MHRS_Dummy_Input.xlsx` | Synthetic example of the input workbook (shows every sheet and column format). Replace its contents with your own data. |
| `MHRS_Dummy_Output.xlsx` | Example of the output produced from the dummy input. (The script names its own output `<input name>_MHRS_output.xlsx` unless you give a name.) |

---

## Contents

1. [What the script does](#1-what-the-script-does)
2. [Requirements and running it](#2-requirements-and-running-it)
3. [The input workbook: overview and general rules](#3-the-input-workbook-overview-and-general-rules)
4. [Sheet specifications (how to build each input sheet)](#4-sheet-specifications-how-to-build-each-input-sheet)
5. [Readiness checklist for an input workbook](#5-readiness-checklist-for-an-input-workbook)
6. [The `Config` sheet: every setting, and how to change them](#6-the-config-sheet-every-setting-and-how-to-change-them)
7. [How the score is calculated](#7-how-the-score-is-calculated)
8. [The output workbook](#8-the-output-workbook)
9. [Self-check against reference results](#9-self-check-against-reference-results)
10. [Validation performed](#10-validation-performed)
11. [Error messages and fixes](#11-error-messages-and-fixes)
12. [FAQ](#12-faq)
13. [Interpretation and limits](#13-interpretation-and-limits)

---

## 1. What the script does

The script reads **one Excel workbook** and writes **one Excel workbook**.

* **Input**: district- and block-level open data for one state (Census 2011 block counts, DCHB 2011 village rows, NFHS-5 district indicators, Census 2011 religion and marriage tables) and, optionally, village religion and caste counts.
* **Output**: for every block, an Access Deprivation Index, an Adolescent Vulnerability Index, a Social Diversity & Structure score, the **Mental Health Risk Score (MHRS)**, its rank, percentiles and a Low / Medium / High category. For every village (if village data are supplied) a Social Diversity Index, the village's block risk category, a priority score and a within-category rank and shortlist.

Everything the calculation needs comes from the input workbook: the district and block names, the data and every setting (weights, risk directions, cut-offs on the `Config` sheet). The script contains no state, district or block names. To run another state, build the same workbook for that state and run the same command.

```
Census_Blocks ──────────┐
DCHB_Villages ──────────┼─► block indices ─► z-scores ─► ADI ─┐
(or DCHB_Blocks)        │                                      ├─► MHRS ─► rank, percentile,
NFHS_Districts ─────────┼─► district indicators ─► AVI ────────┘           Low/Medium/High
Religion_Districts ─────┤                          SDS (weight 0 by default)
Census_Marriage ────────┘
Village_Diversity ──► village SDI ─► + block percentile ─► priority score, strata, shortlist
```

The MHRS is a **relative ranking** of blocks by concentration of risk factors. It is not a prevalence estimate, not a diagnostic tool, and describes places, not individuals (see [section 13](#13-interpretation-and-limits)).

---

## 2. Requirements and running it

**Requirements**: R (4.0 or later) and RStudio (optional). The script needs two packages, `readxl` and `openxlsx`. In RStudio they are installed automatically the first time the script runs (internet needed once). In a terminal, install them once with `install.packages(c("readxl", "openxlsx"))`.

**Run from RStudio**

1. Open `mhrs_compute.R` in RStudio.
2. Click **Source** at the top right of the editor (or press Ctrl/Cmd + Shift + Enter; or select everything with Ctrl/Cmd + A and click **Run**).
3. A file chooser opens. Select the input workbook (for example `MHRS_Dummy_Input.xlsx`).
4. The Console prints progress and ends with `Output written to: …`. The result is saved next to the input workbook as `<input name>_MHRS_output.xlsx`.

Clicking **Run** with the cursor on a single line runs only that line; use Source (or select all first). To score another workbook, click Source again. Advanced use: setting `options(mhrs.autorun = FALSE)` before sourcing loads the functions without starting a run; then call `run_mhrs("input.xlsx", "output.xlsx")`.

**Run from a terminal** (the output path is optional):

```
Rscript mhrs_compute.R  MHRS_Dummy_Input.xlsx  MHRS_Dummy_Output.xlsx
```

If the output path is omitted, the result is written next to the input as `<input name>_MHRS_output.xlsx`.

**Run from the R console** with an explicit path:

```r
source("mhrs_compute.R")        # opens the file chooser and runs
run_mhrs("MHRS_Dummy_Input.xlsx", "MHRS_Dummy_Output.xlsx")   # or name the files yourself
```

**Try it first on the dummy workbook.** The Console should end with:

```
[INFO] DCHB_Villages: 507 village rows aggregated to block level.
[INFO] Blocks read: 15 in 3 district(s).
[INFO] Block scores computed for 15 blocks.
[INFO] Village scores computed for 146 villages.
[INFO] Finished.
Output written to: …/MHRS_Dummy_Input_MHRS_output.xlsx
```

Close the output workbook in Excel before running again; Excel locks an open file and the script cannot overwrite it.

---

## 3. The input workbook: overview and general rules

| Sheet | Required | One row per | Source | Used for |
|---|---|---|---|---|
| `Config` | optional (defaults are used if absent) | setting | — | weights, directions, cut-offs, switches ([section 6](#6-the-config-sheet-every-setting-and-how-to-change-them)) |
| `Census_Blocks` | **yes** | block | Census 2011 Primary Census Abstract | female employment, female literacy, caste diversity; defines the list of blocks scored |
| `DCHB_Villages` | **yes**, unless `DCHB_Blocks` is given | village | DCHB 2011 village directory | health access, school access, infrastructure |
| `DCHB_Blocks` | alternative to `DCHB_Villages` | block | DCHB village data already aggregated to blocks | same as above |
| `NFHS_Districts` | **yes** | district | NFHS-5 district fact sheets | sex-ratio deviation, marriage/fertility, education |
| `Religion_Districts` | **yes** | district | Census 2011 Table C-01 | religious fragmentation |
| `Census_Marriage` | **yes** | district × area type | Census 2011 Table F-7 | marriage/fertility index |
| `Village_Diversity` | optional | village | partner survey or Census village data | village stage; if absent, only block scores are produced |
| `Reference_Stats` | only if `reference_mode = fixed` | indicator | a previous output | fixed reference ([section 12](#12-faq)) |
| `Reference_Block_Outputs`, `Reference_Village_Outputs` | optional | block / village | earlier results | self-check ([section 9](#9-self-check-against-reference-results)) |

**General rules**

1. **Sheet names are exact** (for example `Census_Blocks`). Sheets with other names are ignored.
2. **Row 1 is the header**; data start in row 2. No title rows, no merged cells, no blank rows inside the data.
3. **Column names** are the names in section 4. Capitalisation, spaces and punctuation in a header do not matter (`Total Population Female` is read as `total_population_female`). Extra columns are ignored.
4. **Values, not formulas.** Numbers must be stored as numbers (no "1,234" text, no "%" signs). Percentages are written on a 0–100 scale (`12.5`, not `0.125`).
5. **Names are the identifiers.** `district_name` and `block_name` link the sheets. Matching ignores capitalisation, spaces and punctuation (`Bara Banki`, `BARABANKI` and `Bara-Banki` are the same). Spellings that differ in letters (`Chitaura` / `Chittaura`) are **different**: use one spelling for a block in every sheet. Two districts may each have a block with the same name; blocks are matched within their district.
6. **One state per workbook.** Any number of districts and blocks may be included. Use at least two blocks; for meaningful standardisation use many more.
7. **Blank cells**: in `DCHB_Villages` a blank in a data column means 0 (not available / none). Everywhere else, required numeric cells must be filled. Optional columns may be blank.
8. Grey headers in the dummy workbook mark optional columns; blue headers mark required columns. Hover over a header to read its definition and source.

---

## 4. Sheet specifications (how to build each input sheet)

### 4.1 `Census_Blocks` — one row per block

**Source**: Census of India 2011, Primary Census Abstract (PCA), Community Development Block level, area = **Total** (not rural/urban separately). Field codes in brackets are the standard PCA names.

| Column | Required | Unit | Definition (PCA field) |
|---|---|---|---|
| `district_name` | yes | text | District |
| `block_name` | yes | text | CD Block |
| `total_population_person` | yes | persons | Total population (`TOT_P`) |
| `total_population_female` | yes | persons | Total female population (`TOT_F`) |
| `scheduled_castes_population_person` | yes | persons | Scheduled Castes population, persons (`P_SC`) |
| `scheduled_tribes_population_person` | yes | persons | Scheduled Tribes population, persons (`P_ST`) |
| `literates_population_female` | yes | persons | Female literates (`F_LIT`) |
| `total_workers_female` | yes | persons | Total female workers, main + marginal (`TOT_WORK_F`) |
| `main_workers_female` | yes | persons | Female main workers (`MAINWORK_F`) |
| `marginal_workers_female` | yes | persons | Female marginal workers (`MARGWORK_F`) |
| `district_code`, `block_code` | no | number/text | Carried to the output unchanged |

**Build**: filter the PCA to the CD Block rows of the state, area = Total; keep the columns above; write the district and block names exactly as they will appear in `DCHB_Villages` and `Village_Diversity`. The blocks in this sheet **define the set of blocks that are scored**: every block listed here must have DCHB data.

### 4.2 `DCHB_Villages` — one row per village

**Source**: District Census Handbook (DCHB) 2011, village directory (village amenities). The script aggregates these rows to blocks. **A blank cell counts as 0** and the denominator is **all villages listed for the block**, so list every village of the block, including those with nothing available.

| Column | Required | Values | Definition (DCHB field) |
|---|---|---|---|
| `district_name`, `block_name` | yes | text | District, CD Block (must match `Census_Blocks`) |
| `village_name` | no | text | Village (not used in calculations) |
| `pre_primary_school_available` | yes | 1 / 0 | Pre-primary availability (government or private) |
| `primary_school_available` | yes | 1 / 0 | Primary school availability |
| `middle_school_available` | yes | 1 / 0 | Middle school availability |
| `secondary_school_available` | yes | 1 / 0 | Secondary school availability |
| `senior_secondary_school_available` | yes | 1 / 0 | Senior secondary school availability |
| `private_senior_secondary_school_available` | yes | 1 / 0 | Private senior secondary school (weight 0 by default, but the column must exist) |
| `community_health_centres` | yes | count | Community Health Centres in the village |
| `primary_health_centres` | yes | count | Primary Health Centres |
| `primary_health_sub_centres` | yes | count | Primary Health Sub-Centres |
| `hospitals_allopathic` | yes | count | Allopathic hospitals |
| `hospitals_alternative_medicine` | yes | count | Alternative-medicine hospitals |
| `non_govt_health_facilities` | yes | count | Non-governmental medical facilities |
| `safe_drinking_water_available` | yes | 1 / 0 | 1 if any safe source is available: tap water (treated), covered well, hand pump, tube well / borewell |
| `mobile_phone_coverage` | yes | 1 / 0 | Mobile phone coverage |
| `pucca_road_available` | yes | 1 / 0 | Black-topped (pucca) road |
| `power_supply_summer_hours` | yes | hours/day | Domestic power supply per day, summer (April–September) |

Binary columns accept only `1`, `0` or blank. To use a different connectivity variable (for example internet access), add a 1/0 column such as `internet_access` and set `iai_connectivity_column` on `Config` ([section 6](#6-the-config-sheet-every-setting-and-how-to-change-them)).

**Aggregation to blocks** (shown in the output sheet `DCHB_Block_Aggregates`):

* `count_of_villages` = number of village rows in the block
* `pct_villages_<column>` = 100 × (sum of the 1/0 column) ÷ `count_of_villages`
* `total_<column>` = sum of each facility count over the block's villages
* `avg_power_supply_summer_hours` = mean hours over all the block's villages

**Alternative `DCHB_Blocks` sheet.** If DCHB data are already aggregated, provide a `DCHB_Blocks` sheet (and no `DCHB_Villages` sheet) with one row per block and these columns: `district_name`, `block_name`, `count_of_villages`, `pct_villages_pre_primary_school_available`, `pct_villages_primary_school_available`, `pct_villages_middle_school_available`, `pct_villages_secondary_school_available`, `pct_villages_senior_secondary_school_available`, `pct_villages_private_senior_secondary_school_available`, `pct_villages_safe_drinking_water_available`, `pct_villages_pucca_road_available`, `pct_villages_mobile_phone_coverage`, `total_community_health_centres`, `total_primary_health_centres`, `total_primary_health_sub_centres`, `total_hospitals_allopathic`, `total_hospitals_alternative_medicine`, `total_non_govt_health_facilities`, `avg_power_supply_summer_hours`. This is exactly the layout of the output sheet `DCHB_Block_Aggregates`, so an aggregate produced by one run can be reused as the `DCHB_Blocks` sheet of another. The percentage columns are on a 0–100 scale; the connectivity column is named `pct_villages_<iai_connectivity_column>`.

### 4.3 `NFHS_Districts` — one row per district

**Source**: NFHS-5 district fact sheets (indicator numbers refer to the standard district fact-sheet numbering).

| Column | Required | Unit | NFHS indicator |
|---|---|---|---|
| `district_name` | yes | text | District |
| `sex_ratio_total` | yes | females per 1,000 males | 3 — Sex ratio of the total population |
| `sex_ratio_at_birth` | yes | females per 1,000 males | 4 — Sex ratio at birth for children born in the last five years |
| `pct_women_20_24_married_before_18` | yes | % | 16 — Women age 20–24 years married before age 18 years |
| `pct_women_15_19_mothers_or_pregnant` | yes | % | 18 — Women age 15–19 years who were already mothers or pregnant at the time of the survey |
| `pct_female_6plus_ever_attended_school` | yes | % | 1 — Female population age 6 years and above who ever attended school |
| `pct_women_10plus_years_schooling` | yes | % | 15 — Women with 10 or more years of schooling |

**Build**: the district fact sheets list indicators in rows; transpose them so that each district is one row. Each district is given one value per indicator; the value is assigned to every block of that district.

### 4.4 `Religion_Districts` — one row per district

**Source**: Census 2011, Table C-01 (population by religious community), district level, area = **Total**, **female** columns.

| Column | Required | Definition |
|---|---|---|
| `district_name` | yes | District |
| `females_total` | yes | Total females |
| `females_hindu` | yes | Hindu females |
| `females_muslim` | yes | Muslim females |
| `females_christian` | yes | Christian females |
| `females_buddhist` | yes | Buddhist females |
| `females_not_stated` | yes | Females whose religion is not stated |

Sikh, Jain and "other religions" females are not needed: the index groups Christian, Buddhist and not-stated females together as the third group.

### 4.5 `Census_Marriage` — one row per district × area type

**Source**: Census 2011, Table F-7 (number of women and ever-married women by present age, educational level, number of surviving children). Use the rows for **all educational levels** ("Total").

| Column | Required | Definition |
|---|---|---|
| `district_name` | yes | District |
| `area_type` | yes | `Total`, `Rural` or `Urban` |
| `pct_below15_ever_married` | yes | Ever-married women aged under 15 ÷ all women aged under 15 × 100 |
| `pct_15_19_ever_married` | yes | Ever-married women aged 15–19 ÷ all women aged 15–19 × 100 |

Every district needs a `Total` row. By default the minimum and maximum used to scale these two items are taken over **all** rows (Total, Rural and Urban) of the sheet, so include the Rural and Urban rows. To scale on Total rows only, set `amfri_census_scaling_rows` to `total` on `Config`; the Rural and Urban rows are then not needed.

### 4.6 `Village_Diversity` — one row per village (optional sheet)

**Source**: village religion and caste counts from a partner survey, or from Census village-level tables. If this sheet is absent, the village stage is skipped and only block scores are produced.

| Column | Required | Definition |
|---|---|---|
| `district_name`, `block_name`, `village_name` | yes | Must match `Census_Blocks` for district and block |
| `ngo_presence` | no | Free text or flag (for example `Yes`). Carried to the output; not used in any calculation |
| `no_of_ngo_presence` | no | Number. Carried to the output; not used in any calculation |
| `hindu`, `muslim` | counts* | Persons |
| `buddhist`, `christian`, `jain`, `parsi`, `sikh`, `other_religions` | counts* | Persons; these six together form the "Other" religion group |
| `religion_total` | no | Total population for the religion counts; if blank, the sum of the categories is used |
| `sc`, `st`, `obc`, `general` | counts* | Persons. `sc` and `obc` together form the "SC+OBC" group |
| `caste_total` | no | Total population for the caste counts; if blank, the sum of the categories is used |
| `sdi` | no | **Precomputed** Social Diversity Index; if filled, it replaces the value calculated from counts |
| `dominant_religion` | no | Precomputed: `Hindu`, `Muslim` or `Other` (used where filled) |
| `dominant_caste` | no | Precomputed: `SC+OBC`, `ST` or `General` (used where filled) |

\* Blank counts are read as 0. Each village needs either counts (religion and caste) or the three precomputed columns. Rows in the dummy workbook show both: most rows carry counts, the last four carry precomputed values.

Villages whose district and block are not in `Census_Blocks` keep their diversity values but receive no block score (blank scores and a warning in `Run_Log`).

---

## 5. Readiness checklist for an input workbook

Use this checklist before running. "Checked" means the script stops with a message if the condition fails; the others are for you to confirm.

| # | Criterion | Pass condition | Checked by script |
|---|---|---|---|
| 1 | Sheets present | `Census_Blocks`, `NFHS_Districts`, `Religion_Districts`, `Census_Marriage`, and `DCHB_Villages` or `DCHB_Blocks` exist | yes |
| 2 | Columns present | every required column in section 4 exists | yes |
| 3 | Numbers are numbers | no text in numeric columns; required numeric cells are not blank | yes |
| 4 | Unique blocks | no district/block pair appears twice in `Census_Blocks` | yes |
| 5 | DCHB coverage | every block in `Census_Blocks` has at least one village row (or a row in `DCHB_Blocks`) | yes |
| 6 | District coverage | every district in `Census_Blocks` appears in `NFHS_Districts`; every district in `NFHS_Districts` also appears in `Religion_Districts` and has a `Total` row in `Census_Marriage` | yes |
| 7 | Binary columns | DCHB 1/0 columns contain only 1, 0 or blank | yes |
| 8 | Spelling | each district and block is spelled the same in every sheet (letters must agree) | partly (a mismatch triggers 5 or 6) |
| 9 | Counts are consistent | female ≤ total population; literates ≤ females; SC + ST ≤ population; main + marginal workers ≈ total workers; religion counts ≤ `females_total` | no |
| 10 | Percentages in range | all `pct_` columns between 0 and 100 | no |
| 11 | Sex ratios | plausible females per 1,000 males (roughly 700–1,300) | no |
| 12 | All villages listed | `DCHB_Villages` lists every village of each block (the share of villages with a facility uses all villages as denominator) | no |
| 13 | Village keys | `Village_Diversity` district and block names match `Census_Blocks` | warning only |
| 14 | One state | all districts belong to the same state; data years are consistent (Census/DCHB 2011, NFHS-5) | no |

---

## 6. The `Config` sheet: every setting, and how to change them

`Config` has the columns `group`, `parameter`, `value`, `description`. Only `parameter` and `value` are read. **To change a setting, edit its `value` cell; to restore the default, clear the cell** (an empty value means "use the default"). A parameter missing from the sheet also takes its default. The defaults below reproduce the master workbook.

### 6.1 General

| Parameter | Default | Meaning |
|---|---|---|
| `reference_mode` | `input` | input = scale and standardise on the blocks in this workbook (reproduces the master workbook); fixed = use the Reference_Stats sheet from an earlier run so new blocks are scored on that reference. |
| `percentile_digits` | `3` | Decimals kept when the percentile-rank fraction is rounded. 3 gives percentiles to 1 decimal (as in the master workbook). |
| `tercile_low_cut` | `33` | Percentile at or below which a block (or village MH category) is Low. |
| `tercile_high_cut` | `66` | Percentile at or below which a block is Medium; above it = High. Master workbook uses 66; the report text states the 67th percentile. |
| `amfri_census_scaling_rows` | `all` | all = min and max of the two Census marriage items are taken over all Total, Rural and Urban rows of Census_Marriage (master workbook); total = Total rows only. |

### 6.2 Index weights

| Parameter | Default | Meaning |
|---|---|---|
| `har_basis` | `total` | total = facility counts (master workbook); per_10000_population = counts per 10,000 persons before scaling. |
| `w_har_chc` | `0.233` | Weight of Community Health Centres in the Health Access Score. |
| `w_har_phc` | `0.233` | Weight of Primary Health Centres. |
| `w_har_phsc` | `0.233` | Weight of Primary Health Sub-Centres. |
| `w_har_hosp_allopathic` | `0.1` | Weight of Allopathic hospitals. |
| `w_har_hosp_alt_medicine` | `0.1` | Weight of Alternative-medicine hospitals. |
| `w_har_non_govt` | `0.1` | Weight of Non-governmental health facilities. |
| `w_sas_pre_primary` | `0.2` | Weight of % villages with a pre-primary school in the School Access Score. |
| `w_sas_primary` | `0.3` | Weight of % villages with a primary school. |
| `w_sas_middle` | `0.3` | Weight of % villages with a middle school. |
| `w_sas_secondary` | `0.1` | Weight of % villages with a secondary school. |
| `w_sas_senior_secondary` | `0.1` | Weight of % villages with a senior secondary school. |
| `w_sas_private_senior_secondary` | `0` | Weight of % villages with a private senior secondary school (0 in the master workbook). |
| `iai_connectivity_column` | `mobile_phone_coverage` | Column of DCHB_Villages (0/1) used as the connectivity component. mobile_phone_coverage in the master workbook; use internet_access (add that column) to follow the report text. |
| `w_iai_connectivity` | `0.3` | Weight of connectivity in the Infrastructure & Accessibility Index. |
| `w_iai_power` | `0.25` | Weight of average power-supply hours. |
| `w_iai_pucca_road` | `0.25` | Weight of % villages with a pucca (black-topped) road. |
| `w_iai_safe_water` | `0.2` | Weight of % villages with a safe drinking-water source. |
| `w_fed_literacy` | `0.6` | Weight of block female literacy (Census) in the Female Education Index. |
| `w_fed_ever_attended` | `0.2` | Weight of NFHS % females 6+ who ever attended school. |
| `w_fed_10yr_schooling` | `0.2` | Weight of NFHS % women with 10+ years of schooling. |
| `w_femp_working` | `1` | Relative weight of % females working in the Female Employment Index. |
| `w_femp_main` | `1` | Relative weight of % females who are main workers. |
| `w_femp_marginal_reversed` | `1` | Relative weight of (100 - % females who are marginal workers). |
| `srd_use_absolute` | `TRUE` | TRUE = use the absolute deviation, abs(sex ratio - sex ratio at birth); FALSE = signed difference. |

### 6.3 Risk directions and domain weights

Directions apply after standardisation: the z-score is multiplied by `+1` or `-1` so that a higher value always means higher risk. Domain weights are relative and are rescaled to sum to 1 (for example `0.5 / 0.5 / 0` and `1 / 1 / 0` are the same).

| Parameter | Default | Meaning |
|---|---|---|
| `dir_har` | `-1` | +1 = higher value means higher risk; -1 = higher value means lower risk. Health Access Score. |
| `dir_sas` | `-1` | School Access Score. |
| `dir_iai` | `-1` | Infrastructure & Accessibility Index. |
| `dir_srd` | `1` | Sex Ratio Deviation. |
| `dir_amfri` | `1` | Adolescent Marriage & Fertility Risk Index. |
| `dir_femp` | `-1` | Female Employment Index. |
| `dir_fed` | `-1` | Female Education Index. |
| `dir_cdi` | `1` | Caste Diversity Index (+1 = more diversity treated as more risk, as in the master workbook). |
| `dir_rfi` | `1` | Religious Fragmentation Index (+1 = more diversity treated as more risk). |
| `w_adi_har` | `1` | Relative weight of Health Access in the Access Deprivation Index (weights are rescaled to sum to 1). |
| `w_adi_sas` | `1` | Relative weight of School Access in the Access Deprivation Index. |
| `w_adi_iai` | `1` | Relative weight of Infrastructure in the Access Deprivation Index. |
| `w_avi_srd` | `1` | Relative weight of Sex Ratio Deviation in the Adolescent Vulnerability Index. |
| `w_avi_amfri` | `1` | Relative weight of AMFRI in the Adolescent Vulnerability Index. |
| `w_avi_femp` | `1` | Relative weight of Female Employment in the Adolescent Vulnerability Index. |
| `w_avi_fed` | `1` | Relative weight of Female Education in the Adolescent Vulnerability Index. |
| `w_sds_cdi` | `0.65` | Weight of the Caste Diversity Index in the Social Diversity & Structure score. |
| `w_sds_rfi` | `0.35` | Weight of the Religious Fragmentation Index in the Social Diversity & Structure score. |
| `w_mhrs_adi` | `0.5` | Relative weight of ADI in the Mental Health Risk Score. |
| `w_mhrs_avi` | `0.5` | Relative weight of AVI in the Mental Health Risk Score. |
| `w_mhrs_sds` | `0` | Relative weight of SDS in the Mental Health Risk Score (0 = SDS used only in the village priority stage). |

### 6.4 Village stage and self-check

| Parameter | Default | Meaning |
|---|---|---|
| `village_percentile_basis` | `all_blocks` | Block percentile used for villages: all_blocks (master workbook) or within_district. |
| `village_w_mh_percentile` | `0.7` | Weight of the block MH percentile in the village Base Priority Score. |
| `village_w_sdi` | `0.3` | Weight of the Social Diversity Index in the village Base Priority Score. |
| `diversity_percentile` | `0.33` | Villages with SDI at or above this percentile (0-1) of all village SDI values are Div_High. |
| `minority_religions` | `Muslim,Christian,Other` | Comma-separated dominant religions counted as minority. |
| `marginalised_castes` | `SC+OBC,ST` | Comma-separated dominant caste groups counted as marginalised. |
| `bonus_minority_and_marginalised` | `2` | Priority bonus when both the dominant religion and dominant caste are minority/marginalised. |
| `bonus_either` | `1` | Priority bonus when only one of them is. |
| `shortlist_primary_max` | `7` | Rank within category up to this value = Primary. |
| `shortlist_reserve_max` | `15` | Rank up to this value = Reserve. |
| `shortlist_fallback_max` | `20` | Rank up to this value = Fallback; above = Out. |
| `replication_tolerance` | `0.0001` | Absolute tolerance used when Reference_* sheets are present. |

### 6.5 Common changes

| I want to… | Set on `Config` |
|---|---|
| Use internet access instead of mobile coverage | add a 1/0 column `internet_access` to `DCHB_Villages`; set `iai_connectivity_column` = `internet_access` |
| Put the upper Medium/High cut-off at the 67th percentile | `tercile_high_cut` = `67` |
| Adjust for block size in health access (facilities per 10,000 persons) | `har_basis` = `per_10000_population` |
| Include the Social Diversity & Structure score in the MHRS | `w_mhrs_sds` > 0, for example `w_mhrs_adi` = 0.275, `w_mhrs_avi` = 0.275, `w_mhrs_sds` = 0.45 |
| Change what counts as more or less risk for an indicator | the matching `dir_` parameter (`1` or `-1`) |
| Re-weight the three Access Deprivation components | `w_adi_har`, `w_adi_sas`, `w_adi_iai` |
| Give private senior secondary schools a weight | `w_sas_private_senior_secondary` > 0 (weights of the five school types in the master workbook sum to 1) |
| Scale the Census marriage items on Total rows only | `amfri_census_scaling_rows` = `total` |
| Use the signed instead of the absolute sex-ratio deviation | `srd_use_absolute` = `FALSE` |
| Rank villages on the within-district block percentile | `village_percentile_basis` = `within_district` |
| Re-balance the village priority score | `village_w_mh_percentile`, `village_w_sdi` (the percentile is on a 0–100 scale and the SDI on 0–1, so the SDI term has a small effect at the defaults) |
| Score new blocks on an earlier reference | `reference_mode` = `fixed` plus a `Reference_Stats` sheet ([section 12](#12-faq)) |

Weights of the Health Access Score, School Access Score, Infrastructure Index and Female Education Index are used exactly as entered (they are not rescaled). Weights of domains (ADI, AVI, MHRS) and of the Female Employment components are relative and are rescaled to sum to 1. The two SDS weights (`w_sds_cdi`, `w_sds_rfi`) are used as entered.

---

## 7. How the score is calculated

Notation: `mm(x)` = (x − minimum) ÷ (maximum − minimum) over the blocks in the workbook (over the districts for district-level items); `z(x)` = (x − mean) ÷ population standard deviation over the blocks in the workbook. Defaults are shown.

### 7.1 District-level indicators (assigned to every block of the district)

| Indicator | Calculation |
|---|---|
| Sex Ratio Deviation (SRD) | `abs(sex_ratio_total − sex_ratio_at_birth)` |
| Religious Fragmentation Index (RFI) | `1 − (h² + m² + o²)` with h, m, o = share of females who are Hindu, Muslim, and Christian + Buddhist + not stated (each as a fraction of `females_total`) |
| Adolescent Marriage & Fertility Risk Index (AMFRI) | mean of `mm(pct_below15_ever_married)`, `mm(pct_15_19_ever_married)`, `mm(pct_women_20_24_married_before_18)`, `mm(pct_women_15_19_mothers_or_pregnant)` |

### 7.2 Block-level indices

| Index | Calculation | Higher means |
|---|---|---|
| Female Employment Index | mean of `mm(% females working)`, `mm(% females main workers)`, `mm(100 − % females marginal workers)`; percentages are of `total_population_female` | less risk |
| Female Education Index | `0.6·mm(female literacy %) + 0.2·mm(pct_female_6plus_ever_attended_school) + 0.2·mm(pct_women_10plus_years_schooling)`; female literacy % = `literates_population_female` ÷ `total_population_female` × 100 | less risk |
| Caste Diversity Index (CDI) | `1 − (SC share² + ST share²)`, shares as fractions of `total_population_person` | more diversity |
| Health Access Score | `0.233·mm(CHC) + 0.233·mm(PHC) + 0.233·mm(Sub-centres) + 0.1·mm(allopathic hospitals) + 0.1·mm(alternative-medicine hospitals) + 0.1·mm(non-government facilities)`, using block totals | less risk |
| School Access Score | `0.2·%pre-primary + 0.3·%primary + 0.3·%middle + 0.1·%secondary + 0.1·%senior secondary` (percent of villages with each school type; not rescaled) | less risk |
| Infrastructure & Accessibility Index | `0.30·mm(% villages with connectivity) + 0.25·mm(average power-supply hours) + 0.25·mm(% villages with pucca road) + 0.20·mm(% villages with safe drinking water)` | less risk |

### 7.3 Standardisation and domains

Each of the nine indicators (Health Access, School Access, Infrastructure, SRD, AMFRI, Female Employment, Female Education, CDI, RFI) is standardised as `direction × z(x)` using the `dir_` parameters (`-1` for Health, School, Infrastructure, Female Employment, Female Education; `+1` for the others). Then:

| Score | Calculation (defaults) |
|---|---|
| Access Deprivation Index (ADI) | mean of the standardised Health Access, School Access and Infrastructure |
| Adolescent Vulnerability Index (AVI) | mean of the standardised SRD, AMFRI, Female Employment and Female Education |
| Social Diversity & Structure (SDS) | `0.65·z(CDI) + 0.35·z(RFI)` (more diversity is treated as higher risk) |
| **Mental Health Risk Score (MHRS)** | `0.5·ADI + 0.5·AVI` (SDS has weight 0 unless `w_mhrs_sds` is set) |

### 7.4 Rank, percentile and category

* **Risk rank**: 1 = highest MHRS; ties share the lower rank.
* **Percentile** = (number of blocks with a lower MHRS) ÷ (number of blocks − 1) × 100, rounded half-up to one decimal. Reported across all blocks and within each district (a district with one block has no within-district percentile).
* **Risk category** from the all-blocks percentile: at or below `tercile_low_cut` (33) = Low; at or below `tercile_high_cut` (66) = Medium; above = High.

### 7.5 Village stage

For each village in `Village_Diversity`:

1. **Social Diversity Index (SDI)** = mean of the village RFI and village CDI.
   * Village RFI = `1 − (Hindu%² + Muslim%² + Other%²)`, where Other = Buddhist + Christian + Jain + Parsi + Sikh + other religions, as fractions of `religion_total`.
   * Village CDI = `1 − ((SC+OBC)%² + ST%² + General%²)`, as fractions of `caste_total`.
   * Dominant religion / caste = the largest group (ties resolved in the order Hindu, Muslim, Other; SC+OBC, ST, General).
2. The village takes the **percentile** of its block (all blocks, or within district per `village_percentile_basis`), and `MH_Category` = `MH_Low` / `MH_Medium` / `MH_High` using the same cut-offs as blocks.
3. `Diversity Category` = `Div_High` if SDI ≥ the `diversity_percentile` (0.33) quantile of SDI over all villages in the workbook (Excel `PERCENTILE.INC` definition), else `Div_Low`.
4. `Minority_Category`: `Minority+Marginalized` if the dominant religion is in `minority_religions` **and** the dominant caste is in `marginalised_castes`; `Marginalized` if only one is; otherwise `Dominant`. `Minority_Bonus` = 2, 1 or 0 (`bonus_minority_and_marginalised`, `bonus_either`).
5. `Base_Priority_Score` = `0.7 × block percentile + 0.3 × SDI`; `Base_Minority` = base score + bonus.
6. `Selection_Rationale` = `MH_Category_DiversityCategory_MinorityCategory`. `Rank_Within_Category` ranks villages by `Base_Minority` (highest = 1; ties share the lower rank) within each `Selection_Rationale`.
7. `Category_Shortlist` = `Primary` (rank ≤ 7), `Reserve` (≤ 15), `Fallback` (≤ 20), otherwise `Out`; `Top20_Present` = 1 for ranks ≤ 20.

---

## 8. The output workbook

| Sheet | Content |
|---|---|
| `Block_Scores` | **Main result.** One row per block, sorted by risk rank, with colour-coded risk category |
| `Block_Calculation` | All intermediate values for every block (for audit) |
| `District_Indicators` | District-level indicators: SRD, RFI, AMFRI and their components |
| `DCHB_Block_Aggregates` | DCHB village rows aggregated to blocks (reusable as a `DCHB_Blocks` sheet) |
| `Village_Scores` | One row per village (only if `Village_Diversity` was supplied) |
| `Reference_Stats` | Minimum, maximum, mean and SD used for scaling; can be copied into a later input workbook |
| `Parameters_Used` | Every setting used in the run (defaults plus `Config` overrides) |
| `Replication_Check` | Self-check table (only if `Reference_*` sheets are present) |
| `Run_Log` | Input and output paths, counts, warnings, R version, time of run |

**`Block_Scores`**: `District Name`, `Block Name`, `Access Deprivation Index (ADI)`, `Adolescent Vulnerability Index (AVI)`, `Social Diversity & Structure (SDS)`, `Mental Health Risk Score (MHRS)`, `Risk Rank (1 = highest)`, `Percentile (All Blocks)`, `Percentile (Within District)`, `Risk Category`. ADI, AVI, SDS and MHRS are in standard-deviation units: 0 = average block of the workbook, positive = more risk.

**`Block_Calculation`** columns: `district_name`, `block_name`, `district_code`, `block_code` (if supplied), `count_of_villages`; `pct_females_working`, `pct_females_main`, `pct_females_marginal`, `female_employment_index`; `female_literacy_pct`, `pct_female_6plus_ever_attended_school`, `pct_women_10plus_years_schooling`, `female_education_index`; `pct_sc`, `pct_st`, `caste_diversity_index`; `health_access_score`, `school_access_score`, `infrastructure_index`; `religious_fragmentation_index`, `sex_ratio_deviation`, `amfri`; the standardised (risk-direction) values `z_health_access`, `z_school_access`, `z_infrastructure`, `z_sex_ratio_dev`, `z_amfri`, `z_female_employment`, `z_female_education`, `z_caste_diversity`, `z_religious_fragmentation`; then `ADI`, `AVI`, `SDS`, `MHRS`, `risk_rank`, `percentile_all_blocks`, `percentile_within_district`, `risk_category`.

**`Village_Scores`** columns: `District Name`, `Block Name`, `Village Name`, `NGO Presence`, `No_of_NGO_Presence`, `Mental Health Risk Score (Block)`, `MH Percentile Within District (Block)`, `MH Percentile All Blocks (Block)`, `Block_Rank`, `Social Diversity Index`, `SDI_Rank_Within_District` (rank among the villages supplied for that district; 1 = most diverse), `Dominant Religion`, `Dominant Caste`, `MH_Category`, `Diversity Category`, `Minority_Religion_Dominant`, `Minority_Caste_Dominant`, `Minority_Category`, `Selection_Rationale`, `Minority_Bonus`, `Base_Priority_Score`, `Base_Minority`, `Rank_Within_Category`, `Top20_Present`, `Category_Shortlist`.

**`Reference_Stats`**: `name`, `type` (`minmax` or `zscore`), `min`, `max`, `mean`, `sd`. For `zscore` rows, `mean` and `sd` describe the index **before** the risk direction is applied.

---

## 9. Self-check against reference results

If the input workbook contains the optional sheets below, the script compares its own results with them and writes a `Replication_Check` sheet (item, number compared, largest absolute difference, number outside tolerance, status). The tolerance is `replication_tolerance` (default 0.0001).

* **`Reference_Block_Outputs`**: `district_name`, `block_name` and any of `ADI`, `AVI`, `SDS`, `MHRS`, `risk_rank`, `percentile_all_blocks`, `percentile_within_district`, `risk_category`.
* **`Reference_Village_Outputs`**: one row per village **in the same order as `Village_Diversity`**, with any of `MH_Category`, `Diversity Category`, `Minority_Category`, `Selection_Rationale`, `Minority_Bonus`, `Base_Priority_Score`, `Base_Minority`, `Rank_Within_Category`, `Category_Shortlist`. Villages without a block score are not compared.

Delete these sheets when running new data.

---

## 10. Validation performed

**Replication of the master workbook.** The inputs were rebuilt in this workbook format from the raw Census 2011 block file, the DCHB 2011 village file (8,454 village rows), the NFHS-5 and Census religion files and the Census marriage values held in the master workbook (76 blocks, 5 districts, 2,997 villages), and the results were compared with the master workbook's computed columns.

| Item | Result |
|---|---|
| Access Deprivation Index (76 blocks) | largest difference 3.4e-07 |
| Adolescent Vulnerability Index | largest difference 5.0e-07 |
| Mental Health Risk Score | largest difference 2.4e-07 |
| Risk rank | identical for all 76 blocks |
| Percentile (all blocks) | identical for all 76 blocks |
| Percentile (within district) | identical for the 57 blocks of the four districts in which all blocks take their own district's values |
| Village `MH_Category`, `Diversity Category`, `Minority_Category`, `Selection_Rationale`, `Category_Shortlist` | identical for all 2,989 villages that have a block score |
| Village `Base_Priority_Score`, `Base_Minority` | largest difference 5e-09 |
| Village `Rank_Within_Category` | identical for 2,988 of 2,989 villages (one village differs by one position because two villages have scores that are equal to the last decimal) |
| SDS | largest difference 4.0e-05 (the master's religion shares are stored rounded to three decimals) |

In the master workbook, one block (Hardoi–Sandi) carries the sex-ratio, religion and marriage values of Lucknow, while its education values are Hardoi's. The validation input reproduces that state by giving the block its own district row with those values; this uses only the input sheets. When every block takes its own district's values, as the script does for any workbook, that block and (through the standardisation) all other blocks shift slightly relative to the master workbook.

**Feature tests (23 checks, all passed)**: fixed-reference mode reproduces full-run scores for a subset of blocks to 1e-15; the `DCHB_Blocks` route equals the `DCHB_Villages` route; village diversity calculated from raw counts equals the cached SDI, RFI, CDI and dominant groups of the village workbook for 1,744 villages (two districts); each switch in section 6.5 runs and moves the results; `1`/`0` accepted for TRUE/FALSE; a single-district workbook runs; clear errors for a missing column and a block without DCHB rows; blocks with the same name in different districts score identically to the same data with distinct names; sheets that spell districts and blocks with different capitalisation or punctuation still join.

---

## 11. Error messages and fixes

| Message (start) | Cause | Fix |
|---|---|---|
| `The input workbook has no sheet named 'X'` | required sheet missing or misspelled | add or rename the sheet |
| `Sheet 'S' is missing required column(s): …` | a required header is absent or differently spelled | add/rename the column (section 4) |
| `Sheet 'S', column 'C' has non-numeric value(s) in row(s): …` | text in a numeric column (for example `1,234`, `12%`, `NA`) | enter a plain number |
| `Sheet 'S', column 'C' has blank cells in row(s): …` | required value missing | fill the cell (in `DCHB_Villages` blanks are allowed in data columns) |
| `Census_Blocks has duplicated district/block rows` | the same district and block appears twice | remove the duplicate |
| `No DCHB data for block(s): …` | block spelled differently in `DCHB_Villages` / `DCHB_Blocks`, or no villages listed | use one spelling in both sheets |
| `No district-level data (NFHS/Religion/Census marriage) for district(s): …` | district missing or spelled differently in `NFHS_Districts` | add/align the district |
| `Religion_Districts has no row for district(s): …` | as above for the religion sheet | add/align the district |
| `Census_Marriage has no 'Total' row for district(s): …` | no `Total` row for the district | add it |
| `DCHB_Villages column 'X' must contain only 0 or 1 (blank = 0)` | other values in a 1/0 column | recode to 1/0 |
| `Config parameter 'P' must be …` | invalid setting | use a number, TRUE/FALSE, or one of the listed options |
| `Reference_Stats has no entry 'X'` | `reference_mode = fixed` but the statistics sheet is missing entries | copy the `Reference_Stats` sheet from a full run made with the same settings |
| `Village_Diversity: SDI / dominant religion / dominant caste could not be determined for row(s): …` | a village has neither counts nor precomputed values | fill the counts or the three precomputed columns |
| `[WARN] N village(s) are in block(s) with no block score` | village block/district not found in `Census_Blocks` | align names; the villages are kept with blank scores |
| `[WARN] All values of 'X' are identical - scaled to 0` | the indicator has no variation (for example, only one district) | expected in single-district runs; district-level items then contribute 0 |
| file cannot be written | the output workbook is open in Excel | close it and re-run |

---

## 12. FAQ

**Can I run another state?** Yes. Build the workbook for that state in the same format and run the same command. District and block names, the number of districts and blocks, and all settings come from the workbook.

**Why do scores change when I add blocks or districts?** Scaling and z-scores are relative to the blocks in the workbook, so an MHRS of 0 means "average of these blocks". Scores from different workbooks are not comparable unless they use a fixed reference.

**How do I score new blocks on an earlier reference?** (1) Run the full reference workbook once. (2) Copy its `Reference_Stats` output sheet into the new input workbook under the same name. (3) Set `reference_mode` to `fixed` on the new workbook's `Config`. The new blocks are scaled and standardised with the stored minima, maxima, means and SDs. Use the same settings in both runs. Percentiles and ranks remain relative to the blocks in each run.

**Why mobile coverage and not internet access?** The master workbook uses mobile-phone coverage as the connectivity component. Set `iai_connectivity_column` to follow the report's wording (internet access); see section 6.5.

**Why are SC/ST used for caste diversity at block level but SC+OBC/ST/General at village level?** Block-level data (Census PCA) report SC and ST only; village counts include OBC and General. Both follow the master workbook.

**Why does the village SDI barely change the priority score?** The block percentile is on a 0–100 scale and the SDI on 0–1, so with weights 0.7 and 0.3 the SDI term moves the score by under 0.3 points, less than the minority bonus. This reproduces the master workbook; change `village_w_sdi` to re-balance.

**Which risk category cut-offs are used?** 33 and 66 (the master workbook). The report text describes terciles at the 33rd and 67th percentiles; set `tercile_high_cut` to 67 to follow it.

---

## 13. Interpretation and limits

* The score is a **relative ranking** of blocks by concentration of risk factors. It is **not a measure of prevalence**: a high score does not mean more girls in that block have anxiety or depression.
* It is **not a diagnostic tool** and says nothing about any individual (ecological inference): a high-scoring block contains many girls who are well, and a low-scoring block contains girls who are not.
* Several sources are dated (Census and DCHB 2011; NFHS-5 for district values). District-level indicators are assigned to every block of the district, so they differentiate districts, not blocks within a district.
* Health access uses overall health facilities as a proxy for mental-health service access, because mental-health services are not recorded directly.
* Describing a place as "high risk" can stigmatise the people living there. Present results as a guide to where support and research are most needed, and consider withholding village names from publications.
