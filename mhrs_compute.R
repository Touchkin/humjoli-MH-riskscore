#!/usr/bin/env Rscript
# ==============================================================================
#  MENTAL HEALTH RISK SCORE (MHRS) FOR ADOLESCENT GIRLS - BLOCK & VILLAGE SCORING
# ------------------------------------------------------------------------------
#  Reads ONE input workbook (open-access Census / DCHB / NFHS data + optional
#  village diversity data) and writes ONE output workbook with the computed scores.
#
#  This file is never edited. Everything the calculation needs - district and block
#  identifiers, the data, and every setting (weights, directions, cut-offs) - is
#  taken from the input workbook ('Config' sheet and the data sheets).
#
#  USAGE (RStudio)    Open this file and click "Source" (or select all and click "Run").
#                     A file chooser opens: pick the input workbook (.xlsx). The result is
#                     saved next to it as "<input name>_MHRS_output.xlsx".
#  USAGE (terminal)   Rscript mhrs_compute.R  <input.xlsx>  [<output.xlsx>]
#  USAGE (R console)  source("mhrs_compute.R")    # same as RStudio; or run_mhrs("<input.xlsx>", "<output.xlsx>")
#
#  Packages: readxl, openxlsx (installed automatically on first use in RStudio / the R console).
#
#  The built-in defaults below reproduce the master workbook; any of them is
#  overridden by a value on the 'Config' sheet of the input workbook.
#  Full documentation: MHRS_Instructions.md
# ==============================================================================

local({
  need <- c("readxl", "openxlsx")
  miss <- need[!vapply(need, requireNamespace, logical(1), quietly = TRUE)]
  if (length(miss)) {
    if (interactive()) {
      message("Installing required package(s): ", paste(miss, collapse = ", "), " (one-time, needs internet)")
      install.packages(miss, repos = "https://cloud.r-project.org")
    } else stop("Please install the required package(s) first: install.packages(c(",
                paste0('"', miss, '"', collapse = ", "), "))", call. = FALSE)
  }
})
suppressPackageStartupMessages({
  library(readxl)
  library(openxlsx)
})

# ------------------------------------------------------------------------------
# 0. DEFAULT PARAMETERS  (overridden by the 'Config' sheet)
# ------------------------------------------------------------------------------
DEFAULTS <- list(
  # --- general ---------------------------------------------------------------
  reference_mode            = "input",   # "input": scale/standardise on the blocks supplied; "fixed": use 'Reference_Stats' sheet
  percentile_digits         = 3,         # percentile-rank fraction is rounded to this many decimals (3 -> 1 decimal on the 0-100 scale)
  tercile_low_cut           = 33,        # percentile <= this  -> Low
  tercile_high_cut          = 66,        # percentile <= this  -> Medium, otherwise High
  # --- Adolescent Marriage & Fertility Risk Index -----------------------------
  amfri_census_scaling_rows = "all",     # "all": min/max of the two Census marriage items over all Total+Rural+Urban rows; "total": Total rows only
  # --- Health access (HAR) ----------------------------------------------------
  har_basis                 = "total",   # "total": facility counts; "per_10000_population": counts per 10,000 persons
  w_har_chc = 0.233, w_har_phc = 0.233, w_har_phsc = 0.233,
  w_har_hosp_allopathic = 0.1, w_har_hosp_alt_medicine = 0.1, w_har_non_govt = 0.1,
  # --- School access (SAS) ----------------------------------------------------
  w_sas_pre_primary = 0.2, w_sas_primary = 0.3, w_sas_middle = 0.3,
  w_sas_secondary = 0.1, w_sas_senior_secondary = 0.1, w_sas_private_senior_secondary = 0,
  # --- Infrastructure & accessibility (IAI) -----------------------------------
  iai_connectivity_column   = "mobile_phone_coverage",  # 0/1 column of DCHB_Villages used as the connectivity component
  w_iai_connectivity = 0.30, w_iai_power = 0.25, w_iai_pucca_road = 0.25, w_iai_safe_water = 0.20,
  # --- Female education (FEdI) and employment (FEmpI) -------------------------
  w_fed_literacy = 0.6, w_fed_ever_attended = 0.2, w_fed_10yr_schooling = 0.2,
  w_femp_working = 1, w_femp_main = 1, w_femp_marginal_reversed = 1,
  # --- Sex-ratio deviation ----------------------------------------------------
  srd_use_absolute          = TRUE,      # TRUE: |sex ratio - sex ratio at birth|
  # --- Risk direction of each standardised indicator (+1 higher = more risk, -1 higher = less risk)
  dir_har = -1, dir_sas = -1, dir_iai = -1, dir_srd = 1, dir_amfri = 1,
  dir_femp = -1, dir_fed = -1, dir_cdi = 1, dir_rfi = 1,
  # --- Domain weights (relative; rescaled to sum to 1) --------------------------
  w_adi_har = 1, w_adi_sas = 1, w_adi_iai = 1,
  w_avi_srd = 1, w_avi_amfri = 1, w_avi_femp = 1, w_avi_fed = 1,
  w_sds_cdi = 0.65, w_sds_rfi = 0.35,
  w_mhrs_adi = 0.5, w_mhrs_avi = 0.5, w_mhrs_sds = 0,
  # --- Village stage ----------------------------------------------------------
  village_percentile_basis  = "all_blocks",  # block percentile used by villages: "all_blocks" | "within_district"
  village_w_mh_percentile   = 0.7, village_w_sdi = 0.3,
  diversity_percentile      = 0.33,          # villages with SDI >= this percentile of SDI are 'Div_High'
  minority_religions        = "Muslim,Christian,Other",
  marginalised_castes       = "SC+OBC,ST",
  bonus_minority_and_marginalised = 2, bonus_either = 1,
  shortlist_primary_max = 7, shortlist_reserve_max = 15, shortlist_fallback_max = 20,
  # --- Optional self-check ------------------------------------------------------
  replication_tolerance     = 1e-4
)

# ------------------------------------------------------------------------------
# 1. SMALL HELPERS
# ------------------------------------------------------------------------------
norm_names <- function(x) {                      # "No_of_NGO_Presence " -> "no_of_ngo_presence"
  x <- tolower(trimws(as.character(x)))
  x <- gsub("[^a-z0-9]+", "_", x)
  gsub("^_+|_+$", "", x)
}
key_of   <- function(x) toupper(gsub("[^A-Za-z0-9]", "", as.character(x)))   # join key: ignores case, spaces, punctuation
is_blank <- function(x) is.na(x) | trimws(as.character(x)) == ""

LOG <- new.env()
LOG$lines <- character(0)
log_msg <- function(..., level = "INFO") {
  msg <- paste0(...)
  LOG$lines <- c(LOG$lines, sprintf("[%s] %s", level, msg))
  if (level == "INFO") cat(sprintf("[%s] %s\n", level, msg)) else message(sprintf("[%s] %s", level, msg))
}

# Percentile rank = Excel/Sheets PERCENTRANK.INC (share of values strictly below, scaled 0-100)
pct_rank <- function(x, digits = 3) {
  n <- length(x)
  if (n < 2) return(rep(NA_real_, n))
  r <- vapply(x, function(v) sum(x < v) / (n - 1), numeric(1))
  floor(r * 10^digits + 0.5 + 1e-9) / 10^digits * 100     # round half up, as the spreadsheet does
}
wmean <- function(lst, w) {                      # weighted mean of equally-long vectors, weights rescaled to sum to 1
  if (sum(w) <= 0) stop("Domain weights must sum to a positive number.", call. = FALSE)
  w <- w / sum(w)
  Reduce(`+`, Map(function(x, wi) x * wi, lst, w))
}
tier <- function(p, lo, hi) ifelse(is.na(p), NA_character_, ifelse(p <= lo, "Low", ifelse(p <= hi, "Medium", "High")))

# ------------------------------------------------------------------------------
# 2. READING THE INPUT WORKBOOK
# ------------------------------------------------------------------------------
read_config <- function(path) {
  cfg <- DEFAULTS
  if ("Config" %in% readxl::excel_sheets(path)) {
    raw <- as.data.frame(suppressMessages(readxl::read_excel(path, sheet = "Config", col_types = "text",
                                                             .name_repair = "minimal")), stringsAsFactors = FALSE)
    names(raw) <- norm_names(names(raw))
    if (!all(c("parameter", "value") %in% names(raw)))
      stop("The 'Config' sheet needs the columns 'parameter' and 'value'.", call. = FALSE)
    raw <- raw[!is_blank(raw$parameter), , drop = FALSE]
    for (i in seq_len(nrow(raw))) {
      p <- trimws(raw$parameter[i]); v <- raw$value[i]
      if (is_blank(v)) next
      if (!p %in% names(DEFAULTS)) { log_msg("Unknown Config parameter ignored: ", p, level = "WARN"); next }
      d <- DEFAULTS[[p]]
      if (is.logical(d)) {
        u <- toupper(trimws(v))
        cfg[[p]] <- if (u %in% c("TRUE", "T", "1", "YES")) TRUE else if (u %in% c("FALSE", "F", "0", "NO")) FALSE else NA
        if (is.na(cfg[[p]])) stop("Config parameter '", p, "' must be TRUE or FALSE (found '", v, "').", call. = FALSE)
      } else if (is.numeric(d)) {
        x <- suppressWarnings(as.numeric(v))
        if (is.na(x)) stop("Config parameter '", p, "' must be a number (found '", v, "').", call. = FALSE)
        cfg[[p]] <- x
      } else cfg[[p]] <- trimws(v)
    }
  } else log_msg("No 'Config' sheet found - built-in defaults are used.", level = "WARN")
  chk <- function(p, ok) if (!cfg[[p]] %in% ok)
    stop("Config parameter '", p, "' must be one of: ", paste(ok, collapse = ", "), call. = FALSE)
  chk("reference_mode", c("input", "fixed")); chk("amfri_census_scaling_rows", c("all", "total"))
  chk("har_basis", c("total", "per_10000_population")); chk("village_percentile_basis", c("all_blocks", "within_district"))
  cfg
}

# Reads one sheet as text, normalises headers, checks required columns, converts numbers.
read_input_sheet <- function(path, sheet, text = character(), num = character(),
                             opt_text = character(), opt_num = character(),
                             blank_zero = character(), required = TRUE) {
  if (!(sheet %in% readxl::excel_sheets(path))) {
    if (required) stop("The input workbook has no sheet named '", sheet, "'.", call. = FALSE)
    return(NULL)
  }
  df <- as.data.frame(suppressMessages(readxl::read_excel(path, sheet = sheet, col_types = "text",
                                                          .name_repair = "minimal")), stringsAsFactors = FALSE)
  names(df) <- norm_names(names(df))
  df <- df[, names(df) != "", drop = FALSE]
  if (anyDuplicated(names(df))) stop("Sheet '", sheet, "' has duplicated column names: ",
                                      paste(unique(names(df)[duplicated(names(df))]), collapse = ", "), call. = FALSE)
  if (nrow(df) > 0) {
    blank <- vapply(df, is_blank, logical(nrow(df)))
    if (is.null(dim(blank))) blank <- matrix(blank, nrow = nrow(df))
    df <- df[rowSums(!blank) > 0, , drop = FALSE]
  }
  need <- c(text, num)
  miss <- setdiff(need, names(df))
  if (length(miss)) stop("Sheet '", sheet, "' is missing required column(s): ", paste(miss, collapse = ", "),
                         ".\n  Columns found: ", paste(names(df), collapse = ", "), call. = FALSE)
  if (nrow(df) == 0) stop("Sheet '", sheet, "' contains no data rows.", call. = FALSE)
  for (cn in intersect(c(text, opt_text), names(df))) df[[cn]] <- trimws(df[[cn]])
  for (cn in text) if (any(is_blank(df[[cn]])))
    stop("Sheet '", sheet, "', column '", cn, "' has blank cells in row(s): ",
         paste(head(which(is_blank(df[[cn]])) + 1, 10), collapse = ", "), call. = FALSE)
  for (cn in intersect(c(num, opt_num), names(df))) {
    v <- df[[cn]]; bl <- is_blank(v)
    x <- suppressWarnings(as.numeric(v))
    bad <- !bl & is.na(x)
    if (any(bad)) stop("Sheet '", sheet, "', column '", cn, "' has non-numeric value(s) in row(s): ",
                       paste(head(which(bad) + 1, 10), collapse = ", "), call. = FALSE)
    if (cn %in% blank_zero) x[bl] <- 0
    else if (cn %in% num && any(bl)) stop("Sheet '", sheet, "', column '", cn, "' has blank cells in row(s): ",
                                          paste(head(which(bl) + 1, 10), collapse = ", "), call. = FALSE)
    df[[cn]] <- x
  }
  df
}

# ------------------------------------------------------------------------------
# 3. SCALERS: min-max and z-score, either on the supplied data or on fixed reference statistics
# ------------------------------------------------------------------------------
make_scalers <- function(cfg, fixed = NULL) {
  stats <- list()
  mm <- function(x, name, ref = x) {                 # (x - min) / (max - min)
    if (cfg$reference_mode == "fixed") {
      s <- fixed[[name]]
      if (is.null(s)) stop("Reference_Stats has no entry '", name, "'.", call. = FALSE)
      lo <- s$min; hi <- s$max
    } else { lo <- min(ref, na.rm = TRUE); hi <- max(ref, na.rm = TRUE) }
    stats[[name]] <<- list(type = "minmax", min = lo, max = hi, mean = NA_real_, sd = NA_real_)
    if (isTRUE(hi == lo)) { log_msg("All values of '", name, "' are identical - scaled to 0.", level = "WARN"); return(rep(0, length(x))) }
    (x - lo) / (hi - lo)
  }
  zs <- function(x, name) {                          # (x - mean) / population SD
    if (cfg$reference_mode == "fixed") {
      s <- fixed[[name]]
      if (is.null(s)) stop("Reference_Stats has no entry '", name, "'.", call. = FALSE)
      mu <- s$mean; sdv <- s$sd
    } else { mu <- mean(x, na.rm = TRUE); sdv <- sqrt(mean((x - mu)^2, na.rm = TRUE)) }
    stats[[name]] <<- list(type = "zscore", min = NA_real_, max = NA_real_, mean = mu, sd = sdv)
    if (isTRUE(sdv == 0)) { log_msg("All values of '", name, "' are identical - z-score set to 0.", level = "WARN"); return(rep(0, length(x))) }
    (x - mu) / sdv
  }
  list(mm = mm, zs = zs,
       table = function() data.frame(
         name = names(stats), type = vapply(stats, `[[`, "", "type"),
         min = vapply(stats, `[[`, 0, "min"), max = vapply(stats, `[[`, 0, "max"),
         mean = vapply(stats, `[[`, 0, "mean"), sd = vapply(stats, `[[`, 0, "sd"),
         row.names = NULL, stringsAsFactors = FALSE))
}

read_fixed_stats <- function(path) {
  df <- read_input_sheet(path, "Reference_Stats", text = c("name", "type"),
                         opt_num = c("min", "max", "mean", "sd"))
  setNames(lapply(seq_len(nrow(df)), function(i) list(min = df$min[i], max = df$max[i], mean = df$mean[i], sd = df$sd[i])),
           df$name)
}

# ------------------------------------------------------------------------------
# 4. DCHB: village rows -> block aggregates   (blank village cell = 0; denominator = all villages)
# ------------------------------------------------------------------------------
DCHB_BINARY <- c("pre_primary_school_available", "primary_school_available", "middle_school_available",
                 "secondary_school_available", "senior_secondary_school_available",
                 "private_senior_secondary_school_available", "safe_drinking_water_available", "pucca_road_available")
DCHB_COUNTS <- c("community_health_centres", "primary_health_centres", "primary_health_sub_centres",
                 "hospitals_allopathic", "hospitals_alternative_medicine", "non_govt_health_facilities")
DCHB_HOURS  <- "power_supply_summer_hours"

aggregate_dchb <- function(v, cfg) {
  v$dkey <- key_of(v$district_name); v$bkey <- key_of(v$block_name)
  g  <- paste(v$dkey, v$bkey, sep = "|")
  ug <- unique(g); f <- factor(g, levels = ug)
  first <- match(ug, g)
  out <- data.frame(dkey = v$dkey[first], bkey = v$bkey[first],
                    district_name = v$district_name[first], block_name = v$block_name[first],
                    count_of_villages = as.vector(table(f)), stringsAsFactors = FALSE)
  for (cn in c(DCHB_BINARY, cfg$iai_connectivity_column)) {
    x <- v[[cn]]
    if (any(!x %in% c(0, 1))) stop("DCHB_Villages column '", cn, "' must contain only 0 or 1 (blank = 0).", call. = FALSE)
    out[[paste0("pct_villages_", cn)]] <- 100 * as.vector(tapply(x, f, sum)) / out$count_of_villages
  }
  for (cn in DCHB_COUNTS) out[[paste0("total_", cn)]] <- as.vector(tapply(v[[cn]], f, sum))
  out[[paste0("avg_", DCHB_HOURS)]] <- as.vector(tapply(v[[DCHB_HOURS]], f, mean))
  out
}

read_dchb <- function(path, cfg) {
  vill <- read_input_sheet(path, "DCHB_Villages", text = c("district_name", "block_name"),
                           opt_num = c(DCHB_BINARY, cfg$iai_connectivity_column, DCHB_COUNTS, DCHB_HOURS),
                           blank_zero = c(DCHB_BINARY, cfg$iai_connectivity_column, DCHB_COUNTS, DCHB_HOURS), required = FALSE)
  if (!is.null(vill)) {
    miss <- setdiff(c(DCHB_BINARY, cfg$iai_connectivity_column, DCHB_COUNTS, DCHB_HOURS), names(vill))
    if (length(miss)) stop("Sheet 'DCHB_Villages' is missing required column(s): ", paste(miss, collapse = ", "), call. = FALSE)
    log_msg("DCHB_Villages: ", nrow(vill), " village rows aggregated to block level.")
    return(aggregate_dchb(vill, cfg))
  }
  conn <- paste0("pct_villages_", cfg$iai_connectivity_column)
  blk <- read_input_sheet(path, "DCHB_Blocks", text = c("district_name", "block_name"),
                          num = c("count_of_villages", paste0("pct_villages_", DCHB_BINARY), conn,
                                  paste0("total_", DCHB_COUNTS), paste0("avg_", DCHB_HOURS)), required = FALSE)
  if (is.null(blk)) stop("Provide either a 'DCHB_Villages' sheet (village rows) or a 'DCHB_Blocks' sheet (block aggregates).", call. = FALSE)
  log_msg("DCHB_Blocks: ", nrow(blk), " block rows read (already aggregated).")
  blk$dkey <- key_of(blk$district_name); blk$bkey <- key_of(blk$block_name)
  blk
}

# ------------------------------------------------------------------------------
# 5. DISTRICT-LEVEL INDICATORS  (NFHS, Religion, Census marriage)  -> copied to every block of the district
# ------------------------------------------------------------------------------
compute_districts <- function(nfhs, rel, cm, cfg, sc) {
  d <- nfhs
  d$dkey <- key_of(d$district_name)
  if (anyDuplicated(d$dkey)) stop("NFHS_Districts has duplicated districts.", call. = FALSE)
  # Sex-ratio deviation
  d$srd_signed <- d$sex_ratio_total - d$sex_ratio_at_birth
  d$srd <- if (isTRUE(cfg$srd_use_absolute)) abs(d$srd_signed) else d$srd_signed
  # Religious Fragmentation Index (females): 1 - (Hindu^2 + Muslim^2 + (Christian+Buddhist+Not stated)^2)
  rel$dkey <- key_of(rel$district_name)
  i <- match(d$dkey, rel$dkey)
  if (anyNA(i)) stop("Religion_Districts has no row for district(s): ", paste(d$district_name[is.na(i)], collapse = ", "), call. = FALSE)
  tot <- rel$females_total[i]
  d$pct_hindu  <- 100 * rel$females_hindu[i]  / tot
  d$pct_muslim <- 100 * rel$females_muslim[i] / tot
  d$pct_christian_buddhist_not_stated <- 100 * (rel$females_christian[i] + rel$females_buddhist[i] + rel$females_not_stated[i]) / tot
  d$rfi <- 1 - ((d$pct_hindu / 100)^2 + (d$pct_muslim / 100)^2 + (d$pct_christian_buddhist_not_stated / 100)^2)
  # Adolescent Marriage & Fertility Risk Index: mean of four min-max scaled items
  cm$dkey <- key_of(cm$district_name); cm$area <- tolower(trimws(cm$area_type))
  tr <- cm[cm$area == "total", , drop = FALSE]
  j <- match(d$dkey, tr$dkey)
  if (anyNA(j)) stop("Census_Marriage has no 'Total' row for district(s): ", paste(d$district_name[is.na(j)], collapse = ", "), call. = FALSE)
  ref <- if (cfg$amfri_census_scaling_rows == "all") cm else tr
  d$amfri_below15_married_scaled <- sc$mm(tr$pct_below15_ever_married[j], "amfri_below15_married", ref = ref$pct_below15_ever_married)
  d$amfri_15_19_married_scaled   <- sc$mm(tr$pct_15_19_ever_married[j],   "amfri_15_19_married",   ref = ref$pct_15_19_ever_married)
  d$amfri_20_24_married_scaled   <- sc$mm(d$pct_women_20_24_married_before_18, "amfri_20_24_married_before_18")
  d$amfri_teen_mother_scaled     <- sc$mm(d$pct_women_15_19_mothers_or_pregnant, "amfri_15_19_mothers_pregnant")
  d$pct_below15_ever_married <- tr$pct_below15_ever_married[j]
  d$pct_15_19_ever_married   <- tr$pct_15_19_ever_married[j]
  d$amfri <- (d$amfri_below15_married_scaled + d$amfri_15_19_married_scaled +
                d$amfri_20_24_married_scaled + d$amfri_teen_mother_scaled) / 4
  d
}

# ------------------------------------------------------------------------------
# 6. BLOCK-LEVEL INDICES, STANDARDISATION, DOMAINS AND THE MENTAL HEALTH RISK SCORE
# ------------------------------------------------------------------------------
compute_blocks <- function(census, dchb, dist, cfg, sc) {
  b <- census
  b$dkey <- key_of(b$district_name); b$bkey <- key_of(b$block_name)
  if (anyDuplicated(paste(b$dkey, b$bkey))) stop("Census_Blocks has duplicated district/block rows.", call. = FALSE)
  i <- match(paste(b$dkey, b$bkey), paste(dchb$dkey, dchb$bkey))
  if (anyNA(i)) stop("No DCHB data for block(s): ", paste(head(paste0(b$block_name[is.na(i)], " (", b$district_name[is.na(i)], ")"), 10), collapse = "; "),
                     "\n  Block names must match between Census_Blocks and DCHB (case, spaces and punctuation are ignored).", call. = FALSE)
  h <- dchb[i, , drop = FALSE]
  k <- match(b$dkey, dist$dkey)
  if (anyNA(k)) stop("No district-level data (NFHS/Religion/Census marriage) for district(s): ",
                     paste(unique(b$district_name[is.na(k)]), collapse = ", "), call. = FALSE)
  d <- dist[k, , drop = FALSE]
  conn <- paste0("pct_villages_", cfg$iai_connectivity_column)
  pop_f <- b$total_population_female
  o <- data.frame(district_name = b$district_name, block_name = b$block_name, stringsAsFactors = FALSE)
  for (cn in intersect(c("district_code", "block_code"), names(b))) o[[cn]] <- b[[cn]]
  o$count_of_villages <- h$count_of_villages

  # Female Employment Index: mean of min-max scaled (% working, % main workers, 100 - % marginal workers)
  o$pct_females_working  <- 100 * b$total_workers_female    / pop_f
  o$pct_females_main     <- 100 * b$main_workers_female     / pop_f
  o$pct_females_marginal <- 100 * b$marginal_workers_female / pop_f
  s_w <- sc$mm(o$pct_females_working, "femp_working")
  s_m <- sc$mm(o$pct_females_main, "femp_main")
  s_g <- sc$mm(100 - o$pct_females_marginal, "femp_marginal_reversed")
  o$female_employment_index <- wmean(list(s_w, s_m, s_g), c(cfg$w_femp_working, cfg$w_femp_main, cfg$w_femp_marginal_reversed))

  # Female Education Index: block female literacy (Census) + district NFHS items, each min-max scaled
  o$female_literacy_pct <- 100 * b$literates_population_female / pop_f
  o$pct_female_6plus_ever_attended_school <- d$pct_female_6plus_ever_attended_school
  o$pct_women_10plus_years_schooling      <- d$pct_women_10plus_years_schooling
  o$female_education_index <- cfg$w_fed_literacy * sc$mm(o$female_literacy_pct, "fed_literacy") +
    cfg$w_fed_ever_attended * sc$mm(o$pct_female_6plus_ever_attended_school, "fed_ever_attended") +
    cfg$w_fed_10yr_schooling * sc$mm(o$pct_women_10plus_years_schooling, "fed_10yr_schooling")

  # Caste Diversity Index (block): 1 - (share SC^2 + share ST^2)
  o$pct_sc <- 100 * b$scheduled_castes_population_person / b$total_population_person
  o$pct_st <- 100 * b$scheduled_tribes_population_person / b$total_population_person
  o$caste_diversity_index <- 1 - ((o$pct_sc / 100)^2 + (o$pct_st / 100)^2)

  # Health Access (weighted min-max scaled facility counts)
  tot <- list(chc = h$total_community_health_centres, phc = h$total_primary_health_centres,
              phsc = h$total_primary_health_sub_centres, hosp = h$total_hospitals_allopathic,
              alt = h$total_hospitals_alternative_medicine, ngo = h$total_non_govt_health_facilities)
  if (cfg$har_basis == "per_10000_population") tot <- lapply(tot, function(x) 10000 * x / b$total_population_person)
  o$health_access_score <- cfg$w_har_chc * sc$mm(tot$chc, "har_chc") + cfg$w_har_phc * sc$mm(tot$phc, "har_phc") +
    cfg$w_har_phsc * sc$mm(tot$phsc, "har_phsc") + cfg$w_har_hosp_allopathic * sc$mm(tot$hosp, "har_hosp_allopathic") +
    cfg$w_har_hosp_alt_medicine * sc$mm(tot$alt, "har_hosp_alt_medicine") + cfg$w_har_non_govt * sc$mm(tot$ngo, "har_non_govt")

  # School Access Score (weighted % of villages with each school type; no further scaling)
  o$school_access_score <- cfg$w_sas_pre_primary * h$pct_villages_pre_primary_school_available +
    cfg$w_sas_primary * h$pct_villages_primary_school_available + cfg$w_sas_middle * h$pct_villages_middle_school_available +
    cfg$w_sas_secondary * h$pct_villages_secondary_school_available +
    cfg$w_sas_senior_secondary * h$pct_villages_senior_secondary_school_available +
    cfg$w_sas_private_senior_secondary * h$pct_villages_private_senior_secondary_school_available

  # Infrastructure & Accessibility Index (weighted min-max scaled components)
  o$infrastructure_index <- cfg$w_iai_connectivity * sc$mm(h[[conn]], "iai_connectivity") +
    cfg$w_iai_power * sc$mm(h[[paste0("avg_", DCHB_HOURS)]], "iai_power") +
    cfg$w_iai_pucca_road * sc$mm(h$pct_villages_pucca_road_available, "iai_pucca_road") +
    cfg$w_iai_safe_water * sc$mm(h$pct_villages_safe_drinking_water_available, "iai_safe_water")

  # District-level indicators copied to the blocks
  o$religious_fragmentation_index <- d$rfi
  o$sex_ratio_deviation <- d$srd
  o$amfri <- d$amfri

  # Standardisation (z-score) with risk direction: positive = higher risk
  o$z_health_access      <- cfg$dir_har   * sc$zs(o$health_access_score, "z_health_access")
  o$z_school_access      <- cfg$dir_sas   * sc$zs(o$school_access_score, "z_school_access")
  o$z_infrastructure     <- cfg$dir_iai   * sc$zs(o$infrastructure_index, "z_infrastructure")
  o$z_sex_ratio_dev      <- cfg$dir_srd   * sc$zs(o$sex_ratio_deviation, "z_sex_ratio_dev")
  o$z_amfri              <- cfg$dir_amfri * sc$zs(o$amfri, "z_amfri")
  o$z_female_employment  <- cfg$dir_femp  * sc$zs(o$female_employment_index, "z_female_employment")
  o$z_female_education   <- cfg$dir_fed   * sc$zs(o$female_education_index, "z_female_education")
  o$z_caste_diversity    <- cfg$dir_cdi   * sc$zs(o$caste_diversity_index, "z_caste_diversity")
  o$z_religious_fragmentation <- cfg$dir_rfi * sc$zs(o$religious_fragmentation_index, "z_religious_fragmentation")

  # Domains and final score
  o$ADI  <- wmean(list(o$z_health_access, o$z_school_access, o$z_infrastructure), c(cfg$w_adi_har, cfg$w_adi_sas, cfg$w_adi_iai))
  o$AVI  <- wmean(list(o$z_sex_ratio_dev, o$z_amfri, o$z_female_employment, o$z_female_education),
                  c(cfg$w_avi_srd, cfg$w_avi_amfri, cfg$w_avi_femp, cfg$w_avi_fed))
  o$SDS  <- cfg$w_sds_cdi * o$z_caste_diversity + cfg$w_sds_rfi * o$z_religious_fragmentation
  o$MHRS <- wmean(list(o$ADI, o$AVI, o$SDS), c(cfg$w_mhrs_adi, cfg$w_mhrs_avi, cfg$w_mhrs_sds))

  # Ranking and risk categories
  o$risk_rank <- rank(-o$MHRS, ties.method = "min")
  o$percentile_all_blocks <- pct_rank(o$MHRS, cfg$percentile_digits)
  o$percentile_within_district <- ave(o$MHRS, b$dkey, FUN = function(x) pct_rank(x, cfg$percentile_digits))
  o$risk_category <- tier(o$percentile_all_blocks, cfg$tercile_low_cut, cfg$tercile_high_cut)
  o$dkey <- b$dkey; o$bkey <- b$bkey
  o
}

# ------------------------------------------------------------------------------
# 7. VILLAGE STAGE: Social Diversity Index, join to block score, priority score, ranking
# ------------------------------------------------------------------------------
RELIGION_OTHER <- c("buddhist", "christian", "jain", "parsi", "sikh", "other_religions")

diversity_from_counts <- function(v) {
  z <- function(cn) { x <- if (cn %in% names(v)) v[[cn]] else rep(0, nrow(v)); x[is.na(x)] <- 0; x }
  rel_cats <- c("hindu", "muslim", RELIGION_OTHER)
  rel_sum  <- Reduce(`+`, lapply(rel_cats, z))
  rel_tot  <- if ("religion_total" %in% names(v)) ifelse(is.na(v$religion_total) | v$religion_total == 0, rel_sum, v$religion_total) else rel_sum
  cas_sum  <- z("sc") + z("st") + z("obc") + z("general")
  cas_tot  <- if ("caste_total" %in% names(v)) ifelse(is.na(v$caste_total) | v$caste_total == 0, cas_sum, v$caste_total) else cas_sum
  p_h <- 100 * z("hindu") / rel_tot; p_m <- 100 * z("muslim") / rel_tot
  p_o <- 100 * Reduce(`+`, lapply(RELIGION_OTHER, z)) / rel_tot
  p_sc <- 100 * (z("sc") + z("obc")) / cas_tot; p_st <- 100 * z("st") / cas_tot; p_gen <- 100 * z("general") / cas_tot
  rfi <- 1 - ((p_h / 100)^2 + (p_m / 100)^2 + (p_o / 100)^2)
  cdi <- 1 - ((p_sc / 100)^2 + (p_st / 100)^2 + (p_gen / 100)^2)
  rel_mat <- cbind(p_h, p_m, p_o); cas_mat <- cbind(p_sc, p_st, p_gen)
  dom_r <- c("Hindu", "Muslim", "Other")[apply(rel_mat, 1, function(r) if (anyNA(r)) NA_integer_ else which.max(r))]
  dom_c <- c("SC+OBC", "ST", "General")[apply(cas_mat, 1, function(r) if (anyNA(r)) NA_integer_ else which.max(r))]
  data.frame(pct_hindu = p_h, pct_muslim = p_m, pct_other_religions = p_o, pct_sc_obc = p_sc, pct_st = p_st, pct_general = p_gen,
             village_rfi = rfi, village_cdi = cdi, sdi = (rfi + cdi) / 2,
             dominant_religion = dom_r, dominant_caste = dom_c, stringsAsFactors = FALSE)
}

read_village_diversity <- function(path) {
  cnt <- c("hindu", "muslim", RELIGION_OTHER, "religion_total", "sc", "st", "obc", "general", "caste_total")
  v <- read_input_sheet(path, "Village_Diversity", text = c("district_name", "block_name", "village_name"),
                        opt_text = c("ngo_presence", "dominant_religion", "dominant_caste"),
                        opt_num = c("no_of_ngo_presence", "sdi", cnt), required = FALSE)
  v
}

compute_villages <- function(v, blk, cfg) {
  calc <- diversity_from_counts(v)
  given <- if ("sdi" %in% names(v)) v$sdi else rep(NA_real_, nrow(v))
  sdi <- ifelse(is.na(given), calc$sdi, given)
  dom_r <- if ("dominant_religion" %in% names(v)) ifelse(is_blank(v$dominant_religion), calc$dominant_religion, v$dominant_religion) else calc$dominant_religion
  dom_c <- if ("dominant_caste" %in% names(v)) ifelse(is_blank(v$dominant_caste), calc$dominant_caste, v$dominant_caste) else calc$dominant_caste
  if (anyNA(sdi) || anyNA(dom_r) || anyNA(dom_c))
    stop("Village_Diversity: SDI / dominant religion / dominant caste could not be determined for row(s): ",
         paste(head(which(is.na(sdi) | is.na(dom_r) | is.na(dom_c)) + 1, 10), collapse = ", "),
         ". Provide the religion and caste counts or the precomputed values.", call. = FALSE)
  m <- match(paste(key_of(v$district_name), key_of(v$block_name)), paste(blk$dkey, blk$bkey))
  if (anyNA(m)) {
    u <- unique(paste0(v$block_name[is.na(m)], " (", v$district_name[is.na(m)], ")"))
    log_msg(sum(is.na(m)), " village(s) are in block(s) with no block score and get NA scores: ", paste(head(u, 8), collapse = "; "), level = "WARN")
  }
  mh_pct <- if (cfg$village_percentile_basis == "within_district") blk$percentile_within_district[m] else blk$percentile_all_blocks[m]
  mh_cat <- ifelse(is.na(mh_pct), NA_character_, ifelse(mh_pct <= cfg$tercile_low_cut, "MH_Low",
                                                       ifelse(mh_pct <= cfg$tercile_high_cut, "MH_Medium", "MH_High")))
  thr <- unname(quantile(sdi, cfg$diversity_percentile, type = 7))            # PERCENTILE.INC
  div_cat <- ifelse(sdi >= thr, "Div_High", "Div_Low")
  min_rel <- dom_r %in% trimws(strsplit(cfg$minority_religions, ",")[[1]])
  min_cas <- dom_c %in% trimws(strsplit(cfg$marginalised_castes, ",")[[1]])
  min_cat <- ifelse(min_rel & min_cas, "Minority+Marginalized", ifelse(min_rel | min_cas, "Marginalized", "Dominant"))
  bonus   <- ifelse(min_cat == "Minority+Marginalized", cfg$bonus_minority_and_marginalised,
                    ifelse(min_cat == "Marginalized", cfg$bonus_either, 0))
  base    <- cfg$village_w_mh_percentile * mh_pct + cfg$village_w_sdi * sdi
  base_m  <- base + bonus
  strat   <- ifelse(is.na(mh_cat), NA_character_, paste(mh_cat, div_cat, min_cat, sep = "_"))
  rk <- rep(NA_integer_, length(strat))
  ok <- !is.na(strat)
  rk[ok] <- as.integer(ave(base_m[ok], strat[ok], FUN = function(x) rank(-x, ties.method = "min")))
  shortlist <- ifelse(is.na(rk), NA_character_, ifelse(rk <= cfg$shortlist_primary_max, "Primary",
                      ifelse(rk <= cfg$shortlist_reserve_max, "Reserve", ifelse(rk <= cfg$shortlist_fallback_max, "Fallback", "Out"))))
  out <- data.frame(
    `District Name` = v$district_name, `Block Name` = v$block_name, `Village Name` = v$village_name,
    `NGO Presence` = if ("ngo_presence" %in% names(v)) v$ngo_presence else NA_character_,
    `No_of_NGO_Presence` = if ("no_of_ngo_presence" %in% names(v)) v$no_of_ngo_presence else NA_real_,
    `Mental Health Risk Score (Block)` = blk$MHRS[m],
    `MH Percentile Within District (Block)` = blk$percentile_within_district[m],
    `MH Percentile All Blocks (Block)` = blk$percentile_all_blocks[m],
    `Block_Rank` = blk$risk_rank[m],
    `Social Diversity Index` = sdi,
    `SDI_Rank_Within_District` = as.integer(ave(-sdi, key_of(v$district_name), FUN = function(x) rank(x, ties.method = "min"))),
    `Dominant Religion` = dom_r, `Dominant Caste` = dom_c,
    `MH_Category` = mh_cat, `Diversity Category` = div_cat,
    `Minority_Religion_Dominant` = as.integer(min_rel), `Minority_Caste_Dominant` = as.integer(min_cas),
    `Minority_Category` = min_cat, `Selection_Rationale` = strat, `Minority_Bonus` = bonus,
    `Base_Priority_Score` = base, `Base_Minority` = base_m, `Rank_Within_Category` = rk,
    `Top20_Present` = as.integer(rk <= cfg$shortlist_fallback_max), `Category_Shortlist` = shortlist,
    check.names = FALSE, stringsAsFactors = FALSE)
  out$`Base_Priority_Score`[is.na(mh_pct)] <- NA_real_
  out$`Base_Minority`[is.na(mh_pct)] <- NA_real_
  out$`Minority_Bonus`[is.na(mh_pct)] <- NA_real_
  out
}

# ------------------------------------------------------------------------------
# 8. OPTIONAL SELF-CHECK: compare against 'Reference_Block_Outputs' / 'Reference_Village_Outputs' if present
# ------------------------------------------------------------------------------
replication_check <- function(path, blocks, villages, cfg) {
  rows <- list()
  add <- function(item, n, mx, nbad) rows[[length(rows) + 1]] <<- data.frame(
    item = item, n_compared = n, max_abs_difference = mx, n_outside_tolerance = nbad,
    status = ifelse(nbad == 0, "MATCH", "DIFFERS"), stringsAsFactors = FALSE)
  tol <- cfg$replication_tolerance
  rb <- read_input_sheet(path, "Reference_Block_Outputs", text = c("district_name", "block_name"),
                         opt_num = c("adi", "avi", "sds", "mhrs", "risk_rank", "percentile_all_blocks", "percentile_within_district"),
                         opt_text = "risk_category", required = FALSE)
  if (!is.null(rb)) {
    for (cn in c("ADI", "AVI", "SDS", "MHRS")) if (tolower(cn) %in% names(rb)) names(rb)[names(rb) == tolower(cn)] <- cn
    i <- match(paste(key_of(rb$district_name), key_of(rb$block_name)), paste(blocks$dkey, blocks$bkey))
    for (cn in intersect(c("ADI", "AVI", "SDS", "MHRS", "risk_rank", "percentile_all_blocks", "percentile_within_district"), names(rb))) {
      dlt <- abs(rb[[cn]] - blocks[[cn]][i]); add(paste0("Block: ", cn), sum(!is.na(dlt)), max(dlt, na.rm = TRUE), sum(dlt > tol, na.rm = TRUE) + sum(is.na(dlt)))
    }
    if ("risk_category" %in% names(rb)) add("Block: risk_category (exact)", nrow(rb), 0, sum(rb$risk_category != blocks$risk_category[i], na.rm = TRUE))
  }
  if (!is.null(villages)) {
    rv <- read_input_sheet(path, "Reference_Village_Outputs",
                           opt_num = c("base_priority_score", "base_minority", "rank_within_category", "minority_bonus"),
                           opt_text = c("mh_category", "diversity_category", "minority_category", "selection_rationale",
                                        "category_shortlist"), required = FALSE)
    if (!is.null(rv)) {
      if (nrow(rv) != nrow(villages))
        stop("Reference_Village_Outputs must have the same rows, in the same order, as Village_Diversity.", call. = FALSE)
      vn <- villages; names(vn) <- norm_names(names(vn))
      scored <- !is.na(vn$mh_category)                      # villages without a block score are not compared
      for (cn in intersect(c("base_priority_score", "base_minority", "rank_within_category", "minority_bonus"), names(rv))) {
        dlt <- abs(rv[[cn]] - vn[[cn]])[scored]
        add(paste0("Village: ", cn), sum(!is.na(dlt)), max(dlt, na.rm = TRUE), sum(dlt > tol, na.rm = TRUE) + sum(is.na(dlt)))
      }
      for (cn in intersect(c("mh_category", "diversity_category", "minority_category", "selection_rationale", "category_shortlist"), names(rv)))
        add(paste0("Village: ", cn, " (exact)"), sum(scored), 0, sum(rv[[cn]][scored] != vn[[cn]][scored], na.rm = TRUE))
    }
  }
  if (length(rows)) do.call(rbind, rows) else NULL
}

# ------------------------------------------------------------------------------
# 9. OUTPUT WORKBOOK
# ------------------------------------------------------------------------------
write_output <- function(file, sheets, tier_cols = list()) {
  wb <- createWorkbook()
  hs <- createStyle(fontColour = "#FFFFFF", fgFill = "#1F4E78", textDecoration = "bold", halign = "center",
                    valign = "center", wrapText = TRUE, border = "TopBottomLeftRight")
  for (nm in names(sheets)) {
    df <- sheets[[nm]]
    addWorksheet(wb, nm)
    writeData(wb, nm, df, headerStyle = hs, keepNA = FALSE)
    if (nrow(df) > 0) {
      for (cc in which(vapply(df, is.numeric, logical(1)))) {
        whole <- all(df[[cc]] == round(df[[cc]]), na.rm = TRUE)
        fmt <- if (grepl("percentile", names(df)[cc], ignore.case = TRUE)) "0.0" else if (whole) "#,##0" else "0.000"
        addStyle(wb, nm, createStyle(numFmt = fmt), rows = 2:(nrow(df) + 1), cols = cc, stack = TRUE)
      }
    }
    setColWidths(wb, nm, cols = seq_len(ncol(df)), widths = pmin(pmax(nchar(names(df)) * 0.9 + 4, 12), 40))
    freezePane(wb, nm, firstRow = TRUE)
    if (nm %in% names(tier_cols) && nrow(df) > 0) {
      cc <- match(tier_cols[[nm]], names(df))
      for (lab in list(c("High", "#F8CBAD"), c("Medium", "#FFE699"), c("Low", "#C6E0B4")))
        conditionalFormatting(wb, nm, cols = cc, rows = 2:(nrow(df) + 1), rule = lab[1], type = "contains",
                              style = createStyle(bgFill = lab[2]))
    }
  }
  saveWorkbook(wb, file, overwrite = TRUE)
  repair_xlsx(file)
  invisible(file)
}

# Some openxlsx builds leave references to (empty) drawing parts in the package without writing the parts.
# Remove those dangling references so that the workbook is structurally clean and opens without a repair prompt.
repair_xlsx <- function(file) {
  tryCatch({
    file <- normalizePath(file, mustWork = TRUE)
    td <- tempfile("xlsx_fix"); dir.create(td)
    utils::unzip(file, exdir = td)
    changed <- FALSE
    for (r in list.files(file.path(td, "xl", "worksheets", "_rels"), full.names = TRUE)) {
      x <- paste(readLines(r, warn = FALSE), collapse = "")
      for (rel in regmatches(x, gregexpr("<Relationship [^>]*/>", x))[[1]]) {
        target <- sub('.*Target="([^"]+)".*', "\\1", rel)
        if (grepl("^\\.\\./drawings/", target) && !file.exists(file.path(dirname(dirname(r)), target))) {
          x <- sub(rel, "", x, fixed = TRUE); changed <- TRUE
        }
      }
      writeLines(x, r, useBytes = TRUE)
    }
    ct <- file.path(td, "[Content_Types].xml")
    y <- paste(readLines(ct, warn = FALSE), collapse = "")
    for (ov in regmatches(y, gregexpr("<Override [^>]*/>", y))[[1]]) {
      part <- sub('.*PartName="([^"]+)".*', "\\1", ov)
      if (grepl("^/xl/drawings/", part) && !file.exists(file.path(td, sub("^/", "", part)))) {
        y <- sub(ov, "", y, fixed = TRUE); changed <- TRUE
      }
    }
    writeLines(y, ct, useBytes = TRUE)
    if (changed) {
      files <- list.files(td, recursive = TRUE, all.files = TRUE)
      files <- c("[Content_Types].xml", setdiff(files, "[Content_Types].xml"))
      unlink(file)
      zip::zipr(file, files, root = td, mode = "mirror")
    }
  }, error = function(e) log_msg("Could not post-process the output workbook (", conditionMessage(e), ")", level = "WARN"))
}

# ------------------------------------------------------------------------------
# 10. MAIN
# ------------------------------------------------------------------------------
run_mhrs <- function(input = NULL, output = NULL) {
  LOG$lines <- character(0)
  if (is.null(input)) {
    if (!interactive()) stop("Usage: Rscript mhrs_compute.R <input.xlsx> [<output.xlsx>]", call. = FALSE)
    message("Choose the MHRS input workbook (.xlsx)")
    input <- file.choose()
  }
  if (!file.exists(input)) stop("Input file not found: ", input, call. = FALSE)
  if (is.null(output)) output <- sub("\\.xlsx$", "_MHRS_output.xlsx", input, ignore.case = TRUE)
  log_msg("Input : ", normalizePath(input)); log_msg("Output: ", output)
  cfg <- read_config(input)
  fixed <- if (cfg$reference_mode == "fixed") read_fixed_stats(input) else NULL
  if (cfg$reference_mode == "fixed") log_msg("reference_mode = fixed: scaling and z-scores use the 'Reference_Stats' sheet.")
  sc <- make_scalers(cfg, fixed)

  census <- read_input_sheet(input, "Census_Blocks", text = c("district_name", "block_name"),
    num = c("total_population_person", "total_population_female", "scheduled_castes_population_person",
            "scheduled_tribes_population_person", "literates_population_female", "total_workers_female",
            "main_workers_female", "marginal_workers_female"), opt_num = c("district_code", "block_code"))
  nfhs <- read_input_sheet(input, "NFHS_Districts", text = "district_name",
    num = c("sex_ratio_total", "sex_ratio_at_birth", "pct_women_20_24_married_before_18",
            "pct_women_15_19_mothers_or_pregnant", "pct_female_6plus_ever_attended_school", "pct_women_10plus_years_schooling"))
  rel <- read_input_sheet(input, "Religion_Districts", text = "district_name",
    num = c("females_total", "females_hindu", "females_muslim", "females_christian", "females_buddhist", "females_not_stated"))
  cm <- read_input_sheet(input, "Census_Marriage", text = c("district_name", "area_type"),
    num = c("pct_below15_ever_married", "pct_15_19_ever_married"))
  dchb <- read_dchb(input, cfg)
  log_msg("Blocks read: ", nrow(census), " in ", length(unique(key_of(census$district_name))), " district(s).")

  dist   <- compute_districts(nfhs, rel, cm, cfg, sc)
  blocks <- compute_blocks(census, dchb, dist, cfg, sc)
  log_msg("Block scores computed for ", nrow(blocks), " blocks.")

  vdiv <- read_village_diversity(input)
  villages <- NULL
  if (!is.null(vdiv)) {
    villages <- compute_villages(vdiv, blocks, cfg)
    log_msg("Village scores computed for ", nrow(villages), " villages.")
  } else log_msg("No 'Village_Diversity' sheet - village stage skipped.")

  chk <- replication_check(input, blocks, villages, cfg)
  if (!is.null(chk)) log_msg("Self-check against Reference_* sheets: ", sum(chk$status == "MATCH"), " of ", nrow(chk), " items match.")

  summary_tbl <- data.frame(
    `District Name` = blocks$district_name, `Block Name` = blocks$block_name,
    `Access Deprivation Index (ADI)` = blocks$ADI, `Adolescent Vulnerability Index (AVI)` = blocks$AVI,
    `Social Diversity & Structure (SDS)` = blocks$SDS, `Mental Health Risk Score (MHRS)` = blocks$MHRS,
    `Risk Rank (1 = highest)` = blocks$risk_rank, `Percentile (All Blocks)` = blocks$percentile_all_blocks,
    `Percentile (Within District)` = blocks$percentile_within_district, `Risk Category` = blocks$risk_category,
    check.names = FALSE, stringsAsFactors = FALSE)
  summary_tbl <- summary_tbl[order(summary_tbl$`Risk Rank (1 = highest)`), ]
  calc <- blocks[, setdiff(names(blocks), c("dkey", "bkey"))]
  dist_out <- dist[, setdiff(names(dist), "dkey")]
  params <- data.frame(parameter = names(cfg), value = vapply(cfg, function(x) paste(x, collapse = ","), ""), stringsAsFactors = FALSE)
  log_msg("Finished.")
  log_df <- data.frame(run_info = c(LOG$lines, paste0("[INFO] R ", getRversion(), "; run at ", format(Sys.time(), "%Y-%m-%d %H:%M:%S"))), stringsAsFactors = FALSE)
  dchb_out <- dchb[paste(dchb$dkey, dchb$bkey) %in% paste(blocks$dkey, blocks$bkey), setdiff(names(dchb), c("dkey", "bkey")), drop = FALSE]
  sheets <- list(Block_Scores = summary_tbl, Block_Calculation = calc, District_Indicators = dist_out,
                 DCHB_Block_Aggregates = dchb_out)
  if (!is.null(villages)) sheets$Village_Scores <- villages
  sheets$Reference_Stats <- sc$table()
  sheets$Parameters_Used <- params
  if (!is.null(chk)) sheets$Replication_Check <- chk
  sheets$Run_Log <- log_df
  write_output(output, sheets, tier_cols = list(Block_Scores = "Risk Category", Block_Calculation = "risk_category"))
  cat("Output written to: ", output, "\n", sep = "")
  invisible(list(blocks = blocks, districts = dist, villages = villages, config = cfg, stats = sc$table(), check = chk, output = output))
}

if (!interactive()) {
  if (sys.nframe() == 0L) {                     # executed with Rscript (not when source()d by another script)
    args <- commandArgs(trailingOnly = TRUE)
    if (length(args) < 1) stop("Usage: Rscript mhrs_compute.R <input.xlsx> [<output.xlsx>]", call. = FALSE)
    invisible(run_mhrs(args[1], if (length(args) >= 2) args[2] else NULL))
  }
} else if (isTRUE(getOption("mhrs.autorun", TRUE))) {   # RStudio / R console: Source or Run -> file chooser -> scores
  invisible(run_mhrs())
}
