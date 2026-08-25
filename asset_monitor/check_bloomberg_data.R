# check_bloomberg_data.R
# Validates a populated bloomberg_data.xlsx BEFORE the model runs, so a mistyped
# ticker, an empty column, a date-grid mismatch, or too-short history is caught
# with a clear message instead of a confusing downstream error.
#
# Usage:
#   source("asset_monitor/check_bloomberg_data.R")
#   rep <- check_bloomberg_data("asset_monitor/bloomberg_data.xlsx")
#   if (!rep$ok) stop("Fix the data issues above before running the model.")

check_bloomberg_data <- function(path = "asset_monitor/bloomberg_data.xlsx",
                                 min_weeks = 420, max_missing_frac = 0.50,
                                 verbose = TRUE) {
  if (!requireNamespace("readxl", quietly = TRUE)) install.packages("readxl")

  price_cols <- c("equity_world","equity_europe","equity_em","bonds_eur","bonds_em_hc",
                  "bonds_em_lc","convertibles","commodities","gold","usd_currency",
                  "bonds_global","bonds_short_dur","bonds_med_dur","bonds_long_dur",
                  "credit_ig","credit_hy")
  # REQUIRED core macro (market/rates data that any terminal has)
  macro_straight <- c("yield_2s10s","us_10y_yield",
                      "us_real_yield","us_breakeven","eur_10y_yield","eur_breakeven","dxy",
                      "vix","us_ig_oas","us_hy_oas","eur_ig_oas","em_hc_oas","ecb_dir")
  # OPTIONAL / licence-gated / thin macro — warn (never fail) if missing or empty
  macro_optional <- c("pmi_global","pmi_europe","pmi_em","m2_yoy",
                      "fwd_pe_world","fwd_pe_eu","fwd_pe_em",
                      "eps_ntm_world","eps_ntm_eu","eps_ntm_em","gbiem_yield",
                      "bcom_roll_tr","wti_front","wti_second",
                      "cftc_gold","cftc_comm","put_call_ratio","conv_impl_vol")

  fails <- character(0); warns <- character(0); rows <- list()
  add_fail <- function(m) fails[[length(fails) + 1]] <<- m
  add_warn <- function(m) warns[[length(warns) + 1]] <<- m

  # Read WITHOUT sorting so undated rows keep their original position; this is
  # what lets us detect the mixed-periodicity overflow (values on undated rows).
  rd <- function(sheet, sort = TRUE) {
    df <- as.data.frame(readxl::read_excel(path, sheet = sheet))
    if (!"Date" %in% names(df)) { add_fail(sprintf("[%s] no 'Date' column", sheet)); return(NULL) }
    df$Date <- suppressWarnings(as.Date(df$Date))
    if (sort) df[order(df$Date), ] else df
  }

  sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) NULL)
  if (is.null(sheets)) { message("Cannot open file: ", path); return(list(ok = FALSE)) }
  for (req in c("prices", "macro")) if (!req %in% sheets) add_fail(sprintf("missing sheet '%s'", req))
  if (length(fails)) { report(fails, warns, NULL, verbose); return(list(ok = FALSE, issues = fails)) }

  prices <- rd("prices"); macro <- rd("macro")

  # ── MIXED-PERIODICITY GUARD ────────────────────────────────────────────────
  # A sheet whose date-bearing column is monthly while other columns are weekly
  # writes value rows past the last dated row. Those undated weekly values lose
  # their dates and, unless the loader reconstructs them, get shifted ~10 years
  # early -> recent-year signals silently go flat. Detect it here so it is fixed
  # (regenerate the template with all columns Per=W) rather than trusted blindly.
  check_periodicity <- function(sheet_name) {
    raw <- rd(sheet_name, sort = FALSE)
    if (is.null(raw)) return(invisible())
    n_dated <- sum(!is.na(raw$Date))
    n_rows  <- nrow(raw)
    val_cols <- setdiff(names(raw), "Date")
    # how many columns have finite values on rows beyond the last dated row?
    last_dated <- suppressWarnings(max(which(!is.na(raw$Date))))
    if (!is.finite(last_dated)) return(invisible())
    overflow_cols <- val_cols[vapply(val_cols, function(c) {
      v <- suppressWarnings(as.numeric(raw[[c]]))
      last_dated < n_rows && any(is.finite(v[(last_dated + 1):n_rows]))
    }, logical(1))]
    if (length(overflow_cols) > 0) {
      add_warn(sprintf(paste0("[%s] MIXED PERIODICITY: %d/%d rows are dated but %d column(s) have data on ",
                              "UNDATED rows (e.g. %s). The date column looks monthly while other columns ",
                              "are weekly. The loader will right-align the weekly block to recover it, but ",
                              "you should regenerate the template with EVERY column at Per=W."),
                       sheet_name, n_dated, n_rows, length(overflow_cols),
                       paste(utils::head(overflow_cols, 4), collapse = ", ")))
    }
    invisible()
  }
  check_periodicity("macro")

  # ── history length + date sanity ────────────────────────────────────────────
  n <- nrow(prices)
  if (n < min_weeks) add_fail(sprintf("only %d weeks of prices (need >= %d)", n, min_weeks))
  if (any(is.na(prices$Date))) add_fail("prices Date column has unparseable dates")

  # header alias (some workbooks use pmi_world for the global PMI)
  if ("pmi_world" %in% names(macro) && !"pmi_global" %in% names(macro))
    names(macro)[names(macro) == "pmi_world"] <- "pmi_global"

  # ── date alignment: warn (not fail) — loader row-aligns if grids differ ─────
  miss_dates <- sum(!prices$Date %in% macro$Date)
  if (miss_dates > 0)
    add_warn(sprintf("%d/%d price dates not found in macro sheet — loader will align macro by date/position (OK if macro is a same-range weekly pull)", miss_dates, nrow(prices)))

  # ── per-column coverage check ──────────────────────────────────────────────
  check_cols <- function(df, cols, sheet) {
    for (c in cols) {
      if (!c %in% names(df)) { add_fail(sprintf("[%s] missing column '%s'", sheet, c)); next }
      v <- suppressWarnings(as.numeric(df[[c]]))
      nfin <- sum(is.finite(v)); frac_missing <- 1 - nfin / length(v)
      status <- if (nfin == 0) "FAIL" else if (frac_missing > max_missing_frac) "WARN" else "ok"
      if (status == "FAIL") add_fail(sprintf("[%s] column '%s' is empty / all #N/A (bad ticker?)", sheet, c))
      if (status == "WARN") add_warn(sprintf("[%s] column '%s' %.0f%% missing", sheet, c, 100 * frac_missing))
      rng <- if (nfin > 0) range(df$Date[is.finite(v)], na.rm = TRUE) else c(NA, NA)
      rows[[length(rows) + 1]] <<- data.frame(Sheet = sheet, Column = c, N_finite = nfin,
        Pct_missing = round(100 * frac_missing, 1),
        First = as.character(rng[1]), Last = as.character(rng[2]),
        Status = status, stringsAsFactors = FALSE)
    }
  }

  check_cols(prices, price_cols, "prices")
  check_cols(macro, macro_straight, "macro")

  # ── OPTIONAL sentiment columns: warn (never fail) if missing/empty ──────────
  sentiment_cols <- c("aaii_bull","aaii_bear","naaim_exposure","skew","spx_cftc",
                      "vstoxx","estoxx_put_call","em_fx_rr","em_equity_flows",
                      "usd_cftc","move_index","ust_cftc","lqd_flows","hyg_flows","gold_etf")
  for (c in sentiment_cols) {
    if (!c %in% names(macro)) { add_warn(sprintf("[macro] optional sentiment '%s' absent (bucket will skip it)", c)); next }
    v <- suppressWarnings(as.numeric(macro[[c]])); nfin <- sum(is.finite(v))
    if (nfin == 0) add_warn(sprintf("[macro] optional sentiment '%s' empty/#N/A (will be dropped)", c))
  }
  # OPTIONAL / licence-gated macro: warn (never fail) if missing/empty
  for (c in macro_optional) {
    if (!c %in% names(macro)) { add_warn(sprintf("[macro] optional '%s' absent (signal will contribute nothing)", c)); next }
    v <- suppressWarnings(as.numeric(macro[[c]])); if (sum(is.finite(v)) == 0)
      add_warn(sprintf("[macro] optional '%s' empty/#N/A (will be dropped)", c))
  }
  # roll yield needs bcom_roll_tr OR the wti pair — note if neither is present
  if (!any(c("bcom_roll_tr") %in% names(macro)) &&
      !all(c("wti_front","wti_second") %in% names(macro)))
    add_warn("[macro] no roll-yield source (bcom_roll_tr or CL1/CL2) — roll_yield will be NA")

  tbl <- if (length(rows)) do.call(rbind, rows) else NULL
  ok  <- length(fails) == 0
  report(fails, warns, tbl, verbose)
  if (verbose) cat(if (ok) "\n  RESULT: PASS — safe to run the model.\n\n"
                   else   "\n  RESULT: FAIL — fix the items above before running.\n\n")
  invisible(list(ok = ok, fails = fails, warns = warns, table = tbl))
}

report <- function(fails, warns, tbl, verbose) {
  if (!verbose) return(invisible())
  W <- 92; cat("\n", strrep("=", W), "\n", sep = "")
  cat("  BLOOMBERG DATA CHECK\n"); cat(strrep("=", W), "\n", sep = "")
  if (!is.null(tbl)) {
    bad <- tbl[tbl$Status != "ok", , drop = FALSE]
    show <- if (nrow(bad) > 0) bad else utils::head(tbl, 6)
    cat(sprintf("  %-7s %-16s %8s %9s  %-11s %-11s %s\n",
                "Sheet","Column","N_finite","%missing","First","Last","Status"))
    for (i in seq_len(nrow(show))) { r <- show[i, ]
      cat(sprintf("  %-7s %-16s %8d %8.1f%%  %-11s %-11s %s\n",
                  r$Sheet, r$Column, r$N_finite, r$Pct_missing, r$First, r$Last, r$Status)) }
    if (nrow(bad) == 0) cat("  (all columns ok — showing first rows)\n")
  }
  for (m in warns) cat("  WARN:", m, "\n")
  for (m in fails) cat("  FAIL:", m, "\n")
}
