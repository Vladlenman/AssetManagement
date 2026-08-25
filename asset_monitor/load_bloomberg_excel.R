# load_bloomberg_excel.R
# Reads bloomberg_data.xlsx (populated by the Bloomberg Excel add-in) and
# returns the same list structure as simulate_market_data().
#
# The spreadsheet holds only clean BDH pulls (levels/yields). The few DERIVED
# model inputs are computed here in R, so the Excel file needs no fragile
# in-cell array formulas:
#     eps_rev_world/eu/em  <- 4-week % change of BEST_EPS_NXT12MO (raw: eps_ntm_*)
#     em_carry_diff        <- gbiem_yield  -  us_10y_yield
#     roll_yield           <- 4-week % change of BCMRYTOT level (raw: bcom_roll_tr)
#
# readxl is loaded lazily so sourcing never fails without it.
#
# ── MIXED-PERIODICITY SHEETS (important) ──────────────────────────────────────
# A macro sheet whose FIRST (date-bearing) column is a MONTHLY BDH pull, while
# the other columns are WEEKLY pulls, is pathological: the weekly columns write
# more rows than the monthly date column carries, so their tail rows have NO
# date next to them. Left unhandled, those undated weekly values get row-aligned
# from the top of the weekly grid and land ~10 years too early, leaving every
# recent week NA — which silently flattens every macro/sentiment-driven signal.
# This loader now DETECTS that case and reconstructs the weekly dates by right-
# aligning the weekly macro block to the tail of the weekly price grid (both are
# "as of now", so their ENDS coincide even though macro starts later). The clean
# fix is still to regenerate the template with EVERY macro column at Per=W (see
# create_bloomberg_template.R); this recovery just stops a mis-built sheet from
# corrupting the run.

load_bloomberg_excel <- function(path = "asset_monitor/bloomberg_data.xlsx") {
  if (!requireNamespace("readxl", quietly = TRUE)) install.packages("readxl")
  cat(sprintf("Loading Bloomberg data from: %s\n", path))

  price_cols <- c("equity_world", "equity_europe", "equity_em",
                  "bonds_eur", "bonds_em_hc", "bonds_em_lc",
                  "convertibles", "commodities", "gold", "usd_currency",
                  "bonds_global", "bonds_short_dur", "bonds_med_dur",
                  "bonds_long_dur", "credit_ig", "credit_hy")

  # Columns the model ultimately needs from the macro sheet
  macro_needed <- c("pmi_global","pmi_europe","pmi_em",
                    "yield_2s10s","us_10y_yield","us_real_yield","us_breakeven",
                    "eur_10y_yield","eur_breakeven","dxy","vix",
                    "us_ig_oas","us_hy_oas","eur_ig_oas","em_hc_oas",
                    "m2_yoy","ecb_dir","fwd_pe_world","fwd_pe_eu","fwd_pe_em",
                    "cftc_gold","cftc_comm","put_call_ratio","conv_impl_vol",
                    "eps_rev_world","eps_rev_eu","eps_rev_em",
                    "em_carry_diff","roll_yield")

  rd <- function(sheet) {
    df <- as.data.frame(readxl::read_excel(path, sheet = sheet))
    df$Date <- as.Date(df$Date)
    df[order(df$Date), ]
  }
  # Read WITHOUT reordering by Date: for a mixed-periodicity sheet the undated
  # weekly rows must keep their original top-to-bottom (chronological) position.
  rd_raw <- function(sheet) {
    df <- as.data.frame(readxl::read_excel(path, sheet = sheet))
    df$Date <- suppressWarnings(as.Date(df$Date))
    df
  }

  # helper: k-period % change (for revisions / roll yield)
  pct_change_k <- function(x, k = 4) {
    x <- as.numeric(x); n <- length(x)
    out <- rep(NA_real_, n)
    if (n > k) {
      prev <- x[seq_len(n - k)]
      out[(k + 1):n] <- (x[(k + 1):n] - prev) / pmax(abs(prev), 1e-8)
    }
    out
  }
  ffill <- function(col) { for (i in seq_along(col)) if (is.na(col[i]) && i > 1) col[i] <- col[i-1]; col }

  # ── Prices ─────────────────────────────────────────────────────────────────
  raw_prices <- rd("prices")
  miss <- setdiff(price_cols, names(raw_prices))
  if (length(miss) > 0) stop("prices sheet missing columns: ", paste(miss, collapse = ", "))
  dates  <- raw_prices$Date
  prices <- raw_prices[, price_cols, drop = FALSE]
  # Force every price column numeric. Bloomberg exports non-trading / no-history
  # cells as strings ("#N/A N/A", "", "#N/A Requesting Data...") which make the
  # whole column character; the technical-signal (TTR) functions then fail with
  # "non-numeric argument to binary operator". Coercion turns those into NA, so
  # an index that only has prices from (say) 2010 simply starts at its first
  # real observation — the earliest available date for that asset.
  prices <- as.data.frame(lapply(prices, function(x) suppressWarnings(as.numeric(x))),
                          stringsAsFactors = FALSE)
  names(prices) <- price_cols
  .cov <- vapply(prices, function(x) mean(is.finite(x)), numeric(1))
  if (any(.cov < 1)) {
    for (nm in names(.cov)[.cov < 1]) {
      first <- which(is.finite(prices[[nm]]))[1]
      cat(sprintf("  price '%s': starts %s (%.0f%% of weeks have data)\n",
                  nm, if (is.na(first)) "NEVER" else format(dates[first]), 100 * .cov[[nm]]))
    }
  }

  # ── Macro: align to the weekly price grid (handles mixed weekly/monthly) ──
  # Read RAW (unsorted) so undated weekly rows keep their chronological position;
  # a well-formed weekly sheet is already in order, so this is harmless there.
  m <- rd_raw("macro")
  if ("pmi_world" %in% names(m) && !"pmi_global" %in% names(m))   # header alias
    names(m)[names(m) == "pmi_world"] <- "pmi_global"

  # Last-observation-carried-forward merge of a lower-frequency (e.g. monthly)
  # series onto the weekly grid, using its REAL release dates — a proper step
  # function, not the "first value copied everywhere" that Bloomberg Fill=P gives.
  locf_merge <- function(src_dates, src_vals, tgt_dates) {
    keep <- is.finite(src_vals) & !is.na(src_dates)
    if (!any(keep)) return(rep(NA_real_, length(tgt_dates)))
    o  <- order(src_dates[keep]); sd <- src_dates[keep][o]; sv <- src_vals[keep][o]
    idx <- findInterval(as.numeric(tgt_dates), as.numeric(sd))    # last release <= week
    out <- rep(NA_real_, length(tgt_dates)); ok <- idx >= 1
    out[ok] <- sv[idx[ok]]; out
  }

  md <- as.Date(m$Date)
  if (sum(dates %in% md, na.rm = TRUE) >= 0.8 * length(dates)) {
    # ── Clean weekly sheet: macro Date grid matches the price grid ────────────
    m <- m[match(dates, md), , drop = FALSE]; m$Date <- dates
    m <- as.data.frame(lapply(m, ffill))
  } else {
    # ── Mixed-frequency / mis-built sheet ─────────────────────────────────────
    # The macro Date column is lower-frequency (e.g. monthly) so it does not line
    # up with the weekly price grid. Two kinds of column can appear:
    #   (A) genuinely low-frequency series (all finite values sit on real dated
    #       rows, e.g. monthly PMI)  -> forward-fill by their REAL dates.
    #   (B) weekly series whose dates were lost because they overflow past the
    #       monthly date column (finite values on UNDATED rows) -> reconstruct
    #       weekly dates by RIGHT-aligning the weekly block to the price-grid
    #       tail, then forward-fill by date. Right-align (not left) because both
    #       macro and prices are "as of now": their most-recent rows coincide,
    #       while macro history typically starts later than prices.
    cat(sprintf("  NOTE: macro Date grid is not weekly (%d/%d matched) — merging mixed frequencies.\n",
                sum(dates %in% md, na.rm = TRUE), length(dates)))
    date_rows   <- which(!is.na(md))
    n_date_rows <- if (length(date_rows)) max(date_rows) else 0L

    # Last macro row that carries ANY value defines the weekly block length.
    vcols   <- setdiff(names(m), "Date")
    val_mat <- suppressWarnings(vapply(vcols, function(c) as.numeric(m[[c]]),
                                       numeric(nrow(m))))
    if (is.null(dim(val_mat))) val_mat <- matrix(val_mat, nrow = nrow(m))
    row_has_data <- apply(val_mat, 1, function(r) any(is.finite(r)))
    n_blk <- if (any(row_has_data)) max(which(row_has_data)) else nrow(m)

    # Right-align the weekly block to the tail of the price grid.
    recon_dates <- utils::tail(dates, min(n_blk, length(dates)))
    if (length(recon_dates) < n_blk)                        # macro has more rows
      recon_dates <- c(rep(as.Date(NA), n_blk - length(recon_dates)), recon_dates)

    overflow_any <- n_blk > n_date_rows
    if (overflow_any)
      cat(sprintf(paste0("  WARNING: macro sheet mixes weekly series with a MONTHLY date column",
                         " (%d data rows, only %d dated). Weekly columns lost their dates;\n",
                         "           reconstructing them by right-aligning to the price grid",
                         " (last obs -> %s). Regenerate the template with every macro column at\n",
                         "           Per=W to remove this ambiguity (see create_bloomberg_template.R).\n"),
                  n_blk, n_date_rows, format(dates[length(dates)])))

    out <- data.frame(Date = dates)
    for (c in vcols) {
      v  <- suppressWarnings(as.numeric(m[[c]])); cf <- which(is.finite(v))
      if (length(cf) == 0) { out[[c]] <- NA_real_; next }
      if (max(cf) <= n_date_rows && all(cf %in% date_rows)) {
        # (A) genuinely low-frequency: every value has a real date -> merge by date.
        out[[c]] <- locf_merge(md[cf], v[cf], dates)
        fin <- which(is.finite(out[[c]]))
        cat(sprintf("  dated series '%s' forward-filled by date (%d releases, covers %s to %s).\n",
                    c, length(cf),
                    if (length(fin)) format(dates[min(fin)]) else "NA",
                    if (length(fin)) format(dates[max(fin)]) else "NA"))
      } else {
        # (B) weekly series with lost dates: reconstruct by right-aligned position.
        cf2 <- cf[cf <= length(recon_dates)]
        out[[c]] <- locf_merge(recon_dates[cf2], v[cf2], dates)
        fin <- which(is.finite(out[[c]]))
        cat(sprintf("  weekly series '%s' realigned to price-grid tail (%d obs, covers %s to %s).\n",
                    c, length(cf2),
                    if (length(fin)) format(dates[min(fin)]) else "NA",
                    if (length(fin)) format(dates[max(fin)]) else "NA"))
      }
    }
    m <- as.data.frame(lapply(out, ffill))
  }

  # column present AND has some finite data
  have <- function(col) col %in% names(m) && any(is.finite(suppressWarnings(as.numeric(m[[col]]))))

  # ── Derive model inputs; each is non-fatal (falls back to NA -> contributes
  #    nothing and gets zero backtest weight) so a licence-gated source never
  #    blocks the run. ──
  if (!"eps_rev_world" %in% names(m)) m$eps_rev_world <- if (have("eps_ntm_world")) pct_change_k(as.numeric(m$eps_ntm_world)) else NA_real_
  if (!"eps_rev_eu"    %in% names(m)) m$eps_rev_eu    <- if (have("eps_ntm_eu"))    pct_change_k(as.numeric(m$eps_ntm_eu))    else NA_real_
  if (!"eps_rev_em"    %in% names(m)) m$eps_rev_em    <- if (have("eps_ntm_em"))    pct_change_k(as.numeric(m$eps_ntm_em))    else NA_real_
  if (!"em_carry_diff" %in% names(m)) m$em_carry_diff <- if (have("gbiem_yield"))   as.numeric(m$gbiem_yield) - as.numeric(m$us_10y_yield) else NA_real_
  if (!"roll_yield"    %in% names(m)) {
    m$roll_yield <-
      if (have("bcom_roll_tr")) pct_change_k(as.numeric(m$bcom_roll_tr))
      else if (have("wti_front") && have("wti_second"))
        (as.numeric(m$wti_front) - as.numeric(m$wti_second)) / pmax(abs(as.numeric(m$wti_second)), 1e-8)
      else NA_real_
  }

  # Derive AAII bull-bear spread if raw columns present
  if (!"aaii_bull_bear" %in% names(m) && all(c("aaii_bull","aaii_bear") %in% names(m)))
    m$aaii_bull_bear <- as.numeric(m$aaii_bull) - as.numeric(m$aaii_bear)

  # Required core columns must exist; licence-gated ones are created as NA if absent
  macro_optional <- c("fwd_pe_world","fwd_pe_eu","fwd_pe_em","eps_rev_world","eps_rev_eu",
                      "eps_rev_em","em_carry_diff","roll_yield","m2_yoy",
                      "cftc_gold","cftc_comm","put_call_ratio","conv_impl_vol")
  macro_required <- setdiff(macro_needed, macro_optional)
  miss_req <- setdiff(macro_required, names(m))
  if (length(miss_req) > 0) stop("macro sheet missing REQUIRED columns: ", paste(miss_req, collapse = ", "))
  for (c in setdiff(macro_optional, names(m))) {
    cat(sprintf("  NOTE: macro '%s' unavailable — set NA (signal contributes nothing).\n", c))
    m[[c]] <- NA_real_
  }

  # Keep ALL columns (incl. optional sentiment series) so build_sent() can use them
  macro <- as.data.frame(lapply(m, ffill))

  # ── Highs / lows (optional) ─────────────────────────────────────────────────
  sheets <- readxl::excel_sheets(path)
  hl <- function(sheet) { d <- rd(sheet); d[match(dates, d$Date), price_cols, drop = FALSE] }
  # Any missing high/low value (e.g. bond indices with no intraday range) falls
  # back to that week's close, so DMI/Stochastic degrade gracefully to close-based.
  fill_hl <- function(hl_df) as.data.frame(Map(function(h, c) {
    h <- suppressWarnings(as.numeric(h)); ifelse(is.finite(h), h, as.numeric(c))
  }, hl_df, prices), col.names = price_cols)
  highs <- if ("highs" %in% sheets) fill_hl(hl("highs")) else prices
  lows  <- if ("lows"  %in% sheets) fill_hl(hl("lows"))  else prices
  if (!("highs" %in% sheets)) cat("  No 'highs' sheet — using close as proxy.\n")
  if (!("lows"  %in% sheets)) cat("  No 'lows' sheet — using close as proxy.\n")

  # ── Returns ─────────────────────────────────────────────────────────────────
  returns <- as.data.frame(lapply(prices, function(p) {
    r <- c(0, diff(log(as.numeric(p)))); r[!is.finite(r)] <- 0; r
  }))

  n <- nrow(prices)
  cat(sprintf("  Loaded %d weekly observations (%s to %s).\n", n, format(min(dates)), format(max(dates))))
  if (n < 400) warning("Fewer than 400 weeks — backtest quality may be low.")

  list(dates = dates, prices = prices, returns = returns,
       highs = highs, lows = lows, macro = macro)
}
