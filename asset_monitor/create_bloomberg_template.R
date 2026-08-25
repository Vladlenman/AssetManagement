# create_bloomberg_template.R
# Generates bloomberg_data.xlsx pre-filled with Bloomberg Excel add-in (BDH)
# formulas. Open it on a machine with the Bloomberg Terminal + Excel add-in;
# the formulas auto-populate weekly history. Save, then run main.R with
# USE_BLOOMBERG <- TRUE to produce signals.
#
# WEEKLY UPDATE: the end date is Params!B2 = TODAY(). Just open the file
# (or press F9 / Bloomberg > Refresh Worksheets) each week, save, and re-run R.
#
# ── IMPORTANT: ONE periodicity per sheet ──────────────────────────────────────
# Every column on a data sheet MUST use the SAME Bloomberg periodicity (here:
# weekly, Per=W). If some columns are pulled monthly (Per=M) while the sheet's
# first, date-bearing column is weekly (or vice-versa), the shorter series and
# the date column write different numbers of rows: the longer columns overflow
# past the date column and their tail rows end up with NO date. The R loader can
# no longer tell which week each value belongs to, and a mis-built sheet silently
# flattens every macro/sentiment signal in recent years. So: monthly underliers
# (PMI, M2, etc.) are pulled at Per=W with Fill=P — Bloomberg simply repeats the
# last monthly print each week, giving a proper weekly step function that shares
# the single weekly date grid. (load_bloomberg_excel.R still LOCF-fills, so no
# information is invented.)
#
# Run:  source("asset_monitor/create_bloomberg_template.R")

if (!requireNamespace("openxlsx", quietly = TRUE)) install.packages("openxlsx")
library(openxlsx)

# ── BDH formula builder ───────────────────────────────────────────────────────
# date-bearing (first column of a sheet) shows dates; others hide them (Dts=H).
bdh <- function(ticker, field = "PX_LAST", hide_dates = TRUE, per = "W") {
  sprintf('BDH("%s","%s",Params!$B$1,Params!$B$2,"Per=%s","Days=A","Fill=P"%s)',
          ticker, field, per, if (hide_dates) ',"Dts=H"' else "")
}

# ── Column -> (ticker, field) maps ────────────────────────────────────────────
price_map <- list(
  equity_world="MXWO Index", equity_europe="MXEU Index", equity_em="MXEF Index",
  bonds_eur="LBEATREU Index", bonds_em_hc="JPEIGLBL Index", bonds_em_lc="JGENVUUG Index",
  convertibles="h24641eu Index", # your convertibles index
  commodities="BCOM Index", gold="XAUUSD Curncy",
  usd_currency="DXY Curncy", bonds_global="LEGATRUU Index", bonds_short_dur="LT01TRUU Index",
  bonds_med_dur="IEI US Equity", # iShares 3-7yr Tsy ETF (price)
  bonds_long_dur="LUTLTRUU Index",
  credit_ig="LUACTRUU Index", credit_hy="LF98TRUU Index"
)

# macro: each entry c(ticker, field). Order matters only for layout.
# ALL series are pulled WEEKLY (Per=W). Monthly underliers (PMI etc.) use Fill=P
# so their last print is carried forward each week -> one shared weekly grid.
macro_map <- list(
  # PMI: monthly underliers pulled WEEKLY (Per=W, Fill=P) so they align to the
  # single weekly date grid instead of writing a shorter monthly column that
  # overflows and loses its dates. The R loader forward-fills anyway.
  pmi_global   =c("CFNAI Index","PX_LAST","W"),   # Chicago Fed National Activity Index
  pmi_europe   =c("GRIFPBUS Index","PX_LAST","W"),# German Ifo Business Climate
  pmi_em       =c("CPMINDX Index","PX_LAST","W"), # China Official Manufacturing PMI
  # ── weekly series (Per=W) ──────────────────────────────────────────────────
  yield_2s10s  =c("USYC2Y10 Index","PX_LAST"),
  us_10y_yield =c("USGG10YR Index","PX_LAST"), us_real_yield=c("USGGT10Y Index","PX_LAST"),
  us_breakeven =c("USGGBE10 Index","PX_LAST"), eur_10y_yield=c("GDBR10 Index","PX_LAST"),
  eur_breakeven=c("EUSWI10 BGN Curncy","PX_LAST"), dxy=c("DXY Curncy","PX_LAST"),
  vix          =c("VIX Index","PX_LAST"),      us_ig_oas=c("LUACOAS Index","PX_LAST"),
  us_hy_oas    =c("LF98OAS Index","PX_LAST"),  eur_ig_oas=c("LECPOAS Index","PX_LAST"),
  em_hc_oas    =c("JPGCSOSD Index","PX_LAST"), m2_yoy=c("M2 Index","PX_LAST"),
  ecb_dir      =c("EUORDEPO Index","PX_LAST"), fwd_pe_world=c("MXWO Index","BEST_PE_RATIO"),
  fwd_pe_eu    =c("MXEU Index","BEST_PE_RATIO"), fwd_pe_em=c("MXEF Index","BEST_PE_RATIO"),
  cftc_gold    =c("GC1 Comdty","PX_LAST"),     cftc_comm=c("DN1 Index","PX_LAST"),
  put_call_ratio=c("PCUSEQTR Index","PX_LAST"),
  conv_impl_vol=c("VIX Index","PX_LAST"),
  # raw inputs; R derives eps_rev_*, em_carry_diff, roll_yield from these
  eps_ntm_world=c("MXWO Index","BEST_PE_RATIO"), eps_ntm_eu=c("MXEU Index","BEST_PE_RATIO"),
  eps_ntm_em   =c("MXEF Index","BEST_PE_RATIO"),
  gbiem_yield  =c("EMLC US Equity","YAS_BOND_YLD"),
  wti_front    =c("CL1 Comdty","PX_LAST"),     # roll yield = (front-second)/second (derived in R)
  wti_second   =c("CL2 Comdty","PX_LAST"),
  # ── sentiment (only the ones your terminal carries) ────────────────────────
  aaii_bull    =c("AAIIBULL Index","PX_LAST"), aaii_bear=c("AAIIBEAR Index","PX_LAST"),
  skew         =c("SKEW Index","PX_LAST"),     vstoxx=c("V2X Index","PX_LAST"),
  move_index   =c("MOVE Index","PX_LAST"),
  lqd_flows    =c("LQD US Equity","FUND_FLOW"), hyg_flows=c("HYG US Equity","FUND_FLOW")
)

# ── Styles ────────────────────────────────────────────────────────────────────
hdr <- createStyle(fontColour="#FFFFFF", fgFill="#1F4E79", halign="CENTER",
                   textDecoration="bold", border="Bottom")
note<- createStyle(fontColour="#8a6d00", textDecoration="italic")

# ── Write one BDH data sheet ──────────────────────────────────────────────────
# cols: named list column_name -> ticker (prices) or c(ticker,field) (macro)
write_bdh_sheet <- function(wb, sheet, cols, field_default = "PX_LAST") {
  addWorksheet(wb, sheet)
  headers <- c("Date", names(cols))
  writeData(wb, sheet, t(headers), startRow = 1, startCol = 1, colNames = FALSE)
  addStyle(wb, sheet, hdr, rows = 1, cols = seq_along(headers), gridExpand = TRUE)

  # First data column (B) is date-bearing: its BDH shows dates into column A.
  # Because ALL columns share one periodicity, the date column and every value
  # column write the same number of rows -> no undated overflow rows.
  spec1 <- cols[[1]]
  tick1 <- spec1[1]; fld1 <- if (length(spec1) > 1) spec1[2] else field_default
  per1  <- if (length(spec1) >= 3) spec1[3] else "W"
  writeFormula(wb, sheet, x = bdh(tick1, fld1, hide_dates = FALSE, per = per1), startCol = 1, startRow = 2)

  # Remaining columns C.. : values only (Dts=H)
  for (j in seq_along(cols)[-1]) {
    spec <- cols[[j]]
    tick <- spec[1]; fld <- if (length(spec) > 1) spec[2] else field_default
    per  <- if (length(spec) >= 3) spec[3] else "W"
    writeFormula(wb, sheet, x = bdh(tick, fld, hide_dates = TRUE, per = per),
                 startCol = j + 1, startRow = 2)   # +1 because col A is Date
  }
  setColWidths(wb, sheet, cols = 1, widths = 12)
  setColWidths(wb, sheet, cols = 2:(length(cols) + 1), widths = 16)
  freezePane(wb, sheet, firstActiveRow = 2, firstActiveCol = 2)
}

wb <- createWorkbook()

# Params sheet — start date in B1, end date in B2 (BDH formulas reference $B$1/$B$2).
# NO header row: the date cells MUST be B1 and B2, as real dates (via DATE()/TODAY()).
addWorksheet(wb, "Params")
writeData(wb, "Params", "Start date", startRow = 1, startCol = 1)
writeFormula(wb, "Params", x = "DATE(2010,1,1)", startCol = 2, startRow = 1)   # B1 = start
writeData(wb, "Params", "End date", startRow = 2, startCol = 1)
writeFormula(wb, "Params", x = "TODAY()", startCol = 2, startRow = 2)          # B2 = end
addStyle(wb, "Params", createStyle(numFmt = "yyyy-mm-dd"), rows = 1:2, cols = 2, gridExpand = TRUE)
writeData(wb, "Params", "B1 = start date (edit for more/less history). B2 = TODAY() auto-updates.",
          startRow = 4, startCol = 1)
addStyle(wb, "Params", note, rows = 4, cols = 1)
setColWidths(wb, "Params", cols = 1:2, widths = c(16, 16))

# Price / high / low sheets (same tickers, different field)
price_cols_named <- price_map
write_bdh_sheet(wb, "prices", price_cols_named)   # PX_LAST
# highs / lows: same tickers but PX_HIGH / PX_LOW
to_field <- function(m, f) lapply(m, function(t) c(t, f))
write_bdh_sheet(wb, "highs", to_field(price_map, "PX_HIGH"))
write_bdh_sheet(wb, "lows",  to_field(price_map, "PX_LOW"))

# Macro sheet
write_bdh_sheet(wb, "macro", macro_map)

# Instructions sheet
addWorksheet(wb, "Instructions")
instr <- c(
  "ASSET CLASS MONITOR — Bloomberg data workbook",
  "",
  "HOW IT WORKS",
  " - Each data sheet (prices, highs, lows, macro) uses Bloomberg BDH formulas.",
  " - Column A pulls the weekly dates; each other column pulls one series (PX_LAST,",
  "   or a field like BEST_PE_RATIO / BEST_EPS_NXT12MO / YLD_YTM_MID).",
  " - Periodicity is weekly (Per=W) for EVERY column, non-trading days filled with",
  "   previous value (Fill=P). Monthly underliers (PMI, M2) are also pulled weekly so",
  "   the last monthly print repeats each week and every column shares ONE date grid.",
  "",
  "WHY ONE PERIODICITY MATTERS",
  " - Do NOT mix Per=M and Per=W on the same sheet. A monthly column writes fewer",
  "   rows than the weekly ones; the weekly columns then overflow past the date",
  "   column and their tail rows have no date. The R loader cannot date those values,",
  "   and recent-year signals silently go flat. Keep every column Per=W.",
  "",
  "FIRST USE",
  " 1. Open on a PC with Bloomberg Terminal running + the Excel add-in installed.",
  " 2. Let the formulas populate (Bloomberg menu > Refresh Worksheets, or F9).",
  " 3. Save the file (keep the name bloomberg_data.xlsx).",
  " 4. In main.R set USE_BLOOMBERG <- TRUE, then run it.",
  "",
  "WEEKLY UPDATE",
  " - End date = Params!B2 = TODAY(). Just reopen / refresh each week, save, re-run R.",
  "",
  "DERIVED FIELDS (computed in R, not Excel)",
  " - eps_rev_world/eu/em = 4-week % change of eps_ntm_* (BEST_EPS_NXT12MO)",
  " - em_carry_diff       = gbiem_yield (JGENVUUG YLD_YTM_MID) minus us_10y_yield",
  " - roll_yield          = 4-week % change of bcom_roll_tr (BCMRYTOT)",
  "",
  "CHECK BEFORE FIRST RUN",
  " - conv_impl_vol currently proxies VIX — replace with a convertibles implied-vol",
  "   series if your terminal has one.",
  " - Confirm each ticker resolves on your terminal (type it + <GO>); swap house",
  "   equivalents for any Bloomberg/Barclays aggregate ticker that has been renamed.",
  " - Do NOT delete or rename the header row; the R loader matches columns by name.",
  " - Run check_bloomberg_data() after populating — it flags any column that has",
  "   more filled rows than dated rows (the mixed-periodicity mistake above)."
)
writeData(wb, "Instructions", instr, startRow = 1, startCol = 1, colNames = FALSE)
setColWidths(wb, "Instructions", cols = 1, widths = 95)
worksheetOrder(wb) <- match(c("Instructions","Params","prices","macro","highs","lows"),
                            names(wb))

out_path <- "asset_monitor/bloomberg_data.xlsx"
saveWorkbook(wb, out_path, overwrite = TRUE)
cat(sprintf("Workbook with BDH formulas saved -> %s\n", out_path))
