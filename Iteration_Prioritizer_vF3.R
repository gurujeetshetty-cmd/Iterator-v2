# =========================================================
# 🚀 ITERATION PRIORITIZER vF3.6 — Stratified (Constraint-Aware) Sampling + Live Progress
# =========================================================

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tidyr)
})

# -------------------------
# Optional: lightweight logger creator if you want a file logger elsewhere
# -------------------------
create_flush_logger <- function(path) {
  function(message, level = "INFO") {
    ts <- format(Sys.time(), "%H:%M:%S")
    line <- sprintf("[%s] %s %s\n", ts, level, message)
    con <- file(path, open = "ab", blocking = TRUE)
    writeBin(charToRaw(line), con)
    close(con)
    flush.console()
    invisible(NULL)
  }
}

# =========================================================
# Metric helpers (sourced from dedicated helper scripts)
# =========================================================

source_helper <- function(file) {
  candidates <- character(0)
  append_candidate <- function(path) {
    if (is.null(path) || !nzchar(path)) {
      return()
    }
    normalized <- tryCatch(
      normalizePath(path, winslash = "/", mustWork = FALSE),
      error = function(e) path
    )
    candidates <<- c(candidates, normalized)
  }

  append_candidate(file)
  append_candidate(file.path(getwd(), file))

  frames <- sys.frames()
  if (length(frames)) {
    for (frame in frames) {
      if (exists("ofile", envir = frame, inherits = FALSE)) {
        frame_file <- get("ofile", envir = frame, inherits = FALSE)
        if (!is.null(frame_file) && length(frame_file) && nzchar(frame_file[1])) {
          append_candidate(file.path(dirname(frame_file[1]), file))
        }
      }
    }
  }

  cmd_args <- commandArgs(trailingOnly = FALSE)
  file_arg <- "--file="
  matches <- cmd_args[grepl(file_arg, cmd_args, fixed = TRUE)]
  if (length(matches)) {
    script_path <- sub(file_arg, "", matches[[1]], fixed = TRUE)
    if (nzchar(script_path)) {
      append_candidate(file.path(dirname(script_path), file))
    }
  }

  candidates <- unique(candidates)
  for (candidate in candidates) {
    if (!is.na(candidate) && file.exists(candidate)) {
      source(candidate, local = FALSE)
      return(invisible(TRUE))
    }
  }

  stop(sprintf(
    "Unable to locate helper script '%s'. Checked: %s",
    file,
    paste(candidates, collapse = ", ")
  ))
}

source_helper("metrics_fa.R")
source_helper("metrics_pearson.R")
source_helper("input_utils.R")

# =========================================================
# Shared data helpers
# =========================================================

load_combination_data <- function(data, mandatory_variables_cols, log_fn = function(...) NULL) {
  data_path <- tryCatch(normalizePath(data, winslash = "/", mustWork = TRUE),
                        error = function(e) stop("Unable to normalize input path: ", conditionMessage(e)))

  themes_sheet <- "THEMES"
  input_sheet  <- "INPUT"
  scores_sheet <- "VARIABLES"
  na_values    <- c("NA", ".", "", " ")

  standardized_input <- read_standardized_input(data_path, sheet = input_sheet, na_values = na_values)
  input_metadata <- standardized_input$metadata
  candidate_meta <- input_metadata[input_metadata$input_flag_clean, , drop = FALSE]
  if (!nrow(candidate_meta)) {
    stop("No variables with INPUT_FLAG = 1 found in INPUT sheet.")
  }

  variable_meta <- candidate_meta[candidate_meta$column_index >= 4, , drop = FALSE]
  if (!nrow(variable_meta)) {
    stop("No candidate variables with INPUT_FLAG = 1 found beyond respondent identifier columns.")
  }

  XTab_Data <- standardized_input$data[, variable_meta$variable_id, drop = FALSE]

  variable_scores_raw <- readxl::read_excel(data_path, sheet = scores_sheet, range = "B2:C999")
  colnames(variable_scores_raw) <- c("Variable", "Score")
  variable_scores <- variable_scores_raw %>%
    mutate(Variable = clean_variable_ids(Variable),
           Score    = suppressWarnings(as.numeric(Score))) %>%
    filter(!is.na(Variable) & nzchar(Variable)) %>%
    mutate(Score = ifelse(is.na(Score), 0, Score))

  themes_raw <- readxl::read_excel(data_path, sheet = themes_sheet, range = "B2:D999")
  if (ncol(themes_raw) < 3) {
    stop("THEMES sheet must have 3 columns (Theme, Min, Max).")
  }

  names_lc <- tolower(trimws(names(themes_raw)))

  map_col <- function(patterns) {
    which(vapply(names_lc, function(nm) any(grepl(patterns, nm, ignore.case = TRUE)), logical(1)))
  }

  idx_theme <- map_col("^themes?$|^theme$")
  idx_min   <- map_col("^min\\b|\\bmin\\b|^min .*theme|^.*min .*select")
  idx_max   <- map_col("^max\\b|\\bmax\\b|^max .*theme|^.*max .*select")

  if (length(idx_theme) == 0) idx_theme <- which(toupper(names(themes_raw)) == "THEMES")
  if (length(idx_min)   == 0) idx_min   <- which(toupper(names(themes_raw)) == "MIN VARIABLES TO SELECT FROM THEME")
  if (length(idx_max)   == 0) idx_max   <- which(toupper(names(themes_raw)) == "MAX VARIABLES TO SELECT FROM THEME")

  if (length(idx_theme) != 1 || length(idx_min) != 1 || length(idx_max) != 1) {
    stop("THEMES sheet missing expected columns. Required headers (or equivalents): 'THEMES', 'MIN VARIABLES TO SELECT FROM THEME', 'MAX VARIABLES TO SELECT FROM THEME'.")
  }

  themes_df <- themes_raw[, c(idx_theme, idx_min, idx_max)]
  colnames(themes_df) <- c("THEME", "min", "max")
  themes_df <- themes_df %>%
    mutate(THEME = trimws(as.character(THEME))) %>%
    filter(!is.na(THEME) & nzchar(THEME)) %>%
    mutate(min = suppressWarnings(as.numeric(min)),
           max = suppressWarnings(as.numeric(max)))

  bad <- which(is.na(themes_df$min) | is.na(themes_df$max) | themes_df$max < themes_df$min)
  if (length(bad) > 0) {
    msg <- paste0(
      "Invalid rows in THEMES sheet:\n",
      paste0(" - Row ", bad + 1, " (sheet row offset by header): THEME='",
             themes_df$THEME[bad], "', min=", themes_df$min[bad], ", max=", themes_df$max[bad],
             collapse = "\n")
    )
    stop(msg)
  }

  Theme_Variable_Mapping <- tibble::tibble(
    Variable = variable_meta$variable_id,
    Theme    = variable_meta$theme
  ) %>%
    left_join(variable_scores, by = "Variable") %>%
    mutate(Score = ifelse(is.na(Score), 0, Score),
           Theme = ifelse(is.na(Theme), "", Theme))

  mandatory_variables_cols <- as.integer(mandatory_variables_cols)
  mandatory_variables_cols <- mandatory_variables_cols[!is.na(mandatory_variables_cols)]
  if (length(mandatory_variables_cols) > 0) {
    if (any(mandatory_variables_cols < 1 | mandatory_variables_cols > nrow(variable_meta))) {
      stop("mandatory_variables_cols contains indices outside available candidate variables.")
    }
  }

  input_colnames <- variable_meta$variable_id
  mandatory_vars <- intersect(input_colnames[mandatory_variables_cols], Theme_Variable_Mapping$Variable)

  theme_constraints <- setNames(
    lapply(seq_len(nrow(themes_df)), function(i) c(min = themes_df$min[i], max = themes_df$max[i])),
    themes_df$THEME
  )

  vars_by_theme <- Theme_Variable_Mapping %>%
    group_by(Theme) %>%
    summarise(pool = list(Variable), .groups = "drop") %>%
    { setNames(.$pool, .$Theme) }

  missing_pools <- setdiff(names(theme_constraints), names(vars_by_theme))
  if (length(missing_pools)) {
    stop(sprintf("No variables mapped to theme(s): %s", paste(missing_pools, collapse = ", ")))
  }

  mand_counts <- setNames(integer(length(theme_constraints)), names(theme_constraints))
  if (length(mandatory_vars)) {
    mand_themes <- Theme_Variable_Mapping %>%
      filter(Variable %in% mandatory_vars) %>%
      group_by(Theme) %>%
      summarise(cnt = dplyr::n(), .groups = "drop")
    mand_counts[mand_themes$Theme] <- mand_themes$cnt
  }

  optional_counts <- setNames(
    vapply(names(theme_constraints), function(th) {
      length(setdiff(vars_by_theme[[th]], mandatory_vars))
    }, integer(1)),
    names(theme_constraints)
  )

  theme_summary <- tibble::tibble(
    Theme = names(theme_constraints),
    min = vapply(theme_constraints, function(x) x[["min"]], numeric(1)),
    max = vapply(theme_constraints, function(x) x[["max"]], numeric(1)),
    mandatory = as.integer(mand_counts[names(theme_constraints)]),
    optional = as.integer(optional_counts[names(theme_constraints)])
  ) %>%
    mutate(total_available = mandatory + optional)

  min_total_required <- sum(pmax(theme_summary$min, theme_summary$mandatory))
  max_total_allowed <- sum(pmin(theme_summary$max, theme_summary$mandatory + theme_summary$optional))

  if (min_total_required > max_total_allowed) {
    stop("Theme-level minimums exceed maximums across all themes. Adjust THEME constraints.")
  }

  log_fn(sprintf("Loaded %d themes and %d candidate variables.",
                 nrow(theme_summary), nrow(Theme_Variable_Mapping)))

  list(
    source_path = data_path,
    XTab_Data = XTab_Data,
    Theme_Variable_Mapping = Theme_Variable_Mapping,
    themes_df = themes_df,
    theme_constraints = theme_constraints,
    vars_by_theme = vars_by_theme,
    mandatory_vars = mandatory_vars,
    mandatory_by_theme = mand_counts,
    optional_by_theme = optional_counts,
    theme_summary = theme_summary,
    total_variables = sum(theme_summary$total_available),
    min_total_required = min_total_required,
    max_total_allowed = max_total_allowed
  )
}

count_combinations_for_n <- function(n, context) {
  if (!length(context$theme_summary)) {
    return(0)
  }

  tbl <- context$theme_summary
  mandatory <- tbl$mandatory
  optional <- tbl$optional
  min_bounds <- tbl$min
  max_bounds <- tbl$max

  min_sel <- pmax(min_bounds, mandatory)
  max_sel <- pmin(max_bounds, mandatory + optional)

  total_mandatory <- sum(mandatory)
  if (total_mandatory > n) {
    return(0)
  }

  if (any(min_sel > max_sel)) {
    return(0)
  }

  min_total <- sum(min_sel)
  max_total <- sum(max_sel)
  if (n < min_total || n > max_total) {
    return(0)
  }

  theme_names <- tbl$Theme
  memo <- new.env(parent = emptyenv())

  calc <- function(idx, remaining) {
    if (idx > length(theme_names)) {
      return(ifelse(remaining == 0, 1, 0))
    }

    key <- paste0(idx, "|", remaining)
    if (exists(key, envir = memo, inherits = FALSE)) {
      return(memo[[key]])
    }

    mand <- mandatory[idx]
    opt  <- optional[idx]
    lo   <- min_sel[idx]
    hi   <- max_sel[idx]

    total <- 0
    if (lo > remaining) {
      memo[[key]] <- 0
      return(0)
    }

    hi <- min(hi, remaining)
    for (k in seq.int(lo, hi)) {
      opt_needed <- k - mand
      if (opt_needed < 0 || opt_needed > opt) {
        next
      }
      ways <- suppressWarnings(choose(opt, opt_needed))
      if (!is.finite(ways)) {
        ways <- Inf
      }
      sub <- calc(idx + 1L, remaining - k)
      if (is.infinite(ways) || is.infinite(sub)) {
        total <- Inf
        break
      } else {
        total <- total + (ways * sub)
      }
    }

    memo[[key]] <- total
    total
  }

  calc(1L, n)
}

calculate_capacity_table <- function(context, n_values) {
  if (!length(n_values)) {
    return(tibble::tibble())
  }

  counts <- vapply(n_values, count_combinations_for_n, numeric(1), context = context)
  tibble::tibble(
    n_size = n_values,
    min_required_total = context$min_total_required,
    max_allowed_total = context$max_total_allowed,
    max_possible_iterations = counts,
    feasible = counts > 0
  )
}

resolve_iteration_plan <- function(capacity_df, default_iterations, custom_iterations = NULL) {
  if (!nrow(capacity_df)) {
    capacity_df$requested_iterations <- integer(0)
    capacity_df$iterations_to_run <- integer(0)
    capacity_df$exhausted <- logical(0)
    capacity_df$available_unbounded <- logical(0)
    return(capacity_df)
  }

  default_iterations <- ifelse(is.na(default_iterations), 0L, as.integer(default_iterations))
  default_iterations <- max(default_iterations, 0L)

  requested <- rep(default_iterations, nrow(capacity_df))

  if (!is.null(custom_iterations)) {
    custom_iterations <- custom_iterations[!is.na(custom_iterations)]
    if (length(custom_iterations)) {
      idx <- match(capacity_df$n_size, as.integer(names(custom_iterations)))
      override <- custom_iterations[idx]
      requested <- ifelse(!is.na(override), as.integer(pmax(0, override)), requested)
    }
  }

  requested[is.na(requested)] <- 0L

  max_possible <- capacity_df$max_possible_iterations
  iterations_to_run <- requested
  finite_mask <- is.finite(max_possible)
  iterations_to_run[finite_mask] <- pmin(requested[finite_mask], as.numeric(max_possible[finite_mask]))
  iterations_to_run[!finite_mask] <- requested[!finite_mask]
  iterations_to_run[!is.finite(iterations_to_run)] <- 0

  capacity_df$requested_iterations <- requested
  capacity_df$iterations_to_run <- iterations_to_run
  capacity_df$exhausted <- finite_mask & (requested > max_possible)
  capacity_df$available_unbounded <- !finite_mask & (requested > 0)
  capacity_df
}

summarize_combination_capacity <- function(
  data,
  min_n_size,
  max_n_size,
  mandatory_variables_cols,
  default_iterations,
  custom_iterations = NULL,
  log_fn = function(...) NULL
) {
  if (max_n_size < min_n_size) {
    log_fn(sprintf("Invalid N range: min=%d, max=%d", min_n_size, max_n_size), "WARN")
    return(list(
      plan = tibble::tibble(),
      context = load_combination_data(data, mandatory_variables_cols, log_fn),
      total_iterations = 0L
    ))
  }

  context <- load_combination_data(data, mandatory_variables_cols, log_fn)
  n_values <- seq.int(min_n_size, max_n_size)
  capacity <- calculate_capacity_table(context, n_values)
  plan <- resolve_iteration_plan(capacity, default_iterations, custom_iterations)
  list(
    plan = plan,
    context = context,
    total_iterations = sum(plan$iterations_to_run)
  )
}

# =========================================================
# Main
# =========================================================
# progress_cb: optional function(n, iter_i, iter_max) provided by the dashboard with withProgress()
generate_best_combinations <- function(
  data,
  min_n_size,
  max_n_size,
  max_iterations,                 # iterations PER n (fallback)
  mandatory_variables_cols,       # numeric indices into INPUT (variable columns)
  output_file,
  append_log,                     # function(msg, level="INFO")
  progress_cb = NULL,             # optional function(n, iter_i, iter_max)
  iterations_per_n = NULL,        # optional named vector of requested iterations
  plan_summary = NULL,            # optional precomputed list from summarize_combination_capacity()
  include_fa = FALSE,
  include_corr = FALSE
) {
  log_it <- function(msg, lvl = "INFO") { try(append_log(msg, lvl), silent = TRUE) }

  normalized_path <- tryCatch(normalizePath(data, winslash = "/", mustWork = TRUE),
                              error = function(e) stop("Unable to normalize input path: ", conditionMessage(e)))

  custom_vec <- NULL
  if (!is.null(iterations_per_n)) {
    if (is.null(names(iterations_per_n)) && length(iterations_per_n) == (max_n_size - min_n_size + 1)) {
      names(iterations_per_n) <- as.character(seq.int(min_n_size, max_n_size))
    }
    custom_vec <- suppressWarnings(as.numeric(iterations_per_n))
    names(custom_vec) <- names(iterations_per_n)
    custom_vec <- custom_vec[!is.na(custom_vec)]
  }

  if (is.null(plan_summary) || is.null(plan_summary$plan) || is.null(plan_summary$context)) {
    plan_summary <- summarize_combination_capacity(
      data = normalized_path,
      min_n_size = min_n_size,
      max_n_size = max_n_size,
      mandatory_variables_cols = mandatory_variables_cols,
      default_iterations = max_iterations,
      custom_iterations = custom_vec,
      log_fn = log_it
    )
  } else {
    summary_path <- tryCatch(plan_summary$context$source_path, error = function(e) NULL)
    if (!is.null(summary_path) && !identical(summary_path, normalized_path)) {
      log_it("Provided plan summary does not match requested data path. Recomputing plan.", "WARN")
      plan_summary <- summarize_combination_capacity(
        data = normalized_path,
        min_n_size = min_n_size,
        max_n_size = max_n_size,
        mandatory_variables_cols = mandatory_variables_cols,
        default_iterations = max_iterations,
        custom_iterations = custom_vec,
        log_fn = log_it
      )
    }
  }

  plan <- plan_summary$plan
  context <- plan_summary$context
  total_iterations <- if (!is.null(plan_summary$total_iterations)) {
    plan_summary$total_iterations
  } else {
    sum(plan$iterations_to_run)
  }

  fa_context <- NULL
  corr_context <- NULL

  if (isTRUE(include_fa) || isTRUE(include_corr)) {
    sheets <- tryCatch(readxl::excel_sheets(normalized_path),
                       error = function(e) stop("Unable to list sheets in workbook: ", conditionMessage(e)))
    if (isTRUE(include_fa)) {
      if (!"FA_OP" %in% sheets) {
        stop("FA_OP sheet not found in workbook. Disable factor analysis metrics or add the sheet.")
      }
      fa_raw <- tryCatch(readxl::read_excel(normalized_path, sheet = "FA_OP"),
                         error = function(e) stop("Unable to read FA_OP sheet: ", conditionMessage(e)))
      fa_context <- prepare_fa_context(fa_raw)
      log_it("Factor analysis metrics enabled (FA_OP sheet loaded).")
    }
    if (isTRUE(include_corr)) {
      if (!"RAW_DATA" %in% sheets) {
        stop("RAW_DATA sheet not found in workbook. Disable Pearson metrics or add the sheet.")
      }
      raw_data <- tryCatch(readxl::read_excel(normalized_path, sheet = "RAW_DATA"),
                           error = function(e) stop("Unable to read RAW_DATA sheet: ", conditionMessage(e)))
      corr_context <- prepare_corr_context(raw_data)
      log_it("Pearson correlation metrics enabled (RAW_DATA sheet loaded).")
    }
  }

  Theme_Variable_Mapping <- context$Theme_Variable_Mapping
  theme_constraints <- context$theme_constraints
  vars_by_theme <- context$vars_by_theme
  mandatory_vars <- context$mandatory_vars
  mand_cnt <- context$mandatory_by_theme

  # ---------- Output Header ----------
  base_extra_cols <- c(
    "Ran_Iteration_Flag","ITR_PATH","ITR_FILE_NAME","MAX_N_SIZE","MAX_N_SIZE_PERC",
    "MIN_N_SIZE","MIN_N_SIZE_PERC","SOLUTION_N_SIZE","PROB_95","PROB_90","PROB_80",
    "PROB_75","PROB_LESS_THAN_50","BIMODAL_VARS","PROPER_BUCKETED_VARS",
    "BIMODAL_VARS_PERC","ROV_SD","ROV_RANGE","seg1_diff","seg2_diff","seg3_diff",
    "seg4_diff","seg5_diff","bi_1","perf_1","indT_1","indB_1","bi_2","perf_2",
    "indT_2","indB_2","bi_3","perf_3","indT_3","indB_3","bi_4","perf_4","indT_4",
    "indB_4","bi_5","perf_5","indT_5","indB_5"
  )

  fa_metric_cols <- c(
    "FA_DISTINCT_FACTORS","FA_FACTOR_COUNTS","FA_ENTROPY","FA_MEAN_PRIMARY_LOADING",
    "FA_MIN_PRIMARY_LOADING","FA_MEAN_COMMUNALITY","FA_MIN_COMMUNALITY","FA_CROSSLOAD_RATIO",
    "FA_WEAK_ITEM_COUNT","FA_COVERAGE_MIN_K","FA_MISSING_COUNT","FA_COVERAGE_RATIO"
  )

  corr_metric_cols <- c(
    "P_R_MEAN_ABS","P_R_MEDIAN_ABS","P_R_MAX_ABS","P_EIGVAR_PC1","P_EFF_DIM",
    "P_TOPPAIR1","P_TOPPAIR1_R"
  )

  input_perf_cols <- c(
    "BIMODAL_VARS_INPUT_VARS","INPUT_PERF_SEG_COVERED",
    "INPUT_PERF_FLAG","INPUT_PERF_MINSEG","INPUT_PERF_MAXSEG"
  )

  sort_of_columns <- function(cols) {
    if (!length(cols)) return(cols)

    base_keys <- gsub("^(max|min)_", "", cols)
    base_keys <- sub("_(values|val|diff_max|diff_min|diff)$", "", base_keys, perl = TRUE)
    base_keys <- toupper(base_keys)

    metric_type <- ifelse(grepl("^max_", cols), "legacy_max",
                    ifelse(grepl("^min_", cols), "legacy_min",
                           sub("^OF_[A-Za-z0-9_]+_", "", cols)))

    metric_order <- c(
      values = 1L,
      val = 2L,
      diff_max = 3L,
      diff = 4L,
      diff_min = 5L,
      legacy_max = 6L,
      legacy_min = 7L
    )

    metric_type[metric_type == ""] <- "values"
    metric_type[!(metric_type %in% names(metric_order))] <- "zzz"
    metric_order <- c(metric_order, zzz = 99L)

    order_idx <- metric_order[metric_type]
    cols[order(base_keys, order_idx, cols)]
  }

  existing_of_cols <- character(0)
  if (file.exists(output_file)) {
    suppressWarnings({
      header_preview <- try(read.csv(output_file, nrows = 0, stringsAsFactors = FALSE), silent = TRUE)
      if (!inherits(header_preview, "try-error")) {
        of_pattern <- "^(max|min)_OF_[A-Za-z0-9_]+_diff$|^OF_[A-Za-z0-9_]+_(val|values|diff|diff_min|diff_max)$"
        existing_of_cols <- names(header_preview)[
          grepl(of_pattern, names(header_preview))
        ]
      }
    })
  }

  of_cols <- sort_of_columns(unique(existing_of_cols))
  extra_cols <- unique(c(base_extra_cols, of_cols, input_perf_cols))
  if (isTRUE(include_fa)) {
    extra_cols <- c(extra_cols, fa_metric_cols)
  }
  if (isTRUE(include_corr)) {
    extra_cols <- c(extra_cols, corr_metric_cols)
  }

  if (file.exists(output_file)) {
    log_it(sprintf("Overwriting existing output file: %s", output_file), "WARN")
    try(file.remove(output_file), silent = TRUE)
  }

  write.table(
    t(c("Iteration","n_size","Combination","Total_Score","Avg_Score", extra_cols)),
    file = output_file, sep = ",", col.names = FALSE, row.names = FALSE
  )

  if (!nrow(plan)) {
    log_it("No feasible combinations for requested N range. Output file initialized only.", "WARN")
  }

  log_it(sprintf(
    "Starting combination generation for n=%d..%d (requested iterations total: %s).",
    min_n_size, max_n_size,
    format(total_iterations, big.mark = ",", scientific = FALSE)
  ))

  if (any(plan$exhausted)) {
    limited <- plan[plan$exhausted, , drop = FALSE]
    msg <- paste0(
      "Iteration requests exceed feasible combinations for n sizes: ",
      paste(sprintf("n=%d (requested=%d, feasible=%s)",
                   limited$n_size,
                   limited$requested_iterations,
                   format(limited$max_possible_iterations, scientific = FALSE, big.mark = ",")),
            collapse = "; ")
    )
    log_it(msg, "WARN")
  }

  plan <- plan[order(plan$n_size), , drop = FALSE]

  total_mand <- length(mandatory_vars)

  global_iter_id <- 1L

  for (row_idx in seq_len(nrow(plan))) {
    n <- plan$n_size[row_idx]
    target_iterations <- as.integer(plan$iterations_to_run[row_idx])
    requested_iterations <- as.integer(plan$requested_iterations[row_idx])
    max_possible_for_n <- plan$max_possible_iterations[row_idx]

    if (total_mand > n) {
      log_it(sprintf("Skipping n=%d: mandatory vars (%d) exceed n.", n, total_mand), "WARN")
      next
    }

    if (target_iterations <= 0) {
      log_it(sprintf("Skipping n=%d: zero iterations requested or feasible.", n), "INFO")
      next
    }

    for (th in names(theme_constraints)) {
      maxv <- theme_constraints[[th]]["max"]
      cur  <- if (!is.null(mand_cnt[[th]])) mand_cnt[[th]] else 0L
      if (cur > maxv) {
        stop(sprintf("Theme '%s' violates MAX with mandatory vars: %d > %d", th, cur, maxv))
      }
    }

    log_it(sprintf(
      "n=%d: attempting %d iteration(s) (requested %d, feasible %s)",
      n,
      target_iterations,
      requested_iterations,
      format(max_possible_for_n, scientific = FALSE, big.mark = ",")
    ))

    seen <- new.env(hash = TRUE, parent = emptyenv())
    iter_i <- 1L
    attempts <- 0L
    attempt_cap <- max(1000L, target_iterations * 50L)

    while (iter_i <= target_iterations && attempts < attempt_cap) {
      attempts <- attempts + 1L

      if (!is.null(progress_cb)) {
        try(progress_cb(n, iter_i - 1L, target_iterations), silent = TRUE)
      }

      k_free <- n - total_mand
      cur_counts <- mand_cnt
      cur_counts[setdiff(names(theme_constraints), names(cur_counts))] <- 0L

      need_min <- setNames(
        pmax(0, vapply(names(theme_constraints), function(th)
          theme_constraints[[th]]["min"] - cur_counts[[th]], numeric(1))), names(theme_constraints)
      )
      need_sum <- sum(need_min)
      if (need_sum > k_free) {
        log_it(sprintf(
          "n=%d: remaining free slots %d < sum(min deficits) %d; stopping for this n.",
          n, k_free, need_sum
        ), "WARN")
        break
      }

      alloc <- need_min
      slots_left <- k_free - need_sum

      max_room <- setNames(
        pmax(0, vapply(names(theme_constraints), function(th)
          theme_constraints[[th]]["max"] - (cur_counts[[th]] + alloc[[th]]), numeric(1))),
        names(theme_constraints)
      )

      if (slots_left > 0 && sum(max_room) > 0) {
        themes_vec <- rep(names(max_room), times = pmax(1L, as.integer(max_room)))
        if (length(themes_vec) > 0) {
          add <- table(sample(themes_vec, size = slots_left, replace = length(themes_vec) < slots_left))
          for (th in names(add)) alloc[[th]] <- alloc[[th]] + as.integer(add[[th]])
        }
      }

      pick <- character(0)

      infeasible_allocation <- FALSE
      for (th in names(theme_constraints)) {
        want <- alloc[[th]]
        if (want <= 0) next

        pool_th <- setdiff(vars_by_theme[[th]], mandatory_vars)
        pool_th <- setdiff(pool_th, pick)

        if (length(pool_th) < want) {
          infeasible_allocation <- TRUE
          break
        }
        pick <- c(pick, sample(pool_th, want, replace = FALSE))
      }

      if (infeasible_allocation) {
        Sys.sleep(0.002)
        next
      }

      combo_vars <- sort(unique(c(mandatory_vars, pick)))
      if (length(combo_vars) != n) {
        Sys.sleep(0.001)
        next
      }

      key <- paste(combo_vars, collapse = "|")
      if (exists(key, envir = seen, inherits = FALSE)) {
        next
      }
      assign(key, TRUE, envir = seen)

      rows <- Theme_Variable_Mapping[match(combo_vars, Theme_Variable_Mapping$Variable), , drop = FALSE]
      total_score <- sum(rows$Score, na.rm = TRUE)
      avg_score   <- total_score / n

      row_values <- setNames(rep(NA_character_, length(extra_cols)), extra_cols)
      row_values["Ran_Iteration_Flag"] <- "0"

      # ---- Factor Analysis Metrics Integration ----
      if (isTRUE(include_fa) && !is.null(fa_context)) {
        fa_metrics <- compute_fa_metrics(combo_vars, fa_context)
        row_values[names(fa_metrics)] <- vapply(fa_metrics, function(x) {
          if (is.null(x) || (length(x) == 1 && is.na(x))) {
            NA_character_
          } else {
            as.character(x)
          }
        }, character(1))
      }

      # ---- Pearson Correlation Metrics Integration ----
      if (isTRUE(include_corr) && !is.null(corr_context)) {
        corr_metrics <- compute_corr_metrics(combo_vars, corr_context)
        row_values[names(corr_metrics)] <- vapply(corr_metrics, function(x) {
          if (is.null(x) || (length(x) == 1 && is.na(x))) {
            NA_character_
          } else {
            as.character(x)
          }
        }, character(1))
      }

      write.table(
        t(c(global_iter_id,
            n,
            paste(combo_vars, collapse = ";"),
            total_score,
            round(avg_score, 6),
            row_values)),
        file = output_file, append = TRUE, sep = ",", col.names = FALSE, row.names = FALSE
      )

      if (!is.null(progress_cb)) {
        try(progress_cb(n, iter_i, target_iterations), silent = TRUE)
      }
      if ((iter_i %% 10L) == 0L) {
        log_it(sprintf("Progress: n=%d | %d/%d", n, iter_i, target_iterations))
      }

      global_iter_id <- global_iter_id + 1L
      iter_i <- iter_i + 1L

      if (is.finite(max_possible_for_n)) {
        if (length(ls(envir = seen, all.names = TRUE)) >= max_possible_for_n) {
          if (iter_i <= target_iterations) {
            log_it(sprintf("n=%d: all feasible combinations exhausted (%s).", n,
                           format(max_possible_for_n, scientific = FALSE, big.mark = ",")), "INFO")
          }
          break
        }
      }
    }

    if (iter_i <= target_iterations) {
      log_it(sprintf(
        "n=%d: stopped early after %d/%d iterations (attempts=%d).",
        n, iter_i - 1L, target_iterations, attempts
      ), "WARN")
    }
  }

  # ---------- Post-process: sort by Avg_Score ----------
  out <- read.csv(output_file, header = FALSE, stringsAsFactors = FALSE)
  if (nrow(out) >= 2) {
    colnames(out) <- out[1, ]
    out <- out[-1, , drop = FALSE]
    suppressWarnings({
      out$Avg_Score   <- as.numeric(out$Avg_Score)
      out$Total_Score <- as.numeric(out$Total_Score)
      out$n_size      <- as.integer(out$n_size)
      out$Iteration   <- as.integer(out$Iteration)
    })
    out <- out %>%
      arrange(desc(Avg_Score)) %>%
      mutate(Iteration = seq_len(n()))
    write.csv(out, file = output_file, row.names = FALSE)
  }

  log_it("Combination generation complete (or stopped).", "SUCCESS")
  invisible(TRUE)
}

# For compatibility with other scripts that expect it:
find_variable_columns <- function(input_generator_output, xtabs_input_file) {
  seq_len(length(input_generator_output))
}
