library(readxl)

find_neighbor_file <- function(filename) {
  caller_frames <- sys.frames()
  caller_paths <- vapply(caller_frames, function(fr) {
    if (!is.null(fr$ofile)) return(fr$ofile)
    NA_character_
  }, character(1))
  caller_paths <- caller_paths[!is.na(caller_paths) & nzchar(caller_paths)]

  candidate_dirs <- c(getwd())
  if (length(caller_paths)) {
    candidate_dirs <- c(candidate_dirs, dirname(normalizePath(caller_paths, winslash = "/", mustWork = FALSE)))
  }
  candidate_dirs <- unique(candidate_dirs)

  candidates <- file.path(candidate_dirs, filename)
  existing <- candidates[file.exists(candidates)]
  if (length(existing)) return(existing[[1]])
  NULL
}

if (!exists("run_summary_generator")) {
  helper_path <- find_neighbor_file("summary_generator_helper.R")
  if (!is.null(helper_path)) {
    source(helper_path)
  }
}

# Function to run the iterations
run_iterations <- function(comb_data,input_working_dir,output_working_dir,input_db_name,set_num_segments,File_name,solo, progress_cb = NULL, run_summary_txt = FALSE) {
  print(paste0("SOLO VALUE -4: ",solo))
  # Read the Excel file
  file_path<-file.path(comb_data[[1]],comb_data[[2]])
  input_data <- read.csv(file_path,stringsAsFactors=FALSE)

  safe_progress <- function(current, total, iteration = NA_integer_, status = NULL) {
    if (is.null(progress_cb)) {
      return(invisible(NULL))
    }

    safe_current <- current
    safe_total <- total

    if (!is.finite(safe_current)) safe_current <- 0
    if (!is.finite(safe_total) || safe_total <= 0) {
      safe_total <- 0
      safe_current <- 0
    }

    safe_current <- max(0, min(safe_current, safe_total))

    try(
      progress_cb(
        current = safe_current,
        total = safe_total,
        iteration = iteration,
        status = status
      ),
      silent = TRUE
    )

    invisible(NULL)
  }

  ensure_column <- function(df, name, value = NA) {
    if (!name %in% colnames(df)) {
      df[[name]] <- value
    }
    df
  }

  ensure_columns <- function(df, names, value = NA) {
    for (nm in names) {
      df <- ensure_column(df, nm, value)
    }
    df
  }

  propagate_legacy_of_aliases <- function(df) {
    name_upper <- toupper(names(df))
    canonical_pattern <- "^OF_[A-Z0-9_]+_(VALUES|MINDIFF|MAXDIFF|DIFF_MAX|DIFF_MIN)$"
    canonical_cols <- names(df)[grepl(canonical_pattern, name_upper)]
    if (!length(canonical_cols)) return(df)

    for (col in canonical_cols) {
      col_upper <- toupper(col)
      base <- sub("_(VALUES|MINDIFF|MAXDIFF|DIFF_MAX|DIFF_MIN)$", "", col_upper)

      aliases <- character(0)
      if (endsWith(col_upper, "VALUES")) {
        aliases <- c(aliases, paste0(base, "_VAL"))
      }
      if (endsWith(col_upper, "MAXDIFF") || endsWith(col_upper, "DIFF_MAX")) {
        aliases <- c(aliases, paste0(base, "_DIFF"), paste0("MAX_", base, "_DIFF"))
      }
      if (endsWith(col_upper, "MINDIFF") || endsWith(col_upper, "DIFF_MIN")) {
        aliases <- c(aliases, paste0("MIN_", base, "_DIFF"))
      }

      for (alias in aliases) {
        existing_idx <- match(alias, toupper(names(df)))
        target_name <- if (is.na(existing_idx)) alias else names(df)[existing_idx]
        df[[target_name]] <- df[[col]]
      }
    }

    df
  }

  sort_of_columns <- function(cols) {
    if (!length(cols)) return(cols)

    base_keys <- gsub("^(MAX|MIN)_", "", toupper(cols))
    base_keys <- sub("_(VALUES|VAL|MINDIFF|MAXDIFF|DIFF_MAX|DIFF_MIN|DIFF)$", "", base_keys, perl = TRUE)

    suffix <- sub("^OF_[A-Z0-9_]+_", "", toupper(cols))
    metric_type <- tolower(suffix)

    metric_type[metric_type == "diff_max"] <- "maxdiff"
    metric_type[metric_type == "diff_min"] <- "mindiff"

    metric_order <- c(
      values = 1L,
      val = 2L,
      maxdiff = 3L,
      diff = 4L,
      mindiff = 5L,
      legacy_max = 6L,
      legacy_min = 7L
    )

    metric_type <- ifelse(grepl("^MAX_", cols, ignore.case = TRUE), "legacy_max",
                    ifelse(grepl("^MIN_", cols, ignore.case = TRUE), "legacy_min", metric_type))

    metric_type[metric_type == ""] <- "values"
    metric_type[!(metric_type %in% names(metric_order))] <- "zzz"
    metric_order <- c(metric_order, zzz = 99L)

    order_idx <- metric_order[metric_type]
    cols[order(base_keys, order_idx, cols)]
  }

  input_data <- ensure_column(input_data, "Ran_Iteration_Flag", 0)
  input_data <- ensure_columns(input_data, c("ITR_PATH", "ITR_FILE_NAME", "SUMMARY_TXT_PATH", "SUMMARY_TXT_FILE"), NA)
  input_data$ITR_FILE_NAME <- as.character(input_data$ITR_FILE_NAME)

  rules_checker_cols <- c(
    "MAX_N_SIZE","MAX_N_SIZE_PERC","MIN_N_SIZE","MIN_N_SIZE_PERC","SOLUTION_N_SIZE",
    "PROB_95","PROB_90","PROB_80","PROB_75","PROB_LESS_THAN_50","BIMODAL_VARS",
    "PROPER_BUCKETED_VARS","BIMODAL_VARS_PERC","ROV_SD","ROV_RANGE",
    "seg1_diff","seg2_diff","seg3_diff","seg4_diff","seg5_diff",
    "bi_1","perf_1","indT_1","indB_1","bi_2","perf_2","indT_2","indB_2",
    "bi_3","perf_3","indT_3","indB_3","bi_4","perf_4","indT_4","indB_4",
    "bi_5","perf_5","indT_5","indB_5"
  )
  input_data <- ensure_columns(input_data, rules_checker_cols, NA)

  input_perf_cols <- c("BIMODAL_VARS_INPUT_VARS","INPUT_PERF_SEG_COVERED","INPUT_PERF_FLAG",
                       "INPUT_PERF_MINSEG","INPUT_PERF_MAXSEG")
  input_data <- ensure_columns(input_data, input_perf_cols, NA)

  ordered_base_cols <- c("Iteration","n_size","Combination","Total_Score","Avg_Score")
  if ("Var_cols" %in% names(input_data)) {
    ordered_base_cols <- c(ordered_base_cols, "Var_cols")
  }

  tracker_scaffold_cols <- c(
    "Ran_Iteration_Flag","ITR_PATH","ITR_FILE_NAME","SUMMARY_TXT_PATH","SUMMARY_TXT_FILE",
    "MAX_N_SIZE","MAX_N_SIZE_PERC","MIN_N_SIZE","MIN_N_SIZE_PERC","SOLUTION_N_SIZE",
    "PROB_95","PROB_90","PROB_80","PROB_75","PROB_LESS_THAN_50","BIMODAL_VARS",
    "PROPER_BUCKETED_VARS","BIMODAL_VARS_PERC","ROV_SD","ROV_RANGE","seg1_diff",
    "seg2_diff","seg3_diff","seg4_diff","seg5_diff","bi_1","perf_1","indT_1",
    "indB_1","bi_2","perf_2","indT_2","indB_2","bi_3","perf_3","indT_3",
    "indB_3","bi_4","perf_4","indT_4","indB_4","bi_5","perf_5","indT_5",
    "indB_5"
  )

  ordered_input_perf_cols <- input_perf_cols
  
  write.csv(input_data,file_path, row.names = FALSE)
  input_data$Combination <- as.character(input_data$Combination)
  
  cont_seg_var_list <- c()
  set_maxiter <- 1000
  set_nrep <- 10
  option_summary_only <- 0
  
  total_iterations <- sum(input_data$Ran_Iteration_Flag != 1, na.rm = TRUE)
  if (!is.finite(total_iterations) || total_iterations <= 0) {
    total_iterations <- nrow(input_data)
  }

  safe_progress(0L, total_iterations, NA_integer_, "Preparing iterations...")

  processed_iterations <- 0L

  print(nrow(input_data))
  for (i in 1:nrow(input_data)) {
    iteration_value <- input_data$Iteration[i]
    iteration_value <- ifelse(is.na(iteration_value) || !is.finite(iteration_value), i, iteration_value)

    if (isTRUE(input_data$Ran_Iteration_Flag[i] == 1)) {next}
    processed_iterations <- processed_iterations + 1L
    safe_progress(
      processed_iterations - 1L,
      total_iterations,
      iteration = iteration_value,
      status = sprintf(
        "Running iteration %d (%d of %d)...",
        iteration_value,
        processed_iterations,
        total_iterations
      )
    )
    iteration <- iteration_value
    combination <- unlist(strsplit(input_data$Combination[i], ";"))
    print(combination)
    file_run_and_date <- paste0(File_name, iteration)
    print(file_run_and_date)

    # Run LCA + RULES_Checker
    LCA_SEG_Summary(
      input_working_dir = input_working_dir,
      output_working_dir = output_working_dir,
      input_db_name = input_db_name,
      file_run_and_date = file_run_and_date,
      cat_seg_var_list = combination,
      cont_seg_var_list = cont_seg_var_list,
      set_num_segments = set_num_segments,
      set_maxiter = set_maxiter,
      set_nrep = set_nrep,
      option_summary_only = option_summary_only
    )

    # === DEBUG WRAPPER: detect which RULES_OP element is empty or invalid ===
check_and_assign <- function(df, row_index, iteration_value, col, val, var_name) {
  if (!col %in% colnames(df)) {
    df[[col]] <- NA
  }
  # Detect empty or invalid values
  if (is.null(val) || length(val) == 0) {
    cat("❌ ERROR @ Iteration", iteration_value, "| Variable", var_name, "is NULL or length 0\n")
    return(df)
  }
  if (all(is.na(val))) {
    cat("⚠️ WARNING @ Iteration", iteration_value, "| Variable", var_name, "is all NA\n")
  }
  if (is.list(val)) {
    cat("⚠️ WARNING @ Iteration", iteration_value, "| Variable", var_name, "is a list of length", length(val), "\n")
    print(val)
  }
  # Print data type and value for QC
  cat("✅ Assigning", var_name, "=", paste(val, collapse=", "), 
      "| Type:", typeof(val), "| Length:", length(val), "\n")
  
  target_rows <- which(df$Iteration == iteration_value)
  if (!length(target_rows)) {
    target_rows <- row_index
  }

  # Perform safe assignment
  tryCatch({
    df[[col]][target_rows] <- val
  }, error = function(e) {
    cat("💥 CRASH during assignment of", var_name, ":", conditionMessage(e), "\n")
  })
  return(df)
}

    RULES_OP<- RULES_Checker(output_working_dir,paste0(file_run_and_date,"_SUMMARY",".xlsx"),solo)

    of_pattern <- "^((max|min)_)?OF_[A-Za-z0-9_]+(_VALUES|_VAL|_MINDIFF|_MAXDIFF|_DIFF(_MAX|_MIN)?|_DIFF)$"
    of_metric_cols <- names(RULES_OP)[grepl("^((max|min)_)?OF_[A-Za-z0-9_]+", names(RULES_OP), ignore.case = TRUE)]
    of_metric_cols <- sort_of_columns(of_metric_cols)
    if (length(of_metric_cols)) {
      input_data <- ensure_columns(input_data, of_metric_cols, NA)
    }

    # === Begin debug-safe assignments ===

    input_data <- check_and_assign(input_data, i, iteration, "ITR_PATH", output_working_dir, "ITR_PATH")
    input_data <- check_and_assign(input_data, i, iteration, "ITR_FILE_NAME", paste0(file_run_and_date, "_SUMMARY.xlsx"), "ITR_FILE_NAME")

    summary_txt_path <- NA_character_
    summary_txt_file <- NA_character_

    if (run_summary_txt) {
      summary_workbook_path <- file.path(output_working_dir, paste0(file_run_and_date, "_SUMMARY.xlsx"))
      summary_txt_name <- paste0(file_run_and_date, "_SUMMARY.txt")

      safe_progress(
        processed_iterations,
        total_iterations,
        iteration = iteration,
        status = sprintf("Running SUMMARY_GENERATOR for iteration %d", iteration)
      )

      if (file.exists(summary_workbook_path)) {
        tryCatch({
          summary_result <- run_summary_generator(summary_workbook_path, output_working_dir, summary_txt_name)
          summary_txt_path <- normalizePath(summary_result$output_path, winslash = "/", mustWork = FALSE)
          summary_txt_file <- basename(summary_txt_path)
          cat("📝 SUMMARY_GENERATOR saved:", summary_txt_path, "\n")
        }, error = function(e) {
          cat("⚠️ Failed to run SUMMARY_GENERATOR for iteration", i, ":", conditionMessage(e), "\n")
        })
      } else {
        cat("⚠️ SUMMARY_GENERATOR workbook missing for iteration", i, "at", summary_workbook_path, "\n")
      }
    }

    input_data <- check_and_assign(input_data, i, iteration, "SUMMARY_TXT_PATH", summary_txt_path, "SUMMARY_TXT_PATH")
    input_data <- check_and_assign(input_data, i, iteration, "SUMMARY_TXT_FILE", summary_txt_file, "SUMMARY_TXT_FILE")

    # --- RULES_Checker outputs ---
    input_data <- check_and_assign(input_data, i, iteration, "MAX_N_SIZE", RULES_OP$Max_n_size, "Max_n_size")
    input_data <- check_and_assign(input_data, i, iteration, "MAX_N_SIZE_PERC", RULES_OP$Max_n_size_perc, "Max_n_size_perc")
    input_data <- check_and_assign(input_data, i, iteration, "MIN_N_SIZE", RULES_OP$Min_n_size, "Min_n_size")
    input_data <- check_and_assign(input_data, i, iteration, "MIN_N_SIZE_PERC", RULES_OP$Min_n_size_perc, "Min_n_size_perc")
    input_data <- check_and_assign(input_data, i, iteration, "SOLUTION_N_SIZE", RULES_OP$seg_n, "seg_n")

    input_data <- check_and_assign(input_data, i, iteration, "PROB_95", RULES_OP$prob_95, "prob_95")
    input_data <- check_and_assign(input_data, i, iteration, "PROB_90", RULES_OP$prob_90, "prob_90")
    input_data <- check_and_assign(input_data, i, iteration, "PROB_80", RULES_OP$prob_80, "prob_80")
    input_data <- check_and_assign(input_data, i, iteration, "PROB_75", RULES_OP$prob_75, "prob_75")
    input_data <- check_and_assign(input_data, i, iteration, "PROB_LESS_THAN_50", RULES_OP$prob_low_50, "prob_low_50")

    input_data <- check_and_assign(input_data, i, iteration, "BIMODAL_VARS", RULES_OP$final_bi_n, "final_bi_n")
    input_data <- check_and_assign(input_data, i, iteration, "PROPER_BUCKETED_VARS", RULES_OP$proper_bucketted_vars, "proper_bucketted_vars")
    input_data <- check_and_assign(input_data, i, iteration, "BIMODAL_VARS_PERC", RULES_OP$bi_n_perc, "bi_n_perc")

    input_data <- check_and_assign(input_data, i, iteration, "ROV_SD", RULES_OP$sd_rov, "sd_rov")
    input_data <- check_and_assign(input_data, i, iteration, "ROV_RANGE", RULES_OP$range_rov, "range_rov")

    input_data <- check_and_assign(input_data, i, iteration, "seg1_diff", RULES_OP$seg1_diff, "seg1_diff")
    input_data <- check_and_assign(input_data, i, iteration, "seg2_diff", RULES_OP$seg2_diff, "seg2_diff")
    input_data <- check_and_assign(input_data, i, iteration, "seg3_diff", RULES_OP$seg3_diff, "seg3_diff")
    input_data <- check_and_assign(input_data, i, iteration, "seg4_diff", RULES_OP$seg4_diff, "seg4_diff")
    input_data <- check_and_assign(input_data, i, iteration, "seg5_diff", RULES_OP$seg5_diff, "seg5_diff")

    # --- bi/perf/indT/indB loops ---
    for (k in 1:5) {
      input_data <- check_and_assign(input_data, i, iteration, paste0("bi_", k),   RULES_OP[[paste0("bi_", k)]],   paste0("bi_", k))
      input_data <- check_and_assign(input_data, i, iteration, paste0("perf_", k), RULES_OP[[paste0("perf_", k)]], paste0("perf_", k))
      input_data <- check_and_assign(input_data, i, iteration, paste0("indT_", k), RULES_OP[[paste0("indT_", k)]], paste0("indT_", k))
      input_data <- check_and_assign(input_data, i, iteration, paste0("indB_", k), RULES_OP[[paste0("indB_", k)]], paste0("indB_", k))
    }

    # --- OF metrics (dynamic) ---
    for (col_name in of_metric_cols) {
      input_data <- check_and_assign(input_data, i, iteration, col_name, RULES_OP[[col_name]], col_name)
    }

    input_data <- propagate_legacy_of_aliases(input_data)

    input_data <- check_and_assign(input_data, i, iteration, "BIMODAL_VARS_INPUT_VARS", RULES_OP$final_bi_n_input, "final_bi_n_input")

    # --- new input performance metrics ---
    input_data <- check_and_assign(input_data, i, iteration, "INPUT_PERF_SEG_COVERED", RULES_OP$input_perf_seg_covered, "input_perf_seg_covered")
    input_data <- check_and_assign(input_data, i, iteration, "INPUT_PERF_FLAG", RULES_OP$input_perf_flag, "input_perf_flag")
    input_data <- check_and_assign(input_data, i, iteration, "INPUT_PERF_MINSEG", RULES_OP$input_perf_minseg, "input_perf_minseg")
    input_data <- check_and_assign(input_data, i, iteration, "INPUT_PERF_MAXSEG", RULES_OP$input_perf_maxseg, "input_perf_maxseg")

    # --- wrap up ---
    input_data <- check_and_assign(input_data, i, iteration, "Ran_Iteration_Flag", 1, "Ran_Iteration_Flag")

    dynamic_of_cols <- sort_of_columns(names(input_data)[grepl(of_pattern, names(input_data))])
    ordered_cols <- c(
      ordered_base_cols,
      tracker_scaffold_cols[tracker_scaffold_cols %in% names(input_data)],
      dynamic_of_cols,
      ordered_input_perf_cols[ordered_input_perf_cols %in% names(input_data)]
    )
    leftover_cols <- setdiff(names(input_data), ordered_cols)
    ordered_cols <- c(ordered_cols, leftover_cols)
    input_data <- input_data[, ordered_cols, drop = FALSE]

    write.csv(input_data, file_path, row.names = FALSE)

    safe_progress(
      processed_iterations,
      total_iterations,
      iteration = iteration,
      status = sprintf(
        "Completed iteration %d (%d of %d)",
        iteration,
        processed_iterations,
        total_iterations
      )
    )

    op<-list()
    print("TABLE FROM LOOP:")
    print(RULES_OP$n_size_table)
    if(solo){
       safe_progress(
         total_iterations,
         total_iterations,
         iteration = iteration,
         status = "Output generation complete."
       )
       op<-list(RULES_OP$seg_n_sizes,RULES_OP$final_bi_n,RULES_OP$bi_n_perc,iteration,solo,RULES_OP$n_size_table)
       return(op)
    }
  }

  safe_progress(
    total_iterations,
    total_iterations,
    iteration = NA_integer_,
    status = "Output generation complete."
  )
}

# External status hook
update_status_run <- function(message) {
  assign("external_status_message", message, envir = .GlobalEnv)
}
