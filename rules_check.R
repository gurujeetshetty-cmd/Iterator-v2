library(readxl)
library(dplyr)
library(openxlsx)

RULES_Checker <- function(path, xTABS_name, solo) {
  print("SQC-initial")
  setwd(path)

  # ============================================================
  # CONFIGURATION: Read from iterator.config
  # ============================================================
  iterator_config <- getOption("iterator.config", default = list())
  cfg_num <- function(name, default) {
    val <- iterator_config[[name]]
    if (is.null(val) || !length(val)) return(default)
    val_num <- suppressWarnings(as.numeric(val[[1]]))
    if (is.na(val_num)) default else val_num
  }

  bimodality_hi_threshold <- cfg_num("bimodality_hi_threshold", 110)
  bimodality_low_threshold <- cfg_num("bimodality_low_threshold", 0.2)
  performance_hi_threshold <- cfg_num("performance_hi_threshold", 120)
  performance_lo_threshold <- cfg_num("performance_lo_threshold", 80)
  performance_base_threshold <- cfg_num("performance_base_threshold", 0.2)
  independence_base_threshold <- cfg_num("independence_base_threshold", performance_base_threshold)

  # ============================================================
  # 🧹 Safe data cleaner — removes illegal characters, converts list → string, trims
  # ============================================================
  clean_df_safe <- function(df) {
    df <- as.data.frame(lapply(df, function(col) {
      # Unlist list-columns
      if (is.list(col)) col <- sapply(col, function(x) paste(unlist(x), collapse = " "))
      # Convert to character
      col <- as.character(col)
      # Normalize weird punctuation / spaces
      col <- gsub("[\r\n\t]", " ", col)
      col <- gsub("\u00A0", " ", col)    # non-breaking space
      col <- gsub("[""'']", "\"", col)   # smart quotes
      col <- gsub("[\u2013\u2014]", "-", col) # en/em dash
      col <- gsub("[^[:print:]]", "", col)    # remove hidden non-printables
      trimws(col)
    }), stringsAsFactors = FALSE)
    return(df)
  }

  # Safe numeric conversion
  safe_as_numeric <- function(x) {
    suppressWarnings(as.numeric(gsub("[^0-9\\.-]", "", x)))
  }

  # Safe Excel reader wrapper
  safe_read_excel <- function(file, sheet, col_names = TRUE, na = "") {
    tryCatch({
      df <- read_excel(file, sheet = sheet, col_names = col_names, na = na)
      clean_df_safe(df)
    }, error = function(e) {
      message(sprintf("⚠️ Sheet %s could not be read cleanly: %s", sheet, e$message))
      data.frame()
    })
  }

  # ============================================================
  # Load and clean sheets
  # ============================================================
  segment_mapped <- safe_read_excel(xTABS_name, 1, col_names = TRUE)
  
  # Validate poLCA_SEG column
  if (!"poLCA_SEG" %in% names(segment_mapped)) {
    stop("poLCA_SEG column not found in Sheet 1 of ", xTABS_name)
  }
  segment_mapped$poLCA_SEG <- suppressWarnings(as.numeric(segment_mapped$poLCA_SEG))
  segment_mapped <- segment_mapped[!is.na(segment_mapped$poLCA_SEG), ]
  if (nrow(segment_mapped) == 0) {
    stop("Sheet 1 has no valid segment data in ", xTABS_name)
  }

  seg_n <- max(segment_mapped$poLCA_SEG)

  seg_prob <- safe_read_excel(xTABS_name, 3, col_names = TRUE)
  xTabs <- safe_read_excel(xTABS_name, 2, col_names = FALSE)

  if (nrow(segment_mapped) == 0 || ncol(segment_mapped) == 0) {
    stop("❌ ERROR: 'segment_mapped' sheet could not be read properly.")
  }

  # Calculate segment sizes
  segment_mapped_1 <- segment_mapped %>%
    dplyr::group_by(poLCA_SEG) %>% 
    dplyr::count(poLCA_SEG)
  seg_n_sizes <- segment_mapped_1
  Max_n_size <- max(segment_mapped_1$n)
  Max_n_size_perc <- Max_n_size / sum(segment_mapped_1$n)
  Min_n_size <- min(segment_mapped_1$n)
  Min_n_size_perc <- Min_n_size / sum(segment_mapped_1$n)

  # Calculate probability metrics
  seg_prob[] <- lapply(seg_prob, safe_as_numeric)
  seg_prob$max_prob <- apply(seg_prob[, 2:(seg_n + 1)], 1, max, na.rm = TRUE)
  prob_95 <- mean(seg_prob$max_prob > 0.95, na.rm = TRUE)
  prob_90 <- mean(seg_prob$max_prob > 0.9, na.rm = TRUE)
  prob_80 <- mean(seg_prob$max_prob > 0.8, na.rm = TRUE)
  prob_75 <- mean(seg_prob$max_prob > 0.75, na.rm = TRUE)
  prob_low_50 <- mean(seg_prob$max_prob < 0.5, na.rm = TRUE)

  # ============================================================
  # Split and process xTabs safely
  # ============================================================
  n_size_table <- xTabs[c(1, 4, 5), 2:(seg_n + 2)]

  find_blank_row_indices <- function(df) {
    blank_rows <- apply(df, 1, function(row) all(is.na(row) | row == ""))
    which(blank_rows)
  }

  split_by_blank_rows <- function(df) {
    blank_indices <- find_blank_row_indices(df)
    indices <- c(1, blank_indices + 1, nrow(df) + 1)
    tables <- mapply(function(start, end) df[start:(end - 1), ],
      indices[-length(indices)], indices[-1], SIMPLIFY = FALSE)
    return(tables)
  }

  # Safety checks before slicing xTabs
  if (is.null(xTabs) || nrow(xTabs) == 0) {
    stop("[rules_check.R] ❌ xTabs is empty or could not be read. Check the Excel sheet name and structure.")
  }

  if (is.null(seg_n) || is.na(seg_n) || seg_n <= 0) {
    stop(sprintf("[rules_check.R] ❌ seg_n is invalid (value: %s). Check sheet 1 of %s for poLCA_SEG column.", seg_n, xTABS_name))
  }

  last_row_index <- nrow(xTabs[, 1, drop = FALSE])
  if (is.na(last_row_index) || last_row_index < 28) {
    stop(sprintf("[rules_check.R] ❌ xTabs appears too short (%d rows). Check input file format.", last_row_index))
  }

  col_limit <- 3 + 2 * seg_n
  if (col_limit > ncol(xTabs)) {
    warning(sprintf("[rules_check.R] ⚠️ Expected %d columns based on seg_n=%d, but xTabs has only %d. Adjusting.",
                    col_limit, seg_n, ncol(xTabs)))
    col_limit <- ncol(xTabs)
  }

  # Safe subsetting
  subset_start <- min(28, nrow(xTabs))
  subset_end <- max(28, last_row_index)

  tables <- split_by_blank_rows(xTabs[subset_start:subset_end, 1:col_limit])
  tables_ori <- lapply(tables, clean_df_safe)

  contains_all_patterns <- function(df) {
    patterns <- c("A_", "C_")
    results <- sapply(patterns, function(pattern) {
      any(sapply(df, function(col) any(grepl(pattern, col, ignore.case = TRUE))))
    })
    all(results)
  }

  tables_cleaned_input <- list()
  for (i in seq_along(tables_ori)) {
    ti  <- tables_ori[[i]]
    hdr <- ti[1, 1, drop = TRUE]
    if (contains_all_patterns(ti) && is.character(hdr) && grepl("^00_SEGM_", hdr)) {
      tables_cleaned_input[[length(tables_cleaned_input) + 1]] <- ti
    }
  }

  tables_cleaned <- Filter(contains_all_patterns, tables_ori)
  proper_bucketted_vars <- length(tables_cleaned)

  # ============================================================
  # ROV calc (with NA resilience)
  # ============================================================
  rov_list_input <- numeric()
  for (i in seq_along(tables_ori)) {
    if (grepl("00_SEGM_", tables_ori[[i]][1, 1])) {
      rov_val <- suppressWarnings(safe_as_numeric(sub("ROV=", "", tables_ori[[i]][3, 1])))
      if (!is.na(rov_val)) rov_list_input <- c(rov_list_input, rov_val)
    }
  }
  sd_rov <- ifelse(length(rov_list_input) > 1, sd(rov_list_input), NA)
  range_rov <- ifelse(length(rov_list_input) > 1, diff(range(rov_list_input)), NA)

  # ============================================================
  # OPTIMIZED: Segment-level differentiating variables
  # Using matrices instead of assign/get
  # ============================================================
  cat("[RULES] Calculating segment differentiation metrics...\n")
  
  # Pre-allocate matrices for all metrics
  n_tables <- length(tables_cleaned)
  bi_matrix <- matrix(0, nrow = n_tables, ncol = seg_n)
  perf_matrix <- matrix(0, nrow = n_tables, ncol = seg_n)
  indT_matrix <- matrix(0, nrow = n_tables, ncol = seg_n)
  indB_matrix <- matrix(0, nrow = n_tables, ncol = seg_n)
  seg_diff_flags <- matrix(FALSE, nrow = n_tables, ncol = seg_n)

  # Cache A_row and C_row for each table (avoid repeated pattern matching)
  ab_rows_cache <- vector("list", n_tables)
  for(i in seq_along(tables_cleaned)){
    d1 <- tables_cleaned[[i]]
    ab_rows_cache[[i]] <- list(
      A = d1[apply(d1, 1, function(row) any(grepl("A_", row, fixed = TRUE))), ],
      C = d1[apply(d1, 1, function(row) any(grepl("C_", row, fixed = TRUE))), ]
    )
  }

  # Vectorized calculations
  for(i in seq_along(tables_cleaned)){
    A_row <- ab_rows_cache[[i]]$A
    C_row <- ab_rows_cache[[i]]$C
    
    if (nrow(A_row) == 0 || nrow(C_row) == 0) next
    
    for(j in 1:seg_n){
      # Safe numeric conversion for this segment
      A_hi <- safe_as_numeric(A_row[3 + seg_n + j])
      C_hi <- safe_as_numeric(C_row[3 + seg_n + j])
      A_lo <- safe_as_numeric(A_row[1 + j])
      C_lo <- safe_as_numeric(C_row[1 + j])
      A_base <- safe_as_numeric(A_row[2 + seg_n])
      C_base <- safe_as_numeric(C_row[2 + seg_n])
      
      # Check bimodality
      if (!is.na(A_hi) && !is.na(C_hi) && !is.na(A_lo) && !is.na(C_lo)) {
        if (A_hi > bimodality_hi_threshold && C_hi > bimodality_hi_threshold && 
            A_lo > bimodality_low_threshold && C_lo > bimodality_low_threshold) {
          bi_matrix[i, j] <- 1
        }
      }
      
      # Check performance
      if (!is.na(A_hi) && !is.na(C_hi) && !is.na(A_base) && !is.na(C_base)) {
        if ((A_hi > performance_hi_threshold && C_hi < performance_lo_threshold && A_base > performance_base_threshold) ||
            (A_hi < performance_lo_threshold && C_hi > performance_hi_threshold && C_base > performance_base_threshold)) {
          perf_matrix[i, j] <- 1
        }
      }
      
      # Check independent top
      if (!is.na(A_hi) && !is.na(C_hi) && !is.na(A_base) && !is.na(C_base)) {
        if (((A_hi > performance_hi_threshold && A_base > independence_base_threshold) ||
             (C_hi > performance_hi_threshold && C_base > independence_base_threshold)) && 
            bi_matrix[i, j] != 1) {
          indT_matrix[i, j] <- 1
        }
      }
      
      # Check independent bottom
      if (!is.na(A_hi) && !is.na(C_hi) && !is.na(A_base) && !is.na(C_base)) {
        if (((A_hi < performance_lo_threshold && A_base > independence_base_threshold) ||
             (C_hi < performance_lo_threshold && C_base > independence_base_threshold)) && 
            bi_matrix[i, j] != 1) {
          indB_matrix[i, j] <- 1
        }
      }
    }
  }

  # Calculate segment-level sums
  bi_1 <- sum(bi_matrix[, 1]); perf_1 <- sum(perf_matrix[, 1])
  indT_1 <- sum(indT_matrix[, 1]); indB_1 <- sum(indB_matrix[, 1])
  
  if (seg_n >= 2) {
    bi_2 <- sum(bi_matrix[, 2]); perf_2 <- sum(perf_matrix[, 2])
    indT_2 <- sum(indT_matrix[, 2]); indB_2 <- sum(indB_matrix[, 2])
  } else {
    bi_2 <- 0; perf_2 <- 0; indT_2 <- 0; indB_2 <- 0
  }
  
  if (seg_n >= 3) {
    bi_3 <- sum(bi_matrix[, 3]); perf_3 <- sum(perf_matrix[, 3])
    indT_3 <- sum(indT_matrix[, 3]); indB_3 <- sum(indB_matrix[, 3])
  } else {
    bi_3 <- 0; perf_3 <- 0; indT_3 <- 0; indB_3 <- 0
  }
  
  if (seg_n >= 4) {
    bi_4 <- sum(bi_matrix[, 4]); perf_4 <- sum(perf_matrix[, 4])
    indT_4 <- sum(indT_matrix[, 4]); indB_4 <- sum(indB_matrix[, 4])
  } else {
    bi_4 <- 0; perf_4 <- 0; indT_4 <- 0; indB_4 <- 0
  }
  
  if (seg_n >= 5) {
    bi_5 <- sum(bi_matrix[, 5]); perf_5 <- sum(perf_matrix[, 5])
    indT_5 <- sum(indT_matrix[, 5]); indB_5 <- sum(indB_matrix[, 5])
  } else {
    bi_5 <- 0; perf_5 <- 0; indT_5 <- 0; indB_5 <- 0
  }

  # Calculate seg_diff metrics
  seg_diff <- integer(seg_n)
  diffTB <- 0
  for(j in 1:seg_n){
    for(i in 1:n_tables){
      if (bi_matrix[i, j] == 0 && (perf_matrix[i, j] == 1 || indT_matrix[i, j] == 1 || indB_matrix[i, j] == 1)) {
        seg_diff[j] <- seg_diff[j] + 1
        if (!seg_diff_flags[i, j]) {
          diffTB <- diffTB + 1
          seg_diff_flags[i, j] <- TRUE
        }
        break
      }
    }
  }
  
  seg1_diff <- if (seg_n >= 1) seg_diff[1] else 0
  seg2_diff <- if (seg_n >= 2) seg_diff[2] else 0
  seg3_diff <- if (seg_n >= 3) seg_diff[3] else 0
  seg4_diff <- if (seg_n >= 4) seg_diff[4] else 0
  seg5_diff <- if (seg_n >= 5) seg_diff[5] else 0

  cat("[RULES] ✅ Segment differentiation complete\n")

  # ============================================================
  # Objective Function (OF_*) metrics — dynamic extraction
  # ============================================================

  extract_of_metrics <- function(tbl, seg_n) {
    result <- list(
      values = numeric(0),
      values_display = "-",
      diff_max = 0,
      diff_min = 0
    )

    if (is.null(tbl) || !nrow(tbl) || !ncol(tbl)) return(result)

    tbl <- tbl[rowSums(is.na(tbl)) != ncol(tbl), , drop = FALSE]
    if (nrow(tbl) <= 4) return(result)

    # Drop header rows (first 3 rows)
    drop_count <- min(3, nrow(tbl) - 1)
    if (drop_count > 0) tbl <- tbl[-seq_len(drop_count), , drop = FALSE]
    if (nrow(tbl) <= 1) return(result)

    # Drop last row (totals)
    tbl <- tbl[-nrow(tbl), , drop = FALSE]
    if (!nrow(tbl)) return(result)

    # Convert to numeric
    tbl[] <- lapply(tbl, function(col) suppressWarnings(as.numeric(as.character(col))))
    if (all(vapply(tbl, function(col) all(is.na(col)), logical(1)))) return(result)

    # First row contains weights
    weights <- suppressWarnings(as.numeric(tbl[1, ]))
    if (all(is.na(weights))) return(result)

    values <- numeric(0)
    upper_bound <- min(seg_n + 1, nrow(tbl))
    
    # Calculate weighted average for each segment column
    for (idx in 2:upper_bound) {
      row_vals <- suppressWarnings(as.numeric(tbl[idx, ]))
      if (all(is.na(row_vals))) next

      valid <- !is.na(row_vals) & !is.na(weights)
      if (!any(valid)) next

      row_vals_valid <- row_vals[valid]
      weights_valid <- weights[valid]
      weight_sum <- sum(weights_valid)

      if (is.na(weight_sum) || !is.finite(weight_sum) || weight_sum == 0) {
        sp_v <- mean(row_vals_valid, na.rm = TRUE)
      } else {
        sp_v <- sum(weights_valid * row_vals_valid) / weight_sum
      }

      values <- c(values, sp_v)
    }

    if (!length(values)) return(result)

    values <- values[!is.na(values)]
    if (!length(values)) return(result)

    # Format values
    format_values <- function(vals) {
      if (!length(vals)) return("-")
      formatted <- sprintf("%.2f", vals)
      formatted <- sub("0+$", "", formatted)
      formatted <- sub("\\.$", "", formatted)
      paste(formatted, collapse = " | ")
    }

    result$values <- values
    result$values_display <- format_values(values)
    
    if (length(values) <= 1) {
      result$diff_max <- 0
      result$diff_min <- 0
      return(result)
    }

    # Calculate differences
    pairwise_differences <- outer(values, values, FUN = function(x, y) abs(x - y))
    result$diff_max <- suppressWarnings(max(pairwise_differences, na.rm = TRUE))
    diffs <- pairwise_differences[lower.tri(pairwise_differences)]
    diffs <- diffs[!is.na(diffs) & diffs != 0]
    result$diff_min <- if (length(diffs)) suppressWarnings(min(diffs, na.rm = TRUE)) else 0
    
    if (!is.finite(result$diff_max)) result$diff_max <- 0
    if (!is.finite(result$diff_min)) result$diff_min <- 0

    result
  }

  cat("\n[OF_METRICS] Scanning", length(tables_ori), "tables for OF_ metrics...\n")
  of_metrics <- list()

  for (tbl_idx in seq_along(tables_ori)) {
    tbl <- tables_ori[[tbl_idx]]
    header_val <- tbl[1, 1]
    if (!is.character(header_val) || !length(header_val)) next
    
    # Look for OF_ pattern
    match_loc <- regexpr("OF_[A-Za-z0-9_]+", header_val, ignore.case = TRUE)
    if (match_loc[1] == -1) next
    
    metric_name <- substr(header_val, match_loc[1], match_loc[1] + attr(match_loc, "match.length") - 1)

    # Normalize metric key
    metric_key <- gsub("[^A-Za-z0-9_]", "", metric_name)
    metric_key <- gsub("_+", "_", metric_key)
    metric_key <- sub("_$", "", metric_key)
    metric_key <- sub("^_", "", metric_key)
    if (!nchar(metric_key)) next
    
    metric_key <- toupper(metric_key)
    if (grepl("^OF[A-Z0-9]", metric_key) && !grepl("^OF_", metric_key)) {
      metric_key <- sub("^OF", "OF_", metric_key)
    }
    if (!grepl("^OF_", metric_key)) next

    # Extract base key
    base_key <- sub("(_DIFF.*|_VAL(UES)?|_VALUE.*|_MEAN.*|_MIN.*|_MAX.*)$", "", metric_key)
    base_key <- sub("_+$", "", base_key)
    if (!nchar(base_key)) base_key <- metric_key
    if (base_key %in% names(of_metrics)) next

    cat(sprintf("[OF_METRICS] Table %d: Found metric '%s'\n", tbl_idx, base_key))
    
    # Extract metrics
    metrics <- extract_of_metrics(tbl, seg_n)
    of_metrics[[base_key]] <- metrics
    
    cat(sprintf("[OF_METRICS] Stored: %s_VALUES = '%s'\n", base_key, metrics$values_display))
  }

  cat(sprintf("\n[OF_METRICS] Total OF_ metrics found: %d\n", length(of_metrics)))

  # ============================================================
  # Check for bimodality across cleaned vars
  # ============================================================
  bi_n <- 0
  for(i in 1:length(tables_cleaned)){
    for(j in 1:seg_n){
      if (bi_matrix[i, j] == 1) {
        bi_n <- bi_n + 1
        break
      }
    }
  }
  final_bi_n <-  bi_n
  bi_n_perc <- bi_n / proper_bucketted_vars
  
  n_size_table <- as.data.frame(n_size_table)
  colnames(n_size_table) <- n_size_table[1,]
  n_size_table <- n_size_table[-1,]

  # ============================================================
  # Check for bimodality among input variables
  # ============================================================
  bi_n_input <- 0

  if (length(tables_cleaned_input) == 0) {
    cat("⚠️ WARNING: No valid input tables found for bimodality check.\n")
  } else {
    for (i in seq_along(tables_cleaned_input)) {
      d1 <- tables_cleaned_input[[i]]
      if (ncol(d1) < (3 + 2 * seg_n)) {
        next
      }
      for (j in 1:seg_n) {
        A_row <- d1[apply(d1, 1, function(row) any(grepl("A_", row, fixed = TRUE))), , drop = FALSE]
        C_row <- d1[apply(d1, 1, function(row) any(grepl("C_", row, fixed = TRUE))), , drop = FALSE]
        if (nrow(A_row) == 0 || nrow(C_row) == 0) next

        A_hi <- suppressWarnings(as.numeric(A_row[3 + seg_n + j]))
        C_hi <- suppressWarnings(as.numeric(C_row[3 + seg_n + j]))
        A_lo <- suppressWarnings(as.numeric(A_row[1 + j]))
        C_lo <- suppressWarnings(as.numeric(C_row[1 + j]))

        if (any(is.na(c(A_hi, C_hi, A_lo, C_lo)))) next

        if ((A_hi > bimodality_hi_threshold) && (C_hi > bimodality_hi_threshold) && 
            (A_lo > bimodality_low_threshold) && (C_lo > bimodality_low_threshold)) {
          bi_n_input <- bi_n_input + 1
          break
        }
      }
    }
  }

  final_bi_n_input <- if (is.null(bi_n_input) || length(bi_n_input) == 0 || is.na(bi_n_input)) 0 else bi_n_input

  # ============================================================
  # Input Performance Coverage Metrics
  # ============================================================
  input_perf_seg_covered <- 0
  input_perf_flag <- 0
  input_perf_minseg <- 0
  input_perf_maxseg <- 0

  if (length(tables_cleaned_input) > 0) {
    covered_by_segment <- integer(seg_n)
    perf_count_by_segment <- integer(seg_n)
    
    for (j in 1:seg_n) {
      covered <- FALSE
      perf_count <- 0
      
      for (i in seq_along(tables_cleaned_input)) {
        d1 <- tables_cleaned_input[[i]]
        A_row <- d1[apply(d1, 1, function(row) any(grepl("A_", row, fixed = TRUE))), ]
        C_row <- d1[apply(d1, 1, function(row) any(grepl("C_", row, fixed = TRUE))), ]
        
        if (nrow(A_row)==0||nrow(C_row)==0) next
        if (ncol(A_row)<(3+seg_n+j)||ncol(C_row)<(3+seg_n+j)) next
        
        A_hi <- as.numeric(A_row[3+seg_n+j])
        C_hi <- as.numeric(C_row[3+seg_n+j])
        A_base <- as.numeric(A_row[2+seg_n])
        C_base <- as.numeric(C_row[2+seg_n])
        
        if (!any(is.na(c(A_hi, C_hi, A_base, C_base)))) {
          if (((A_hi > performance_hi_threshold && C_hi < performance_lo_threshold && A_base > performance_base_threshold) ||
               (A_hi < performance_lo_threshold && C_hi > performance_hi_threshold && C_base > performance_base_threshold))) {
            covered <- TRUE
            perf_count <- perf_count + 1
          }
        }
      }
      
      covered_by_segment[j] <- as.integer(covered)
      perf_count_by_segment[j] <- perf_count
    }
    
    input_perf_seg_covered <- sum(covered_by_segment, na.rm = TRUE)
    input_perf_flag <- as.integer(input_perf_seg_covered == seg_n)
    input_perf_minseg <- min(perf_count_by_segment, na.rm = TRUE)
    input_perf_maxseg <- max(perf_count_by_segment, na.rm = TRUE)
  }

  # ============================================================
  # FINAL OUTPUT - UPPERCASE NAMING FOR COMPATIBILITY
  # ============================================================
  normalize_metric_value <- function(x, fallback = NA_real_) {
    if (is.null(x) || !length(x)) return(fallback)
    x_val <- x[1]
    if (is.na(x_val)) return(fallback)
    if (!is.finite(x_val)) return(fallback)
    x_val
  }

  normalize_metric_string <- function(x, fallback = "-") {
    if (is.null(x) || !length(x)) return(fallback)
    x_val <- x[1]
    if (is.na(x_val)) return(fallback)
    if (!nchar(x_val)) return(fallback)
    x_val
  }

  # CRITICAL: Use UPPERCASE for OF_ metric names to match Iteration_loop.R expectations
  of_output <- list()
  metric_keys <- sort(names(of_metrics))
  for (metric_key in metric_keys) {
    metric <- of_metrics[[metric_key]]
    display <- normalize_metric_string(metric$values_display)
    diff_max <- normalize_metric_value(round(metric$diff_max, 2), 0)
    diff_min <- normalize_metric_value(round(metric$diff_min, 2), 0)

    # UPPERCASE suffixes: VALUES, MAXDIFF, MINDIFF
    of_output[[paste0(metric_key, "_VALUES")]] <- display
    of_output[[paste0(metric_key, "_MAXDIFF")]] <- diff_max
    of_output[[paste0(metric_key, "_MINDIFF")]] <- diff_min
  }

  # Build final output
  ITR_analysis_output <- c(
    list(
      bi_n_perc=bi_n_perc, final_bi_n=final_bi_n, proper_bucketted_vars=proper_bucketted_vars,
      seg_n=seg_n, Max_n_size=Max_n_size, Max_n_size_perc=Max_n_size_perc, Min_n_size=Min_n_size,
      Min_n_size_perc=Min_n_size_perc, prob_95=prob_95, prob_90=prob_90, prob_80=prob_80, prob_75=prob_75,
      prob_low_50=prob_low_50, seg_n_sizes=seg_n_sizes, solo=solo, sd_rov=sd_rov, range_rov=range_rov,
      n_size_table=n_size_table, seg1_diff=seg1_diff, seg2_diff=seg2_diff, seg3_diff=seg3_diff,
      seg4_diff=seg4_diff, seg5_diff=seg5_diff, bi_1=bi_1, perf_1=perf_1, indT_1=indT_1, indB_1=indB_1,
      bi_2=bi_2, perf_2=perf_2, indT_2=indT_2, indB_2=indB_2, bi_3=bi_3, perf_3=perf_3, indT_3=indT_3,
      indB_3=indB_3, bi_4=bi_4, perf_4=perf_4, indT_4=indT_4, indB_4=indB_4, bi_5=bi_5, perf_5=perf_5,
      indT_5=indT_5, indB_5=indB_5
    ),
    of_output,
    list(
      final_bi_n_input=final_bi_n_input,
      input_perf_seg_covered=input_perf_seg_covered, input_perf_flag=input_perf_flag,
      input_perf_minseg=input_perf_minseg, input_perf_maxseg=input_perf_maxseg
    )
  )
  
  return(ITR_analysis_output)
}
