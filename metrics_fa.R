# =========================================================
# metrics_fa.R — Factor analysis metric helpers
# Author: Iterator Automation
# Date: 2025-11-06
# Summary: Utilities to parse FA_OP sheets and compute factor analysis metrics.
# =========================================================

sanitize_factor_label <- function(label, idx) {
  digits <- regmatches(label, regexpr("\\d+", label))
  if (length(digits) && nzchar(digits[1])) {
    return(paste0("F", digits[1]))
  }
  clean <- gsub("[^A-Za-z0-9]", "", label)
  if (!nzchar(clean)) {
    clean <- paste0("Factor", idx)
  }
  clean
}

prepare_fa_context <- function(fa_raw) {
  if (ncol(fa_raw) < 2) {
    stop("FA_OP sheet must contain variable identifiers and at least one factor column.")
  }

  fa_df <- as.data.frame(fa_raw, stringsAsFactors = FALSE)
  variables <- trimws(as.character(fa_df[[1]]))
  variables[is.na(variables)] <- ""

  if (!length(variables)) {
    stop("FA_OP sheet does not contain any variables.")
  }

  loadings_df <- fa_df[, -1, drop = FALSE]
  lower_names <- tolower(names(loadings_df))
  communality_idx <- which(grepl("commun", lower_names))

  communality <- rep(NA_real_, length(variables))
  if (length(communality_idx)) {
    idx <- communality_idx[1]
    communality <- suppressWarnings(as.numeric(loadings_df[[idx]]))
    loadings_df <- loadings_df[, -idx, drop = FALSE]
  }

  if (!ncol(loadings_df)) {
    stop("FA_OP sheet missing factor loading columns.")
  }

  loadings_list <- lapply(loadings_df, function(col) suppressWarnings(as.numeric(col)))
  loadings_mat <- tryCatch({
    do.call(cbind, loadings_list)
  }, error = function(e) NULL)

  if (is.null(loadings_mat) || !nrow(loadings_mat)) {
    stop("Unable to parse factor loading columns in FA_OP sheet.")
  }

  colnames(loadings_mat) <- names(loadings_df)
  abs_loadings <- abs(loadings_mat)

  primary_idx <- rep(NA_integer_, length(variables))
  primary_loading <- rep(NA_real_, length(variables))
  secondary_loading <- rep(0, length(variables))

  for (i in seq_len(nrow(abs_loadings))) {
    row_vals <- abs_loadings[i, ]
    if (all(is.na(row_vals))) {
      next
    }
    order_idx <- order(row_vals, decreasing = TRUE, na.last = NA)
    if (!length(order_idx)) {
      next
    }
    primary_idx[i] <- order_idx[1]
    primary_loading[i] <- row_vals[order_idx[1]]
    if (length(order_idx) >= 2) {
      secondary_loading[i] <- row_vals[order_idx[2]]
    } else {
      secondary_loading[i] <- 0
    }
  }

  factor_labels <- vapply(seq_along(colnames(loadings_mat)), function(i) {
    sanitize_factor_label(colnames(loadings_mat)[i], i)
  }, character(1))
  factor_labels <- make.unique(factor_labels)

  primary_label <- ifelse(is.na(primary_idx), NA_character_, factor_labels[primary_idx])
  cross_ratio <- rep(NA_real_, length(primary_loading))
  valid_primary <- !is.na(primary_loading) & primary_loading > 0
  cross_ratio[valid_primary] <- secondary_loading[valid_primary] / primary_loading[valid_primary]
  cross_ratio[!valid_primary] <- NA_real_

  lookup <- data.frame(
    variable = variables,
    primary_factor = primary_label,
    primary_loading = primary_loading,
    secondary_loading = secondary_loading,
    cross_ratio = cross_ratio,
    communality = communality,
    stringsAsFactors = FALSE
  )

  list(
    lookup = lookup,
    factor_labels = factor_labels
  )
}

compute_fa_metrics <- function(combo_vars, ctx) {
  lookup <- ctx$lookup
  idx <- match(combo_vars, lookup$variable)
  found_mask <- !is.na(idx)
  found <- lookup[idx[found_mask], , drop = FALSE]
  total_vars <- length(combo_vars)
  found_cnt <- nrow(found)
  missing_cnt <- total_vars - found_cnt

  base_metrics <- list(
    FA_DISTINCT_FACTORS = 0L,
    FA_FACTOR_COUNTS = "",
    FA_ENTROPY = 0,
    FA_MEAN_PRIMARY_LOADING = NA_real_,
    FA_MIN_PRIMARY_LOADING = NA_real_,
    FA_MEAN_COMMUNALITY = NA_real_,
    FA_MIN_COMMUNALITY = NA_real_,
    FA_CROSSLOAD_RATIO = NA_real_,
    FA_WEAK_ITEM_COUNT = 0L,
    FA_COVERAGE_MIN_K = 0L,
    FA_MISSING_COUNT = missing_cnt,
    FA_COVERAGE_RATIO = if (total_vars > 0) found_cnt / total_vars else NA_real_
  )

  if (!found_cnt) {
    return(base_metrics)
  }

  factor_counts <- table(found$primary_factor, useNA = "no")
  factor_counts <- factor_counts[names(factor_counts) != ""]
  distinct_factors <- length(factor_counts)
  counts_str <- if (length(factor_counts)) {
    paste(sprintf("%s:%d", names(factor_counts), as.integer(factor_counts)), collapse = ", ")
  } else {
    ""
  }

  entropy <- 0
  if (length(factor_counts)) {
    probs <- as.numeric(factor_counts) / sum(factor_counts)
    good <- probs > 0 & is.finite(probs)
    if (any(good)) {
      entropy <- -sum(probs[good] * log(probs[good]))
    }
  }

  primary_loading <- found$primary_loading
  mean_primary <- if (any(!is.na(primary_loading))) mean(primary_loading, na.rm = TRUE) else NA_real_
  min_primary <- if (any(!is.na(primary_loading))) min(primary_loading, na.rm = TRUE) else NA_real_

  communalities <- found$communality
  mean_comm <- if (any(!is.na(communalities))) mean(communalities, na.rm = TRUE) else NA_real_
  min_comm <- if (any(!is.na(communalities))) min(communalities, na.rm = TRUE) else NA_real_

  cross_ratio <- found$cross_ratio
  cross_ratio_avg <- if (any(!is.na(cross_ratio))) mean(cross_ratio, na.rm = TRUE) else NA_real_

  weak_flags <- rep(FALSE, nrow(found))
  weak_flags <- weak_flags | (ifelse(is.na(primary_loading), FALSE, primary_loading < 0.4))
  weak_flags <- weak_flags | (ifelse(is.na(cross_ratio), FALSE, cross_ratio > 0.75))
  weak_flags <- weak_flags | (ifelse(is.na(communalities), FALSE, communalities < 0.3))
  weak_cnt <- sum(weak_flags)

  min_k <- if (length(factor_counts)) min(as.integer(factor_counts)) else 0L

  base_metrics$FA_DISTINCT_FACTORS <- as.integer(distinct_factors)
  base_metrics$FA_FACTOR_COUNTS <- counts_str
  base_metrics$FA_ENTROPY <- entropy
  base_metrics$FA_MEAN_PRIMARY_LOADING <- mean_primary
  base_metrics$FA_MIN_PRIMARY_LOADING <- min_primary
  base_metrics$FA_MEAN_COMMUNALITY <- mean_comm
  base_metrics$FA_MIN_COMMUNALITY <- min_comm
  base_metrics$FA_CROSSLOAD_RATIO <- cross_ratio_avg
  base_metrics$FA_WEAK_ITEM_COUNT <- as.integer(weak_cnt)
  base_metrics$FA_COVERAGE_MIN_K <- as.integer(min_k)
  base_metrics
}
