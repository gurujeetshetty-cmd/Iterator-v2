# =========================================================
# metrics_pearson.R — Pearson correlation metric helpers
# Author: Iterator Automation
# Date: 2025-11-06
# Summary: Utilities to parse RAW_DATA sheets and compute correlation-based metrics.
# =========================================================

prepare_corr_context <- function(raw_df) {
  if (ncol(raw_df) < 2) {
    stop("RAW_DATA sheet must contain an ID column and at least one variable column.")
  }

  df <- as.data.frame(raw_df, stringsAsFactors = FALSE)
  if (ncol(df) == 1) {
    stop("RAW_DATA sheet does not include variable columns beyond the identifier.")
  }

  data_cols <- df[, -1, drop = FALSE]
  numeric_cols <- lapply(data_cols, function(col) suppressWarnings(as.numeric(col)))
  numeric_df <- as.data.frame(numeric_cols, stringsAsFactors = FALSE)
  colnames(numeric_df) <- colnames(data_cols)

  keep <- vapply(numeric_df, function(col) any(!is.na(col)), logical(1))
  numeric_df <- numeric_df[, keep, drop = FALSE]

  if (!ncol(numeric_df)) {
    stop("RAW_DATA sheet does not contain usable numeric variables.")
  }

  correlation <- suppressWarnings(stats::cor(numeric_df, use = "pairwise.complete.obs"))
  if (is.null(correlation) || !nrow(correlation)) {
    stop("Unable to compute correlation matrix from RAW_DATA sheet.")
  }

  diag(correlation) <- 1

  list(
    correlation = correlation,
    variables = colnames(correlation)
  )
}

compute_corr_metrics <- function(combo_vars, ctx) {
  available <- intersect(combo_vars, ctx$variables)
  metrics <- list(
    P_R_MEAN_ABS = NA_real_,
    P_R_MEDIAN_ABS = NA_real_,
    P_R_MAX_ABS = NA_real_,
    P_EIGVAR_PC1 = NA_real_,
    P_EFF_DIM = NA_real_,
    P_TOPPAIR1 = NA_character_,
    P_TOPPAIR1_R = NA_real_
  )

  if (length(available) < 2) {
    return(metrics)
  }

  sub_cor <- ctx$correlation[available, available, drop = FALSE]
  pair_vals <- sub_cor[upper.tri(sub_cor)]
  finite_mask <- is.finite(pair_vals)
  finite_pairs <- pair_vals[finite_mask]

  if (length(finite_pairs)) {
    abs_pairs <- abs(finite_pairs)
    metrics$P_R_MEAN_ABS <- mean(abs_pairs)
    metrics$P_R_MEDIAN_ABS <- stats::median(abs_pairs)
    metrics$P_R_MAX_ABS <- max(abs_pairs)

    tri_idx <- which(upper.tri(sub_cor), arr.ind = TRUE)
    tri_idx <- tri_idx[finite_mask, , drop = FALSE]
    if (nrow(tri_idx)) {
      order_idx <- order(abs_pairs, decreasing = TRUE)
      top_idx <- tri_idx[order_idx[1], ]
      v1 <- available[top_idx[1]]
      v2 <- available[top_idx[2]]
      metrics$P_TOPPAIR1 <- paste(v1, v2, sep = ";")
      metrics$P_TOPPAIR1_R <- sub_cor[top_idx[1], top_idx[2]]
    }
  }

  if (all(is.finite(sub_cor))) {
    eig_vals <- tryCatch({
      stats::eigen(sub_cor, symmetric = TRUE, only.values = TRUE)$values
    }, error = function(e) NULL)
    if (!is.null(eig_vals)) {
      total <- sum(eig_vals)
      if (total > 0) {
        metrics$P_EIGVAR_PC1 <- eig_vals[1] / total
      }
      denom <- sum(eig_vals^2)
      if (denom > 0) {
        metrics$P_EFF_DIM <- (total^2) / denom
      }
    }
  }

  metrics
}
