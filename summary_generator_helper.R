# summary_generator_helper.R
# Author: OpenAI Assistant
# Date: 2024-11-17
# Utility functions for SUMMARY_GENERATOR module.

# ---- Formatting helpers ----
format_percent <- function(value, digits = 0) {
  if (is.null(value) || is.na(value)) {
    return("0%")
  }
  sprintf(paste0("%.", digits, "f%%"), round(as.numeric(value), digits))
}

normalize_theme_label <- function(theme_label, fallback = "General Insights") {
  if (is.null(theme_label) || all(is.na(theme_label))) {
    return(fallback)
  }
  label <- trimws(as.character(theme_label))
  if (!nzchar(label)) {
    return(fallback)
  }
  label
}

is_not_clear_theme <- function(theme_label) {
  if (is.null(theme_label) || all(is.na(theme_label))) {
    return(FALSE)
  }
  label <- trimws(as.character(theme_label))
  if (!nzchar(label)) {
    return(FALSE)
  }
  grepl("not\\s*clear", label, ignore.case = TRUE)
}

safe_numeric <- function(x) {
  vals <- suppressWarnings(as.numeric(x))
  vals[is.na(vals)] <- NA_real_
  vals
}

weighted_mean_safe <- function(x, w) {
  x <- safe_numeric(x)
  w <- safe_numeric(w)
  valid <- !is.na(x) & !is.na(w)
  if (!any(valid)) {
    return(NA_real_)
  }
  sum(x[valid] * w[valid]) / sum(w[valid])
}

weighted_sd_safe <- function(x, w) {
  x <- safe_numeric(x)
  w <- safe_numeric(w)
  valid <- !is.na(x) & !is.na(w)
  if (!any(valid)) {
    return(NA_real_)
  }
  m <- weighted_mean_safe(x[valid], w[valid])
  sqrt(sum(w[valid] * (x[valid] - m)^2) / sum(w[valid]))
}

# ---- Segment size calculator ----
compute_segment_percentages <- function(data, seg_col = "SEGM", weight_col = "Weights") {
  if (!seg_col %in% names(data)) stop(sprintf("Missing segment column: %s", seg_col))
  if (!weight_col %in% names(data)) stop(sprintf("Missing weight column: %s", weight_col))

  data[[weight_col]] <- safe_numeric(data[[weight_col]])
  grouped <- split(seq_len(nrow(data)), data[[seg_col]])
  totals <- vapply(grouped, function(idx) sum(data[[weight_col]][idx], na.rm = TRUE), numeric(1))
  overall <- sum(totals, na.rm = TRUE)
  perc <- if (overall > 0) (totals / overall) * 100 else rep(0, length(totals))
  data.frame(segment = names(totals), percent = perc, stringsAsFactors = FALSE)
}

# ---- Box score calculations ----
calculate_box_shares <- function(values, weights, top_values, bottom_values = NULL) {
  values <- safe_numeric(values)
  weights <- safe_numeric(weights)
  valid <- !is.na(values) & !is.na(weights)
  if (!any(valid)) {
    return(list(top = 0, bottom = 0))
  }
  total_w <- sum(weights[valid])
  top_share <- if (length(top_values)) sum(weights[valid & values %in% top_values]) / total_w * 100 else 0
  bottom_share <- if (length(bottom_values)) sum(weights[valid & values %in% bottom_values]) / total_w * 100 else 0
  list(top = top_share, bottom = bottom_share)
}

calculate_rating_metrics <- function(values, weights, type) {
  values <- safe_numeric(values)
  weights <- safe_numeric(weights)
  if (toupper(type) == "RATING7") {
    t2b <- calculate_box_shares(values, weights, c(6, 7), c(1, 2))
    t3b <- calculate_box_shares(values, weights, c(5, 6, 7), c(1, 2, 3))
    list(t2b = t2b$top, b2b = t2b$bottom, t3b = t3b$top, b3b = t3b$bottom)
  } else {
    t2b <- calculate_box_shares(values, weights, c(4, 5), c(1, 2))
    list(t2b = t2b$top, b2b = t2b$bottom, t3b = NA_real_, b3b = NA_real_)
  }
}

select_rating_phrase <- function(metrics, type) {
  eval_phrase <- function(top_share, bottom_share) {
    if (!is.na(top_share) && top_share >= 70) {
      return(list(phrase = "strongly agrees that", metric_value = top_share, metric_type = "top"))
    }
    if (!is.na(top_share) && top_share >= 50) {
      return(list(phrase = "somewhat agrees that", metric_value = top_share, metric_type = "top"))
    }
    if (!is.na(bottom_share) && bottom_share >= 50) {
      return(list(phrase = "strongly disagrees that", metric_value = bottom_share, metric_type = "bottom"))
    }
    if (!is.na(bottom_share) && bottom_share >= 30) {
      return(list(phrase = "somewhat disagrees that", metric_value = bottom_share, metric_type = "bottom"))
    }
    if (!is.na(top_share) && !is.na(bottom_share) && top_share >= 30 && bottom_share >= 30) {
      return(list(phrase = "is divided on whether", metric_value = top_share, metric_type = "top"))
    }
    list(phrase = "is uncertain about whether", metric_value = top_share, metric_type = "top")
  }

  first_pass <- eval_phrase(metrics$t2b, metrics$b2b)
  if (toupper(type) == "RATING7" && identical(first_pass$phrase, "is uncertain about whether")) {
    fallback <- eval_phrase(metrics$t3b, metrics$b3b)
    fallback$metric_type <- "top3"
    fallback$used_fallback <- TRUE
    return(fallback)
  }
  first_pass$used_fallback <- FALSE
  first_pass$metric_type <- ifelse(first_pass$metric_type == "top", "top2", first_pass$metric_type)
  first_pass
}

# ---- Index and layer comparisons ----
compute_index <- function(metric_value, overall_metric) {
  if (is.na(overall_metric) || overall_metric <= 0) {
    return(NA_real_)
  }
  metric_value / overall_metric
}

passes_layer_rules <- function(metric_s, metric_x, overall_metric) {
  if (is.na(metric_s) || is.na(metric_x)) {
    return(0L)
  }
  gap <- metric_s - metric_x
  idx_s <- compute_index(metric_s, overall_metric)
  idx_x <- compute_index(metric_x, overall_metric)

  # Layer A
  if (gap >= 15) return(1L)
  if (gap <= -15) return(-1L)

  # Layer B
  if (!is.na(overall_metric) && overall_metric >= 20) {
    if (!is.na(idx_s) && !is.na(idx_x)) {
      if (idx_s >= 1.20 && idx_s >= idx_x) return(1L)
      if (idx_s <= 0.80 && idx_s <= idx_x) return(-1L)
    }
  }

  # Layer C
  if (!is.na(overall_metric) && overall_metric < 20) {
    if (gap >= 8 && !is.na(idx_s) && idx_s >= 1.50) return(1L)
    if (gap <= -8 && !is.na(idx_s) && idx_s <= 0.67) return(-1L)
  }

  0L
}

cross_segment_phrase <- function(target_segment, metrics, overall_metric, digits = 0, format_fn = format_percent) {
  if (!length(metrics)) return("")
  seg_names <- names(metrics)
  target_value <- metrics[[target_segment]]
  others <- seg_names[seg_names != target_segment]
  higher <- character(0)
  lower <- character(0)
  other_metric <- NA_real_
  if (length(others)) {
    other_values <- safe_numeric(unlist(metrics[others]))
    valid_other <- !is.na(other_values)
    if (any(valid_other)) {
      other_metric <- mean(other_values[valid_other])
    }
  }

  for (seg in others) {
    cmp <- passes_layer_rules(target_value, metrics[[seg]], overall_metric)
    if (cmp > 0) {
      higher <- c(higher, seg)
    } else if (cmp < 0) {
      lower <- c(lower, seg)
    }
  }

  fmt <- function(seg) sprintf("Segment %s (%s)", seg, format_fn(metrics[[seg]], digits))

  if (length(higher) && length(higher) == length(others)) {
    return(" compared to all other segments")
  }
  if (length(lower) && length(lower) == length(others)) {
    if (!is.na(other_metric)) {
      return(paste0(" much less so than other segments (", format_fn(other_metric, digits), ")"))
    }
    return(" much less so than other segments")
  }

  if (length(higher) >= 2) {
    segs <- higher[seq_len(min(2, length(higher)))]
    return(paste0(" more than ", paste(vapply(segs, fmt, character(1)), collapse = " and ")))
  }

  if (length(lower) >= 2) {
    segs <- lower[seq_len(min(2, length(lower)))]
    return(paste0(" less than ", paste(vapply(segs, fmt, character(1)), collapse = " and ")))
  }

  ""
}

# ---- Mode and selection metrics ----
calculate_mode_share <- function(values, weights) {
  values <- as.character(values)
  weights <- safe_numeric(weights)
  valid <- !is.na(values) & !is.na(weights) & nzchar(values)
  if (!any(valid)) {
    return(list(mode = NA_character_, share = 0))
  }
  totals <- tapply(weights[valid], values[valid], sum, na.rm = TRUE)
  mode_val <- names(totals)[which.max(totals)]
  share <- totals[[mode_val]] / sum(totals) * 100
  list(mode = mode_val, share = share)
}

calculate_selection_share <- function(values, weights, selected_value = 1) {
  values <- safe_numeric(values)
  weights <- safe_numeric(weights)
  valid <- !is.na(values) & !is.na(weights)
  if (!any(valid)) {
    return(0)
  }
  total <- sum(weights[valid])
  if (total <= 0) return(0)
  sum(weights[valid & values == selected_value]) / total * 100
}

# ---- Narrative builders ----

affix_question <- function(prefix, description, include_prefix = TRUE) {
  prefix <- ifelse(is.na(prefix), "", prefix)
  description <- ifelse(is.na(description), "", description)
  label <- if (include_prefix) paste0(prefix, description) else description
  trimmed <- trimws(label)
  if (!nzchar(trimmed)) return("")
  sprintf("\"%s\"", trimmed)
}

build_rating_narratives <- function(data, meta_row, segments, overall_metrics) {
  narratives <- list()
  segment_metrics <- list()

  for (seg in segments) {
    seg_data <- data[data$SEGM == seg, ]
    metrics <- calculate_rating_metrics(seg_data[[meta_row$variable_id]], seg_data$Weights, meta_row$variable_type)
    phrase_info <- select_rating_phrase(metrics, meta_row$variable_type)

    display_value <- if (phrase_info$metric_type %in% c("top", "top2")) metrics$t2b else metrics$b2b
    display_value <- if (phrase_info$metric_type == "top3") metrics$t3b else display_value
    comparison_value <- if (phrase_info$metric_type == "top3") metrics$t3b else metrics$t2b
    segment_metrics[[seg]] <- comparison_value

    statement <- sprintf(
      "Segment %s %s %s (%s)",
      seg,
      phrase_info$phrase,
      affix_question(meta_row$question_prefix, meta_row$answer_description, include_prefix = TRUE),
      format_percent(display_value)
    )
    narratives[[seg]] <- list(statement = statement, metric_type = phrase_info$metric_type)
  }

  result <- list()
  for (seg in segments) {
    metric_type <- narratives[[seg]]$metric_type
    overall_metric <- if (metric_type == "top3") overall_metrics$top3 else overall_metrics$top2
    cross <- cross_segment_phrase(seg, segment_metrics, overall_metric)
    result[[seg]] <- paste0(narratives[[seg]]$statement, cross, ".")
  }
  result
}

build_singleselect_narratives <- function(data, meta_row, segments, overall_mode_share) {
  narratives <- list()
  mode_shares <- list()

  for (seg in segments) {
    seg_data <- data[data$SEGM == seg, ]
    mode_info <- calculate_mode_share(seg_data[[meta_row$variable_id]], seg_data$Weights)
    mode_shares[[seg]] <- mode_info$share

    phrase <- if (mode_info$share >= 60) {
      "is most likely to choose"
    } else if (mode_info$share >= 40) {
      "is somewhat more likely to choose"
    } else {
      "does not show a clear preference for"
    }

    narratives[[seg]] <- list(
      statement = sprintf(
        "Segment %s %s %s when asked about %s (%s)",
        seg,
        phrase,
        affix_question(meta_row$question_prefix, meta_row$answer_description, include_prefix = FALSE),
        meta_row$question_prefix,
        format_percent(mode_info$share)
      )
    )
  }

  result <- list()
  for (seg in segments) {
    cross <- cross_segment_phrase(seg, mode_shares, overall_mode_share)
    result[[seg]] <- paste0(narratives[[seg]]$statement, cross, ".")
  }
  result
}

build_multiselect_narratives <- function(data, meta_row, segments, overall_selection) {
  narratives <- list()
  selection <- list()

  for (seg in segments) {
    seg_data <- data[data$SEGM == seg, ]
    pct <- calculate_selection_share(seg_data[[meta_row$variable_id]], seg_data$Weights, 1)
    selection[[seg]] <- pct

    phrase <- if (pct >= 60) {
      "most often selects"
    } else if (pct >= 40) {
      "is more likely to select"
    } else {
      "is less likely to select"
    }

    narratives[[seg]] <- sprintf(
      "Segment %s %s %s when considering %s (%s)",
      seg,
      phrase,
      affix_question(meta_row$question_prefix, meta_row$answer_description, include_prefix = FALSE),
      meta_row$question_prefix,
      format_percent(pct)
    )
  }

  result <- list()
  for (seg in segments) {
    cross <- cross_segment_phrase(seg, selection, overall_selection)
    result[[seg]] <- paste0(narratives[[seg]], cross, ".")
  }
  result
}

build_numeric_narratives <- function(data, meta_row, segments, overall_mean, overall_sd) {
  narratives <- list()
  means <- list()

  for (seg in segments) {
    seg_data <- data[data$SEGM == seg, ]
    seg_mean <- weighted_mean_safe(seg_data[[meta_row$variable_id]], seg_data$Weights)
    means[[seg]] <- seg_mean

    phrase <- if (!is.na(seg_mean) && !is.na(overall_sd) && seg_mean >= overall_mean + 0.30 * overall_sd) {
      "has the highest levels of"
    } else if (!is.na(seg_mean) && !is.na(overall_sd) && seg_mean >= overall_mean + 0.10 * overall_sd) {
      "is higher than average on"
    } else if (!is.na(seg_mean) && !is.na(overall_sd) && seg_mean <= overall_mean - 0.30 * overall_sd) {
      "has the lowest levels of"
    } else if (!is.na(seg_mean) && !is.na(overall_sd) && seg_mean <= overall_mean - 0.10 * overall_sd) {
      "is lower than average on"
    } else {
      "is similar to other segments on"
    }

    narratives[[seg]] <- sprintf(
      "Segment %s %s %s (mean = %.2f)",
      seg,
      phrase,
      affix_question(meta_row$question_prefix, meta_row$answer_description, include_prefix = TRUE),
      seg_mean
    )
  }

  result <- list()
  for (seg in segments) {
    cross <- cross_segment_phrase(
      seg,
      means,
      overall_mean,
      digits = 2,
      format_fn = function(x, digits = 2) sprintf(paste0("%.", digits, "f"), round(as.numeric(x), digits))
    )
    result[[seg]] <- paste0(narratives[[seg]], cross, ".")
  }
  result
}

build_ranking_narratives <- function(data, meta_row, segments, overall_rank1) {
  narratives <- list()
  rank1 <- list()

  for (seg in segments) {
    seg_data <- data[data$SEGM == seg, ]
    values <- safe_numeric(seg_data[[meta_row$variable_id]])
    weights <- safe_numeric(seg_data$Weights)
    valid <- !is.na(values) & !is.na(weights)
    total <- sum(weights[valid])
    if (total <= 0) {
      p1 <- 0
      p3 <- 0
    } else {
      p1 <- sum(weights[valid & values == 1]) / total * 100
      p3 <- sum(weights[valid & values %in% 1:3]) / total * 100
    }
    rank1[[seg]] <- p1

    phrase <- if (p1 >= 40) {
      sprintf("is most likely to rank %s as the highest priority for %s", affix_question(meta_row$question_prefix, meta_row$answer_description, include_prefix = FALSE), meta_row$question_prefix)
    } else if (p3 >= 60) {
      sprintf("is more likely to include %s among their top priorities for %s", affix_question(meta_row$question_prefix, meta_row$answer_description, include_prefix = FALSE), meta_row$question_prefix)
    } else {
      sprintf("is less likely to treat %s as a key priority for %s", affix_question(meta_row$question_prefix, meta_row$answer_description, include_prefix = FALSE), meta_row$question_prefix)
    }

    narratives[[seg]] <- sprintf("Segment %s %s (%s)", seg, phrase, format_percent(p1))
  }

  result <- list()
  for (seg in segments) {
    cross <- cross_segment_phrase(seg, rank1, overall_rank1)
    result[[seg]] <- paste0(narratives[[seg]], cross, ".")
  }
  result
}

# ---- TXT export ----
assemble_segment_text <- function(segment_sizes, narratives) {
  segments <- segment_sizes$segment
  output_lines <- character(0)
  for (seg in segments) {
    header <- sprintf("Segment %s (%s)", seg, format_percent(segment_sizes$percent[segment_sizes$segment == seg]))
    output_lines <- c(output_lines, header)
    seg_narratives <- narratives[[seg]]
    if (is.null(seg_narratives) || !length(seg_narratives)) {
      output_lines <- c(output_lines, "")
      next
    }
    output_lines <- c(output_lines, "")
    for (theme_name in names(seg_narratives)) {
      statements <- seg_narratives[[theme_name]]
      if (!length(statements)) next
      output_lines <- c(output_lines, sprintf("%s:", theme_name), statements, "")
    }
    output_lines <- c(output_lines, "")
  }
  output_lines
}

write_summary_txt <- function(lines, output_dir, file_name) {
  if (!dir.exists(output_dir)) {
    stop(sprintf("Output directory does not exist: %s", output_dir))
  }
  output_path <- file.path(output_dir, file_name)
  writeLines(lines, output_path)
  output_path
}

# ---- Main orchestrator ----
parse_summary_input <- function(path) {
  raw <- readxl::read_excel(path, sheet = "SUMMARY_GENERATOR", col_names = FALSE)
  if (nrow(raw) < 7) stop("SUMMARY_GENERATOR sheet missing data rows.")

  headers <- as.character(unlist(raw[1, ]))
  data_rows <- raw[-seq_len(6), ]
  colnames(data_rows) <- headers

  meta <- data.frame(
    variable_id = headers[4:length(headers)],
    theme = as.character(unlist(raw[2, 4:length(headers)])),
    answer_description = as.character(unlist(raw[3, 4:length(headers)])),
    variable_type = as.character(unlist(raw[4, 4:length(headers)])),
    question_prefix = as.character(unlist(raw[5, 4:length(headers)])),
    input_flag = as.character(unlist(raw[6, 4:length(headers)])),
    stringsAsFactors = FALSE
  )

  data_rows$Weights <- safe_numeric(data_rows$Weights)
  list(data = data_rows, metadata = meta)
}

normalize_meta <- function(meta) {
  allowed <- c("RATING7", "RATING5", "SINGLESELECT", "MULTISELECT", "NUMERIC", "RANKING")
  meta$variable_type <- toupper(trimws(meta$variable_type))
  meta <- meta[meta$variable_type %in% allowed, ]
  meta$input_flag <- tolower(trimws(meta$input_flag))
  meta$include <- ifelse(meta$input_flag %in% c("0", "n", "no"), FALSE, TRUE)
  meta[meta$include, ]
}

build_all_narratives <- function(parsed) {
  data <- parsed$data
  meta <- normalize_meta(parsed$metadata)
  segment_sizes <- compute_segment_percentages(data)
  segments <- segment_sizes$segment
  narratives_by_segment <- setNames(lapply(segments, function(x) list()), segments)

  for (i in seq_len(nrow(meta))) {
    meta_row <- meta[i, ]
    if (is_not_clear_theme(meta_row$theme)) {
      next
    }
    theme_label <- normalize_theme_label(meta_row$theme)
    overall_values <- calculate_rating_metrics(data[[meta_row$variable_id]], data$Weights, meta_row$variable_type)
    type <- meta_row$variable_type

    if (type == "RATING7") {
      var_narratives <- build_rating_narratives(data, meta_row, segments, list(top2 = overall_values$t2b, top3 = overall_values$t3b))
    } else if (type == "RATING5") {
      var_narratives <- build_rating_narratives(data, meta_row, segments, list(top2 = overall_values$t2b, top3 = NA_real_))
    } else if (type == "SINGLESELECT") {
      overall_mode <- calculate_mode_share(data[[meta_row$variable_id]], data$Weights)
      var_narratives <- build_singleselect_narratives(data, meta_row, segments, overall_mode$share)
    } else if (type == "MULTISELECT") {
      overall_selection <- calculate_selection_share(data[[meta_row$variable_id]], data$Weights, 1)
      var_narratives <- build_multiselect_narratives(data, meta_row, segments, overall_selection)
    } else if (type == "NUMERIC") {
      overall_mean <- weighted_mean_safe(data[[meta_row$variable_id]], data$Weights)
      overall_sd <- weighted_sd_safe(data[[meta_row$variable_id]], data$Weights)
      var_narratives <- build_numeric_narratives(data, meta_row, segments, overall_mean, overall_sd)
    } else if (type == "RANKING") {
      values <- safe_numeric(data[[meta_row$variable_id]])
      weights <- safe_numeric(data$Weights)
      valid <- !is.na(values) & !is.na(weights)
      total <- sum(weights[valid])
      overall_rank1 <- if (total > 0) sum(weights[valid & values == 1]) / total * 100 else 0
      var_narratives <- build_ranking_narratives(data, meta_row, segments, overall_rank1)
    } else {
      next
    }

    for (seg in segments) {
      seg_bucket <- narratives_by_segment[[seg]][[theme_label]]
      if (is.null(seg_bucket)) {
        seg_bucket <- character(0)
      }
      seg_bucket <- c(seg_bucket, var_narratives[[seg]])
      narratives_by_segment[[seg]][[theme_label]] <- seg_bucket
    }
  }

  list(segment_sizes = segment_sizes, narratives_by_segment = narratives_by_segment)
}

run_summary_generator <- function(file_path, output_dir, output_name) {
  parsed <- parse_summary_input(file_path)
  summary_content <- build_all_narratives(parsed)
  lines <- assemble_segment_text(summary_content$segment_sizes, summary_content$narratives_by_segment)
  output_path <- write_summary_txt(lines, output_dir, output_name)
  list(
    output_path = output_path,
    lines = lines,
    segment_sizes = summary_content$segment_sizes
  )
}

