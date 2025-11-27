# summary_generator_helper.R
# Author: OpenAI Assistant
# Date: 2024-11-17
# Utility functions for SUMMARY_GENERATOR module.

find_neighbor_file <- function(filename) {
  caller_frames <- sys.frames()
  caller_paths <- vapply(caller_frames, function(fr) {
    if (!is.null(fr$ofile)) {
      return(fr$ofile)
    }
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
  if (length(existing)) {
    return(existing[[1]])
  }
  NULL
}

input_utils_path <- find_neighbor_file("input_utils.R")
if (!is.null(input_utils_path)) {
  source(input_utils_path)
} else {
  stop("input_utils.R not found. Please place it alongside summary_generator_helper.R.")
}

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
  if (is.null(theme_label)) {
    return(FALSE)
  }

  labels <- trimws(as.character(theme_label))
  labels[is.na(labels)] <- ""
  if (!length(labels)) {
    return(FALSE)
  }

  nzchar(labels) & grepl("not\\s*clear", labels, ignore.case = TRUE)
}

safe_numeric <- function(x) {
  vals <- suppressWarnings(as.numeric(x))
  vals[is.na(vals)] <- NA_real_
  vals
}

MIN_DIRECTIONAL_GAP <- 10
MIN_MEANINGFUL_SHARE <- 15
MIN_DIFFERENTIATION_GAP <- 5
TOP_DIFFERENTIATOR_COUNT <- 5
INDEX_THRESHOLD_HIGH <- 1.20
INDEX_THRESHOLD_LOW <- 0.80
INDEX_THRESHOLD_SMALL_SEG_HIGH <- 1.50
INDEX_THRESHOLD_SMALL_SEG_LOW <- 0.67

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
compute_segment_percentages <- function(data, seg_col = "SEGM", weight_col = "WEIGHTS") {
  if (!seg_col %in% names(data)) stop(sprintf("Missing segment column: %s", seg_col))
  if (!weight_col %in% names(data)) stop(sprintf("Missing weight column: %s", weight_col))

  data[[weight_col]] <- safe_numeric(data[[weight_col]])
  grouped <- split(seq_len(nrow(data)), data[[seg_col]])
  totals <- vapply(grouped, function(idx) sum(data[[weight_col]][idx], na.rm = TRUE), numeric(1))
  overall <- sum(totals, na.rm = TRUE)
  perc <- if (overall > 0) (totals / overall) * 100 else rep(0, length(totals))
  data.frame(segment = names(totals), percent = perc, stringsAsFactors = FALSE)
}

# ---- Bucket parsing helpers ----
is_prebucketed_rating <- function(values, pattern = "^[A-C]_[TMB]") {
  vals <- as.character(values)
  any(grepl(pattern, vals, ignore.case = TRUE), na.rm = TRUE)
}

parse_rating_bucket <- function(values) {
  vals <- as.character(values)
  bucket <- rep(NA_character_, length(vals))
  bucket[grepl("^A_T", vals, ignore.case = TRUE)] <- "top"
  bucket[grepl("^B_M", vals, ignore.case = TRUE)] <- "middle"
  bucket[grepl("^C_B", vals, ignore.case = TRUE)] <- "bottom"
  bucket
}

count_bucket_span <- function(values, prefix) {
  vals <- as.character(values)
  codes <- vals[grepl(paste0("^", prefix), vals, ignore.case = TRUE) & nzchar(vals)]
  if (!length(codes)) {
    return(0L)
  }
  digits <- gsub("^.{3}", "", unique(codes))
  digit_chars <- unlist(strsplit(paste(digits, collapse = ""), ""))
  length(unique(digit_chars[grepl("[0-9]", digit_chars)]))
}

calculate_prebucketed_rating_metrics <- function(values, weights, type) {
  weights <- safe_numeric(weights)
  buckets <- parse_rating_bucket(values)
  valid <- !is.na(buckets) & !is.na(weights)
  if (!any(valid)) {
    return(list(t2b = NA_real_, b2b = NA_real_, t3b = NA_real_, b3b = NA_real_))
  }
  total_w <- sum(weights[valid])
  share <- function(label) sum(weights[valid & buckets == label]) / total_w * 100
  top_share <- share("top")
  bottom_share <- share("bottom")

  top_span <- count_bucket_span(values, "A_T")
  bottom_span <- count_bucket_span(values, "C_B")

  if (toupper(type) == "RATING7") {
    list(
      t2b = top_share,
      b2b = bottom_share,
      t3b = if (top_span >= 3) top_share else NA_real_,
      b3b = if (bottom_span >= 3) bottom_share else NA_real_
    )
  } else {
    list(t2b = top_share, b2b = bottom_share, t3b = NA_real_, b3b = NA_real_)
  }
}

# ---- Box score calculations ----
calculate_rating_metrics <- function(values, weights, type) {
  if (!is_prebucketed_rating(values)) {
    warning("Rating metrics require pre-bucketed A_T/B_M/C_B inputs; raw ratings are not supported.")
    return(list(t2b = NA_real_, b2b = NA_real_, t3b = NA_real_, b3b = NA_real_))
  }

  calculate_prebucketed_rating_metrics(values, weights, type)
}

select_rating_phrase <- function(metrics, type) {
  # Prefer the narrow boxes (top2/bottom2); fall back to top3/bottom3 for 7-pt scales when needed
  top_share <- metrics$t2b
  bottom_share <- metrics$b2b
  metric_label <- "top2"

  if (toupper(type) == "RATING7" && is.na(top_share) && !is.na(metrics$t3b)) {
    top_share <- metrics$t3b
    metric_label <- "top3"
  }
  if (toupper(type) == "RATING7" && is.na(bottom_share) && !is.na(metrics$b3b)) {
    bottom_share <- metrics$b3b
  }

  gap <- top_share - bottom_share
  top_meaningful <- !is.na(top_share) && top_share >= MIN_MEANINGFUL_SHARE
  bottom_meaningful <- !is.na(bottom_share) && bottom_share >= MIN_MEANINGFUL_SHARE

  # Bimodal: both ends have meaningful share and the gap between them is small
  if (top_meaningful && bottom_meaningful && abs(gap) < MIN_DIRECTIONAL_GAP) {
    return(list(
      phrase = "bimodal split",
      metric_value = top_share,
      secondary_value = bottom_share,
      metric_type = metric_label,
      skip = FALSE,
      variant = "bimodal"
    ))
  }

  if (top_meaningful && top_share >= 70) {
    return(list(phrase = "strong agreement", metric_value = top_share, metric_type = metric_label, skip = FALSE, variant = "directional"))
  }
  if (top_meaningful && top_share >= 50) {
    return(list(phrase = "somewhat agreement", metric_value = top_share, metric_type = metric_label, skip = FALSE, variant = "directional"))
  }
  if (bottom_meaningful && bottom_share >= 70) {
    return(list(phrase = "strong disagreement", metric_value = bottom_share, metric_type = "bottom", skip = FALSE, variant = "directional"))
  }
  if (bottom_meaningful && bottom_share >= 50) {
    return(list(phrase = "somewhat disagreement", metric_value = bottom_share, metric_type = "bottom", skip = FALSE, variant = "directional"))
  }

  list(phrase = "", metric_value = NA_real_, metric_type = metric_label, skip = TRUE, variant = "unclear")
}

parse_dual_statements <- function(answer_description) {
  right <- NA_character_
  left <- NA_character_
  if (!is.null(answer_description) && !is.na(answer_description)) {
    lines <- unlist(strsplit(answer_description, "\n"))
    right_line <- lines[grepl("^\\s*RIGHT:", lines, ignore.case = TRUE)]
    left_line <- lines[grepl("^\\s*LEFT:", lines, ignore.case = TRUE)]
    if (length(right_line)) right <- trimws(sub("^\\s*RIGHT:\\s*", "", right_line[1], ignore.case = TRUE))
    if (length(left_line)) left <- trimws(sub("^\\s*LEFT:\\s*", "", left_line[1], ignore.case = TRUE))
  }
  list(right = right, left = left)
}

calculate_dual_metrics <- function(values, weights) {
  weights <- safe_numeric(weights)
  vals <- as.character(values)
  valid <- !is.na(vals) & !is.na(weights)
  if (!any(valid)) {
    return(list(right = NA_real_, middle = NA_real_, left = NA_real_, net = NA_real_))
  }
  total <- sum(weights[valid])
  share <- function(pattern) sum(weights[valid & grepl(pattern, vals, ignore.case = TRUE)]) / total * 100
  right_share <- share("^A_R")
  middle_share <- share("^B_M")
  left_share <- share("^C_L")
  net <- right_share - left_share
  list(right = right_share, middle = middle_share, left = left_share, net = net)
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

  threshold_high <- getOption("narrative.cross_segment.index_high", INDEX_THRESHOLD_HIGH)
  threshold_low <- getOption("narrative.cross_segment.index_low", INDEX_THRESHOLD_LOW)
  threshold_small_high <- getOption("narrative.cross_segment.small_seg_index_high", INDEX_THRESHOLD_SMALL_SEG_HIGH)
  threshold_small_low <- getOption("narrative.cross_segment.small_seg_index_low", INDEX_THRESHOLD_SMALL_SEG_LOW)

  # Layer A
  if (gap >= 15) return(1L)
  if (gap <= -15) return(-1L)

  # Layer B
  if (!is.na(overall_metric) && overall_metric >= 20) {
    if (!is.na(idx_s) && !is.na(idx_x)) {
      if (idx_s >= threshold_high && idx_s >= idx_x) return(1L)
      if (idx_s <= threshold_low && idx_s <= idx_x) return(-1L)
    }
  }

  # Layer C
  if (!is.na(overall_metric) && overall_metric < 20) {
    if (gap >= 8 && !is.na(idx_s) && idx_s >= threshold_small_high) return(1L)
    if (gap <= -8 && !is.na(idx_s) && idx_s <= threshold_small_low) return(-1L)
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

derive_metric_value <- function(data, meta_row, segment = NULL) {
  subset_data <- if (is.null(segment)) data else data[data$SEGM == segment, , drop = FALSE]
  type <- toupper(meta_row$variable_type)

  if (type %in% c("RATING7", "RATING5")) {
    metrics <- calculate_rating_metrics(subset_data[[meta_row$variable_id]], subset_data$WEIGHTS, type)
    primary <- metrics$t2b
    if (is.na(primary)) primary <- metrics$t3b
    return(primary)
  }
  if (type == "RATING_DUAL") {
    metrics <- calculate_dual_metrics(subset_data[[meta_row$variable_id]], subset_data$WEIGHTS)
    return(metrics$net)
  }
  if (type == "SINGLESELECT") {
    mode_info <- calculate_mode_share(subset_data[[meta_row$variable_id]], subset_data$WEIGHTS)
    return(mode_info$share)
  }
  if (type == "MULTISELECT") {
    return(calculate_selection_share(subset_data[[meta_row$variable_id]], subset_data$WEIGHTS, 1))
  }
  if (type == "NUMERIC") {
    return(weighted_mean_safe(subset_data[[meta_row$variable_id]], subset_data$WEIGHTS))
  }
  if (type == "RANKING") {
    values <- safe_numeric(subset_data[[meta_row$variable_id]])
    weights <- safe_numeric(subset_data$WEIGHTS)
    valid <- !is.na(values) & !is.na(weights)
    total <- sum(weights[valid])
    if (total <= 0) return(NA_real_)
    return(sum(weights[valid & values == 1]) / total * 100)
  }
  NA_real_
}

compute_top_differentiators <- function(data, meta, segments, top_n = TOP_DIFFERENTIATOR_COUNT) {
  theme_map <- setNames(lapply(segments, function(x) list()), segments)
  meta$normalized_theme <- vapply(meta$theme, normalize_theme_label, character(1))

  for (theme_label in unique(meta$normalized_theme)) {
    theme_rows <- meta[meta$normalized_theme == theme_label, , drop = FALSE]
    for (seg in segments) {
      diff_frame <- data.frame(variable_id = character(0), diff = numeric(0), stringsAsFactors = FALSE)
      for (i in seq_len(nrow(theme_rows))) {
        row <- theme_rows[i, ]
        overall_metric <- derive_metric_value(data, row)
        seg_metric <- derive_metric_value(data, row, segment = seg)
        gap <- seg_metric - overall_metric
        abs_gap <- abs(gap)
        if (is.na(abs_gap) || abs_gap < MIN_DIFFERENTIATION_GAP) {
          abs_gap <- -Inf
        }
        diff_frame <- rbind(
          diff_frame,
          data.frame(variable_id = row$variable_id, diff = gap, abs_diff = abs_gap, stringsAsFactors = FALSE)
        )
      }
      meaningful <- diff_frame[diff_frame$abs_diff > -Inf, , drop = FALSE]

      if (!nrow(meaningful)) {
        theme_map[[seg]][[theme_label]] <- character(0)
        next
      }

      meaningful <- meaningful[order(-meaningful$abs_diff), , drop = FALSE]
      pos <- meaningful[meaningful$diff > 0, , drop = FALSE]
      neg <- meaningful[meaningful$diff < 0, , drop = FALSE]

      pos <- pos[order(-pos$abs_diff), , drop = FALSE]
      neg <- neg[order(-neg$abs_diff), , drop = FALSE]

      selected <- character(0)
      if (nrow(pos)) selected <- c(selected, pos$variable_id[1])
      if (nrow(neg) && length(selected) < top_n) selected <- c(selected, neg$variable_id[1])

      remaining <- meaningful$variable_id[!meaningful$variable_id %in% selected]
      if (length(remaining) && length(selected) < top_n) {
        selected <- c(selected, head(remaining, top_n - length(selected)))
      }

      theme_map[[seg]][[theme_label]] <- head(selected, top_n)
    }
  }
  theme_map
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
  segment_values <- list()

  # First pass: collect phrase info for each segment
  for (seg in segments) {
    seg_data <- data[data$SEGM == seg, ]
    metrics <- calculate_rating_metrics(seg_data[[meta_row$variable_id]], seg_data$WEIGHTS, meta_row$variable_type)
    segment_values[[seg]] <- select_rating_phrase(metrics, meta_row$variable_type)
  }

  # Second pass: construct narratives with overall metrics available
  for (seg in segments) {
    phrase_info <- segment_values[[seg]]

    if (isTRUE(phrase_info$skip)) {
      narratives[[seg]] <- list(statement = character(0), metric_type = phrase_info$metric_type)
      next
    }

    question_text <- affix_question(meta_row$question_prefix, meta_row$answer_description, include_prefix = TRUE)

    # Select the right overall comparison metric based on whether we are discussing agreement or disagreement.
    overall_value <- NA_real_
    comparison_label <- "overall avg"
    if (identical(phrase_info$metric_type, "bottom")) {
      if (toupper(meta_row$variable_type) == "RATING7" && is.na(overall_metrics$bottom2) && !is.na(overall_metrics$bottom3)) {
        overall_value <- overall_metrics$bottom3
        comparison_label <- "overall avg disagreement (bottom3)"
      } else {
        overall_value <- overall_metrics$bottom2
        comparison_label <- "overall avg disagreement"
      }
    } else if (identical(phrase_info$metric_type, "top3")) {
      overall_value <- overall_metrics$top3
      comparison_label <- "overall avg (top3)"
    } else {
      overall_value <- overall_metrics$top2
    }

    if (identical(phrase_info$variant, "bimodal")) {
      statement <- sprintf(
        "Segment %s shows a bimodal split with statement %s (agreeing: %s, disagreeing: %s) compared to %s (agreeing: %s, disagreeing: %s)",
        seg,
        question_text,
        format_percent(phrase_info$metric_value),
        format_percent(phrase_info$secondary_value),
        comparison_label,
        if (!is.na(overall_value)) format_percent(overall_value) else "N/A",
        if (!is.na(overall_metrics$bottom2)) format_percent(overall_metrics$bottom2) else if (!is.na(overall_metrics$bottom3)) format_percent(overall_metrics$bottom3) else "N/A"
      )
    } else {
      overindex_note <- ""
      if (!is.na(overall_value)) {
        gap <- phrase_info$metric_value - overall_value
        if (!is.na(gap) && gap >= MIN_DIFFERENTIATION_GAP) {
          overindex_note <- sprintf(
            " overindexes on %s compared to %s (%s, +%s)",
            if (identical(phrase_info$metric_type, "bottom")) "disagreement" else "agreement",
            comparison_label,
            format_percent(overall_value),
            format_percent(gap)
          )
        }
      }

      statement <- sprintf(
        "Segment %s has %s (%s) with statement %s%s",
        seg,
        phrase_info$phrase,
        format_percent(phrase_info$metric_value),
        question_text,
        overindex_note
      )
    }

    narratives[[seg]] <- list(statement = statement, metric_type = phrase_info$metric_type)
  }

  result <- list()
  for (seg in segments) {
    if (!length(narratives[[seg]]$statement)) {
      result[[seg]] <- character(0)
      next
    }
    result[[seg]] <- paste0(narratives[[seg]]$statement, ".")
  }
  result
}

build_dual_narratives <- function(data, meta_row, segments, overall_metrics) {
  narratives <- list()
  segment_metrics <- list()
  statements <- parse_dual_statements(meta_row$answer_description)
  right_text <- ifelse(nzchar(statements$right), affix_question("", statements$right, include_prefix = FALSE), "the right-hand statement")
  left_text <- ifelse(nzchar(statements$left), affix_question("", statements$left, include_prefix = FALSE), "the left-hand statement")

  for (seg in segments) {
    seg_data <- data[data$SEGM == seg, ]
    metrics <- calculate_dual_metrics(seg_data[[meta_row$variable_id]], seg_data$WEIGHTS)
    segment_metrics[[seg]] <- metrics$net

    if (is.na(metrics$net) || (max(metrics$right, metrics$left, na.rm = TRUE) < MIN_MEANINGFUL_SHARE && abs(metrics$net) < MIN_DIRECTIONAL_GAP)) {
      narratives[[seg]] <- list(statement = character(0), metric_type = "net_dual")
      next
    }

    phrase <- "is divided between"
    if (!is.na(metrics$net) && metrics$net >= 25 && metrics$right >= MIN_MEANINGFUL_SHARE) {
      phrase <- sprintf("strongly prefers %s over", right_text)
    } else if (!is.na(metrics$net) && metrics$net >= 10 && metrics$right >= MIN_MEANINGFUL_SHARE) {
      phrase <- sprintf("leans toward %s over", right_text)
    } else if (!is.na(metrics$net) && metrics$net <= -25 && metrics$left >= MIN_MEANINGFUL_SHARE) {
      phrase <- sprintf("strongly prefers %s over", left_text)
    } else if (!is.na(metrics$net) && metrics$net <= -10 && metrics$left >= MIN_MEANINGFUL_SHARE) {
      phrase <- sprintf("leans toward %s over", left_text)
    } else if (!is.na(metrics$net) && abs(metrics$net) < MIN_DIRECTIONAL_GAP) {
      phrase <- "is divided between"
    }

    statement <- sprintf(
      "Segment %s %s %s and %s (Right: %s, Left: %s, Net: %s)",
      seg,
      phrase,
      right_text,
      left_text,
      format_percent(metrics$right),
      format_percent(metrics$left),
      format_percent(metrics$net)
    )
    narratives[[seg]] <- list(statement = statement, metric_type = "net_dual")
  }

  result <- list()
  for (seg in segments) {
    if (!length(narratives[[seg]]$statement)) {
      result[[seg]] <- character(0)
      next
    }
    cross <- cross_segment_phrase(seg, segment_metrics, overall_metrics$net, digits = 0, format_fn = format_percent)
    result[[seg]] <- paste0(narratives[[seg]]$statement, cross, ".")
  }
  result
}

build_singleselect_narratives <- function(data, meta_row, segments, overall_mode_share) {
  narratives <- list()
  mode_shares <- list()

  for (seg in segments) {
    seg_data <- data[data$SEGM == seg, ]
    mode_info <- calculate_mode_share(seg_data[[meta_row$variable_id]], seg_data$WEIGHTS)
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
    pct <- calculate_selection_share(seg_data[[meta_row$variable_id]], seg_data$WEIGHTS, 1)
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
    seg_mean <- weighted_mean_safe(seg_data[[meta_row$variable_id]], seg_data$WEIGHTS)
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
    weights <- safe_numeric(seg_data$WEIGHTS)
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
  parsed <- read_standardized_input(path, sheet = "SUMMARY_GENERATOR", na_values = c("NA", ".", "", " "))
  data_rows <- parsed$data
  metadata <- parsed$metadata
  metadata <- metadata[metadata$column_index >= 4, , drop = FALSE]
  if (!nrow(metadata)) {
    stop("SUMMARY_GENERATOR sheet does not contain variable columns beyond respondent metadata.")
  }

  if ("WEIGHTS" %in% names(data_rows)) {
    data_rows$WEIGHTS <- safe_numeric(data_rows$WEIGHTS)
  }

  list(data = data_rows, metadata = metadata)
}

normalize_meta <- function(meta, use_input_flag = TRUE) {
  allowed <- c("RATING7", "RATING5", "RATING_DUAL", "SINGLESELECT", "MULTISELECT", "NUMERIC", "RANKING")
  meta$variable_type <- clean_variable_types(meta$variable_type)
  meta <- meta[meta$variable_type %in% allowed, ]
  meta$input_flag <- tolower(trimws(meta$input_flag))
  if (isTRUE(use_input_flag)) {
    meta$include <- ifelse(meta$input_flag %in% c("0", "n", "no", "false"), FALSE, TRUE)
    meta <- meta[meta$include, ]
  }
  meta
}

build_all_narratives <- function(parsed) {
  data <- parsed$data
  meta <- normalize_meta(parsed$metadata, use_input_flag = FALSE)
  segment_sizes <- compute_segment_percentages(data)
  segments <- segment_sizes$segment
  meta$normalized_theme <- vapply(meta$theme, normalize_theme_label, character(1))
  top_n <- getOption("narrative.top_differentiators", TOP_DIFFERENTIATOR_COUNT)
  differentiators <- compute_top_differentiators(data, meta, segments, top_n)
  narratives_by_segment <- setNames(lapply(segments, function(x) list()), segments)

  for (i in seq_len(nrow(meta))) {
    meta_row <- meta[i, ]
    if (is_not_clear_theme(meta_row$theme)) {
      next
    }
    theme_label <- meta_row$normalized_theme
    overall_values <- NULL
    type <- meta_row$variable_type

    if (type == "RATING7") {
      overall_values <- calculate_rating_metrics(data[[meta_row$variable_id]], data$WEIGHTS, meta_row$variable_type)
      var_narratives <- build_rating_narratives(
        data,
        meta_row,
        segments,
        list(
          top2 = overall_values$t2b,
          top3 = overall_values$t3b,
          bottom2 = overall_values$b2b,
          bottom3 = overall_values$b3b
        )
      )
    } else if (type == "RATING5") {
      overall_values <- calculate_rating_metrics(data[[meta_row$variable_id]], data$WEIGHTS, meta_row$variable_type)
      var_narratives <- build_rating_narratives(
        data,
        meta_row,
        segments,
        list(
          top2 = overall_values$t2b,
          top3 = NA_real_,
          bottom2 = overall_values$b2b,
          bottom3 = NA_real_
        )
      )
    } else if (type == "SINGLESELECT") {
      overall_mode <- calculate_mode_share(data[[meta_row$variable_id]], data$WEIGHTS)
      var_narratives <- build_singleselect_narratives(data, meta_row, segments, overall_mode$share)
    } else if (type == "MULTISELECT") {
      overall_selection <- calculate_selection_share(data[[meta_row$variable_id]], data$WEIGHTS, 1)
      var_narratives <- build_multiselect_narratives(data, meta_row, segments, overall_selection)
    } else if (type == "NUMERIC") {
      overall_mean <- weighted_mean_safe(data[[meta_row$variable_id]], data$WEIGHTS)
      overall_sd <- weighted_sd_safe(data[[meta_row$variable_id]], data$WEIGHTS)
      var_narratives <- build_numeric_narratives(data, meta_row, segments, overall_mean, overall_sd)
    } else if (type == "RANKING") {
      values <- safe_numeric(data[[meta_row$variable_id]])
      weights <- safe_numeric(data$WEIGHTS)
      valid <- !is.na(values) & !is.na(weights)
      total <- sum(weights[valid])
      overall_rank1 <- if (total > 0) sum(weights[valid & values == 1]) / total * 100 else 0
      var_narratives <- build_ranking_narratives(data, meta_row, segments, overall_rank1)
    } else if (type == "RATING_DUAL") {
      overall_dual <- calculate_dual_metrics(data[[meta_row$variable_id]], data$WEIGHTS)
      var_narratives <- build_dual_narratives(data, meta_row, segments, overall_dual)
    } else {
      next
    }

    for (seg in segments) {
      allowed_vars <- differentiators[[seg]][[theme_label]]
      if (is.null(allowed_vars) || !meta_row$variable_id %in% allowed_vars) {
        next
      }
      seg_bucket <- narratives_by_segment[[seg]][[theme_label]]
      if (is.null(seg_bucket)) {
        seg_bucket <- character(0)
      }
      seg_bucket <- c(seg_bucket, var_narratives[[seg]])
      narratives_by_segment[[seg]][[theme_label]] <- seg_bucket
    }
  }

  # Add explicit messages when no differentiating metrics exist for a theme/segment pair.
  candidate_themes <- unique(meta$normalized_theme[!is_not_clear_theme(meta$theme)])
  for (seg in segments) {
    for (theme_label in candidate_themes) {
      allowed_vars <- differentiators[[seg]][[theme_label]]
      existing <- narratives_by_segment[[seg]][[theme_label]]
      if ((is.null(existing) || !length(existing)) && (is.null(allowed_vars) || !length(allowed_vars))) {
        narratives_by_segment[[seg]][[theme_label]] <- sprintf("No differentiating metrics for theme '%s'.", theme_label)
      }
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

