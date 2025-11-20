# input_utils.R
# Author: OpenAI Assistant
# Date: 2024-11-17
# Helpers for standardizing INPUT-like sheets across modules.

if (!requireNamespace("readxl", quietly = TRUE)) {
  stop("Package 'readxl' is required to load standardized input sheets.")
}

clean_variable_ids <- function(ids) {
  ids <- as.character(ids)
  ids <- trimws(ids)
  ids <- gsub("\\s+", "_", ids)
  ids <- gsub("[^A-Za-z0-9_]+", "_", ids)
  ids <- toupper(ids)
  ids[ids == "_"] <- ""
  ids <- gsub("_+", "_", ids)
  ids <- gsub("^_+|_+$", "", ids)
  ids <- make.unique(ids, sep = "_")
  ids
}

clean_variable_types <- function(types) {
  allowed <- c("RATING7", "RATING5", "SINGLESELECT", "MULTISELECT", "NUMERIC", "RANKING")
  normalized <- toupper(trimws(as.character(types)))
  normalized <- gsub("\\s+", "", normalized)
  normalized[normalized %in% c("LIKERT7", "RATING_7", "RATING07")] <- "RATING7"
  normalized[normalized %in% c("LIKERT5", "RATING_5", "RATING05")] <- "RATING5"
  normalized[!normalized %in% allowed] <- NA_character_
  normalized
}

normalize_input_flag <- function(flags) {
  val <- tolower(trimws(as.character(flags)))
  val %in% c("1", "y", "yes", "true", "include")
}

read_standardized_input <- function(path, sheet = "INPUT", na_values = c("NA", ".", "", " "), drop_empty_cols = TRUE) {
  raw <- readxl::read_excel(path, sheet = sheet, na = na_values, col_names = FALSE)
  if (nrow(raw) < 7) {
    stop(sprintf("%s sheet must contain at least 6 metadata rows and data rows.", sheet))
  }

  variable_id <- clean_variable_ids(unlist(raw[1, ]))
  theme <- as.character(unlist(raw[2, ]))
  answer_description <- as.character(unlist(raw[3, ]))
  variable_type <- clean_variable_types(unlist(raw[4, ]))
  question_prefix <- as.character(unlist(raw[5, ]))
  input_flag <- as.character(unlist(raw[6, ]))

  metadata <- data.frame(
    variable_id = variable_id,
    theme = trimws(theme),
    answer_description = trimws(answer_description),
    variable_type = variable_type,
    question_prefix = trimws(question_prefix),
    input_flag = trimws(input_flag),
    stringsAsFactors = FALSE
  )
  metadata$input_flag_clean <- normalize_input_flag(metadata$input_flag)
  metadata$column_index <- seq_len(nrow(metadata))

  if (isTRUE(drop_empty_cols)) {
    keep <- nzchar(metadata$variable_id)
    metadata <- metadata[keep, , drop = FALSE]
    raw <- raw[, keep, drop = FALSE]
  }

  colnames(raw) <- metadata$variable_id
  data_rows <- raw[-seq_len(6), , drop = FALSE]
  data_rows <- as.data.frame(data_rows, stringsAsFactors = FALSE)
  names(data_rows) <- metadata$variable_id

  list(data = data_rows, metadata = metadata)
}
