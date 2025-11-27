# ai_summarizer_prompt.R
# Author: Iterator Automation
# Date: 2025-02-06
# Summary: Prompt construction and response parsing for AI summarizer tab.

assert_packages <- function(pkgs) {
  missing <- pkgs[!vapply(pkgs, requireNamespace, logical(1), quietly = TRUE)]
  if (length(missing)) {
    stop(sprintf("Missing required packages: %s", paste(missing, collapse = ", ")), call. = FALSE)
  }
}

assert_packages(c("jsonlite", "purrr", "tibble"))

summarizer_output_schema <- function() {
  paste(
    "Return a compact JSON object using this schema:",
    "{",
    '  "coherence_score": integer (0-100),',
    '  "segment_definition": string (exactly 15 words describing how segments are defined),',
    '  "segment_issues": [',
    '    { "segment_label": string, "issues": [string, ...] },',
    '    ...',
    '  ],',
    '  "detailed_story": string (rich narrative that maximises differentiation while staying coherent)',
    "}",
    sep = "\n"
  )
}

default_ai_summarizer_sections <- function() {
  list(
    role = "You are an expert segment summarizer.",
    instructions = c(
      "Use the provided context to judge the coherence of the segment stories.",
      "Respond ONLY with JSON, no prose."
    ),
    guidance_header = "Guidance:",
    guidance = c(
      "- Penalize contradictions or unclear linkages between segments when scoring coherence.",
      "- Keep the segment_definition to exactly 15 words.",
      "- Ensure segment_issues enumerates potential risks, gaps, or contradictions for each segment you infer.",
      "- The detailed_story must tell a coherent narrative while maximizing differentiation across segments."
    ),
    context_header = "Context provided by user:",
    narrative_header = "Narrative to summarize:"
  )
}

collect_ai_summarizer_sections <- function() {
  overrides <- getOption("iterator.ai_summarizer_prompt", list())
  defaults <- default_ai_summarizer_sections()
  if (!is.list(overrides) || is.null(names(overrides))) {
    return(defaults)
  }
  utils::modifyList(defaults, overrides)
}

build_ai_summarizer_prompt <- function(narrative_text, context_prompt) {
  narrative_text <- narrative_text %||% ""
  context_prompt <- context_prompt %||% ""

  sections <- collect_ai_summarizer_sections()

  paste(
    sections$role,
    paste(sections$instructions, collapse = "\n"),
    summarizer_output_schema(),
    sections$guidance_header,
    paste(sections$guidance, collapse = "\n"),
    sections$context_header,
    context_prompt,
    sections$narrative_header,
    narrative_text,
    sep = "\n\n"
  )
}

process_ai_summarizer_output <- function(output_text) {
  if (!nzchar(output_text)) {
    return(list(error = "Empty model response", raw_json = output_text))
  }

  parsed <- tryCatch(jsonlite::fromJSON(output_text, simplifyVector = FALSE), error = function(e) NULL)
  if (is.null(parsed)) {
    return(list(error = "Response was not valid JSON", raw_json = output_text))
  }

  score <- suppressWarnings(as.numeric(parsed$coherence_score))
  score <- ifelse(is.na(score), NA_real_, score)

  definition <- parsed$segment_definition %||% ""
  detailed_story <- parsed$detailed_story %||% ""

  issues <- parsed$segment_issues
  issues_df <- data.frame(segment_label = character(0), issues = character(0), stringsAsFactors = FALSE)
  if (length(issues)) {
    issues_df <- purrr::map_dfr(issues, function(item) {
      seg <- item$segment_label %||% ""
      txt <- item$issues
      if (is.null(txt)) {
        txt <- ""
      }
      tibble::tibble(segment_label = seg, issues = paste(unlist(txt), collapse = "; "))
    })
  }

  list(
    coherence_score = score,
    segment_definition = definition,
    detailed_story = detailed_story,
    segment_issues = issues_df,
    raw_json = jsonlite::toJSON(parsed, auto_unbox = TRUE, pretty = TRUE)
  )
}

