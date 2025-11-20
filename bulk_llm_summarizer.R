# bulk_llm_summarizer.R
# Author: OpenAI Assistant
# Date: 2025-02-10
# Summary: Helpers for bulk LLM summarization of combination outputs.

if (!exists("%||%")) {
  `%||%` <- function(x, y) {
    if (is.null(x)) return(y)
    if (is.character(x) && !length(x)) return(y)
    if (is.character(x) && !nzchar(x[1])) return(y)
    x
  }
}

assert_packages <- function(pkgs) {
  missing <- pkgs[!vapply(pkgs, requireNamespace, logical(1), quietly = TRUE)]
  if (length(missing)) {
    stop(sprintf("Missing required packages: %s", paste(missing, collapse = ", ")), call. = FALSE)
  }
}

assert_packages(c("stringr"))

build_bulk_summary_prompt <- function(segment_text, prompt_text = "", context_prompt = "") {
  segment_text <- segment_text %||% ""
  prompt_text <- prompt_text %||% ""
  context_prompt <- context_prompt %||% ""

  required_layout <- paste(
    "Return the response in EXACTLY this layout:",
    "segment <number>: <Short segment name based on their information>",
    "<A key pointer on segment behavior, a 10–20 word pointer for each theme, summarizing their key behaviour/attitudes specific to that theme>",
    "(Repeat the above block for all segments in that solution)",
    "POSSIBLE ISSUES WITH SEGMENTATION:",
    "<List down all possible contradictions, weaknesses, inconsistencies or issues with the segmentation, in bullet points or short sentences, so that the human can quickly understand what might be problematic in this solution.>",
    "CLARITY_SCORE_0_100: <single integer from 0–100 reflecting clarity and coherence across the full solution>",
    "SOLUTION_ISSUES_BRIEF: <1–3 sentences summarizing the main concerns/issues with this solution>",
    sep = "\n"
  )

  paste(
    "You are validating a segmentation solution. Follow any user-provided prompt verbatim before applying the required layout.",
    "Incorporate the context prompt to refine your assessment. Do not add any extra sections or JSON.",
    "PROMPT FROM COMBINATION OUTPUT:",
    prompt_text,
    "ADDITIONAL CONTEXT PROMPT (user provided):",
    context_prompt,
    "SEGMENT SUMMARY TEXT (an entire solution):",
    segment_text,
    required_layout,
    "Rules:",
    "- Keep the structure identical to the required layout; do not add new headers.",
    "- Every segment must be listed as 'segment <number>:' followed by its short name and 10–20 word pointers.",
    "- CLARITY_SCORE_0_100 must be an integer between 0 and 100.",
    "- SOLUTION_ISSUES_BRIEF must be concise and directly reflect the segmentation quality issues.",
    sep = "\n\n"
  )
}

validate_bulk_output <- function(output_text) {
  if (!nzchar(output_text)) {
    stop("Model response is empty.", call. = FALSE)
  }

  if (!grepl("(?i)segment\\s+\\d+", output_text, perl = TRUE)) {
    stop("Model response is missing the required segment lines.", call. = FALSE)
  }

  if (!grepl("(?i)POSSIBLE ISSUES WITH SEGMENTATION", output_text, perl = TRUE)) {
    stop("Model response is missing the 'POSSIBLE ISSUES WITH SEGMENTATION' section.", call. = FALSE)
  }

  clarity_match <- stringr::str_match(output_text, "(?i)CLARITY_SCORE_0_100\\s*[:=]\\s*(\\d{1,3})")
  clarity <- suppressWarnings(as.integer(clarity_match[, 2]))
  if (is.na(clarity) || clarity < 0 || clarity > 100) {
    stop("Model response is missing a valid CLARITY_SCORE_0_100 between 0 and 100.", call. = FALSE)
  }

  issues_match <- stringr::str_match(output_text, "(?is)SOLUTION_ISSUES_BRIEF\\s*[:=]\\s*(.+)")
  issues_brief <- issues_match[, 2] %||% ""
  issues_brief <- trimws(issues_brief)
  if (!nzchar(issues_brief)) {
    stop("Model response is missing SOLUTION_ISSUES_BRIEF.", call. = FALSE)
  }

  tokens_estimate <- ceiling(nchar(output_text) / 4)

  list(
    clarity_score = clarity,
    issues_brief = issues_brief,
    tokens_used = tokens_estimate
  )
}

run_llm_summary <- function(segment_text,
                            prompt_text = "",
                            context_prompt = "",
                            provider = "openai",
                            model = NULL,
                            temperature = 0.2,
                            max_tokens = 800) {
  model <- model %||% llm_default_models()[[tolower(provider %||% "")]]
  prompt <- build_bulk_summary_prompt(segment_text, prompt_text, context_prompt)

  response <- call_llm_model(
    provider = provider,
    model = model,
    prompt = prompt,
    temperature = temperature,
    max_tokens = max_tokens
  )

  validation <- validate_bulk_output(response$text)

  list(
    full_text = response$text,
    clarity_score = validation$clarity_score,
    issues_brief = validation$issues_brief,
    tokens_used = validation$tokens_used,
    prompt_used = prompt,
    provider = response$provider,
    model = response$model
  )
}

read_text_safely <- function(path) {
  if (!nzchar(path) || !file.exists(path)) {
    return("")
  }
  paste(readLines(path, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
}

