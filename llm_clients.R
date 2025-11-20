# llm_clients.R
# Author: Iterator Automation
# Date: 2025-02-06
# Summary: Provider-agnostic LLM client helpers for OpenAI, Groq, and Gemini.

`%||%` <- function(x, y) {
  if (is.null(x)) return(y)
  if (is.character(x) && !length(x)) return(y)
  if (is.character(x) && !nzchar(x[1])) return(y)
  x
}

assert_packages <- function(pkgs) {
  missing <- pkgs[!vapply(pkgs, requireNamespace, logical(1), quietly = TRUE)]
  if (length(missing)) {
    stop(sprintf("Missing required packages: %s", paste(missing, collapse = ", ")), call. = FALSE)
  }
}

assert_packages(c("httr", "jsonlite"))

llm_default_models <- function() {
  list(openai = "gpt-4.1-mini", groq = "llama-3.1-70b-versatile", gemini = "gemini-1.5-pro")
}

llm_env_var <- function(provider) {
  switch(tolower(provider),
    openai = "OPENAI_API_KEY",
    groq = "GROQ_API_KEY",
    gemini = "GEMINI_API_KEY",
    stop(sprintf("Unsupported provider: %s", provider), call. = FALSE)
  )
}

set_llm_api_key <- function(provider, key) {
  provider <- tolower(provider %||% "")
  if (!nzchar(provider)) {
    stop("Provider is required to set an API key.", call. = FALSE)
  }

  env_var <- llm_env_var(provider)
  if (!nzchar(key)) {
    stop(sprintf("Key for %s is empty; supply a non-empty value.", provider), call. = FALSE)
  }

  Sys.setenv(structure(key, names = env_var))
  invisible(env_var)
}

read_env_key <- function(provider) {
  key <- Sys.getenv(llm_env_var(provider), unset = "")
  if (!nzchar(key)) {
    stop(sprintf("API key for %s missing. Set %s in your environment.", provider, llm_env_var(provider)), call. = FALSE)
  }
  key
}

parse_openai_like <- function(resp_text) {
  parsed <- jsonlite::fromJSON(resp_text, simplifyVector = FALSE)
  choices <- parsed$choices
  if (!length(choices)) {
    stop("No choices returned from provider.", call. = FALSE)
  }
  content <- choices[[1]]$message$content
  if (is.null(content)) {
    stop("Provider response did not include message content.", call. = FALSE)
  }
  list(text = paste(content, collapse = "\n"), raw = parsed)
}

parse_gemini <- function(resp_text) {
  parsed <- jsonlite::fromJSON(resp_text, simplifyVector = FALSE)
  candidates <- parsed$candidates
  if (!length(candidates)) {
    stop("No candidates returned from Gemini.", call. = FALSE)
  }
  parts <- candidates[[1]]$content$parts
  if (!length(parts)) {
    stop("Gemini response did not include content parts.", call. = FALSE)
  }
  text <- vapply(parts, function(part) part$text %||% "", character(1))
  list(text = paste(text, collapse = "\n"), raw = parsed)
}

call_llm_model <- function(provider, model, prompt, temperature = 0.2, max_tokens = 800) {
  if (!nzchar(prompt)) {
    stop("Prompt is empty; cannot call LLM.", call. = FALSE)
  }

  provider <- tolower(provider %||% "")
  defaults <- llm_default_models()
  model <- model %||% defaults[[provider]] %||% ""

  key <- read_env_key(provider)

  if (identical(provider, "openai")) {
    url <- "https://api.openai.com/v1/chat/completions"
    body <- list(
      model = model,
      messages = list(list(role = "user", content = prompt)),
      temperature = temperature,
      max_tokens = max_tokens
    )
    resp <- httr::POST(
      url,
      httr::add_headers(Authorization = paste("Bearer", key)),
      httr::content_type_json(),
      body = body,
      encode = "json"
    )
    httr::stop_for_status(resp)
    parsed <- parse_openai_like(httr::content(resp, as = "text", encoding = "UTF-8"))
    return(c(parsed, list(provider = provider, model = model)))
  }

  if (identical(provider, "groq")) {
    url <- "https://api.groq.com/openai/v1/chat/completions"
    body <- list(
      model = model,
      messages = list(list(role = "user", content = prompt)),
      temperature = temperature,
      max_tokens = max_tokens
    )
    resp <- httr::POST(
      url,
      httr::add_headers(Authorization = paste("Bearer", key)),
      httr::content_type_json(),
      body = body,
      encode = "json"
    )
    httr::stop_for_status(resp)
    parsed <- parse_openai_like(httr::content(resp, as = "text", encoding = "UTF-8"))
    return(c(parsed, list(provider = provider, model = model)))
  }

  if (identical(provider, "gemini")) {
    base_url <- sprintf("https://generativelanguage.googleapis.com/v1beta/models/%s:generateContent", model)
    body <- list(
      contents = list(list(parts = list(list(text = prompt)))),
      generationConfig = list(
        temperature = temperature,
        maxOutputTokens = max_tokens
      )
    )
    resp <- httr::POST(
      base_url,
      query = list(key = key),
      httr::content_type_json(),
      body = body,
      encode = "json"
    )
    httr::stop_for_status(resp)
    parsed <- parse_gemini(httr::content(resp, as = "text", encoding = "UTF-8"))
    return(c(parsed, list(provider = provider, model = model)))
  }

  stop(sprintf("Unsupported provider: %s", provider), call. = FALSE)
}

