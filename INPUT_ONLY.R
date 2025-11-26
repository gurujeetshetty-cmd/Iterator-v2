# INPUT_ONLY.R
# Author: OpenAI Assistant
# Date: 2024-11-21
# Central configuration hub for Iterator. Update variables here and run this file to launch the app.

# -------------------------------
# 🗂️ Core paths
# -------------------------------
iterator_home <- "C:/Users/yourname/Iterator"  # 🔧 Edit this to your preferred root folder

# Define tweakable configuration values in a single list.
iterator_config <- list(
  iterator_home = iterator_home,                              # Root directory containing all Iterator scripts
  input_directory = file.path(iterator_home, "input"),       # Example: "C:/Users/yourname/Iterator/input"
  output_directory = file.path(iterator_home, "output"),     # Example: "C:/Users/yourname/Iterator/output"
  tracker_directory = file.path(iterator_home, "output"),    # Example: "C:/Users/yourname/Iterator/output"
  summary_directory = file.path(iterator_home, "summaries"), # Example: "C:/Users/yourname/Iterator/summaries"

  # Default file names
  input_excel = file.path(iterator_home, "INPUT.xlsx"),
  tracker_csv = file.path(iterator_home, "Iteration_Tracker.csv"),
  output_filename = "combinations_output.csv",

  # Iteration generator defaults
  min_n_size = 8L,
  max_n_size = 10L,
  max_iterations = 100L,
  mandatory_vars = "5,7",                     # Column numbers or names, comma-separated
  include_fa = FALSE,                          # Include factor analysis metrics
  include_corr = FALSE,                        # Include Pearson correlation metrics

  # Thresholds & tuning knobs
  bimodality_hi_threshold = 110,               # High score cutoff for bimodality detection
  bimodality_low_threshold = 0.2,              # Low score cutoff for bimodality detection
  performance_hi_threshold = 120,              # High score cutoff for performance checks
  performance_lo_threshold = 80,               # Low score cutoff for performance checks
  performance_base_threshold = 0.2,            # Base rate floor to count perf/indT/indB
  independence_base_threshold = 0.2,           # Base rate floor to count indT/indB

  # LLM + AI knobs (used by summarizers)
  ai_provider = "openai",                     # openai | groq | gemini
  ai_model = "gpt-4o-mini",                   # Model/deployment name
  azure_deployment = NA_character_,            # Azure deployment name if using Azure OpenAI
  azure_endpoint = NA_character_,              # Azure endpoint
  openai_api_key = NA_character_,              # Set to your OPENAI_API_KEY to avoid exporting env vars manually
  groq_api_key = NA_character_,                # Set to your GROQ_API_KEY if using Groq
  gemini_api_key = NA_character_,              # Set to your GEMINI_API_KEY if using Gemini
  azure_api_key = NA_character_,               # Set to your AZURE_OPENAI_API_KEY if using Azure OpenAI

  # Shiny runtime options
  shiny_host = "0.0.0.0",
  shiny_port = 8080L
)

# Ensure configurable directories exist before launching the dashboard.
make_sure_dir <- function(path) {
  if (!dir.exists(path)) {
    dir.create(path, recursive = TRUE, showWarnings = FALSE)
  }
  invisible(normalizePath(path, winslash = "/", mustWork = FALSE))
}

iterator_config$input_directory <- make_sure_dir(iterator_config$input_directory)
iterator_config$output_directory <- make_sure_dir(iterator_config$output_directory)
iterator_config$tracker_directory <- make_sure_dir(iterator_config$tracker_directory)
iterator_config$summary_directory <- make_sure_dir(iterator_config$summary_directory)

# Expose configuration for the dashboard to consume.
options(iterator.config = iterator_config)
Sys.setenv(ITERATOR_HOME = iterator_config$iterator_home)
set_env_if_present <- function(var_name, value) {
  if (!is.null(value) && !is.na(value) && nzchar(value)) {
    Sys.setenv(structure(value, names = var_name))
  }
}

if (!is.na(iterator_config$azure_deployment)) {
  Sys.setenv(AZURE_OPENAI_DEPLOYMENT = iterator_config$azure_deployment)
}
if (!is.na(iterator_config$azure_endpoint)) {
  Sys.setenv(AZURE_OPENAI_ENDPOINT = iterator_config$azure_endpoint)
}
set_env_if_present("OPENAI_API_KEY", iterator_config$openai_api_key)
set_env_if_present("GROQ_API_KEY", iterator_config$groq_api_key)
set_env_if_present("GEMINI_API_KEY", iterator_config$gemini_api_key)
set_env_if_present("AZURE_OPENAI_API_KEY", iterator_config$azure_api_key)

# Optionally set Shiny runtime values globally.
if (!is.null(iterator_config$shiny_host)) options(shiny.host = iterator_config$shiny_host)
if (!is.null(iterator_config$shiny_port)) options(shiny.port = iterator_config$shiny_port)

# Launch the dashboard directly from this file.
dashboard_path <- file.path(iterator_config$iterator_home, "Iterator_Dashboard.R")
if (!file.exists(dashboard_path)) {
  stop(sprintf("Cannot locate Iterator_Dashboard.R at %s", dashboard_path))
}

message("⚙️  Loading Iterator dashboard from: ", dashboard_path)
source(dashboard_path)
