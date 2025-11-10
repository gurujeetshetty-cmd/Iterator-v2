# =========================================================
# 🚀 ITERATOR DASHBOARD (V7.6 — Modal Progress + Clean UI)
# =========================================================

# ---- Load Packages ----
packages <- c(
  "shiny", "shinydashboard", "DT", "shinyWidgets", "shinyFiles", "readxl", "shinyjs",
  "openxlsx", "dplyr", "tidyr", "lubridate", "plotly", "RcppAlgos",
  "combinat", "parallel", "gRbase", "ggplot2", "weights", "poLCA",
  "stringi", "utf8", "purrr", "Hmisc", "checkmate", "jomo", "pan", "gtools",
  "future", "promises"
)
for (pkg in packages) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    message(sprintf("⚠️ Package %s not installed.", pkg))
  }
}

# ---- Explicit imports for lintr ----
# (Silences “no visible global function” warnings)
utils::globalVariables(c(
  "observeEvent", "req", "showModal", "removeModal", "modalDialog",
  "div", "br", "showNotification", "renderText",
  "updateProgressBar", "generate_best_combinations", "run_iterations"
))

options(shiny.autoreload = TRUE)

# ---- Directory & Source ----
DIRECTORY_PATH <- "C:/Users/gs36325/Documents/ENHERTU"

get_script_path <- function() {
  cmd_args <- commandArgs(trailingOnly = FALSE)
  file_arg <- "--file="
  matches <- cmd_args[grepl(file_arg, cmd_args, fixed = TRUE)]
  if (!length(matches)) {
    return(NULL)
  }
  sub(file_arg, "", matches[[1]], fixed = TRUE)
}

normalize_if_exists <- function(path) {
  tryCatch(normalizePath(path, winslash = "/", mustWork = FALSE), error = function(e) path)
}

`%||%` <- function(x, y) {
  if (is.null(x)) return(y)
  if (is.character(x) && !length(x)) return(y)
  if (is.character(x) && !nzchar(x[1])) return(y)
  x
}

required_scripts <- c(
  "Iteration_Prioritizer_vF3.R",
  "metrics_fa.R",
  "metrics_pearson.R",
  "Function_poLCA.R",
  "rules_check.R",
  "Iteration_loop.R",
  "Iteration_Runner.R"
)

script_dir <- get_script_path()
if (!is.null(script_dir)) {
  script_dir <- normalize_if_exists(script_dir)
}
env_dir <- Sys.getenv("ITERATOR_HOME", unset = "")

candidate_sources <- normalize_if_exists(c(
  getwd(),
  file.path(getwd(), "Iterator-main vF.4"),
  DIRECTORY_PATH,
  file.path(DIRECTORY_PATH, "X"),
  env_dir,
  file.path(env_dir, "X"),
  dirname(script_dir %||% ""),
  file.path(dirname(script_dir %||% ""), "X")
))

candidate_sources <- candidate_sources[!is.na(candidate_sources) & nzchar(candidate_sources)]
candidate_dirs <- unique(candidate_sources[dir.exists(candidate_sources)])

if (!length(candidate_dirs)) {
  stop("❌ Unable to locate Iterator directory. Please set ITERATOR_HOME or update DIRECTORY_PATH.")
}

find_repo_root <- function(paths, files) {
  for (p in paths) {
    if (all(file.exists(file.path(p, files)))) {
      return(p)
    }
  }
  NULL
}

repo_root <- find_repo_root(candidate_dirs, required_scripts)

if (is.null(repo_root)) {
  stop("❌ Required Iterator scripts not found in candidate directories.")
}

if (!identical(normalize_if_exists(getwd()), repo_root)) {
  setwd(repo_root)
}

message(sprintf("📁 Using Iterator directory: %s", repo_root))

if (!dir.exists(DIRECTORY_PATH)) {
  DIRECTORY_PATH <- repo_root
  message(sprintf("ℹ️ DIRECTORY_PATH not found; defaulting to %s", DIRECTORY_PATH))
}

safe_source <- function(file) {
  full_path <- file.path(repo_root, file)
  if (!file.exists(full_path)) {
    stop(sprintf("❌ Missing required script: %s", full_path))
  }
  source(full_path)
}

invisible(lapply(required_scripts, safe_source))

if (!exists("run_iterations")) stop("❌ run_iterations() not found — check Iteration_loop.R")
message("✅ run_iterations() successfully loaded.")

# =========================================================
# ✅ Logger (console only)
# =========================================================
append_log <- function(message, level = "INFO") {
  ts <- format(Sys.time(), "%H:%M:%S")
  cat(sprintf("[%s] %s %s\n", ts, level, message))
}

# =========================================================
# 🧠 USER INTERFACE
# =========================================================
ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(title = "Iterator Dashboard"),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Input Generator", tabName = "input_generator", icon = icon("cogs"), selected = TRUE),
      menuItem("Output Generator", tabName = "output_generator", icon = icon("chart-line")),
      menuItem("Iteration Runner", tabName = "iteration_runner", icon = icon("play"))
    )
  ),
  dashboardBody(
    useShinyjs(),
    tags$head(
      tags$style(HTML('
        body, .content-wrapper { background-color: #f5f7fb; }
        .box {
          border-top: 2px solid #2563eb;
          border-radius: 10px;
          box-shadow: 0 10px 24px rgba(15, 23, 42, 0.06);
          background-color: #ffffff;
        }
        .box .box-title { font-weight: 600; letter-spacing: 0.4px; }
        .box-intro {
          background: rgba(37, 99, 235, 0.06);
          border-radius: 12px;
          padding: 18px 20px;
          margin-bottom: 18px;
        }
        .box-intro h4 { font-weight: 700; margin: 0 0 6px 0; }
        .box-intro p { margin: 0; color: #475569; }
        .btn-primary {
          background: #2563eb;
          border-color: #2563eb;
          border-radius: 8px;
          box-shadow: none;
        }
        .btn-primary:hover, .btn-primary:focus {
          background: #1d4ed8;
          border-color: #1d4ed8;
        }
        .modal-content {
          border-radius: 16px;
          border: none;
          box-shadow: 0 26px 60px rgba(15, 23, 42, 0.35);
        }
        .progress { height: 40px !important; border-radius: 10px; }
        .progress-bar {
          font-size: 16px;
          line-height: 40px;
          background: linear-gradient(135deg, #38bdf8 0%, #2563eb 100%);
        }
        .input-strip {
          display: flex;
          flex-wrap: wrap;
          gap: 16px;
          align-items: stretch;
        }
        .input-block {
          flex: 1 1 220px;
          min-width: 200px;
          padding: 12px 14px 10px 14px;
          border: 1px solid rgba(148, 163, 184, 0.4);
          border-radius: 10px;
          background: #ffffff;
        }
        .input-block .form-group { margin-bottom: 0; }
        .input-block label { font-weight: 600; color: #1e293b; }
        .upload-block {
          display: flex;
          flex-direction: column;
          justify-content: center;
          border-style: dashed;
          border-color: rgba(37, 99, 235, 0.45);
          background: rgba(37, 99, 235, 0.04);
        }
        .upload-block .form-group { margin-bottom: 6px; }
        .metrics-toggle-panel {
          margin-top: 18px;
          border-radius: 10px;
          padding: 18px 20px 6px 20px;
          background: rgba(14, 165, 233, 0.08);
          border: 1px solid rgba(14, 165, 233, 0.25);
        }
        .metrics-toggle-panel h5 {
          font-weight: 600;
          margin: 0 0 6px 0;
        }
        .toggle-row {
          display: flex;
          flex-wrap: wrap;
          gap: 22px;
        }
        .toggle-row .form-group { margin-bottom: 12px; }
        .custom-iterations-panel {
          margin-top: 20px;
          border-radius: 10px;
          padding: 18px 20px 8px 20px;
          background: rgba(79, 70, 229, 0.08);
          border: 1px solid rgba(79, 70, 229, 0.18);
        }
        .custom-iterations-panel h4 {
          margin-top: 0;
          font-weight: 600;
        }
        .action-area {
          margin-top: 24px;
          display: flex;
          flex-wrap: wrap;
          align-items: center;
          gap: 12px;
        }
        .summary-card-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
          gap: 16px;
          margin-top: 18px;
        }
        .summary-card {
          padding: 18px 20px;
          border-radius: 12px;
          border: 1px solid rgba(148, 163, 184, 0.35);
          background: #f8fafc;
        }
        .summary-card small { display: block; font-size: 12px; color: #475569; margin-top: 4px; }
        .output-status-callout {
          margin-top: 18px;
          padding: 16px 20px;
          border-radius: 10px;
          border: 1px solid rgba(37, 99, 235, 0.25);
          background: rgba(37, 99, 235, 0.08);
        }
        .output-status-callout strong { display: block; margin-bottom: 4px; letter-spacing: 0.6px; font-size: 12px; color: #1e293b; }
        .output-status-callout span { font-size: 14px; color: #334155; }
      '))
    ),
    tabItems(

      # ------------------------------------------------------
      # TAB 1: INPUT GENERATOR
      # ------------------------------------------------------
      tabItem(tabName = "input_generator",
        fluidRow(
          column(width = 12,
            box(title = "Input Generator", status = "primary", solidHeader = TRUE, width = 12,
              div(class = "box-intro",
                  h4("Design a balanced combination plan"),
                  p(class = "text-muted", style = "margin-bottom:0;",
                    "Upload your poLCA workbook and fine-tune iteration targets to instantly preview feasibility ",
                    "across N sizes and theme constraints.")),
              div(class = "input-strip",
                  div(class = "input-block upload-block",
                      fileInput("input_file", "Upload Excel File", accept = ".xlsx")),
                  div(class = "input-block",
                      numericInput("min_n_size", "Minimum N Size", value = 8)),
                  div(class = "input-block",
                      numericInput("max_n_size", "Maximum N Size", value = 10)),
                  div(class = "input-block",
                      numericInput("max_iterations", "Iterations per N Size", value = 100)),
                  div(class = "input-block",
                      textInput("mandatory_vars", "Mandatory Variable Columns (comma-separated)", value = "5,7")),
                  div(class = "input-block",
                      textInput("output_filename", "Output Filename", value = "combinations_output.csv"))
              ),
              div(class = "metrics-toggle-panel",
                  h5("Optional metric enrichments"),
                  p(class = "text-muted", "Select the metric suites to append during combination generation."),
                  div(class = "toggle-row",
                      checkboxInput("include_fa", "Include Factor Analysis Metrics", value = FALSE),
                      checkboxInput("include_corr", "Include Pearson Correlation Metrics", value = FALSE)
                  )
              ),
              div(class = "custom-iterations-panel",
                  tags$h4(class = "mt-0", icon("sliders"), span(" Customise iteration volume", style = "margin-left:6px;")),
                  p(class = "text-muted", "Toggle bespoke iteration counts for each N size whenever you need more control over sampling intensity."),
                  checkboxInput("use_custom_iterations", "Specify iterations per N size", value = FALSE),
                  uiOutput("custom_iterations_inputs")),
              div(class = "action-area",
                  actionButton("run_input_generator", "Generate Combinations",
                               class = "btn btn-primary btn-lg"),
                  tags$small(class = "text-muted",
                             icon("info-circle"),
                             span(" Previewed capacity updates live below as soon as your file is processed.",
                                  style = "margin-left:6px;"))),
              uiOutput("input_summary_cards")
            )
          )
        ),
        fluidRow(
          column(width = 12,
            box(title = "Combination Capacity Overview", status = "primary", solidHeader = TRUE, width = 12,
              uiOutput("combination_plan_summary"),
              DT::dataTableOutput("combination_plan_table"),
              br(),
              plotlyOutput("combination_plan_plot", height = "320px"),
              br(),
              DT::dataTableOutput("theme_summary_table")
            )
          )
        )
      ),

      # ------------------------------------------------------
      # TAB 2: OUTPUT GENERATOR
      # ------------------------------------------------------
      tabItem(tabName = "output_generator",
        fluidRow(
          box(title = "Output Generator", status = "primary", solidHeader = TRUE, width = 12,
            div(class = "box-intro",
                h4("Transform prioritised combinations into polished ITRs"),
                p(class = "text-muted", style = "margin-bottom:0;",
                  "Pair your poLCA input with a combinations plan, choose your destination, and follow the live status feed while results export.")),
            div(class = "input-strip",
                div(class = "input-block upload-block",
                    shinyFiles::shinyFilesButton(
                      "xtabs_input_file", "Select poLCA input file", "Browse...",
                      multiple = FALSE
                    ),
                    textOutput(
                      "xtabs_input_file_display",
                      container = function(...) tags$small(class = "text-muted d-block mt-2", ...)
                    )
                ),
                div(class = "input-block upload-block",
                    shinyFiles::shinyFilesButton(
                      "combinations_file", "Select combinations sheet", "Browse...",
                      multiple = FALSE
                    ),
                    textOutput(
                      "combinations_file_display",
                      container = function(...) tags$small(class = "text-muted d-block mt-2", ...)
                    )
                ),
                div(class = "input-block",
                    textInput("file_name", "Enter starting text for ITRs", placeholder = "TEVA_ITR_")),
                div(class = "input-block",
                    textInput("output_directiory", "Output Directory",
                              value = "C:/Users/gs36325/Documents/22 Saphnelo Segmentation/09 Iterator Final/AAKASH V2")),
                div(class = "input-block",
                    numericInput("num_segments", "# of Segments", value = 4, min = 1, max = 10))
            ),
            div(class = "action-area",
                actionButton("run_output_generator", "Run Output Generator", class = "btn btn-primary btn-lg"),
                tags$small(class = "text-muted",
                           icon("star"),
                           span(" Progress updates appear in the modal and the status panel below.", style = "margin-left:6px;"))),
            uiOutput("output_generator_cards"),
            div(class = "output-status-callout",
                strong("Live status"),
                textOutput("status_run", container = function(...) tags$span(style = "display:block;font-weight:600;font-size:15px;", ...)))
          )
        )
      ),

      # ------------------------------------------------------
      # TAB 3: ITERATION RUNNER
      # ------------------------------------------------------
      tabItem(tabName = "iteration_runner",
        fluidRow(
          box(
            title = "Manual Iteration Setup",
            status = "primary",
            solidHeader = TRUE,
            width = 4,
            textInput(
              "input_file_path",
              "poLCA Input File Path",
              value = "C:/Users/GS36325/Documents/14 TRUQAP mBC HCP Segmentation/09 Iterator_CLEAN/poLCA_input.xlsx"
            ),
            actionButton("load_files", "Load Input", class = "btn btn-primary"),
            tags$small(class = "text-muted d-block mt-2", "Provide the full path to poLCA_input.xlsx and click Load."),
            br(),
            textOutput("files_loaded"),
            hr(),
            shinyWidgets::pickerInput(
              "manual_variables",
              "Select Variables",
              choices = NULL,
              selected = NULL,
              multiple = TRUE,
              options = list(
                `live-search` = TRUE,
                `actions-box` = TRUE,
                size = 10,
                `selected-text-format` = "count > 2"
              )
            ),
            DT::DTOutput("manual_selected_table"),
            br(),
            numericInput("manual_num_segments", "# of Segments", value = 4, min = 2, max = 10),
            actionButton("run_manual_iteration", "Run Iteration", class = "btn btn-primary btn-block"),
            br(),
            htmlOutput("manual_save_info")
          ),
          box(
            title = "Metrics & Results",
            status = "primary",
            solidHeader = TRUE,
            width = 8,
            shinyWidgets::pickerInput(
              "manual_metrics",
              "Metrics Selector",
              choices = c(
                "Segment Incidence",
                "OF Spread",
                "Variable Differentiation",
                "Variable Summaries"
              ),
              selected = c(
                "Segment Incidence",
                "OF Spread",
                "Variable Differentiation",
                "Variable Summaries"
              ),
              multiple = TRUE,
              options = list(
                `actions-box` = TRUE,
                `selected-text-format` = "count > 2"
              )
            ),
            uiOutput("manual_results_ui")
          )
        )
      )
    )
  )
)

# =========================================================
# ⚙️ SERVER LOGIC
# =========================================================
server <- function(input, output, session) {
  base_volumes <- shinyFiles::getVolumes()()
  home_volume <- normalizePath("~", winslash = "/", mustWork = FALSE)
  volumes <- if (length(base_volumes) && home_volume %in% base_volumes) {
    base_volumes
  } else {
    c(Home = home_volume, base_volumes)
  }

  rv_paths <- reactiveValues(
    xtabs_path = NULL,
    xtabs_dir = NULL,
    xtabs_name = NULL,
    combos_path = NULL,
    combos_dir = NULL,
    combos_name = NULL
  )

  manual_state <- reactiveValues(
    metadata = NULL,
    input_path = NULL,
    input_dir = NULL,
    segment_incidence = NULL,
    of_metrics = tibble::tibble(),
    variable_diff = NULL,
    variable_tables = list(),
    rendered_var_ids = character(0),
    output_path = NULL,
    run_name = NULL,
    last_run_time = NULL,
    running = FALSE
  )

  status_text <- reactiveVal("Output generator idle.")

  output$status_run <- renderText({ status_text() })

summary_card <- function(title, value, subtitle = NULL, icon_tag = NULL, status_class = NULL) {
  classes <- c("summary-card", if (!is.null(status_class)) paste0("summary-card-", status_class))
  tags$div(
    class = paste(classes, collapse = " "),
    tags$div(
      class = "summary-card-content",
      if (!is.null(icon_tag)) tags$div(class = "summary-card-icon", icon_tag),
      tags$div(class = "summary-card-title", title),
      tags$div(class = "summary-card-value", value),
      if (!is.null(subtitle)) tags$div(class = "summary-card-subtitle", subtitle)
    )
  )
}

manual_metric_options <- c(
  "Segment Incidence",
  "OF Spread",
  "Variable Differentiation",
  "Variable Summaries"
)

build_manual_metadata <- function(path) {
  na_values <- c("NA", ".", "", " ")
  raw <- readxl::read_excel(path, sheet = "INPUT", na = na_values, col_names = FALSE)

  if (!nrow(raw) || ncol(raw) <= 3) {
    stop("INPUT sheet does not contain variable metadata in expected format.")
  }

  var_cols <- seq(4, ncol(raw))
  variables <- as.character(unlist(raw[1, var_cols]))
  themes <- if (nrow(raw) >= 2) as.character(unlist(raw[2, var_cols])) else rep(NA_character_, length(var_cols))
  descriptions <- if (nrow(raw) >= 3) as.character(unlist(raw[3, var_cols])) else rep(NA_character_, length(var_cols))

  normalize_text <- function(x) {
    vals <- ifelse(is.na(x), NA_character_, trimws(as.character(x)))
    vals
  }

  tibble::tibble(
    variable = normalize_text(variables),
    theme = normalize_text(themes),
    description = normalize_text(descriptions),
    column_index = var_cols
  ) %>%
    dplyr::filter(!is.na(variable) & nzchar(variable))
}

parse_of_display_values <- function(display, seg_n) {
  if (is.null(display) || !length(display) || is.na(display)) {
    return(rep(NA_real_, seg_n))
  }

  tokens <- unlist(strsplit(display, "\\|", fixed = FALSE))
  tokens <- trimws(tokens)
  nums <- suppressWarnings(as.numeric(tokens))
  if (!length(nums)) {
    return(rep(NA_real_, seg_n))
  }

  if (length(nums) < seg_n) {
    nums <- c(nums, rep(NA_real_, seg_n - length(nums)))
  } else if (length(nums) > seg_n) {
    nums <- nums[seq_len(seg_n)]
  }
  nums
}

extract_of_metric_table <- function(rules_result, seg_n) {
  value_cols <- grep("^OF_.*_values$", names(rules_result), value = TRUE)
  if (!length(value_cols)) {
    return(tibble::tibble())
  }

  metric_tables <- lapply(value_cols, function(col) {
    base <- sub("_values$", "", col)
    values <- parse_of_display_values(rules_result[[col]], seg_n)
    tibble::tibble(
      Metric = base,
      Segment = paste0("Segment ", seq_len(seg_n)),
      Value = values,
      DiffMax = rep(suppressWarnings(as.numeric(rules_result[[paste0(base, "_diff_max")]])), seg_n),
      DiffMin = rep(suppressWarnings(as.numeric(rules_result[[paste0(base, "_diff_min")]])), seg_n)
    )
  })

  dplyr::bind_rows(metric_tables)
}

sanitize_manual_id <- function(x) {
  id <- gsub("[^A-Za-z0-9]", "_", x)
  id <- gsub("_+", "_", id)
  id <- sub("^_+", "", id)
  id <- sub("_+$", "", id)
  paste0("manual_var_table_", tolower(id))
}

extract_manual_variable_tables <- function(summary_path, selected_vars) {
  if (!length(selected_vars) || !file.exists(summary_path)) {
    return(list())
  }

  safe_read <- function(sheet_idx) {
    tryCatch(
      readxl::read_excel(summary_path, sheet = sheet_idx, col_names = FALSE),
      error = function(e) NULL
    )
  }

  xtabs <- safe_read(2)
  if (is.null(xtabs) || !nrow(xtabs)) {
    return(list())
  }

  find_blank_rows <- function(df) {
    blank_rows <- apply(df, 1, function(row) all(is.na(row) | trimws(as.character(row)) == ""))
    which(blank_rows)
  }

  split_tables <- function(df) {
    blanks <- find_blank_rows(df)
    idx <- c(1, blanks + 1, nrow(df) + 1)
    res <- lapply(seq_len(length(idx) - 1), function(i) {
      df[idx[i]:(idx[i + 1] - 1), , drop = FALSE]
    })
    res
  }

  raw_tables <- split_tables(xtabs)
  norm_key <- function(x) tolower(gsub("[^a-z0-9]", "", x))

  tables_by_var <- list()
  for (tbl in raw_tables) {
    header_val <- tbl[1, 1]
    header_val <- ifelse(is.na(header_val), "", as.character(header_val))
    if (!nzchar(header_val)) next
    var_name <- sub("^00_SEGM_", "", header_val)
    var_name <- trimws(var_name)
    if (!nzchar(var_name)) next
    norm_var <- norm_key(var_name)
    matches <- norm_key(selected_vars)
    if (!norm_var %in% matches) next

    cleaned <- tbl
    cleaned[] <- lapply(cleaned, function(col) {
      if (is.list(col)) {
        col <- sapply(col, function(x) paste(unlist(x), collapse = " "))
      }
      trimws(as.character(col))
    })

    cleaned <- cleaned[rowSums(cleaned == "" | is.na(cleaned)) != ncol(cleaned), , drop = FALSE]
    if (!nrow(cleaned)) next
    colnames(cleaned) <- as.character(unlist(cleaned[1, ]))
    cleaned <- cleaned[-1, , drop = FALSE]
    cleaned <- cleaned[, colSums(cleaned == "" | is.na(cleaned)) < nrow(cleaned), drop = FALSE]

    tables_by_var[[var_name]] <- tibble::as_tibble(cleaned)
  }

  if (!length(tables_by_var)) {
    return(list())
  }

  ordered <- list()
  seen_ids <- character(0)
  for (var in selected_vars) {
    key <- norm_key(var)
    matched_name <- names(tables_by_var)[norm_key(names(tables_by_var)) == key]
    if (!length(matched_name)) next
    chosen <- matched_name[1]
    data_tbl <- tables_by_var[[chosen]]
    output_id <- sanitize_manual_id(paste0(var, "_", length(ordered) + 1))
    while (output_id %in% seen_ids) {
      output_id <- paste0(output_id, sample.int(9, 1))
    }
    seen_ids <- c(seen_ids, output_id)
    ordered[[length(ordered) + 1]] <- list(
      variable = var,
      data = data_tbl,
      output_id = output_id
    )
  }
  ordered
}

  shinyFiles::shinyFileChoose(
    input, "xtabs_input_file",
    roots = volumes,
    session = session,
    filetypes = c("xlsx", "xls", "xlsm")
  )
  shinyFiles::shinyFileChoose(
    input, "combinations_file",
    roots = volumes,
    session = session,
    filetypes = c("csv")
  )


  # --------------------------------------------------------
  # INPUT GENERATOR
  # --------------------------------------------------------

  parse_mandatory_indices <- function(raw) {
    if (is.null(raw) || !nzchar(raw)) {
      return(integer(0))
    }
    tokens <- unlist(strsplit(raw, "[,;\\s]+"))
    vals <- suppressWarnings(as.integer(tokens))
    vals[!is.na(vals)]
  }

  format_eta <- function(remaining_seconds, finish_time) {
    if (!is.finite(remaining_seconds) || remaining_seconds < 0) {
      return("Estimated completion: calculating...")
    }

    format_component <- function(value, unit) {
      if (value <= 0) return(NULL)
      sprintf("%d %s", value, if (value == 1) substr(unit, 1, nchar(unit) - 1) else unit)
    }

    remaining_seconds <- round(remaining_seconds)
    mins <- remaining_seconds %/% 60
    secs <- remaining_seconds %% 60
    hours <- mins %/% 60
    mins <- mins %% 60

    parts <- c(
      format_component(hours, "hours"),
      format_component(mins, "minutes"),
      if (hours == 0 || secs > 0) format_component(secs, "seconds") else NULL
    )
    duration_text <- paste(parts, collapse = " ")
    sprintf("Estimated completion: %s (≈ %s)",
            if (nzchar(duration_text)) duration_text else "< 1 second",
            format(finish_time, "%H:%M:%S"))
  }

  plan_info <- reactive({
    req(input$input_file)
    data_path <- tryCatch(normalizePath(input$input_file$datapath, mustWork = TRUE),
                          error = function(e) NULL)
    if (is.null(data_path)) {
      return(NULL)
    }

    min_n_size <- suppressWarnings(as.integer(input$min_n_size))
    max_n_size <- suppressWarnings(as.integer(input$max_n_size))
    if (is.na(min_n_size) || is.na(max_n_size)) {
      return(NULL)
    }

    default_iterations <- suppressWarnings(as.integer(input$max_iterations))
    if (is.na(default_iterations) || default_iterations < 0) {
      default_iterations <- 0L
    }

    mandatory_vars <- parse_mandatory_indices(input$mandatory_vars)

    custom_iterations <- NULL
    if (isTRUE(input$use_custom_iterations) && max_n_size >= min_n_size) {
      n_values <- seq.int(min_n_size, max_n_size)
      custom_vals <- vapply(n_values, function(n) {
        raw_val <- input[[paste0("iter_n_", n)]]
        if (is.null(raw_val) || is.na(raw_val)) {
          return(default_iterations)
        }
        val <- suppressWarnings(as.integer(raw_val))
        if (is.na(val) || val < 0) {
          return(0L)
        }
        val
      }, integer(1))
      names(custom_vals) <- n_values
      custom_iterations <- custom_vals
    }

    summarize_combination_capacity(
      data = data_path,
      min_n_size = min_n_size,
      max_n_size = max_n_size,
      mandatory_variables_cols = mandatory_vars,
      default_iterations = default_iterations,
      custom_iterations = custom_iterations,
      log_fn = append_log
    )
  })

  output$custom_iterations_inputs <- renderUI({
    if (!isTRUE(input$use_custom_iterations)) {
      return(NULL)
    }

    min_n <- suppressWarnings(as.integer(input$min_n_size))
    max_n <- suppressWarnings(as.integer(input$max_n_size))

    if (is.na(min_n) || is.na(max_n) || max_n < min_n) {
      return(div(class = "text-danger mt-2", "Enter a valid minimum and maximum N size to customise iterations."))
    }

    n_values <- seq.int(min_n, max_n)
    defaults <- suppressWarnings(as.integer(input$max_iterations))
    if (is.na(defaults) || defaults < 0) {
      defaults <- 0L
    }

    tags$div(
      class = "mt-2",
      lapply(n_values, function(n) {
        input_id <- paste0("iter_n_", n)
        current <- input[[input_id]]
        if (is.null(current) || is.na(current)) {
          current <- defaults
        }
        numericInput(
          inputId = input_id,
          label = sprintf("Iterations for n = %d", n),
          value = max(0L, as.integer(current)),
          min = 0,
          step = 1
        )
      })
    )
  })

  format_number <- function(x) {
    ifelse(is.finite(x), format(x, big.mark = ",", trim = TRUE, scientific = FALSE), "∞")
  }

  output$combination_plan_summary <- renderUI({
    info <- plan_info()
    if (is.null(info)) {
      return(div(class = "text-muted", "Upload an input file to preview combination capacity."))
    }

    plan <- as.data.frame(info$plan)
    context <- info$context

    if (!nrow(plan)) {
      return(div(class = "text-warning", "No feasible combinations for the current N range and constraints."))
    }

    total_requested <- sum(plan$requested_iterations)
    total_to_run <- sum(plan$iterations_to_run)
    limited_ns <- plan$n_size[plan$exhausted]
    unlimited_ns <- plan$n_size[plan$available_unbounded]

    msgs <- list(
      div(class = "mb-1",
          tags$strong(sprintf("Total iterations to run: %s", format(total_to_run, big.mark = ",", scientific = FALSE))),
          span(sprintf(" (requested %s)", format(total_requested, big.mark = ",", scientific = FALSE)))
      ),
      div(class = "mb-1 text-muted",
          sprintf("Theme-constrained feasible N range: %s to %s",
                  format(context$min_total_required, big.mark = ",", scientific = FALSE),
                  format(context$max_total_allowed, big.mark = ",", scientific = FALSE)))
    )

    if (length(limited_ns)) {
      msgs <- append(msgs, list(
        div(class = "text-warning",
            sprintf("Requested iterations exceeded feasible combinations for n = %s.",
                    paste(sort(unique(limited_ns)), collapse = ", ")))
      ))
    }

    if (length(unlimited_ns)) {
      msgs <- append(msgs, list(
        div(class = "text-info",
            sprintf("Unlimited combinations detected for n = %s (iterations capped by request).",
                    paste(sort(unique(unlimited_ns)), collapse = ", ")))
      ))
    }

    do.call(tagList, msgs)
  })

  output$input_summary_cards <- renderUI({
    info <- plan_info()
    if (is.null(info)) {
      placeholder <- summary_card(
        title = "Awaiting workbook",
        value = "Upload to begin",
        subtitle = "Select a poLCA workbook to unlock live feasibility metrics.",
        icon_tag = icon("upload"),
        status_class = "muted"
      )
      return(div(class = "summary-card-grid", placeholder))
    }

    plan <- as.data.frame(info$plan)
    if (!nrow(plan)) {
      no_plan <- summary_card(
        title = "No feasible plan",
        value = "Adjust configuration",
        subtitle = "Try widening the N range or relaxing theme minimums to reveal combinations.",
        icon_tag = icon("exclamation-triangle"),
        status_class = "warning"
      )
      return(div(class = "summary-card-grid", no_plan))
    }

    total_requested <- sum(plan$requested_iterations)
    total_to_run <- sum(plan$iterations_to_run)
    limited_ns <- plan$n_size[which(plan$exhausted)]
    unlimited_ns <- plan$n_size[which(plan$available_unbounded)]
    theme_summary <- tryCatch(as.data.frame(info$context$theme_summary), error = function(e) data.frame())
    theme_count <- if (!nrow(theme_summary)) 0L else nrow(theme_summary)
    mandatory_count <- length(info$context$mandatory_vars %||% character(0))
    range_text <- sprintf(
      "%s – %s",
      format(info$context$min_total_required, big.mark = ",", scientific = FALSE),
      format(info$context$max_total_allowed, big.mark = ",", scientific = FALSE)
    )

    cards <- list(
      summary_card(
        title = "N sizes evaluated",
        value = format(length(unique(plan$n_size)), big.mark = ",", scientific = FALSE),
        subtitle = "Feasibility computed across your selected range.",
        icon_tag = icon("th-large"),
        status_class = "info"
      ),
      summary_card(
        title = "Scheduled iterations",
        value = format(total_to_run, big.mark = ",", scientific = FALSE),
        subtitle = sprintf("Requested %s total iterations.", format(total_requested, big.mark = ",", scientific = FALSE)),
        icon_tag = icon("play"),
        status_class = "success"
      ),
      summary_card(
        title = "Mandatory variables",
        value = format(mandatory_count, big.mark = ",", scientific = FALSE),
        subtitle = sprintf("Spanning %d themed pools.", theme_count),
        icon_tag = icon("lock"),
        status_class = NULL
      ),
      summary_card(
        title = "Theme-feasible selections",
        value = range_text,
        subtitle = "Minimum vs. maximum picks allowed across all themes.",
        icon_tag = icon("sliders"),
        status_class = NULL
      )
    )

    if (length(limited_ns)) {
      cards <- append(cards, list(
        summary_card(
          title = "Limited N sizes",
          value = paste0("n = ", paste(sort(unique(limited_ns)), collapse = ", ")),
          subtitle = "Iterations capped by feasible combinations.",
          icon_tag = icon("exclamation-triangle"),
          status_class = "warning"
        )
      ))
    }

    if (length(unlimited_ns)) {
      cards <- append(cards, list(
        summary_card(
          title = "Unlimited combinations",
          value = paste0("n = ", paste(sort(unique(unlimited_ns)), collapse = ", ")),
          subtitle = "Iterations restricted only by your requested volume.",
          icon_tag = icon("repeat"),
          status_class = "info"
        )
      ))
    }

    div(class = "summary-card-grid", do.call(tagList, cards))
  })

  output$combination_plan_table <- DT::renderDataTable({
    info <- plan_info()
    if (is.null(info)) {
      return(DT::datatable(
        data.frame(Message = "Upload an input file to preview plan."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    plan <- as.data.frame(info$plan)
    if (!nrow(plan)) {
      return(DT::datatable(
        data.frame(Message = "No feasible combinations for the selected configuration."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    display <- data.frame(
      `n` = plan$n_size,
      `Requested Iterations` = format(plan$requested_iterations, big.mark = ",", scientific = FALSE),
      `Feasible Combinations` = format_number(plan$max_possible_iterations),
      `Iterations to Run` = format(plan$iterations_to_run, big.mark = ",", scientific = FALSE),
      `Feasible?` = ifelse(plan$iterations_to_run > 0, "Yes", "No"),
      `Status` = ifelse(plan$exhausted, "Limited", "OK")
    )

    DT::datatable(
      display,
      options = list(pageLength = 5, searching = FALSE, lengthChange = FALSE, ordering = FALSE),
      rownames = FALSE
    )
  })

  output$combination_plan_plot <- plotly::renderPlotly({
    info <- plan_info()
    validate(need(!is.null(info), "Upload an input file to preview combinations."))
    plan <- as.data.frame(info$plan)
    validate(need(nrow(plan) > 0, "No feasible combinations to plot."))

    feasible <- plan$max_possible_iterations
    feasible_plot <- feasible
    feasible_plot[!is.finite(feasible_plot)] <- NA

    df <- data.frame(
      n = plan$n_size,
      Requested = plan$requested_iterations,
      Planned = plan$iterations_to_run,
      Feasible = feasible_plot
    )

    plotly::plot_ly(df, x = ~n) %>%
      plotly::add_bars(y = ~Requested, name = "Requested", marker = list(color = "#6c757d")) %>%
      plotly::add_bars(y = ~Feasible, name = "Feasible max", marker = list(color = "#17a2b8")) %>%
      plotly::add_bars(y = ~Planned, name = "Scheduled", marker = list(color = "#007bff")) %>%
      plotly::layout(
        barmode = "group",
        xaxis = list(title = "n size"),
        yaxis = list(title = "Iterations"),
        legend = list(orientation = "h", x = 0.1, y = -0.2)
      )
  })

  output$theme_summary_table <- DT::renderDataTable({
    info <- plan_info()
    if (is.null(info)) {
      return(DT::datatable(
        data.frame(Message = "Upload an input file to view theme-level constraints."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    theme_df <- as.data.frame(info$context$theme_summary)
    if (!nrow(theme_df)) {
      return(DT::datatable(
        data.frame(Message = "No theme constraints detected."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    display <- theme_df %>%
      mutate(
        min = as.integer(min),
        max = as.integer(max),
        mandatory = as.integer(mandatory),
        optional = as.integer(optional),
        total_available = as.integer(total_available)
      ) %>%
      rename(
        `Theme` = Theme,
        `Min Required` = min,
        `Max Allowed` = max,
        `Mandatory Vars` = mandatory,
        `Optional Vars` = optional,
        `Total Available` = total_available
      )

    DT::datatable(
      display,
      options = list(pageLength = 10, searching = FALSE, lengthChange = FALSE, ordering = TRUE),
      rownames = FALSE
    )
  })

  observeEvent(input$run_input_generator, {
    req(input$input_file)

    plan_details <- isolate(plan_info())
    if (is.null(plan_details)) {
      showNotification("❌ Unable to compute iteration plan. Check your inputs and try again.", type = "error", duration = 8)
      return()
    }

    plan_df <- as.data.frame(plan_details$plan)
    if (!nrow(plan_df)) {
      showNotification("⚠️ No feasible combinations for the current configuration.", type = "warning", duration = 8)
      return()
    }

    iterations_to_run <- sum(plan_df$iterations_to_run)
    if (iterations_to_run <= 0) {
      showNotification("⚠️ Zero iterations scheduled. Adjust requested iterations or constraints.", type = "warning", duration = 8)
      return()
    }

    data_path <- normalizePath(input$input_file$datapath, mustWork = TRUE)
    min_n_size <- as.integer(input$min_n_size)
    max_n_size <- as.integer(input$max_n_size)
    default_iterations <- suppressWarnings(as.integer(input$max_iterations))
    if (is.na(default_iterations) || default_iterations < 0) {
      default_iterations <- 0L
    }
    mandatory_vars <- parse_mandatory_indices(input$mandatory_vars)
    output_file_name <- input$output_filename
    include_fa <- isTRUE(input$include_fa)
    include_corr <- isTRUE(input$include_corr)

    cumulative_targets <- cumsum(plan_df$iterations_to_run)
    n_sequence <- plan_df$n_size
    total_steps <- iterations_to_run
    completed <- 0L

    start_time <- Sys.time()
    showModal(modalDialog(
      easyClose = FALSE, fade = TRUE, footer = NULL,
      size = "l",
      title = div(style = "font-size:22px;font-weight:bold;text-align:center;",
                  "Generating Combinations..."),
      div(
        style = "width:100%;text-align:center;padding:20px;",
        shinyWidgets::progressBar(
          id = "progressBarModal",
          value = 0,
          total = 100,
          display_pct = FALSE,
          striped = TRUE,
          status = "info"
        ),
        br(),
        div(id = "progressTextModal",
            style = "font-size:18px;margin-top:20px;",
            "Initializing..."),
        div(id = "etaTextModal",
            style = "font-size:16px;margin-top:12px;color:#1d4ed8;font-weight:500;",
            "Estimated completion: calculating...")
      )
    ))

    progress_cb <- function(n, iter_i, iter_max) {
      idx <- match(n, n_sequence)
      prior <- if (!is.na(idx) && idx > 1) cumulative_targets[idx - 1] else 0
      current <- prior + iter_i
      completed <<- max(completed, current)
      pct <- if (total_steps > 0) round(min(100, (completed / total_steps) * 100)) else 100
      shinyWidgets::updateProgressBar(session, "progressBarModal", value = pct)
      shinyjs::html(
        "progressTextModal",
        sprintf("Processing: n = %d | Iteration %d / %d", n, iter_i, iter_max)
      )
      if (total_steps > 0 && completed > 0) {
        elapsed <- as.numeric(difftime(Sys.time(), start_time, units = "secs"))
        est_total <- (elapsed / completed) * total_steps
        remaining <- max(0, est_total - elapsed)
        shinyjs::html("etaTextModal", format_eta(remaining, Sys.time() + remaining))
      } else {
        shinyjs::html("etaTextModal", "Estimated completion: calculating...")
      }
      flush.console()
    }

    tryCatch({
      generate_best_combinations(
        data = data_path,
        min_n_size = min_n_size,
        max_n_size = max_n_size,
        max_iterations = default_iterations,
        mandatory_variables_cols = mandatory_vars,
        output_file = file.path(DIRECTORY_PATH, output_file_name),
        append_log = append_log,
        progress_cb = progress_cb,
        plan_summary = plan_details,
        include_fa = include_fa,
        include_corr = include_corr
      )
      shinyWidgets::updateProgressBar(session, "progressBarModal", value = 100)
      shinyjs::html("progressTextModal", "Finalizing output...")
      shinyjs::html(
        "etaTextModal",
        sprintf("Estimated completion: completed at %s", format(Sys.time(), "%H:%M:%S"))
      )
      Sys.sleep(0.5)
      removeModal()
      if (any(plan_df$exhausted)) {
        showNotification("⚠️ Generation complete with some iteration requests capped by feasible combinations.", type = "warning", duration = 8)
      } else {
        showNotification("✅ Combination generation complete!", type = "message", duration = 5)
      }
    }, error = function(e) {
      removeModal()
      showNotification(paste("❌ Error:", e$message), type = "error", duration = 10)
    })
  })

  # --------------------------------------------------------
  # OUTPUT GENERATOR
  # --------------------------------------------------------
  observeEvent(input$xtabs_input_file, {
    if (is.null(input$xtabs_input_file)) {
      return()
    }

    selected <- shinyFiles::parseFilePaths(volumes, input$xtabs_input_file)
    if (!nrow(selected)) {
      return()
    }

    selected_path <- as.character(selected$datapath[1])
    if (!length(selected_path) || is.na(selected_path)) {
      return()
    }

    normalized_path <- tryCatch(
      normalizePath(selected_path, winslash = "/", mustWork = FALSE),
      error = function(e) selected_path
    )
    normalized_dir <- tryCatch(
      normalizePath(dirname(normalized_path), winslash = "/", mustWork = FALSE),
      error = function(e) dirname(normalized_path)
    )

    rv_paths$xtabs_path <- normalized_path
    rv_paths$xtabs_dir <- normalized_dir
    rv_paths$xtabs_name <- basename(normalized_path)

    append_log(sprintf("poLCA input set to: %s", normalized_path))

    if (!is.null(rv_paths$xtabs_dir) && nzchar(rv_paths$xtabs_dir)) {
      updateTextInput(session, "output_directiory", value = rv_paths$xtabs_dir)
    }
  })

  observeEvent(input$combinations_file, {
    if (is.null(input$combinations_file)) {
      return()
    }

    selected <- shinyFiles::parseFilePaths(volumes, input$combinations_file)
    if (!nrow(selected)) {
      return()
    }

    selected_path <- as.character(selected$datapath[1])
    if (!length(selected_path) || is.na(selected_path)) {
      return()
    }

    normalized_path <- tryCatch(
      normalizePath(selected_path, winslash = "/", mustWork = FALSE),
      error = function(e) selected_path
    )
    normalized_dir <- tryCatch(
      normalizePath(dirname(normalized_path), winslash = "/", mustWork = FALSE),
      error = function(e) dirname(normalized_path)
    )

    rv_paths$combos_path <- normalized_path
    rv_paths$combos_dir <- normalized_dir
    rv_paths$combos_name <- basename(normalized_path)

    append_log(sprintf("Combinations sheet set to: %s", normalized_path))
  })

  output$xtabs_input_file_display <- renderText({
    if (is.null(rv_paths$xtabs_path)) {
      "Selected poLCA file: (none)"
    } else {
      paste0("Selected poLCA file: ", rv_paths$xtabs_path)
    }
  })

  output$combinations_file_display <- renderText({
    if (is.null(rv_paths$combos_path)) {
      "Selected combinations sheet: (none)"
    } else {
      paste0("Selected combinations sheet: ", rv_paths$combos_path)
    }
  })

  output$output_generator_cards <- renderUI({
    safe_value <- function(x, fallback = "Not selected") {
      if (is.null(x) || !length(x)) {
        return(fallback)
      }
      if (is.na(x)) {
        return(fallback)
      }
      if (is.character(x) && !nzchar(x[1])) {
        return(fallback)
      }
      x
    }

    xtabs_name <- safe_value(rv_paths$xtabs_name)
    xtabs_dir <- safe_value(rv_paths$xtabs_dir, "Choose the workbook to unlock exports.")
    combos_name <- safe_value(rv_paths$combos_name)
    combos_dir <- safe_value(rv_paths$combos_dir, "Link the combinations sheet generated from the input planner.")

    output_dir_raw <- input$output_directiory
    if (is.null(output_dir_raw)) {
      output_dir_raw <- ""
    }
    output_dir_clean <- safe_value(output_dir_raw, "")
    output_dir_label <- if (identical(output_dir_clean, "")) "Not selected" else basename(output_dir_clean)
    output_dir_subtitle <- if (identical(output_dir_clean, "")) {
      "Select a destination to save generated ITR files."
    } else {
      output_dir_clean
    }
    output_dir_status <- if (identical(output_dir_clean, "")) "muted" else "success"

    segments <- safe_value(input$num_segments, "-")

    cards <- list(
      summary_card(
        title = "poLCA input",
        value = xtabs_name,
        subtitle = xtabs_dir,
        icon_tag = icon("database"),
        status_class = if (identical(xtabs_name, "Not selected")) "muted" else "info"
      ),
      summary_card(
        title = "Combination plan",
        value = combos_name,
        subtitle = combos_dir,
        icon_tag = icon("sitemap"),
        status_class = if (identical(combos_name, "Not selected")) "muted" else "info"
      ),
      summary_card(
        title = "Output directory",
        value = output_dir_label,
        subtitle = output_dir_subtitle,
        icon_tag = icon("folder-open"),
        status_class = output_dir_status
      ),
      summary_card(
        title = "Segments",
        value = as.character(segments),
        subtitle = "Number of segments to generate per iteration run.",
        icon_tag = icon("users"),
        status_class = NULL
      ),
      summary_card(
        title = "Generator status",
        value = status_text(),
        subtitle = "This mirrors the live updates shown in the notification bar.",
        icon_tag = icon("info-circle"),
        status_class = NULL
      )
    )

    div(class = "summary-card-grid", do.call(tagList, cards))
  })

  observeEvent(input$run_output_generator, {
    req(rv_paths$xtabs_path, rv_paths$combos_path)

    if (!file.exists(rv_paths$xtabs_path)) {
      status_text("poLCA input file not found. Please reselect and try again.")
      showNotification("The selected poLCA input file could not be found.", type = "error")
      return()
    }

    if (!file.exists(rv_paths$combos_path)) {
      status_text("Combinations sheet not found. Please reselect and try again.")
      showNotification("The selected combinations sheet could not be found.", type = "error")
      return()
    }

    xtabs_input_path <- rv_paths$xtabs_dir
    xtabs_input_db_name <- rv_paths$xtabs_name
    file_name <- input$file_name
    num_segments <- input$num_segments

    if (is.null(xtabs_input_path) || !dir.exists(xtabs_input_path)) {
      status_text("Directory for poLCA input not found.")
      showNotification("The directory for the selected poLCA input could not be found.", type = "error")
      return()
    }

    output_working_dir <- input$output_directiory
    if (is.null(output_working_dir)) {
      output_working_dir <- ""
    }
    output_working_dir <- tryCatch(
      normalizePath(output_working_dir, winslash = "/", mustWork = FALSE),
      error = function(e) output_working_dir
    )

    if (!nzchar(output_working_dir) || !dir.exists(output_working_dir)) {
      status_text("Please choose a valid output directory.")
      showNotification("Please select a valid output directory before running the generator.", type = "error")
      return()
    }

    combos_dir <- rv_paths$combos_dir
    if (is.null(combos_dir) || !nzchar(combos_dir)) {
      combos_dir <- dirname(rv_paths$combos_path)
    }
    if (!dir.exists(combos_dir)) {
      status_text("Directory for combinations sheet not found.")
      showNotification("The directory for the combinations sheet could not be found.", type = "error")
      return()
    }
    comb_input <- list(combos_dir, rv_paths$combos_name)

    append_log("Starting Output Generator...")
    shinyjs::disable("run_output_generator")
    on.exit(shinyjs::enable("run_output_generator"), add = TRUE)

    status_text("Running output generator...")

    start_time <- Sys.time()
    showModal(modalDialog(
      easyClose = FALSE, fade = TRUE, footer = NULL,
      size = "l",
      title = div(style = "font-size:22px;font-weight:bold;text-align:center;",
                  "Running Output Generator..."),
      div(
        style = "width:100%;text-align:center;padding:20px;",
        shinyWidgets::progressBar(
          id = "outputProgressBarModal",
          value = 0,
          total = 100,
          display_pct = FALSE,
          striped = TRUE,
          status = "info"
        ),
        br(),
        div(id = "outputProgressTextModal",
            style = "font-size:18px;margin-top:20px;",
            "Initializing output generator..."),
        div(id = "outputEtaTextModal",
            style = "font-size:16px;margin-top:12px;color:#1d4ed8;font-weight:500;",
            "Estimated completion: calculating...")
      )
    ))
    on.exit(removeModal(), add = TRUE)

    progress_update <- function(current, total, iteration, status) {
      pct <- 0
      if (!is.null(total) && is.finite(total) && total > 0) {
        pct <- round(max(0, min(100, (current / total) * 100)))
      } else if (!is.null(status) && !identical(status, NA) && grepl("complete", status, ignore.case = TRUE)) {
        pct <- 100
      }
      shinyWidgets::updateProgressBar(session, "outputProgressBarModal", value = pct)

      message_text <- NULL
      if (!is.null(status) && !identical(status, NA) && nzchar(status)) {
        message_text <- status
      } else if (!is.null(iteration) && !identical(iteration, NA)) {
        message_text <- sprintf("Processing iteration %s", iteration)
      } else if (!is.null(total) && is.finite(total) && total > 0) {
        message_text <- sprintf("Completed %d of %d iterations", current, total)
      } else {
        message_text <- "Processing iterations..."
      }

      shinyjs::html("outputProgressTextModal", message_text)

      if (!is.null(total) && is.finite(total) && total > 0 &&
          !is.null(current) && is.finite(current) && current > 0) {
        elapsed <- as.numeric(difftime(Sys.time(), start_time, units = "secs"))
        est_total <- (elapsed / current) * total
        remaining <- max(0, est_total - elapsed)
        shinyjs::html(
          "outputEtaTextModal",
          format_eta(remaining, Sys.time() + remaining)
        )
      } else {
        shinyjs::html("outputEtaTextModal", "Estimated completion: calculating...")
      }
    }

    progress_update(0, 1, NA, "Preparing output generator...")

    tryCatch({
      run_iterations(
        comb_input,
        xtabs_input_path,
        output_working_dir,
        xtabs_input_db_name,
        num_segments,
        file_name,
        solo = FALSE,
        progress_cb = progress_update
      )
      append_log("Output generation completed successfully.")
      status_text("Output generation completed successfully.")
      shinyWidgets::updateProgressBar(session, "outputProgressBarModal", value = 100)
      shinyjs::html("outputProgressTextModal", "Finalizing output files...")
      shinyjs::html(
        "outputEtaTextModal",
        sprintf("Estimated completion: completed at %s", format(Sys.time(), "%H:%M:%S"))
      )
      Sys.sleep(0.5)
      showNotification("✅ Output generation completed!", type = "message", duration = 5)
    }, error = function(e) {
      append_log(paste("❌ Error:", e$message))
      status_text(paste("Error in output generator:", e$message))
      showNotification(paste("❌ Error in Output Generator:", e$message), type = "error", duration = 10)
    })
  })

  # --------------------------------------------------------
  # ITERATION RUNNER
  # --------------------------------------------------------
  output$manual_selected_table <- DT::renderDT({
    metadata <- manual_state$metadata
    if (is.null(metadata) || !nrow(metadata)) {
      return(DT::datatable(
        data.frame(Message = "Load a poLCA workbook to populate variables."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    selected <- input$manual_variables
    if (is.null(selected) || !length(selected)) {
      return(DT::datatable(
        data.frame(Message = "Select variables to preview details."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    display <- metadata %>%
      dplyr::filter(variable %in% selected) %>%
      dplyr::mutate(order = match(variable, selected)) %>%
      dplyr::arrange(order) %>%
      dplyr::transmute(
        Variable = variable,
        Description = ifelse(is.na(description) | !nzchar(description), "-", description),
        Theme = ifelse(is.na(theme) | !nzchar(theme), "-", theme)
      )

    DT::datatable(
      display,
      options = list(dom = "t", paging = FALSE, ordering = FALSE),
      rownames = FALSE,
      escape = FALSE
    )
  })

  output$manual_save_info <- renderUI({
    if (is.null(manual_state$output_path)) {
      return(tags$span(class = "text-muted", "Run an iteration to generate manual outputs."))
    }

    tags$div(
      tags$strong("Last manual run"),
      tags$br(),
      tags$span(sprintf("Run name: %s", manual_state$run_name)),
      tags$br(),
      tags$span(sprintf("Saved CSV: %s", basename(manual_state$output_path))),
      tags$br(),
      tags$span(sprintf("Location: %s", manual_state$output_path)),
      if (!is.null(manual_state$last_run_time)) {
        tags$br()
      },
      if (!is.null(manual_state$last_run_time)) {
        tags$span(sprintf("Completed at: %s", format(manual_state$last_run_time, "%Y-%m-%d %H:%M:%S")))
      }
    )
  })

  observe({
    ready <- !is.null(manual_state$metadata) &&
      length(input$manual_variables) > 0 &&
      !isTRUE(manual_state$running)
    shinyjs::toggleState("run_manual_iteration", condition = ready)
  })

  observeEvent(input$load_files, {
    manual_state$metadata <- NULL
    manual_state$input_path <- NULL
    manual_state$input_dir <- NULL
    manual_state$segment_incidence <- NULL
    manual_state$of_metrics <- tibble::tibble()
    manual_state$variable_diff <- NULL
    manual_state$variable_tables <- list()
    if (length(manual_state$rendered_var_ids)) {
      for (id in manual_state$rendered_var_ids) {
        output[[id]] <- NULL
      }
    }
    manual_state$rendered_var_ids <- character(0)
    manual_state$output_path <- NULL
    manual_state$run_name <- NULL
    manual_state$last_run_time <- NULL
    manual_state$running <- FALSE

    shinyWidgets::updatePickerInput(
      session,
      "manual_variables",
      choices = setNames(character(0), character(0)),
      selected = character(0)
    )

    path_raw <- input$input_file_path
    if (is.null(path_raw) || !nzchar(path_raw)) {
      output$files_loaded <- renderText("Please provide a valid path to poLCA_input.xlsx.")
      return()
    }

    normalized <- normalize_if_exists(path_raw)
    if (!file.exists(normalized)) {
      append_log(sprintf("Manual runner: file not found at %s", normalized), "ERROR")
      output$files_loaded <- renderText(sprintf("❌ File not found: %s", path_raw))
      showNotification(sprintf("❌ File not found: %s", path_raw), type = "error", duration = 8)
      return()
    }

    tryCatch({
      metadata <- build_manual_metadata(normalized)
      if (!nrow(metadata)) {
        stop("No variables detected in INPUT sheet.")
      }

      manual_state$metadata <- metadata
      manual_state$input_path <- normalized
      manual_state$input_dir <- dirname(normalized)

      shinyWidgets::updatePickerInput(
        session,
        "manual_variables",
        choices = setNames(metadata$variable, metadata$variable),
        selected = character(0)
      )

      output$files_loaded <- renderText(sprintf(
        "Loaded %d variables from %s. Select variables below.",
        nrow(metadata),
        basename(normalized)
      ))
      append_log(sprintf("Manual runner: loaded %d variables from %s", nrow(metadata), normalized))
      showNotification("✅ poLCA workbook loaded. Select variables to run a manual iteration.", type = "message", duration = 5)
    }, error = function(e) {
      manual_state$metadata <- NULL
      manual_state$input_path <- NULL
      manual_state$input_dir <- NULL
      output$files_loaded <- renderText(sprintf("❌ %s", e$message))
      append_log(sprintf("Manual runner load error: %s", e$message), "ERROR")
      showNotification(paste("❌", e$message), type = "error", duration = 8)
    })
  })

  observeEvent(input$run_manual_iteration, {
    req(manual_state$metadata, manual_state$input_path)

    selected <- input$manual_variables
    if (is.null(selected) || !length(selected)) {
      showNotification("Select at least one variable before running.", type = "warning", duration = 5)
      return()
    }

    manual_state$running <- TRUE
    shinyjs::disable("run_manual_iteration")
    on.exit({
      manual_state$running <- FALSE
      shinyjs::enable("run_manual_iteration")
    }, add = TRUE)

    tryCatch({
      withProgress(message = "Running manual iteration...", value = 0, {
        incProgress(0.05, detail = "Preparing selection")
        metadata <- manual_state$metadata
        selection <- metadata %>%
          dplyr::filter(variable %in% selected) %>%
          dplyr::mutate(order = match(variable, selected)) %>%
          dplyr::arrange(order)

        if (!nrow(selection)) {
          stop("Selected variables were not found in the metadata table.")
        }

        column_indices <- selection$column_index
        manual_run_dir <- file.path(manual_state$input_dir %||% dirname(manual_state$input_path), "MANUAL_RUNS")
        if (!dir.exists(manual_run_dir)) {
          dir.create(manual_run_dir, recursive = TRUE, showWarnings = FALSE)
        }

        incProgress(0.15, detail = "Loading variable scores")
        variable_scores_data <- tryCatch(
          readxl::read_excel(manual_state$input_path, sheet = "VARIABLES", range = "B2:C999"),
          error = function(e) stop("Unable to read VARIABLES sheet: ", e$message)
        )

        if (ncol(variable_scores_data) < 2) {
          stop("VARIABLES sheet does not contain the expected columns.")
        }

        colnames(variable_scores_data)[1:2] <- c("Variable", "Score")
        variable_scores_data <- variable_scores_data %>%
          dplyr::mutate(
            Variable = trimws(as.character(Variable)),
            Score = suppressWarnings(as.numeric(Score))
          ) %>%
          dplyr::filter(!is.na(Variable) & nzchar(Variable))

        selection_scores <- selection %>%
          dplyr::mutate(Variable = trimws(as.character(variable))) %>%
          dplyr::left_join(variable_scores_data, by = "Variable")

        scores <- selection_scores$Score
        total_score <- sum(scores, na.rm = TRUE)
        avg_score <- if (length(scores)) total_score / length(scores) else NA_real_

        run_timestamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
        run_name <- paste0("MANUAL_RUN_", run_timestamp)
        summary_basename <- paste0(run_name, "_SUMMARY.xlsx")
        summary_path <- file.path(manual_run_dir, summary_basename)

        incProgress(0.35, detail = "Running segmentation")
        old_wd <- getwd()
        on.exit(setwd(old_wd), add = TRUE)
        LCA_SEG_Summary(
          input_working_dir = manual_state$input_dir,
          output_working_dir = manual_run_dir,
          input_db_name = basename(manual_state$input_path),
          file_run_and_date = run_name,
          cat_seg_var_list = selection$variable,
          cont_seg_var_list = c(),
          set_num_segments = as.integer(input$manual_num_segments),
          set_maxiter = 1000,
          set_nrep = 10,
          option_summary_only = 0
        )
        setwd(old_wd)

        incProgress(0.6, detail = "Evaluating metrics")
        old_wd_rules <- getwd()
        on.exit(setwd(old_wd_rules), add = TRUE)
        rules_result <- RULES_Checker(manual_run_dir, summary_basename, solo = TRUE)
        setwd(old_wd_rules)

        seg_n <- suppressWarnings(as.integer(rules_result$seg_n))
        if (!is.finite(seg_n) || seg_n <= 0) {
          seg_n <- length(unique(selection$variable))
        }

        seg_sizes_df <- tryCatch(as.data.frame(rules_result$seg_n_sizes), error = function(e) data.frame())
        if (!nrow(seg_sizes_df)) {
          seg_incidence <- tibble::tibble()
        } else {
          if (!"poLCA_SEG" %in% names(seg_sizes_df)) {
            seg_sizes_df$poLCA_SEG <- seq_len(nrow(seg_sizes_df))
          }
          if (!"n" %in% names(seg_sizes_df)) {
            seg_sizes_df$n <- 0
          }
          total_n <- sum(as.numeric(seg_sizes_df$n), na.rm = TRUE)
          seg_incidence <- seg_sizes_df %>%
            dplyr::mutate(
              Segment = paste0("Segment ", poLCA_SEG),
              Count = as.numeric(n),
              Percent = if (total_n > 0) (as.numeric(n) / total_n) * 100 else NA_real_
            ) %>%
            dplyr::select(Segment, Count, Percent)
        }

        manual_state$segment_incidence <- seg_incidence
        manual_state$of_metrics <- extract_of_metric_table(rules_result, max(1, seg_n))

        fetch_metric <- function(fmt) {
          vapply(seq_len(max(1, seg_n)), function(idx) {
            val <- rules_result[[sprintf(fmt, idx)]]
            if (is.null(val) || !length(val)) {
              return(NA_real_)
            }
            suppressWarnings(as.numeric(val))
          }, numeric(1))
        }

        seg_labels <- paste0("Segment ", seq_len(max(1, seg_n)))
        manual_state$variable_diff <- tibble::tibble(
          Segment = seg_labels,
          `Differentiating Vars` = fetch_metric("seg%d_diff"),
          `Bimodal Vars` = fetch_metric("bi_%d"),
          `Performance Vars` = fetch_metric("perf_%d"),
          `Indicator Top` = fetch_metric("indT_%d"),
          `Indicator Bottom` = fetch_metric("indB_%d")
        )

        incProgress(0.75, detail = "Extracting variable summaries")
        var_tables <- extract_manual_variable_tables(summary_path, selection$variable)

        if (length(manual_state$rendered_var_ids)) {
          for (id in manual_state$rendered_var_ids) {
            output[[id]] <- NULL
          }
        }

        manual_state$rendered_var_ids <- vapply(var_tables, function(tbl) tbl$output_id, character(1))
        manual_state$variable_tables <- var_tables

        for (entry in var_tables) {
          local({
            data_tbl <- entry$data
            output_id <- entry$output_id
            output[[output_id]] <- DT::renderDT({
              DT::datatable(
                data_tbl,
                options = list(pageLength = 10, scrollX = TRUE, dom = "t"),
                rownames = FALSE,
                escape = FALSE
              )
            })
          })
        }

        incProgress(0.85, detail = "Writing manual run output")

        manual_row <- list(
          Iteration = 1L,
          n_size = length(selected),
          Combination = paste(selection$variable, collapse = ";"),
          Total_Score = total_score,
          Avg_Score = avg_score,
          Var_cols = paste(column_indices, collapse = ","),
          Ran_Iteration_Flag = 1L,
          ITR_PATH = manual_run_dir,
          ITR_FILE_NAME = summary_basename
        )

        metric_map <- c(
          MAX_N_SIZE = "Max_n_size",
          MAX_N_SIZE_PERC = "Max_n_size_perc",
          MIN_N_SIZE = "Min_n_size",
          MIN_N_SIZE_PERC = "Min_n_size_perc",
          SOLUTION_N_SIZE = "seg_n",
          PROB_95 = "prob_95",
          PROB_90 = "prob_90",
          PROB_80 = "prob_80",
          PROB_75 = "prob_75",
          PROB_LESS_THAN_50 = "prob_low_50",
          BIMODAL_VARS = "final_bi_n",
          PROPER_BUCKETED_VARS = "proper_bucketted_vars",
          BIMODAL_VARS_PERC = "bi_n_perc",
          ROV_SD = "sd_rov",
          ROV_RANGE = "range_rov",
          BIMODAL_VARS_INPUT_VARS = "final_bi_n_input",
          INPUT_PERF_SEG_COVERED = "input_perf_seg_covered",
          INPUT_PERF_FLAG = "input_perf_flag",
          INPUT_PERF_MINSEG = "input_perf_minseg",
          INPUT_PERF_MAXSEG = "input_perf_maxseg"
        )

        for (nm in names(metric_map)) {
          src <- metric_map[[nm]]
          manual_row[[nm]] <- if (!is.null(rules_result[[src]])) rules_result[[src]] else NA
        }

        for (idx in 1:5) {
          manual_row[[paste0("seg", idx, "_diff")]] <- if (!is.null(rules_result[[paste0("seg", idx, "_diff")]])) rules_result[[paste0("seg", idx, "_diff")]] else NA
          manual_row[[paste0("bi_", idx)]] <- if (!is.null(rules_result[[paste0("bi_", idx)]])) rules_result[[paste0("bi_", idx)]] else NA
          manual_row[[paste0("perf_", idx)]] <- if (!is.null(rules_result[[paste0("perf_", idx)]])) rules_result[[paste0("perf_", idx)]] else NA
          manual_row[[paste0("indT_", idx)]] <- if (!is.null(rules_result[[paste0("indT_", idx)]])) rules_result[[paste0("indT_", idx)]] else NA
          manual_row[[paste0("indB_", idx)]] <- if (!is.null(rules_result[[paste0("indB_", idx)]])) rules_result[[paste0("indB_", idx)]] else NA
        }

        of_cols <- grep("^OF_", names(rules_result), value = TRUE)
        for (nm in of_cols) {
          manual_row[[nm]] <- rules_result[[nm]]
        }

        manual_df <- as.data.frame(manual_row, stringsAsFactors = FALSE)

        base_add_cols <- c(
          "Ran_Iteration_Flag", "ITR_PATH", "ITR_FILE_NAME", "MAX_N_SIZE", "MAX_N_SIZE_PERC",
          "MIN_N_SIZE", "MIN_N_SIZE_PERC", "SOLUTION_N_SIZE", "PROB_95", "PROB_90", "PROB_80",
          "PROB_75", "PROB_LESS_THAN_50", "BIMODAL_VARS", "PROPER_BUCKETED_VARS",
          "BIMODAL_VARS_PERC", "ROV_SD", "ROV_RANGE", "seg1_diff", "seg2_diff", "seg3_diff",
          "seg4_diff", "seg5_diff", "bi_1", "perf_1", "indT_1", "indB_1", "bi_2", "perf_2",
          "indT_2", "indB_2", "bi_3", "perf_3", "indT_3", "indB_3", "bi_4", "perf_4", "indT_4",
          "indB_4", "bi_5", "perf_5", "indT_5", "indB_5"
        )

        input_perf_cols <- c("BIMODAL_VARS_INPUT_VARS", "INPUT_PERF_SEG_COVERED", "INPUT_PERF_FLAG", "INPUT_PERF_MINSEG", "INPUT_PERF_MAXSEG")

        ordered_cols <- unique(c(
          "Iteration", "n_size", "Combination", "Total_Score", "Avg_Score", "Var_cols",
          base_add_cols,
          sort(grep("^OF_", names(manual_df), value = TRUE)),
          input_perf_cols
        ))

        ordered_cols <- intersect(ordered_cols, names(manual_df))
        leftover_cols <- setdiff(names(manual_df), ordered_cols)
        manual_df <- manual_df[, c(ordered_cols, leftover_cols), drop = FALSE]

        manual_output_path <- file.path(manual_run_dir, paste0(run_name, "_SUMMARY.csv"))
        utils::write.table(manual_df, file = manual_output_path, sep = ",", row.names = FALSE, col.names = TRUE, na = "")

        manual_state$output_path <- manual_output_path
        manual_state$run_name <- run_name
        manual_state$last_run_time <- Sys.time()

        incProgress(1, detail = "Manual run complete")
        append_log(sprintf("Manual iteration completed. Output saved to %s", manual_output_path))
        showNotification(sprintf("✅ Manual iteration complete. Saved to %s", basename(manual_output_path)), type = "message", duration = 6)
      })
    }, error = function(e) {
      append_log(sprintf("Manual iteration error: %s", e$message), "ERROR")
      showNotification(paste("❌ Manual iteration failed:", e$message), type = "error", duration = 10)
    })
  })

  output$manual_results_ui <- renderUI({
    if (is.null(manual_state$run_name)) {
      return(div(class = "text-muted", "Run a manual iteration to see results."))
    }

    metrics <- input$manual_metrics %||% character(0)
    if (!length(metrics)) {
      return(div(class = "text-muted", "Select one or more metrics to display."))
    }

    sections <- list()

    if ("Segment Incidence" %in% metrics) {
      sections <- append(sections, list(
        tags$div(
          class = "manual-metric-section",
          tags$h4("Segment Incidence"),
          DT::DTOutput("manual_segment_incidence")
        )
      ))
    }

    if ("OF Spread" %in% metrics) {
      sections <- append(sections, list(
        tags$div(
          class = "manual-metric-section",
          tags$h4("OF Spread"),
          plotly::plotlyOutput("manual_of_spread")
        )
      ))
    }

    if ("Variable Differentiation" %in% metrics) {
      sections <- append(sections, list(
        tags$div(
          class = "manual-metric-section",
          tags$h4("Variable Differentiation"),
          DT::DTOutput("manual_variable_diff")
        )
      ))
    }

    if ("Variable Summaries" %in% metrics) {
      var_sections <- lapply(manual_state$variable_tables, function(tbl) {
        tags$div(
          class = "manual-metric-section",
          tags$h4(tbl$variable),
          DT::DTOutput(tbl$output_id)
        )
      })
      if (!length(var_sections)) {
        var_sections <- list(tags$div(class = "text-muted", "Variable summaries unavailable for the selected run."))
      }
      sections <- append(sections, var_sections)
    }

    if (!length(sections)) {
      return(div(class = "text-muted", "Select one or more metrics to display."))
    }

    do.call(tagList, sections)
  })

  output$manual_segment_incidence <- DT::renderDT({
    req(manual_state$run_name)
    data <- manual_state$segment_incidence
    if (is.null(data) || !nrow(data)) {
      return(DT::datatable(
        data.frame(Message = "Segment incidence unavailable for this run."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    display <- data %>%
      dplyr::mutate(
        Percent = ifelse(is.na(Percent), "-", sprintf("%.1f%%", Percent))
      )

    DT::datatable(
      display,
      options = list(dom = "t", paging = FALSE, ordering = FALSE),
      rownames = FALSE
    )
  })

  output$manual_of_spread <- plotly::renderPlotly({
    req(manual_state$run_name)
    df <- manual_state$of_metrics
    validate(need(nrow(df) > 0, "Objective function metrics unavailable for this run."))
    df <- df %>% dplyr::filter(!is.na(Value))
    validate(need(nrow(df) > 0, "No objective function values to plot."))

    plotly::plot_ly(df, x = ~Segment, y = ~Value, color = ~Metric, type = "scatter", mode = "lines+markers",
                    fill = "tozeroy", text = ~paste0(
                      "Metric: ", Metric,
                      "<br>Segment: ", Segment,
                      "<br>Value: ", round(Value, 2),
                      "<br>Diff Max: ", round(DiffMax, 2),
                      "<br>Diff Min: ", round(DiffMin, 2)
                    ), hoverinfo = "text") %>%
      plotly::layout(
        xaxis = list(title = "Segment"),
        yaxis = list(title = "OF Score"),
        legend = list(orientation = "h", x = 0, y = -0.2)
      )
  })

  output$manual_variable_diff <- DT::renderDT({
    req(manual_state$run_name)
    data <- manual_state$variable_diff
    if (is.null(data) || !nrow(data)) {
      return(DT::datatable(
        data.frame(Message = "Variable differentiation metrics unavailable."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    display <- data %>%
      dplyr::mutate(dplyr::across(-Segment, ~ ifelse(is.na(.), "-", as.character(round(., 2)))))

    DT::datatable(
      display,
      options = list(dom = "t", paging = FALSE, ordering = FALSE, scrollX = TRUE),
      rownames = FALSE
    )
  })
}

# =========================================================
# 🚀 Launch App
# =========================================================
if (!exists("ui", inherits = FALSE)) {
  stop("❌ UI failed to initialize. Check earlier log messages for details.")
}
if (!exists("server", inherits = FALSE)) {
  stop("❌ Server failed to initialize. Check earlier log messages for details.")
}

shinyApp(ui, server)
