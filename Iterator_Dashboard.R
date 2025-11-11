# =========================================================
# 🚀 ITERATOR DASHBOARD (V7.6 — Modal Progress + Clean UI)
# =========================================================

# ---- Load Packages ----
packages <- c(
  "shiny", "shinydashboard", "DT", "shinyWidgets", "shinyFiles", "readxl", "shinyjs",
  "openxlsx", "dplyr", "tidyr", "lubridate", "plotly", "htmlwidgets", "RcppAlgos",
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
        .of-metric-panel {
          background: #ffffff;
          border: 1px solid rgba(148, 163, 184, 0.28);
          border-radius: 16px;
          padding: 18px 20px 20px 20px;
          margin-bottom: 22px;
          box-shadow: 0 12px 32px rgba(15, 23, 42, 0.08);
        }
        .of-metric-summary {
          display: flex;
          flex-wrap: wrap;
          gap: 8px;
          margin-bottom: 14px;
        }
        .of-stat-chip {
          display: inline-flex;
          align-items: center;
          padding: 6px 12px;
          border-radius: 999px;
          background: rgba(37, 99, 235, 0.08);
          color: #1e293b;
          font-size: 12px;
          font-weight: 600;
          letter-spacing: 0.2px;
        }
        .of-hist-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
          gap: 14px;
        }
        .of-mini-hist {
          background: #f8fafc;
          border-radius: 14px;
          border: 1px solid rgba(148, 163, 184, 0.2);
          padding: 14px 16px 12px 16px;
          box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.6);
        }
        .of-mini-title {
          font-size: 13px;
          font-weight: 600;
          margin-bottom: 8px;
          color: #0f172a;
        }
        .of-metric-title {
          font-weight: 700;
          margin-bottom: 12px;
          color: #1e293b;
          letter-spacing: 0.3px;
        }
        .of-table-wrapper {
          margin-top: 18px;
          padding: 16px 18px 10px 18px;
          border-radius: 14px;
          border: 1px solid rgba(148, 163, 184, 0.24);
          background: rgba(248, 250, 252, 0.8);
        }
        .manual-metric-section {
          background: rgba(237, 242, 247, 0.6);
          border: 1px solid rgba(148, 163, 184, 0.24);
          border-radius: 16px;
          padding: 18px 20px;
          margin-bottom: 22px;
        }
        .manual-metric-section h4 {
          font-weight: 600;
          color: #0f172a;
          margin-top: 0;
        }
        .of-picker-row {
          display: flex;
          flex-wrap: wrap;
          gap: 14px;
          margin-bottom: 16px;
        }
        .of-picker-row .form-group {
          flex: 1 1 260px;
          margin-bottom: 0;
        }
        table.dataTable td.segment-column,
        table.dataTable th.segment-column {
          font-weight: 600;
        }
        table.dataTable td[class*="segment-cell-"] {
          border-radius: 6px;
        }
        table.dataTable th[class*="segment-header-"] {
          border-radius: 6px 6px 0 0;
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
            div(
              class = "input-block",
              shinyFiles::shinyFilesButton(
                "manual_input_file",
                "Select poLCA workbook",
                "Browse...",
                multiple = FALSE
              ),
              textOutput(
                "manual_input_file_display",
                container = function(...) tags$small(class = "text-muted d-block mt-2", ...)
              )
            ),
            div(
              class = "input-block",
              actionButton("load_files", "Load Input", class = "btn btn-primary"),
              tags$small(
                class = "text-muted d-block mt-2",
                "Choose the poLCA_input workbook and click Load to populate variables."
              )
            ),
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
            div(
              class = "input-block",
              textInput(
                "manual_summary_file",
                "Manual iteration tracker filename",
                value = "MANUAL_RUN_SUMMARY.csv"
              )
            ),
            numericInput("manual_num_segments", "# of Segments", value = 4, min = 2, max = 10),
            actionButton("run_manual_iteration", "Run Iteration", class = "btn btn-primary btn-block"),
            br(),
            htmlOutput("manual_save_info"),
            br(),
            shinyjs::disabled(actionButton("open_summary", "Open Summary Workbook", class = "btn btn-default btn-block"))
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
    segment_summary = NULL,
    of_metrics_long = tibble::tibble(),
    of_metrics_table = tibble::tibble(),
    variable_diff = NULL,
    variable_diff_table = tibble::tibble(),
    variable_tables = list(),
    metrics_table = tibble::tibble(),
    available_of_metrics = character(0),
    of_distributions = list(),
    rendered_var_ids = character(0),
    rendered_hist_ids = character(0),
    category_choices = character(0),
    output_path = NULL,
    summary_workbook = NULL,
    run_name = NULL,
    last_run_time = NULL,
    running = FALSE,
    updating_vars = FALSE
  )

  segment_colors <- reactive({
    segs <- manual_state$segment_summary$Segment
    if (is.null(segs) || !length(segs)) {
      segs <- setdiff(names(manual_state$metrics_table), "Metric")
    }
    segs <- segs[nzchar(segs)]
    palette <- segment_color_palette(segs)
    if (!length(palette)) {
      return(list(header = character(0), cell = character(0)))
    }

    cell_cols <- vapply(palette, lighten_color, character(1), amount = 0.88)
    list(header = palette, cell = cell_cols)
  })

  status_text <- reactiveVal("Output generator idle.")

  output$status_run <- renderText({ status_text() })

  output$manual_segment_styles <- renderUI({
    colors <- segment_colors()
    header_cols <- colors$header %||% character(0)
    if (!length(header_cols)) {
      return(NULL)
    }

    cell_cols <- colors$cell %||% character(0)

    css_rules <- purrr::map_chr(names(header_cols), function(seg) {
      cls <- segment_css_class(seg)
      header_col <- header_cols[[seg]]
      cell_col <- cell_cols[[seg]] %||% lighten_color(header_col, 0.82)
      header_text <- segment_text_contrast(header_col)
      cell_text <- segment_text_contrast(cell_col)
      sprintf(
        ".segment-cell-%s { background-color: %s !important; color: %s !important; }\n.segment-header-%s { background-color: %s !important; color: %s !important; }\n        ",
        cls, cell_col, cell_text, cls, header_col, header_text
      )
    })

    tags$style(HTML(paste(css_rules, collapse = "\n")))
  })

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

normalize_of_metric_names <- function(metrics) {
  if (is.null(metrics) || !length(metrics)) {
    return(character(0))
  }

  upper <- toupper(metrics)
  unique_metrics <- unique(upper[grepl("^OF_", upper)])
  sort(unique_metrics)
}

filter_of_metric_distributions <- function(distributions) {
  if (!length(distributions)) {
    return(list())
  }

  names(distributions) <- toupper(names(distributions))
  valid_idx <- grepl("^OF_", names(distributions))
  distributions[valid_idx]
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

  sheets <- tryCatch(readxl::excel_sheets(summary_path), error = function(...) character(0))
  if (!length(sheets)) {
    return(list())
  }

  norm_key <- function(x) tolower(gsub("[^a-z0-9]", "", x))
  selected_norm <- unique(norm_key(selected_vars))
  if (!length(selected_norm)) {
    return(list())
  }

  candidate_sheets <- unique(c(
    sheets[seq_len(min(5, length(sheets)))],
    sheets[grep("seg|sum|var", tolower(sheets))]
  ))

  read_sheet <- function(sheet) {
    tryCatch(
      readxl::read_excel(summary_path, sheet = sheet, col_names = FALSE),
      error = function(...) NULL
    )
  }

  is_blank_row <- function(row) {
    all(is.na(row) | trimws(as.character(row)) == "")
  }

  split_tables <- function(df) {
    if (is.null(df) || !nrow(df)) return(list())
    blank_idx <- which(apply(df, 1, is_blank_row))
    cuts <- unique(c(1, blank_idx + 1, nrow(df) + 1))
    purrr::map(seq_len(length(cuts) - 1), function(i) {
      df[cuts[i]:(cuts[i + 1] - 1), , drop = FALSE]
    })
  }

  tidy_table <- function(tbl) {
    if (is.null(tbl) || !nrow(tbl)) {
      return(NULL)
    }

    cleaned <- tbl
    cleaned[] <- lapply(cleaned, function(col) {
      if (is.list(col)) {
        col <- vapply(col, function(x) paste(unlist(x), collapse = " "), character(1))
      }
      trimws(as.character(col))
    })

    cleaned <- cleaned[rowSums(cleaned == "" | is.na(cleaned)) != ncol(cleaned), , drop = FALSE]
    if (nrow(cleaned) < 2) {
      return(NULL)
    }

    header <- cleaned[1, , drop = TRUE]
    header <- ifelse(is.na(header) | header == "", paste0("Column_", seq_along(header)), header)
    header <- make.unique(as.character(header))

    cleaned <- cleaned[-1, , drop = FALSE]
    if (!nrow(cleaned)) {
      return(NULL)
    }

    cleaned <- cleaned[, colSums(cleaned == "" | is.na(cleaned)) < nrow(cleaned), drop = FALSE]
    if (!ncol(cleaned)) {
      return(NULL)
    }

    colnames(cleaned) <- header[seq_len(ncol(cleaned))]

    tibble::as_tibble(cleaned)
  }

  tables_by_var <- list()

  for (sheet in candidate_sheets) {
    raw <- read_sheet(sheet)
    if (is.null(raw) || !nrow(raw)) {
      next
    }

    purrr::walk(split_tables(raw), function(tbl) {
      if (is.null(tbl) || !nrow(tbl)) {
        return()
      }

      header_val <- tbl[1, 1]
      header_val <- ifelse(is.na(header_val), "", as.character(header_val))
      if (!nzchar(header_val)) {
        return()
      }

      var_name <- sub("^00_SEGM_", "", header_val)
      var_name <- trimws(var_name)
      norm_var <- norm_key(var_name)
      if (!nzchar(var_name) || !nzchar(norm_var) || !norm_var %in% selected_norm) {
        return()
      }

      cleaned_tbl <- tidy_table(tbl)
      if (is.null(cleaned_tbl) || !nrow(cleaned_tbl)) {
        return()
      }

      numeric_cols <- setdiff(names(cleaned_tbl), names(cleaned_tbl)[1])
      cleaned_tbl[numeric_cols] <- lapply(cleaned_tbl[numeric_cols], function(col) {
        vals <- suppressWarnings(as.numeric(gsub("[,%]", "", col)))
        if (all(is.na(vals))) {
          return(col)
        }
        vals
      })

      tables_by_var[[var_name]] <<- cleaned_tbl
    })
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

order_selected_first <- function(choices, selected) {
  if (!length(choices)) return(choices)
  selected <- unique(selected[selected %in% choices])
  c(selected, setdiff(choices, selected))
}

sanitize_metric_column_name <- function(metric, segment) {
  sanitize <- function(x) {
    key <- gsub("[^A-Za-z0-9]+", "_", x)
    key <- gsub("_+", "_", key)
    key <- sub("^_+", "", key)
    key <- sub("_+$", "", key)
    toupper(key)
  }
  paste0(sanitize(metric), "__", sanitize(segment))
}

normalize_key <- function(x) {
  tolower(gsub("[^a-z0-9]", "", x))
}

segment_color_palette <- function(labels) {
  labels <- unique(labels)
  labels <- labels[nzchar(labels)]
  if (!length(labels)) {
    return(character(0))
  }

  base_palette <- c(
    "#1D4ED8", "#0EA5E9", "#7C3AED", "#059669", "#D97706",
    "#DC2626", "#0891B2", "#2563EB", "#16A34A", "#6D28D9",
    "#BE123C", "#0F766E", "#9D174D", "#EA580C", "#0369A1"
  )

  if (length(labels) > length(base_palette)) {
    palette_fn <- grDevices::colorRampPalette(base_palette)
    colors <- palette_fn(length(labels))
  } else {
    colors <- base_palette[seq_len(length(labels))]
  }

  soften_palette <- function(cols, amount = 0.4) {
    vapply(cols, lighten_color, character(1), amount = amount)
  }
  colors <- soften_palette(colors, amount = 0.4)

  hashed <- vapply(labels, function(lbl) {
    ints <- utf8ToInt(lbl)
    if (!length(ints)) return(0L)
    sum((ints * seq_along(ints)) %% 9973L)
  }, integer(1))

  order_idx <- order((hashed %% 7919L) + seq_along(labels) / 1000)
  shuffled <- colors[order_idx]
  assigned <- shuffled[match(seq_along(labels), order_idx)]

  stats::setNames(assigned, labels)
}

mix_colors <- function(base, overlay, weight = 0.5) {
  safe_col <- function(col) {
    tryCatch(grDevices::col2rgb(col) / 255, error = function(...) matrix(c(0, 0, 0), nrow = 3))
  }

  b <- safe_col(base)
  o <- safe_col(overlay)
  mixed <- b * (1 - weight) + o * weight
  grDevices::rgb(mixed[1, ], mixed[2, ], mixed[3, ])
}

lighten_color <- function(color, amount = 0.65) {
  amount <- max(0, min(1, amount))
  mix_colors(color, "#FFFFFF", amount)
}

segment_css_class <- function(label) {
  key <- gsub("[^a-z0-9]", "-", tolower(label))
  gsub("-+", "-", key)
}

segment_text_contrast <- function(hex_color) {
  if (is.null(hex_color) || !nzchar(hex_color)) {
    return("#0F172A")
  }

  hex <- gsub("#", "", toupper(hex_color))
  if (nchar(hex) == 3) {
    hex <- paste(rep.int(substring(hex, 1:3, 1:3), each = 2), collapse = "")
  }
  if (nchar(hex) != 6) {
    return("#0F172A")
  }

  rgb_vals <- as.numeric(strtoi(substring(hex, c(1, 3, 5), c(2, 4, 6)), 16L)) / 255
  luminance <- 0.2126 * rgb_vals[1] + 0.7152 * rgb_vals[2] + 0.0722 * rgb_vals[3]
  if (is.na(luminance)) {
    return("#0F172A")
  }
  if (luminance < 0.55) "#FFFFFF" else "#0F172A"
}

build_segment_column_defs <- function(all_columns, segment_columns) {
  if (!length(segment_columns)) {
    return(list())
  }

  purrr::map(segment_columns, function(seg) {
    idx <- match(seg, all_columns)
    if (is.na(idx)) {
      return(NULL)
    }
    cls <- segment_css_class(seg)
    list(
      targets = idx - 1,
      className = paste("segment-column", paste0("segment-header-", cls)),
      createdCell = htmlwidgets::JS(sprintf(
        "function(td){td.classList.add('segment-cell-%s');}",
        cls
      ))
    )
  }) %>%
    purrr::compact()
}

build_segment_summary_table <- function(seg_incidence) {
  if (is.null(seg_incidence) || !nrow(seg_incidence)) {
    return(tibble::tibble())
  }

  seg_incidence <- seg_incidence %>%
    dplyr::mutate(
      Segment = as.character(Segment),
      Count = as.numeric(Count),
      Percent = as.numeric(Percent)
    )

  segment_cols <- seg_incidence$Segment
  tbl <- tibble::tibble(Metric = c("Segment Size", "Percent"))
  for (idx in seq_along(segment_cols)) {
    seg_name <- segment_cols[[idx]]
    tbl[[seg_name]] <- c(seg_incidence$Count[[idx]], seg_incidence$Percent[[idx]])
  }

  tbl
}

build_of_metrics_table <- function(of_long) {
  if (is.null(of_long) || !nrow(of_long)) {
    return(tibble::tibble())
  }

  of_long %>%
    dplyr::select(Metric, Segment, Value) %>%
    dplyr::mutate(
      Metric = as.character(Metric),
      Segment = as.character(Segment),
      Value = as.numeric(Value)
    ) %>%
    tidyr::pivot_wider(names_from = Segment, values_from = Value) %>%
    dplyr::arrange(Metric)
}

build_variable_diff_table <- function(diff_tbl) {
  if (is.null(diff_tbl) || !nrow(diff_tbl)) {
    return(tibble::tibble())
  }

  diff_tbl %>%
    tidyr::pivot_longer(-Segment, names_to = "Metric", values_to = "Value") %>%
    dplyr::mutate(
      Metric = as.character(Metric),
      Segment = as.character(Segment),
      Value = as.numeric(Value)
    ) %>%
    tidyr::pivot_wider(names_from = Segment, values_from = Value) %>%
    dplyr::arrange(Metric)
}

build_metric_category_summary <- function(var_tables, selected_metrics, segment_labels) {
  if (!length(var_tables) || !length(selected_metrics) || !length(segment_labels)) {
    return(tibble::tibble())
  }

  seg_norm <- normalize_key(segment_labels)
  seg_map <- stats::setNames(segment_labels, seg_norm)

  entries <- purrr::keep(var_tables, ~ .x$variable %in% selected_metrics)
  if (!length(entries)) {
    return(tibble::tibble())
  }

  processed <- purrr::map(entries, function(entry) {
    data_tbl <- entry$data
    if (is.null(data_tbl) || !nrow(data_tbl)) {
      return(NULL)
    }

    df <- as.data.frame(data_tbl, stringsAsFactors = FALSE)
    if (!ncol(df)) {
      return(NULL)
    }

    colnames(df) <- trimws(colnames(df))
    category_col <- colnames(df)[1]
    categories <- trimws(as.character(df[[category_col]]))

    seg_values <- lapply(names(seg_map), function(seg_norm_key) {
      seg_label <- seg_map[[seg_norm_key]]
      col_idx <- which(normalize_key(colnames(df)) == seg_norm_key)
      if (!length(col_idx)) {
        return(rep(NA_real_, nrow(df)))
      }
      vals <- suppressWarnings(as.numeric(df[[col_idx[1]]]))
      vals
    })

    if (!length(seg_values)) {
      return(NULL)
    }

    seg_df <- as.data.frame(seg_values, stringsAsFactors = FALSE)
    colnames(seg_df) <- segment_labels

    tibble::tibble(
      Metric = entry$variable,
      Category = categories
    ) %>%
      dplyr::bind_cols(seg_df)
  })

  processed <- purrr::compact(processed)
  if (!length(processed)) {
    return(tibble::tibble())
  }

  combined <- dplyr::bind_rows(processed)
  combined <- combined %>%
    dplyr::mutate(
      Category = ifelse(is.na(Category) | !nzchar(Category), "Unspecified", Category)
    )

  segment_cols <- segment_labels[nzchar(segment_labels)]
  segment_cols <- unique(segment_cols)
  for (col in segment_cols) {
    if (!col %in% names(combined)) {
      combined[[col]] <- NA_real_
    }
  }

  combined[, c("Metric", "Category", segment_cols), drop = FALSE]
}

align_metric_table <- function(tbl, segment_cols) {
  if (is.null(tbl) || !nrow(tbl)) {
    return(tibble::tibble())
  }

  tbl <- tbl %>% dplyr::mutate(Metric = as.character(Metric))
  for (col in segment_cols) {
    if (!col %in% names(tbl)) {
      tbl[[col]] <- NA_real_
    }
  }
  tbl[, c("Metric", segment_cols), drop = FALSE]
}

blank_segment_row <- function(segment_cols) {
  if (!length(segment_cols)) {
    return(tibble::tibble())
  }
  tibble::as_tibble(
    c(list(Metric = ""), stats::setNames(as.list(rep(NA_real_, length(segment_cols))), segment_cols))
  )
}

build_combined_metrics_table <- function(seg_table, of_table, diff_table) {
  segment_cols <- unique(c(names(seg_table), names(of_table), names(diff_table)))
  segment_cols <- setdiff(segment_cols, "Metric")
  segment_cols <- segment_cols[nzchar(segment_cols)]

  seg_table <- align_metric_table(seg_table, segment_cols)
  of_table <- align_metric_table(of_table, segment_cols)
  diff_table <- align_metric_table(diff_table, segment_cols)

  pieces <- list()
  if (nrow(seg_table)) {
    pieces <- append(pieces, list(seg_table))
  }
  if (nrow(of_table)) {
    pieces <- append(pieces, list(blank_segment_row(segment_cols), of_table))
  }
  if (nrow(diff_table)) {
    pieces <- append(pieces, list(blank_segment_row(segment_cols), diff_table))
  }

  if (!length(pieces)) {
    return(tibble::tibble())
  }

  dplyr::bind_rows(pieces)
}

extract_of_distributions <- function(summary_path, segment_labels) {
  if (!file.exists(summary_path) || !length(segment_labels)) {
    return(list())
  }

  sheets <- tryCatch(readxl::excel_sheets(summary_path), error = function(...) character(0))
  if (!length(sheets)) {
    return(list())
  }

  seg_norm_map <- normalize_key(segment_labels)
  seg_lookup <- stats::setNames(segment_labels, seg_norm_map)
  seg_lookup <- c(seg_lookup, stats::setNames(segment_labels, as.character(seq_along(segment_labels))))

  candidate_sheets <- unique(c(
    sheets[seq_len(min(5, length(sheets)))],
    sheets[grep("seg|res|manual|summary", tolower(sheets))]
  ))

  distributions <- list()

  for (sheet in candidate_sheets) {
    raw <- tryCatch(readxl::read_excel(summary_path, sheet = sheet), error = function(...) NULL)
    if (is.null(raw) || !nrow(raw)) {
      next
    }

    names(raw) <- make.unique(as.character(names(raw)))
    name_keys <- tolower(gsub("\\s+", "", names(raw)))
    seg_col_idx <- which(name_keys %in% c("polcaseg", "segment", "segmentid", "segment_id"))
    if (!length(seg_col_idx)) {
      next
    }

    seg_values <- raw[[seg_col_idx[1]]]
    seg_norm <- normalize_key(as.character(seg_values))
    seg_numeric <- suppressWarnings(as.integer(as.character(seg_values)))
    seg_labels <- seg_lookup[seg_norm]
    seg_labels_numeric <- segment_labels[seq_along(segment_labels)][seg_numeric]
    seg_labels <- ifelse(!is.na(seg_labels), seg_labels, seg_labels_numeric)

    if (all(is.na(seg_labels))) {
      next
    }

    of_cols <- grep("^OF_", names(raw), ignore.case = TRUE, value = TRUE)
    if (!length(of_cols)) {
      next
    }

    for (col in of_cols) {
      metric_key <- toupper(col)
      values <- suppressWarnings(as.numeric(gsub("[,%]", "", raw[[col]])))
      tbl <- tibble::tibble(Segment = seg_labels, Value = values) %>%
        dplyr::filter(!is.na(Segment) & nzchar(Segment) & !is.na(Value))
      if (!nrow(tbl)) {
        next
      }
      if (metric_key %in% names(distributions)) {
        distributions[[metric_key]] <- dplyr::bind_rows(distributions[[metric_key]], tbl)
      } else {
        distributions[[metric_key]] <- tbl
      }
    }
  }

  if (!length(distributions)) {
    return(list())
  }

  purrr::imap(distributions, function(tbl, nm) {
    tbl %>%
      dplyr::mutate(
        Segment = as.character(Segment),
        Value = as.numeric(Value)
      ) %>%
      dplyr::filter(!is.na(Segment) & nzchar(Segment) & !is.na(Value)) %>%
      dplyr::distinct()
  })
}

prepare_of_distribution_data <- function(distributions) {
  if (!length(distributions)) {
    return(list())
  }

  cleaned <- purrr::imap(distributions, function(tbl, nm) {
    if (is.null(tbl)) {
      return(NULL)
    }

    df <- tibble::as_tibble(tbl)
    segment_col <- names(df)[tolower(names(df)) %in% c("segment", "polca_seg", "segment_id", "segmentid")][1]
    value_col <- names(df)[tolower(names(df)) %in% c("value", "metric_value", "score")][1]

    if (is.na(segment_col) || is.na(value_col)) {
      return(NULL)
    }

    seg_vals <- as.character(df[[segment_col]])
    val_vals <- suppressWarnings(as.numeric(df[[value_col]]))
    df <- tibble::tibble(Segment = seg_vals, Value = val_vals) %>%
      dplyr::filter(!is.na(Segment) & nzchar(Segment) & !is.na(Value))

    if (!nrow(df)) {
      return(NULL)
    }

    df
  })

  cleaned <- purrr::compact(cleaned)
  if (!length(cleaned)) {
    return(list())
  }

  cleaned
}

write_manual_summary_workbook <- function(summary_path, metrics_table, var_tables, category_summary) {
  if (!nzchar(summary_path) || !file.exists(summary_path)) {
    return(invisible(FALSE))
  }

  try({
    wb <- openxlsx::loadWorkbook(summary_path)

    sheet_name <- "Manual Metrics"
    if (sheet_name %in% openxlsx::getSheetNames(wb)) {
      openxlsx::removeWorksheet(wb, sheet_name)
    }
    if (!is.null(metrics_table) && nrow(metrics_table)) {
      openxlsx::addWorksheet(wb, sheet_name)
      openxlsx::writeData(
        wb,
        sheet_name,
        as.data.frame(metrics_table, stringsAsFactors = FALSE)
      )
    }

    sheet_name <- "Manual Variable Summaries"
    if (sheet_name %in% openxlsx::getSheetNames(wb)) {
      openxlsx::removeWorksheet(wb, sheet_name)
    }
    if (length(var_tables)) {
      openxlsx::addWorksheet(wb, sheet_name)

      start_row <- 1
      for (entry in var_tables) {
        data_tbl <- entry$data
        if (is.null(data_tbl) || !nrow(data_tbl)) {
          next
        }

        openxlsx::writeData(
          wb,
          sheet = sheet_name,
          x = entry$variable,
          startRow = start_row,
          startCol = 1,
          colNames = FALSE
        )

        openxlsx::writeData(
          wb,
          sheet = sheet_name,
          x = as.data.frame(data_tbl, stringsAsFactors = FALSE),
          startRow = start_row + 1,
          startCol = 1,
          withFilter = FALSE
        )

        start_row <- start_row + nrow(data_tbl) + 3
      }
    }

    sheet_name <- "Manual Category Summary"
    if (sheet_name %in% openxlsx::getSheetNames(wb)) {
      openxlsx::removeWorksheet(wb, sheet_name)
    }
    if (!is.null(category_summary) && nrow(category_summary)) {
      openxlsx::addWorksheet(wb, sheet_name)
      openxlsx::writeData(
        wb,
        sheet = sheet_name,
        x = as.data.frame(category_summary, stringsAsFactors = FALSE)
      )
    }

    openxlsx::saveWorkbook(wb, summary_path, overwrite = TRUE)
    TRUE
  }, silent = TRUE)
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

  shinyFiles::shinyFileChoose(
    input, "manual_input_file",
    roots = volumes,
    session = session,
    filetypes = c("xlsx", "xls", "xlsm")
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

  observeEvent(input$manual_input_file, {
    if (is.null(input$manual_input_file)) {
      return()
    }

    selected <- shinyFiles::parseFilePaths(volumes, input$manual_input_file)
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

    manual_state$input_path <- normalized_path
    manual_state$input_dir <- normalized_dir

    append_log(sprintf("Manual runner workbook selected: %s", normalized_path))
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

  output$manual_input_file_display <- renderText({
    if (is.null(manual_state$input_path)) {
      "Selected poLCA workbook: (none)"
    } else {
      paste0("Selected poLCA workbook: ", manual_state$input_path)
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
    if (is.null(manual_state$output_path) && is.null(manual_state$summary_workbook)) {
      return(tags$span(class = "text-muted", "Run an iteration to generate manual outputs."))
    }

    tracker_path <- manual_state$output_path %||% ""
    summary_path <- manual_state$summary_workbook %||% ""
    tracker_label <- if (nzchar(tracker_path)) basename(tracker_path) else "Not saved"
    summary_label <- if (nzchar(summary_path)) basename(summary_path) else "Not saved"
    location <- if (nzchar(summary_path)) dirname(summary_path) else dirname(tracker_path)
    if (!nzchar(location) || identical(location, ".")) {
      location <- ""
    }

    tags$div(
      tags$strong("Last manual run"),
      tags$br(),
      tags$span(sprintf("Run name: %s", manual_state$run_name %||% "-")),
      tags$br(),
      tags$span(sprintf("Summary workbook: %s", summary_label)),
      tags$br(),
      tags$span(sprintf("Tracker file: %s", tracker_label)),
      if (nzchar(location)) {
        tagList(tags$br(), tags$span(sprintf("Output location: %s", location)))
      },
      if (!is.null(manual_state$last_run_time)) {
        tagList(tags$br(), tags$span(sprintf("Completed at: %s", format(manual_state$last_run_time, "%Y-%m-%d %H:%M:%S"))))
      }
    )
  })

  observe({
    ready <- !is.null(manual_state$metadata) &&
      length(input$manual_variables) > 0 &&
      !isTRUE(manual_state$running)
    shinyjs::toggleState("run_manual_iteration", condition = ready)
  })

  observe({
    summary_path <- manual_state$summary_workbook %||% ""
    shinyjs::toggleState("open_summary", condition = nzchar(summary_path) && file.exists(summary_path))
  })

  observeEvent(input$open_summary, {
    summary_path <- manual_state$summary_workbook %||% ""
    if (!nzchar(summary_path) || !file.exists(summary_path)) {
      showNotification("Summary workbook is unavailable for the current run.", type = "warning", duration = 5)
      return()
    }

    tryCatch({
      utils::browseURL(summary_path)
    }, error = function(e) {
      showNotification(sprintf("Unable to open summary workbook: %s", e$message), type = "error", duration = 6)
    })
  })

  observeEvent(input$load_files, {
    manual_state$metadata <- NULL
    manual_state$segment_summary <- NULL
    manual_state$of_metrics_long <- tibble::tibble()
    manual_state$of_metrics_table <- tibble::tibble()
    manual_state$metrics_table <- tibble::tibble()
    manual_state$variable_diff <- NULL
    manual_state$variable_diff_table <- tibble::tibble()
    manual_state$variable_tables <- list()
    manual_state$category_choices <- character(0)
    if (length(manual_state$rendered_var_ids)) {
      for (id in manual_state$rendered_var_ids) {
        output[[id]] <- NULL
      }
    }
    manual_state$rendered_var_ids <- character(0)
    if (length(manual_state$rendered_hist_ids)) {
      for (id in manual_state$rendered_hist_ids) {
        output[[id]] <- NULL
      }
    }
    manual_state$rendered_hist_ids <- character(0)
    manual_state$available_of_metrics <- character(0)
    manual_state$of_distributions <- list()
    manual_state$output_path <- NULL
    manual_state$summary_workbook <- NULL
    manual_state$run_name <- NULL
    manual_state$last_run_time <- NULL
    manual_state$running <- FALSE

    shinyWidgets::updatePickerInput(
      session,
      "manual_variables",
      choices = setNames(character(0), character(0)),
      selected = character(0)
    )
    shinyWidgets::updatePickerInput(
      session,
      "manual_metric_category_select",
      choices = setNames(character(0), character(0)),
      selected = character(0)
    )

    normalized <- manual_state$input_path
    if (is.null(normalized) || !nzchar(normalized)) {
      output$files_loaded <- renderText("Select a poLCA workbook using Browse before loading.")
      return()
    }

    if (!file.exists(normalized)) {
      append_log(sprintf("Manual runner: file not found at %s", normalized), "ERROR")
      output$files_loaded <- renderText(sprintf("❌ File not found: %s", normalized))
      showNotification(sprintf("❌ File not found: %s", normalized), type = "error", duration = 8)
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

      manual_state$updating_vars <- TRUE
      shinyWidgets::updatePickerInput(
        session,
        "manual_variables",
        choices = setNames(metadata$variable, metadata$variable),
        selected = character(0)
      )
      manual_state$updating_vars <- FALSE

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
      manual_state$available_of_metrics <- character(0)
      manual_state$of_distributions <- list()
      manual_state$of_metrics_long <- tibble::tibble()
      manual_state$of_metrics_table <- tibble::tibble()
      manual_state$metrics_table <- tibble::tibble()
      manual_state$variable_diff <- NULL
      manual_state$variable_diff_table <- tibble::tibble()
      manual_state$summary_workbook <- NULL
      manual_state$output_path <- NULL
      output$files_loaded <- renderText(sprintf("❌ %s", e$message))
      append_log(sprintf("Manual runner load error: %s", e$message), "ERROR")
      showNotification(paste("❌", e$message), type = "error", duration = 8)
    })
  })

  observeEvent(input$manual_variables, {
    metadata <- manual_state$metadata
    if (is.null(metadata) || !nrow(metadata)) {
      return()
    }

    if (isTRUE(manual_state$updating_vars)) {
      return()
    }

    choices <- metadata$variable
    selected <- input$manual_variables %||% character(0)
    ordered <- order_selected_first(choices, selected)
    if (identical(ordered, choices)) {
      return()
    }

    manual_state$updating_vars <- TRUE
    shinyWidgets::updatePickerInput(
      session,
      "manual_variables",
      choices = setNames(ordered, ordered),
      selected = selected
    )
    manual_state$updating_vars <- FALSE
  }, ignoreNULL = FALSE)

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

        manual_state$segment_summary <- seg_incidence
        segment_labels <- if (nrow(seg_incidence)) {
          as.character(seg_incidence$Segment)
        } else {
          paste0("Segment ", seq_len(max(1, seg_n)))
        }

        of_long <- extract_of_metric_table(rules_result, max(1, seg_n))
        if (nrow(of_long)) {
          of_long <- of_long %>%
            dplyr::mutate(
              Metric = toupper(Metric),
              Segment = as.character(Segment)
            )
        }
        manual_state$of_metrics_long <- of_long

        fetch_metric <- function(fmt) {
          vapply(seq_len(max(1, seg_n)), function(idx) {
            val <- rules_result[[sprintf(fmt, idx)]]
            if (is.null(val) || !length(val)) {
              return(NA_real_)
            }
            suppressWarnings(as.numeric(val))
          }, numeric(1))
        }

        variable_diff_tbl <- tibble::tibble(
          Segment = segment_labels,
          `Differentiating Vars` = fetch_metric("seg%d_diff"),
          `Bimodal Vars` = fetch_metric("bi_%d"),
          `Performance Vars` = fetch_metric("perf_%d"),
          `Indicator Top` = fetch_metric("indT_%d"),
          `Indicator Bottom` = fetch_metric("indB_%d")
        )

        manual_state$variable_diff <- variable_diff_tbl

        seg_table <- build_segment_summary_table(seg_incidence)
        of_table <- build_of_metrics_table(of_long)
        manual_state$of_metrics_table <- of_table
        var_diff_table <- build_variable_diff_table(variable_diff_tbl)
        manual_state$variable_diff_table <- var_diff_table
        combined_table <- build_combined_metrics_table(seg_table, of_table, var_diff_table)
        manual_state$metrics_table <- combined_table

        metric_keys <- if (nrow(of_long)) normalize_of_metric_names(of_long$Metric) else character(0)

        distributions <- extract_of_distributions(summary_path, segment_labels)
        if (!length(distributions) && nrow(of_long)) {
          distributions <- split(
            of_long %>% dplyr::select(Metric, Segment, Value),
            of_long$Metric
          )
        }

        normalized_distributions <- prepare_of_distribution_data(distributions)
        normalized_distributions <- filter_of_metric_distributions(normalized_distributions)
        if (length(normalized_distributions)) {
          manual_state$of_distributions <- normalized_distributions
          manual_state$available_of_metrics <- normalize_of_metric_names(c(metric_keys, names(normalized_distributions)))
        } else {
          manual_state$of_distributions <- list()
          manual_state$available_of_metrics <- metric_keys
        }

        incProgress(0.75, detail = "Extracting variable summaries")
        var_tables <- extract_manual_variable_tables(summary_path, selection$variable)

        if (length(manual_state$rendered_var_ids)) {
          for (id in manual_state$rendered_var_ids) {
            output[[id]] <- NULL
          }
        }

        manual_state$rendered_var_ids <- vapply(var_tables, function(tbl) tbl$output_id, character(1))
        manual_state$variable_tables <- var_tables
        manual_state$category_choices <- unique(vapply(var_tables, function(tbl) tbl$variable, character(1)))
        manual_state$category_choices <- manual_state$category_choices[nzchar(manual_state$category_choices)]

        category_summary <- build_metric_category_summary(var_tables, manual_state$category_choices, segment_labels)

        write_manual_summary_workbook(summary_path, combined_table, var_tables, category_summary)

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

        available_categories <- sort(manual_state$category_choices)
        shinyWidgets::updatePickerInput(
          session,
          "manual_metric_category_select",
          choices = setNames(available_categories, available_categories),
          selected = head(available_categories, min(3, length(available_categories)))
        )

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

        if (nrow(seg_incidence)) {
          for (idx in seq_len(nrow(seg_incidence))) {
            manual_row[[paste0("Segment_", idx, "_Count")]] <- seg_incidence$Count[idx]
            manual_row[[paste0("Segment_", idx, "_Percent")]] <- seg_incidence$Percent[idx]
          }
        }

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

        segment_metric_cols <- grep("__SEGMENT_", names(manual_df), value = TRUE)
        if (length(segment_metric_cols)) {
          manual_df <- manual_df[, setdiff(names(manual_df), segment_metric_cols), drop = FALSE]
        }

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
          sort(grep("^Segment_\\d+_", names(manual_df), value = TRUE)),
          sort(grep("^OF_", names(manual_df), value = TRUE)),
          input_perf_cols
        ))

        ordered_cols <- intersect(ordered_cols, names(manual_df))
        leftover_cols <- setdiff(names(manual_df), ordered_cols)
        manual_df <- manual_df[, c(ordered_cols, leftover_cols), drop = FALSE]

        tracker_name <- input$manual_summary_file %||% ""
        tracker_name <- trimws(if (is.null(tracker_name)) "" else tracker_name)
        if (!nzchar(tracker_name)) {
          tracker_name <- "MANUAL_RUN_SUMMARY.csv"
        }
        if (!grepl("\\.csv$", tracker_name, ignore.case = TRUE)) {
          tracker_name <- paste0(tracker_name, ".csv")
        }
        tracker_name <- basename(tracker_name)

        manual_tracker_path <- file.path(manual_run_dir, tracker_name)

        if (file.exists(manual_tracker_path)) {
          existing_df <- tryCatch(
            utils::read.csv(manual_tracker_path, stringsAsFactors = FALSE, check.names = FALSE),
            error = function(e) NULL
          )
          if (!is.null(existing_df) && nrow(existing_df)) {
            drop_cols <- grep("__SEGMENT_", names(existing_df), value = TRUE)
            if (length(drop_cols)) {
              existing_df <- existing_df[, setdiff(names(existing_df), drop_cols), drop = FALSE]
            }
          }
        } else {
          existing_df <- NULL
        }

        if (!is.null(existing_df) && nrow(existing_df)) {
          missing_in_existing <- setdiff(names(manual_df), names(existing_df))
          for (col in missing_in_existing) {
            existing_df[[col]] <- NA
          }
          missing_in_new <- setdiff(names(existing_df), names(manual_df))
          for (col in missing_in_new) {
            manual_df[[col]] <- NA
          }
          manual_df <- manual_df[, names(existing_df), drop = FALSE]
          combined_df <- rbind(existing_df, manual_df)
        } else {
          combined_df <- manual_df
        }

        utils::write.table(combined_df, file = manual_tracker_path, sep = ",", row.names = FALSE, col.names = TRUE, na = "")

        manual_state$output_path <- manual_tracker_path
        manual_state$summary_workbook <- summary_path
        manual_state$run_name <- run_name
        manual_state$last_run_time <- Sys.time()

        incProgress(0.95, detail = "Updating iteration tracker")
        incProgress(1, detail = "Manual run complete")
        append_log(sprintf("Manual iteration completed. Summary: %s | Tracker: %s", summary_basename, basename(manual_tracker_path)))
        showNotification(
          sprintf(
            "✅ Manual iteration complete. Summary workbook: %s. Tracker updated: %s",
            basename(summary_path),
            basename(manual_tracker_path)
          ),
          type = "message",
          duration = 6
        )
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

    sections <- list(uiOutput("manual_segment_styles"))

    if ("Segment Incidence" %in% metrics) {
      sections <- append(sections, list(
        tags$div(
          class = "manual-metric-section",
          tags$h4("Segment & Objective Summary"),
          DT::DTOutput("manual_metrics_table")
        )
      ))
    }

    if ("OF Spread" %in% metrics) {
      of_choices <- manual_state$available_of_metrics
      if (!length(of_choices)) {
        sections <- append(sections, list(
          tags$div(
            class = "manual-metric-section",
            tags$h4("Objective Function Metrics"),
            tags$span(class = "text-muted", "Objective function metrics unavailable for this run.")
          )
        ))
      } else {
        selected_metrics <- input$manual_of_metrics_selected %||% of_choices
        selected_metrics <- intersect(toupper(selected_metrics), of_choices)
        if (!length(selected_metrics)) selected_metrics <- of_choices

        display_choices <- gsub("_", " ", of_choices)
        named_choices <- stats::setNames(of_choices, display_choices)

        selected_table_metrics <- input$manual_of_table_metrics %||% selected_metrics
        selected_table_metrics <- intersect(toupper(selected_table_metrics), of_choices)
        if (!length(selected_table_metrics)) {
          selected_table_metrics <- selected_metrics
        }

        sections <- append(sections, list(
          tags$div(
            class = "manual-metric-section",
            tags$h4("Objective Function Metrics"),
            div(
              class = "of-picker-row",
              shinyWidgets::pickerInput(
                "manual_of_metrics_selected",
                "Select objective function metrics",
                choices = named_choices,
                selected = selected_metrics,
                multiple = TRUE,
                options = list(`actions-box` = TRUE, `live-search` = TRUE)
              ),
              shinyWidgets::pickerInput(
                "manual_of_table_metrics",
                "Select metrics for summary tables",
                choices = named_choices,
                selected = selected_table_metrics,
                multiple = TRUE,
                options = list(`actions-box` = TRUE, `live-search` = TRUE)
              )
            ),
            uiOutput("manual_of_histograms"),
            div(
              class = "of-table-wrapper",
              DT::DTOutput("manual_of_metric_table")
            )
          )
        ))

        category_choices <- manual_state$category_choices
        if (length(category_choices)) {
          category_named <- stats::setNames(category_choices, category_choices)
          default_categories <- category_choices[seq_len(min(3, length(category_choices)))]
          selected_categories <- input$manual_metric_category_select %||% default_categories
          selected_categories <- intersect(category_choices, selected_categories)
          if (!length(selected_categories) && length(default_categories)) {
            selected_categories <- default_categories
          }

          sections <- append(sections, list(
            tags$div(
              class = "manual-metric-section",
              tags$h4("Metric Category Summary"),
              shinyWidgets::pickerInput(
                "manual_metric_category_select",
                "Select metrics for category summary",
                choices = category_named,
                selected = selected_categories,
                multiple = TRUE,
                options = list(`actions-box` = TRUE, `live-search` = TRUE)
              ),
              DT::DTOutput("manual_metric_category_table")
            )
          ))
        }
      }
    }

    if ("Variable Differentiation" %in% metrics) {
      sections <- append(sections, list(
        tags$div(
          class = "manual-metric-section",
          tags$h4("Variable Differentiation"),
          DT::DTOutput("manual_variable_diff_table")
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

  format_manual_metric_table <- function(data_tbl) {
    if (is.null(data_tbl) || !nrow(data_tbl)) {
      return(NULL)
    }

    display <- data_tbl
    segment_cols <- setdiff(names(display), "Metric")
    if (!length(segment_cols)) {
      return(display)
    }

    for (col in segment_cols) {
      raw_vals <- display[[col]]
      numeric_vals <- suppressWarnings(as.numeric(raw_vals))
      formatted <- as.character(raw_vals)
      formatted[is.na(formatted)] <- "-"
      blank_rows <- !nzchar(display$Metric)
      formatted[blank_rows] <- ""

      percent_rows <- tolower(display$Metric) == "percent" & !is.na(numeric_vals)
      size_rows <- tolower(display$Metric) == "segment size" & !is.na(numeric_vals)
      numeric_rows <- !percent_rows & !size_rows & !is.na(numeric_vals)

      formatted[percent_rows] <- sprintf("%.1f%%", numeric_vals[percent_rows])
      formatted[size_rows] <- format(round(numeric_vals[size_rows]), big.mark = ",", trim = TRUE)
      formatted[numeric_rows] <- sprintf("%.2f", numeric_vals[numeric_rows])
      formatted[is.na(numeric_vals) & !blank_rows] <- "-"

      display[[col]] <- formatted
    }

    display
  }

  output$manual_metrics_table <- DT::renderDT({
    req(manual_state$run_name)
    data_tbl <- manual_state$metrics_table
    if (is.null(data_tbl) || !nrow(data_tbl)) {
      return(DT::datatable(
        data.frame(Message = "Metrics unavailable for this run."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    display <- format_manual_metric_table(data_tbl)

    segment_cols <- setdiff(names(display), "Metric")
    dt <- DT::datatable(
      display,
      options = list(dom = "t", paging = FALSE, ordering = FALSE),
      rownames = FALSE,
      escape = FALSE
    )

    colors <- segment_colors()
    cell_cols <- colors$cell %||% character(0)
    header_cols <- colors$header %||% character(0)
    applicable <- intersect(segment_cols, names(cell_cols))
    for (seg in applicable) {
      cell_col <- cell_cols[[seg]]
      dt <- DT::formatStyle(
        dt,
        columns = seg,
        backgroundColor = cell_col,
        color = segment_text_contrast(cell_col)
      )
    }

    header_applicable <- intersect(segment_cols, names(header_cols))
    col_defs <- build_segment_column_defs(names(display), header_applicable)
    if (length(col_defs)) {
      dt$x$options$columnDefs <- c(dt$x$options$columnDefs, col_defs)
    }

    dt
  })

  output$manual_variable_diff_table <- DT::renderDT({
    req(manual_state$run_name)
    diff_table <- manual_state$variable_diff_table
    if (is.null(diff_table) || !nrow(diff_table)) {
      return(DT::datatable(
        data.frame(Message = "Variable differentiation metrics unavailable."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    segment_cols <- setdiff(names(diff_table), "Metric")
    display <- diff_table
    for (col in segment_cols) {
      numeric_vals <- suppressWarnings(as.numeric(display[[col]]))
      formatted <- ifelse(is.na(numeric_vals), "-", sprintf("%.2f", numeric_vals))
      display[[col]] <- formatted
    }

    segment_cols <- setdiff(names(display), "Metric")
    dt <- DT::datatable(
      display,
      options = list(dom = "t", paging = FALSE, ordering = FALSE),
      rownames = FALSE,
      escape = FALSE
    )

    colors <- segment_colors()
    cell_cols <- colors$cell %||% character(0)
    header_cols <- colors$header %||% character(0)
    applicable <- intersect(segment_cols, names(cell_cols))
    for (seg in applicable) {
      cell_col <- cell_cols[[seg]]
      dt <- DT::formatStyle(
        dt,
        columns = seg,
        backgroundColor = cell_col,
        color = segment_text_contrast(cell_col)
      )
    }

    header_applicable <- intersect(segment_cols, names(header_cols))
    col_defs <- build_segment_column_defs(names(display), header_applicable)
    if (length(col_defs)) {
      dt$x$options$columnDefs <- c(dt$x$options$columnDefs, col_defs)
    }

    dt
  })

  output$manual_metric_category_table <- DT::renderDT({
    req(manual_state$run_name)

    selected_metrics <- input$manual_metric_category_select %||% character(0)
    if (!length(selected_metrics)) {
      return(DT::datatable(
        data.frame(Message = "Select one or more metrics to view category summaries."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    segments <- manual_state$segment_summary$Segment
    if (is.null(segments) || !length(segments)) {
      segments <- setdiff(names(manual_state$metrics_table), "Metric")
    }

    summary_tbl <- build_metric_category_summary(manual_state$variable_tables, selected_metrics, segments)
    if (is.null(summary_tbl) || !nrow(summary_tbl)) {
      return(DT::datatable(
        data.frame(Message = "Category summaries unavailable for the selected metrics."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    segment_cols <- setdiff(names(summary_tbl), c("Metric", "Category"))
    for (col in segment_cols) {
      numeric_vals <- suppressWarnings(as.numeric(summary_tbl[[col]]))
      if (all(is.na(numeric_vals))) {
        summary_tbl[[col]] <- "-"
        next
      }
      scaled_vals <- numeric_vals
      max_val <- suppressWarnings(max(numeric_vals, na.rm = TRUE))
      if (is.finite(max_val) && max_val <= 1) {
        scaled_vals <- numeric_vals * 100
      }
      summary_tbl[[col]] <- ifelse(is.na(scaled_vals), "-", sprintf("%.1f%%", scaled_vals))
    }

    dt <- DT::datatable(
      summary_tbl,
      options = list(dom = "t", paging = FALSE, ordering = FALSE),
      rownames = FALSE,
      escape = FALSE
    )

    colors <- segment_colors()
    cell_cols <- colors$cell %||% character(0)
    header_cols <- colors$header %||% character(0)
    applicable <- intersect(segment_cols, names(cell_cols))
    for (seg in applicable) {
      cell_col <- cell_cols[[seg]]
      dt <- DT::formatStyle(
        dt,
        columns = seg,
        backgroundColor = cell_col,
        color = segment_text_contrast(cell_col)
      )
    }

    header_applicable <- intersect(segment_cols, names(header_cols))
    col_defs <- build_segment_column_defs(names(summary_tbl), header_applicable)
    if (length(col_defs)) {
      dt$x$options$columnDefs <- c(dt$x$options$columnDefs, col_defs)
    }

    dt
  })

  output$manual_of_metric_table <- DT::renderDT({
    req(manual_state$run_name)

    data_tbl <- manual_state$of_metrics_table
    if (is.null(data_tbl) || !nrow(data_tbl)) {
      return(DT::datatable(
        data.frame(Message = "Objective function summaries unavailable for this run."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    available <- manual_state$available_of_metrics
    if (!length(available)) {
      available <- normalize_of_metric_names(data_tbl$Metric)
    }

    selected <- input$manual_of_table_metrics
    if (is.null(selected) || !length(selected)) {
      selected <- input$manual_of_metrics_selected
    }

    selected <- toupper(selected %||% character(0))
    selected <- intersect(selected, available)
    if (!length(selected)) {
      selected <- head(available, min(3, length(available)))
    }

    if (!length(selected)) {
      return(DT::datatable(
        data.frame(Message = "Select one or more metrics to view their summaries."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    display <- data_tbl
    display$MetricKey <- normalize_of_metric_names(display$Metric)
    order_idx <- match(selected, display$MetricKey)
    order_idx <- order_idx[!is.na(order_idx)]
    if (!length(order_idx)) {
      return(DT::datatable(
        data.frame(Message = "Selected metrics do not have summary tables."),
        options = list(dom = "t"),
        rownames = FALSE
      ))
    }

    display <- display[order_idx, , drop = FALSE]
    display$Metric <- tools::toTitleCase(tolower(gsub("_", " ", display$Metric)))
    display$MetricKey <- NULL

    segment_cols <- setdiff(names(display), "Metric")
    for (col in segment_cols) {
      numeric_vals <- suppressWarnings(as.numeric(display[[col]]))
      formatted <- ifelse(is.na(numeric_vals), "-", sprintf("%.2f", numeric_vals))
      display[[col]] <- formatted
    }

    dt <- DT::datatable(
      display,
      options = list(dom = "t", paging = FALSE, ordering = FALSE),
      rownames = FALSE,
      escape = FALSE
    )

    colors <- segment_colors()
    cell_cols <- colors$cell %||% character(0)
    header_cols <- colors$header %||% character(0)
    applicable <- intersect(segment_cols, names(cell_cols))
    for (seg in applicable) {
      cell_col <- cell_cols[[seg]]
      dt <- DT::formatStyle(
        dt,
        columns = seg,
        backgroundColor = cell_col,
        color = segment_text_contrast(cell_col)
      )
    }

    header_applicable <- intersect(segment_cols, names(header_cols))
    col_defs <- build_segment_column_defs(names(display), header_applicable)
    if (length(col_defs)) {
      dt$x$options$columnDefs <- c(dt$x$options$columnDefs, col_defs)
    }

    dt
  })

  observeEvent(
    list(input$manual_of_metrics_selected, manual_state$of_distributions, manual_state$run_name),
    {
      distributions <- manual_state$of_distributions

      if (length(manual_state$rendered_hist_ids)) {
        for (id in manual_state$rendered_hist_ids) {
          output[[id]] <- NULL
        }
      }

      if (is.null(distributions) || !length(distributions)) {
        manual_state$rendered_hist_ids <- character(0)
        output$manual_of_histograms <- renderUI({
          div(class = "text-muted", "Objective function distributions unavailable for this run.")
        })
        return()
      }

      available <- manual_state$available_of_metrics
      if (!length(available)) {
        available <- normalize_of_metric_names(names(distributions))
      }

      selected <- input$manual_of_metrics_selected %||% available
      selected <- intersect(toupper(selected), available)
      if (!length(selected)) {
        selected <- available
      }

      if (!length(selected)) {
        manual_state$rendered_hist_ids <- character(0)
        output$manual_of_histograms <- renderUI({
          div(class = "text-muted", "Select objective function metrics to view distributions.")
        })
        return()
      }

      palette <- segment_colors()
      header_cols <- palette$header %||% character(0)
      cell_cols <- palette$cell %||% character(0)

      hist_ids <- character(0)
      plot_blocks <- lapply(seq_along(selected), function(idx) {
        metric_key <- selected[[idx]]
        dist_tbl <- distributions[[metric_key]]
        metric_label <- gsub("_", " ", metric_key)

        if (is.null(dist_tbl) || !nrow(dist_tbl)) {
          return(NULL)
        }

        metric_values <- suppressWarnings(as.numeric(dist_tbl$Value))
        metric_values <- metric_values[is.finite(metric_values)]
        if (!length(metric_values)) {
          return(NULL)
        }

        range_min <- min(metric_values)
        range_max <- max(metric_values)
        if (!is.finite(range_min) || !is.finite(range_max)) {
          return(NULL)
        }

        stat_mean <- mean(metric_values)
        stat_median <- stats::median(metric_values)
        stat_sd <- if (length(metric_values) > 1) stats::sd(metric_values) else 0
        stat_n <- length(metric_values)

        if (identical(range_min, range_max)) {
          padding <- if (range_min == 0) 0.5 else abs(range_min) * 0.05
          if (!is.finite(padding) || padding <= 0) {
            padding <- 0.5
          }
          range_min <- range_min - padding
          range_max <- range_max + padding
        }

        bucket_count <- 20L
        bin_edges <- seq(range_min, range_max, length.out = bucket_count + 1)
        bin_size <- diff(bin_edges)[1]
        if (!is.finite(bin_size) || bin_size <= 0) {
          bin_size <- 1
        }
        bin_start <- bin_edges[1]
        bin_end <- bin_edges[length(bin_edges)]

        segments <- unique(dist_tbl$Segment)
        segments <- segments[!is.na(segments) & nzchar(segments)]
        if (!length(segments)) {
          return(NULL)
        }

        mini_plots <- lapply(seq_along(segments), function(seg_idx) {
          seg <- segments[[seg_idx]]
          seg_values <- dist_tbl %>%
            dplyr::filter(Segment == seg) %>%
            dplyr::pull(Value)
          cleaned <- suppressWarnings(as.numeric(seg_values))
          cleaned <- cleaned[!is.na(cleaned)]
          if (!length(cleaned)) {
            return(NULL)
          }

          seg_header <- header_cols[[seg]] %||% "#2563EB"
          seg_fill <- cell_cols[[seg]] %||% lighten_color(seg_header, 0.9)

          raw_id <- paste0("hist_", metric_key, "_", idx, "_", seg)
          plot_id <- sanitize_manual_id(raw_id)
          if (plot_id %in% hist_ids) {
            plot_id <- paste0(plot_id, "_", seg_idx)
          }
          hist_ids <<- c(hist_ids, plot_id)

          local({
            fill_col <- seg_fill
            border_col <- seg_header
            data_vals <- cleaned
            bin_start_local <- bin_start
            bin_end_local <- bin_end
            bin_size_local <- bin_size
            output[[plot_id]] <- plotly::renderPlotly({
              plotly::plot_ly(
                x = data_vals,
                type = "histogram",
                autobinx = FALSE,
                xbins = list(
                  start = bin_start_local,
                  end = bin_end_local,
                  size = bin_size_local
                ),
                marker = list(
                  color = fill_col,
                  line = list(color = border_col, width = 1.1)
                ),
                opacity = 0.95,
                showlegend = FALSE
              ) %>%
                plotly::layout(
                  bargap = 0.08,
                  margin = list(l = 10, r = 4, t = 6, b = 26),
                  paper_bgcolor = "rgba(0,0,0,0)",
                  plot_bgcolor = "rgba(0,0,0,0)",
                  xaxis = list(
                    title = "",
                    zeroline = FALSE,
                    showgrid = FALSE,
                    range = c(bin_start_local, bin_end_local)
                  ),
                  yaxis = list(title = "", zeroline = FALSE, showgrid = FALSE)
                )
            })
          })

          tags$div(
            class = "of-mini-hist",
            tags$div(class = "of-mini-title", seg),
            plotly::plotlyOutput(plot_id, height = "120px")
          )
        })

        mini_plots <- purrr::compact(mini_plots)
        if (!length(mini_plots)) {
          return(NULL)
        }

        summary_row <- tags$div(
          class = "of-metric-summary",
          tags$span(class = "of-stat-chip", sprintf("Observations: %d", stat_n)),
          tags$span(class = "of-stat-chip", sprintf("Mean: %.2f", stat_mean)),
          tags$span(class = "of-stat-chip", sprintf("Median: %.2f", stat_median)),
          tags$span(class = "of-stat-chip", sprintf("Std Dev: %.2f", stat_sd))
        )

        tags$div(
          class = "of-metric-panel",
          tags$h5(class = "of-metric-title", metric_label),
          summary_row,
          tags$div(class = "of-hist-grid", mini_plots)
        )
      })

      plot_blocks <- purrr::compact(plot_blocks)
      manual_state$rendered_hist_ids <- unique(hist_ids)
      output$manual_of_histograms <- renderUI({
        if (!length(plot_blocks)) {
          div(class = "text-muted", "Select objective function metrics to view distributions.")
        } else {
          tagList(plot_blocks)
        }
      })
    },
    ignoreNULL = FALSE
  )
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
