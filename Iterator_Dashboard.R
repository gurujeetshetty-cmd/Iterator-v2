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
          box(title = "Iteration Runner", status = "primary", solidHeader = TRUE, width = 12,
            textInput("input_file_path", "Input File Path",
                      value = "C:/Users/GS36325/Documents/14 TRUQAP mBC HCP Segmentation/09 Iterator_CLEAN/poLCA_input.xlsx"),
            textInput("tracker_file_path", "Tracker File Path",
                      value = "C:/Users/GS36325/Documents/14 TRUQAP mBC HCP Segmentation/09 Iterator_CLEAN/combinations_MANUAL_2.csv"),
            actionButton("load_files", "Load Files", class = "btn btn-primary"),
            textOutput("files_loaded")
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
            "Initializing...")
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
            "Initializing output generator...")
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
  observeEvent(input$load_files, {
    input_file_path <- input$input_file_path
    tracker_file_path <- input$tracker_file_path
    if (file.exists(input_file_path) && file.exists(tracker_file_path)) {
      append_log("Files loaded successfully for iteration run.")
      output$files_loaded <- renderText("Files have been loaded successfully!!")
    } else {
      append_log("Error loading files. Please check paths.")
      output$files_loaded <- renderText("Error loading files. Please check the file paths.")
    }
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
