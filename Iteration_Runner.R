#ITERATION RUNNER - KEEP TAB OF YOUR ITERATIONS

# Load necessary libraries
library(readxl)
library(dplyr)
library(tidyr)


Iteration_running_function <- function(data_path, variables, tracker_path, output_file,seg_n,op_dic,solo) {
  

  print(paste0("SOLO VALUE -2: ",solo))
  variables_chosen<-as.numeric(unlist(strsplit(variables, ",")))
  themes_sheet <- "THEMES"
  input_sheet <- "INPUT"
  scores_sheet <- "VARIABLES"
  na_values <- c("NA", ".", "", " ")
  XTab_Data <- read_excel(data_path, sheet = input_sheet, na = na_values)
  variable_scores_data <- read_excel(data_path, sheet = scores_sheet, range = "B2:C999")
  variable_scores <- variable_scores_data
  colnames(variable_scores) <- c("Variable", "Score")
  variable_scores <- variable_scores[complete.cases(variable_scores), ]
  variables_chosen_name<-colnames(XTab_Data[,variables_chosen])
  variables_chosen_name<-as.data.frame(variables_chosen_name)
  colnames(variables_chosen_name)<-"Variable"
  variables_chosen_name <- variables_chosen_name %>% left_join(variable_scores, by = "Variable")
   
        
        
   tracker<-read.csv(tracker_path)

   colnames(tracker)
    -
    if(max(tracker$Iteration)==-Inf){itr_n=1} else{itr_n<-max(tracker$Iteration)+1}

    n_size<-nrow(variables_chosen_name)
    combination_data<-paste(variables_chosen_name$Variable, collapse = ";")
    total_score <- sum(variables_chosen_name$Score, na.rm = TRUE)
   avg_score <- total_score / nrow(variables_chosen_name)
   Var_cols<- variables

   sort_of_columns <- function(cols) {
     if (!length(cols)) return(cols)

     base_keys <- gsub("^(max|min)_", "", cols)
     base_keys <- sub("_(values|val|diff_max|diff_min|diff)$", "", base_keys, perl = TRUE)
     base_keys <- toupper(base_keys)

     metric_type <- ifelse(grepl("^max_", cols), "legacy_max",
                     ifelse(grepl("^min_", cols), "legacy_min",
                            sub("^OF_[A-Za-z0-9_]+_", "", cols)))

     metric_order <- c(
       values = 1L,
       val = 2L,
       diff_max = 3L,
       diff = 4L,
       diff_min = 5L,
       legacy_max = 6L,
       legacy_min = 7L
     )

     metric_type[metric_type == ""] <- "values"
     metric_type[!(metric_type %in% names(metric_order))] <- "zzz"
     metric_order <- c(metric_order, zzz = 99L)

     order_idx <- metric_order[metric_type]
     cols[order(base_keys, order_idx, cols)]
   }

   base_add_cols<-c("Ran_Iteration_Flag", "ITR_PATH", "ITR_FILE_NAME", "MAX_N_SIZE", "MAX_N_SIZE_PERC", "MIN_N_SIZE", "MIN_N_SIZE_PERC", "SOLUTION_N_SIZE", "PROB_95", "PROB_90", "PROB_80", "PROB_75", "PROB_LESS_THAN_50", "BIMODAL_VARS", "PROPER_BUCKETED_VARS", "BIMODAL_VARS_PERC","ROV_SD", "ROV_RANGE", "seg1_diff", "seg2_diff", "seg3_diff", "seg4_diff", "seg5_diff", "bi_1", "perf_1", "indT_1", "indB_1", "bi_2", "perf_2", "indT_2", "indB_2", "bi_3", "perf_3", "indT_3", "indB_3", "bi_4", "perf_4", "indT_4", "indB_4", "bi_5", "perf_5", "indT_5", "indB_5")
   input_perf_cols<-c("BIMODAL_VARS_INPUT_VARS","INPUT_PERF_SEG_COVERED","INPUT_PERF_FLAG","INPUT_PERF_MINSEG","INPUT_PERF_MAXSEG")

   of_pattern <- "^(max|min)_OF_[A-Za-z0-9_]+_diff$|^OF_[A-Za-z0-9_]+_(val|values|diff|diff_min|diff_max)$"
   existing_of_cols<-names(tracker)[grepl(of_pattern, names(tracker))]
   add_cols<-c(base_add_cols, sort_of_columns(unique(existing_of_cols)), input_perf_cols)
   add_cols<-unique(add_cols)

   iteration_data <- c(itr_n, n_size, combination_data, total_score, avg_score, Var_cols,0,rep(NA,length(add_cols)-1))
   comb_data<-list()
   comb_data[1]<- dirname(tracker_path)
   comb_data[2]<- basename(tracker_path) 
   input_working_dir<- dirname(data_path)  
   input_db_name<- basename(data_path) 
   comb_data$V2
   setwd(as.character(comb_data[1]))
      
   write.table(t(iteration_data), file = as.character(comb_data[2]), append = TRUE, sep = ",", col.names = FALSE, row.names = FALSE)
    print(paste0("SOLO VALUE -3: ",solo))
   op<-run_iterations(comb_data,input_working_dir,op_dic,input_db_name,seg_n,output_file,solo)
   return(op)

}
