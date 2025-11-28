###################################################
###################################################
#-PROGRAM-INFO:

###-PROGRAM:      LCA & SEG Summary Program
###-VERSION:      03.26.15
###-AUTHOR:       Kevin Shaney
##################################################
##################################################
###UPDATED BY:    Gurujeet Shetty
###-VERSION:      06.25.2024
##################################################
##################################################
###-PACKAGE-DEPENDENCIES(2) (E.G. INSTALL PRIOR TO RUNNING):
# install.packages("poLCA")
# install.packages("weights")
# install.packages("ggplot2", dep = TRUE)
# install.packages("colorspace")
# install.packages("openxlsx")
# install.packages("zip")
###################################################
# Load necessary libraries
library(openxlsx)
library(poLCA)
library(weights)
library(ggplot2)
library(readxl)

if (file.exists("input_utils.R")) {
  source("input_utils.R")
} else {
  stop("input_utils.R not found. Please place it alongside Function_poLCA.R.")
}

#################################################


# Define the function
LCA_SEG_Summary <- function(input_working_dir, output_working_dir, input_db_name, file_run_and_date, 
                            cat_seg_var_list, cont_seg_var_list, set_num_segments=4, set_maxiter, set_nrep, 
                            option_summary_only) {
 
#Setting the seed.
set.seed(25)
themes_sheet <- "THEMES"
input_sheet <- "INPUT"
scores_sheet <- "VARIABLES"


remove_neutrals <- c(); #Remove neutral (middle coded) values from these variables in clustering
time_start <- Sys.time()

library(poLCA)
library(weights)
setwd(input_working_dir)
na_values <- c("NA", ".", "", " ")
input_path <- file.path(input_working_dir, input_db_name)
standardized_input <- read_standardized_input(input_path, sheet = input_sheet, na_values = na_values)
input_metadata <- standardized_input$metadata
XTab_Data <- as.data.frame(standardized_input$data, stringsAsFactors = FALSE)

#If variable names were provided instead of numbers, convert to numbers
if (is.character(cat_seg_var_list)) cat_seg_var_list <- match(cat_seg_var_list, names(XTab_Data))
if (is.character(cont_seg_var_list)) cont_seg_var_list <- match(cont_seg_var_list, names(XTab_Data))
cat_seg_var_list <- cat_seg_var_list[!is.na(cat_seg_var_list)]
cont_seg_var_list <- cont_seg_var_list[!is.na(cont_seg_var_list)]

##############################################

xtab_var_col <- c(2)
xtab_var_rows <- c(4:ncol(XTab_Data))

##############################################
input_seg_var_list <- c(cat_seg_var_list,cont_seg_var_list)
XTab_Data_Orig_Ord <- XTab_Data
XTab_Data_Labels <- rbind(input_metadata$question_prefix, input_metadata$answer_description)
colnames(XTab_Data_Labels) <- names(XTab_Data)

##############################################
###-FORMAT-SEG-INPUT-DB
print("DD222-2") 
if(option_summary_only != 1){
  
  Seg_Data_Formula <- paste0("as.data.frame(cbind(",paste0("XTab_Data$",names(XTab_Data_Orig_Ord)[input_seg_var_list],collapse=","),"))",collapse="");
  print("DD222-2-1")
  Seg_Data <- eval(parse(text=Seg_Data_Formula))
  names(Seg_Data) <- names(XTab_Data_Orig_Ord[,input_seg_var_list])
  print("DD222-2-2")
  for(col in seq(1,ncol(Seg_Data),1)){  
    Seg_Data[,col] <- as.vector(Seg_Data[,col])
    set_uniq <- unique(Seg_Data[,col])
    set_uniq <- sort(set_uniq[!is.na(set_uniq)])
    if( length(set_uniq) < 10 ){
      for(integ_val in seq(1,length(set_uniq),1)){      
        Seg_Data[(Seg_Data[,col] == set_uniq[integ_val] & !is.na(Seg_Data[,col])),col] <- integ_val
      }
    }
  }
  print("DD222-2-3")
  Seg_Data <- data.frame(lapply(Seg_Data,as.numeric))
  Seg_Data[remove_neutrals] <- lapply(Seg_Data[remove_neutrals],function(x) {(x[which(x == 2)] <- NA); return(x);})
  
  ##############################################
  print("DD222-2-4")
  cont_names = c()
  if(length(cont_seg_var_list) == 0) { #Why do we ever convert this to numbers anyway? TODO: Fix
    cont_names <- c(1);
  }else{
    cont_names <- names(XTab_Data_Orig_Ord[cont_seg_var_list])
  }

  SegFn <- as.formula(paste0("cbind(",paste0(names(XTab_Data_Orig_Ord[cat_seg_var_list]),collapse=","),")~",paste0(cont_names,collapse="+")));
  
  SegProc <- poLCA(SegFn, Seg_Data, nclass=as.numeric(set_num_segments) , maxiter=as.numeric(set_maxiter), graphs=FALSE, tol=1e-10, na.rm=FALSE, probs.start=NULL, nrep=as.numeric(set_nrep), verbose=TRUE, calc.se=TRUE)
  
  if(option_summary_only != 1){XTab_Data[,2] <- SegProc$predclass}
  
  names(XTab_Data)[2] <- "SEGM_poLCA"
  XTab_Data <- cbind( data.frame(XTab_Data[,c(1:3)]) , data.frame(lapply(XTab_Data[,c(4:ncol(XTab_Data))],factor)) )
  
}  ###-END-IF-SUMMARY-ONLY-CONTROL-###
print("DD222-3") 
##############################################
##############################################

uniq_segs <- unique(XTab_Data[,xtab_var_col])
uniq_segs <- uniq_segs[order(as.numeric(uniq_segs))]

SEG_N_WTD <- vector()
SEG_PCT_WTD <- vector()
SEG_N_UNWTD <- vector()
SEG_PCT_UNWTD <- vector()

i=1
for(seg in uniq_segs){
  SEG_N_WTD[i] <- sum(as.numeric(XTab_Data[XTab_Data[,xtab_var_col]==seg,3]))
  SEG_N_UNWTD[i] <- length( which( XTab_Data[,xtab_var_col]==seg )) 
  i=i+1                           
}
print("DD222-4")
objects <- ls()
sizes <- sapply(objects, function(x) object.size(get(x)))
large_objects <- objects[sizes > 1e+09]  # Adjust the threshold as needed
print(large_objects)
print(length(SEG_N_UNWTD)) 
SEG_PCT_WTD <- SEG_N_WTD / sum(SEG_N_WTD)
SEG_PCT_UNWTD <- SEG_N_UNWTD / sum(SEG_N_UNWTD)

SEG_N_WTD[length(SEG_N_WTD)+1] <- sum(SEG_N_WTD) 
SEG_N_UNWTD[length(SEG_N_UNWTD)+1] <- sum(SEG_N_UNWTD)
SEG_PCT_WTD[length(SEG_PCT_WTD)+1] <- sum(SEG_PCT_WTD) 
SEG_PCT_UNWTD[length(SEG_PCT_UNWTD)+1] <- sum(SEG_PCT_UNWTD)
uniq_segs[length(uniq_segs)+1] <- "Total"

empty_row <- rep(NA,length(SEG_N_UNWTD))

##############################################
###-APPEND-POSTERIOR-PROBABILITIES
print("DD222-5") 
if(option_summary_only != 1){
  
  uniq_segs <- uniq_segs[c(1:(length(uniq_segs)-1))]
  
  postprob_table <- cbind(SEGM_poLCA=SegProc$predclass,SegProc$posterior)
  
  for(seq_int in seq(1,length(unique(XTab_Data[,xtab_var_col])),1)){
    colnames(postprob_table)[seq_int+1] <- seq_int
  }
  
  SEG_POST_PROB_MEAN <- vector()
  SEG_POST_PROB_PCT_VEC <- vector()
  SEG_POST_PROB_PCT_MAT <- matrix(NA,11,length(uniq_segs)+1)
  i=1
  for(seg in uniq_segs){
    SEG_POST_PROB_MEAN[i] <- mean( postprob_table[ postprob_table[,1] == seg , as.numeric(seg)+1 ] )
    SEG_POST_PROB_PCT <- quantile(postprob_table[ postprob_table[,1] == seg , as.numeric(seg)+1 ] , c(seq(0.0,1.0,0.1)))
    SEG_POST_PROB_PCT_MAT[,i] <- t(SEG_POST_PROB_PCT) 
    i=i+1
  }
  SEG_POST_PROB_MEAN[length(uniq_segs)+1] <- crossprod(SEG_POST_PROB_MEAN,SEG_PCT_UNWTD[-length(SEG_PCT_UNWTD)])  
  rownames(SEG_POST_PROB_PCT_MAT) <- names(SEG_POST_PROB_PCT)
  
  for(row in seq(1,nrow(SEG_POST_PROB_PCT_MAT),1)){
    SEG_POST_PROB_PCT_MAT[row,ncol(SEG_POST_PROB_PCT_MAT)] <- crossprod(SEG_POST_PROB_PCT_MAT[c(row),-ncol(SEG_POST_PROB_PCT_MAT)],SEG_PCT_UNWTD[-length(SEG_PCT_WTD)])
  }
  
  xtab_output <- rbind(empty_row,empty_row,SEG_N_UNWTD,SEG_PCT_UNWTD,empty_row,SEG_N_WTD,SEG_PCT_WTD,empty_row,empty_row,empty_row,empty_row,empty_row,SEG_POST_PROB_MEAN,empty_row,SEG_POST_PROB_PCT_MAT)
  
  uniq_segs[length(uniq_segs)+1] <- "Total"
  
}else{xtab_output <- rbind(empty_row,empty_row,SEG_N_UNWTD,SEG_PCT_UNWTD,empty_row,SEG_N_WTD,SEG_PCT_WTD,empty_row)} ###-END-IF-###

  colnames(xtab_output) <- uniq_segs

##############################################
print("DD222-6") 
for (var in xtab_var_rows){
  
  uniq_set <- unique(XTab_Data[,var])
  uniq_set <- uniq_set[!is.na(uniq_set)]
  uniq_set <- uniq_set[order(uniq_set)]
  if(length(uniq_set)==0) next;
  xtab_table_iter <- matrix(NA,length(uniq_set),length(uniq_segs))
  
  i=1
  for(col in uniq_segs[-length(uniq_segs)]){
    for(row in seq(1,length(uniq_set),1)){
      xtab_table_iter[row,i] <- sum( as.numeric( XTab_Data[ XTab_Data[,var] == uniq_set[row] & XTab_Data[,xtab_var_col] == col & !is.na(XTab_Data[,var]), 3 ] ) )     
    }
    i=i+1
  }
  
  for(row in seq(1,length(uniq_set),1)){
    xtab_table_iter[row,ncol(xtab_table_iter)] <- sum(xtab_table_iter[row,-ncol(xtab_table_iter)])

##############################################
###-FORMAT-SEG-INPUT-DB
print("DD222-2") 
if(option_summary_only != 1){
  
  Seg_Data_Formula <- paste0("as.data.frame(cbind(",paste0("XTab_Data$",names(XTab_Data_Orig_Ord)[input_seg_var_list],collapse=","),"))",collapse="");
  print("DD222-2-1")
  Seg_Data <- eval(parse(text=Seg_Data_Formula))
  names(Seg_Data) <- names(XTab_Data_Orig_Ord[,input_seg_var_list])
  print("DD222-2-2")
  for(col in seq(1,ncol(Seg_Data),1)){  
    Seg_Data[,col] <- as.vector(Seg_Data[,col])
    set_uniq <- unique(Seg_Data[,col])
    set_uniq <- sort(set_uniq[!is.na(set_uniq)])
    if( length(set_uniq) < 10 ){
      for(integ_val in seq(1,length(set_uniq),1)){      
        Seg_Data[(Seg_Data[,col] == set_uniq[integ_val] & !is.na(Seg_Data[,col])),col] <- integ_val
      }
    }
  }
  print("DD222-2-3")
  Seg_Data <- data.frame(lapply(Seg_Data,as.numeric))
  Seg_Data[remove_neutrals] <- lapply(Seg_Data[remove_neutrals],function(x) {(x[which(x == 2)] <- NA); return(x);})
  
  ##############################################
  print("DD222-2-4")
  cont_names = c()
  if(length(cont_seg_var_list) == 0) { #Why do we ever convert this to numbers anyway? TODO: Fix
    cont_names <- c(1);
  }else{
    cont_names <- names(XTab_Data_Orig_Ord[cont_seg_var_list])
  }

  SegFn <- as.formula(paste0("cbind(",paste0(names(XTab_Data_Orig_Ord[cat_seg_var_list]),collapse=","),")~",paste0(cont_names,collapse="+")));
  
  SegProc <- poLCA(SegFn, Seg_Data, nclass=as.numeric(set_num_segments) , maxiter=as.numeric(set_maxiter), graphs=FALSE, tol=1e-10, na.rm=FALSE, probs.start=NULL, nrep=as.numeric(set_nrep), verbose=TRUE, calc.se=TRUE)
  
  if(option_summary_only != 1){XTab_Data[,2] <- SegProc$predclass}
  
  names(XTab_Data)[2] <- "SEGM_poLCA"
  XTab_Data <- cbind( data.frame(XTab_Data[,c(1:3)]) , data.frame(lapply(XTab_Data[,c(4:ncol(XTab_Data))],factor)) )
  
}  ###-END-IF-SUMMARY-ONLY-CONTROL-###
print("DD222-3") 
##############################################
##############################################

uniq_segs <- unique(XTab_Data[,xtab_var_col])
uniq_segs <- uniq_segs[order(as.numeric(uniq_segs))]

SEG_N_WTD <- vector()
SEG_PCT_WTD <- vector()
SEG_N_UNWTD <- vector()
SEG_PCT_UNWTD <- vector()

i=1
for(seg in uniq_segs){
  SEG_N_WTD[i] <- sum(as.numeric(XTab_Data[XTab_Data[,xtab_var_col]==seg,3]))
  SEG_N_UNWTD[i] <- length( which( XTab_Data[,xtab_var_col]==seg )) 
  i=i+1                           
}
print("DD222-4")
objects <- ls()
sizes <- sapply(objects, function(x) object.size(get(x)))
large_objects <- objects[sizes > 1e+09]  # Adjust the threshold as needed
print(large_objects)
print(length(SEG_N_UNWTD)) 
SEG_PCT_WTD <- SEG_N_WTD / sum(SEG_N_WTD)
SEG_PCT_UNWTD <- SEG_N_UNWTD / sum(SEG_N_UNWTD)

SEG_N_WTD[length(SEG_N_WTD)+1] <- sum(SEG_N_WTD) 
SEG_N_UNWTD[length(SEG_N_UNWTD)+1] <- sum(SEG_N_UNWTD)
SEG_PCT_WTD[length(SEG_PCT_WTD)+1] <- sum(SEG_PCT_WTD) 
SEG_PCT_UNWTD[length(SEG_PCT_UNWTD)+1] <- sum(SEG_PCT_UNWTD)
uniq_segs[length(uniq_segs)+1] <- "Total"

empty_row <- rep(NA,length(SEG_N_UNWTD))

##############################################
###-APPEND-POSTERIOR-PROBABILITIES
print("DD222-5") 
if(option_summary_only != 1){
  
  uniq_segs <- uniq_segs[c(1:(length(uniq_segs)-1))]
  
  postprob_table <- cbind(SEGM_poLCA=SegProc$predclass,SegProc$posterior)
  
  for(seq_int in seq(1,length(unique(XTab_Data[,xtab_var_col])),1)){
    colnames(postprob_table)[seq_int+1] <- seq_int
  }
  
  SEG_POST_PROB_MEAN <- vector()
  SEG_POST_PROB_PCT_VEC <- vector()
  SEG_POST_PROB_PCT_MAT <- matrix(NA,11,length(uniq_segs)+1)
  i=1
  for(seg in uniq_segs){
    SEG_POST_PROB_MEAN[i] <- mean( postprob_table[ postprob_table[,1] == seg , as.numeric(seg)+1 ] )
    SEG_POST_PROB_PCT <- quantile(postprob_table[ postprob_table[,1] == seg , as.numeric(seg)+1 ] , c(seq(0.0,1.0,0.1)))
    SEG_POST_PROB_PCT_MAT[,i] <- t(SEG_POST_PROB_PCT) 
    i=i+1
  }
  SEG_POST_PROB_MEAN[length(uniq_segs)+1] <- crossprod(SEG_POST_PROB_MEAN,SEG_PCT_UNWTD[-length(SEG_PCT_UNWTD)])  
  rownames(SEG_POST_PROB_PCT_MAT) <- names(SEG_POST_PROB_PCT)
  
  for(row in seq(1,nrow(SEG_POST_PROB_PCT_MAT),1)){
    SEG_POST_PROB_PCT_MAT[row,ncol(SEG_POST_PROB_PCT_MAT)] <- crossprod(SEG_POST_PROB_PCT_MAT[c(row),-ncol(SEG_POST_PROB_PCT_MAT)],SEG_PCT_UNWTD[-length(SEG_PCT_WTD)])
  }
  
  xtab_output <- rbind(empty_row,empty_row,SEG_N_UNWTD,SEG_PCT_UNWTD,empty_row,SEG_N_WTD,SEG_PCT_WTD,empty_row,empty_row,empty_row,empty_row,empty_row,SEG_POST_PROB_MEAN,empty_row,SEG_POST_PROB_PCT_MAT)
  
  uniq_segs[length(uniq_segs)+1] <- "Total"
  
}else{xtab_output <- rbind(empty_row,empty_row,SEG_N_UNWTD,SEG_PCT_UNWTD,empty_row,SEG_N_WTD,SEG_PCT_WTD,empty_row)} ###-END-IF-###

  colnames(xtab_output) <- uniq_segs

##############################################
print("DD222-6")

# Helper: build a clean xtab block for each variable (percent share per segment + metadata rows)
build_xtab_block <- function(var_idx, uniq_segs, seg_weights) {
  var_values <- XTab_Data[, var_idx]
  uniq_set <- sort(unique(var_values[!is.na(var_values)]))
  if (!length(uniq_set)) return(NULL)

  base_counts <- sapply(uniq_segs[-length(uniq_segs)], function(seg) {
    vapply(uniq_set, function(val) {
      sum(as.numeric(XTab_Data[XTab_Data[, var_idx] == val &
                                XTab_Data[, xtab_var_col] == seg &
                                !is.na(XTab_Data[, var_idx]), 3]), na.rm = TRUE)
    }, numeric(1))
  })

  if (is.null(dim(base_counts))) {
    base_counts <- matrix(base_counts, ncol = 1)
  }
  base_counts <- as.matrix(base_counts)
  base_counts <- cbind(base_counts, rowSums(base_counts))
  colnames(base_counts) <- uniq_segs
  rownames(base_counts) <- uniq_set

  col_totals <- colSums(base_counts, na.rm = TRUE)
  col_totals[col_totals == 0] <- NA
  xtab_pct <- sweep(base_counts, 2, col_totals, FUN = "/")
  xtab_pct[!is.finite(xtab_pct)] <- 0

  vname <- as.character(XTab_Data_Labels[1, var_idx])
  dname <- as.character(XTab_Data_Labels[2, var_idx])
  var_id <- names(XTab_Data)[var_idx]
  if (var_idx %in% input_seg_var_list) {
    var_id <- paste0("00_SEGM_", var_id)
  }
  var_description <- ifelse(nzchar(dname), dname, vname)

  chisq_stat <- wtd.chi.sq(XTab_Data[, var_idx], XTab_Data[, xtab_var_col],
                           weight = as.numeric(XTab_Data[, 3]), na.rm = TRUE)
  chisq_df <- (length(unique(XTab_Data[, var_idx])) - 1) *
    (length(unique(XTab_Data[, xtab_var_col])) - 1)
  chisq_val <- if (length(chisq_stat)) chisq_stat[[1]] else NA_real_
  rov_stat <- if (is.finite(chisq_val) && chisq_df > 0) {
    (chisq_val - chisq_df) / sqrt(2 * chisq_df)
  } else { NA_real_ }

  header_rows <- matrix("", nrow = 5, ncol = ncol(xtab_pct))
  header_rows[2, 1] <- var_id
  header_rows[3, 1] <- var_description
  header_rows[4, 1] <- vname
  header_rows[5, 1] <- paste0("ROV=", round(rov_stat, 3))

  missing_counts <- vapply(uniq_segs[-length(uniq_segs)], function(seg) {
    sum(as.numeric(XTab_Data[XTab_Data[, xtab_var_col] == seg &
                              is.na(var_values), 3]), na.rm = TRUE)
  }, numeric(1))
  total_missing <- sum(missing_counts)
  denom <- seg_weights
  denom[denom == 0] <- NA
  missing_pct <- c(missing_counts / denom[-length(denom)], total_missing / denom[length(denom)])

  block <- rbind(header_rows, xtab_pct, missing_pct)
  rownames(block)[nrow(block)] <- "%_MSG"
  colnames(block) <- uniq_segs
  block
}

for (var in xtab_var_rows) {
  xtab_block <- build_xtab_block(var, uniq_segs, SEG_N_WTD)
  if (is.null(xtab_block)) next
  xtab_output <- rbind(as.matrix(xtab_output), as.matrix(xtab_block))
}

print("DD222-7")
##############################################

for(col in seq(1,ncol(xtab_output)-1,1)){
  colnames(xtab_output)[col] <- paste0("SG-",colnames(xtab_output)[col])
}   ###-END-FOR-###

##############################################

xtab_output_final <- xtab_output
empty_col <- rep(NA,nrow(xtab_output))
xtab_output_final <- cbind(xtab_output_final,empty_col)

##############################################

for(col in seq(1,ncol(xtab_output)-1,1)){
  xtab_output_index <-  100 * xtab_output[,col] / xtab_output[,ncol(xtab_output)]
  xtab_output_final <- cbind(xtab_output_final,xtab_output_index)
  colnames(xtab_output_final)[ncol(xtab_output_final)] <- paste0("INDEX_",colnames(xtab_output)[col])
}   ###-END-FOR-###

##############################################
print("DD222-8") 
xtab_output_final[is.na(xtab_output_final)] <- c("")

row.names(xtab_output_final)[rownames(xtab_output_final)=='empty_row'] <- c("")
colnames(xtab_output_final)[colnames(xtab_output_final)=='empty_col'] <- c("")



if(option_summary_only != 1){

  SegProc_OUTPUT <- capture.output(SegProc)  
  #print(xtab_output_final)
 #####EXCEL OUTPUT########
  # Create an Excel workbook and add the CSV files as sheets
  wb <- createWorkbook()
  # Add F1_SEG_IDs to Excel
  addWorksheet(wb, "F1_SEG_IDs")
  writeData(wb, "F1_SEG_IDs", data.frame(XTab_Data[1],poLCA_SEG=SegProc$predclass))
  # Add F2_SEG_xTABS to Excel
  addWorksheet(wb, "F2_SEG_xTABS")
  writeData(wb, "F2_SEG_xTABS", cbind(rownames(xtab_output_final),xtab_output_final))  
  # Add F3_POST_PROBs to Excel
  addWorksheet(wb, "F3_POST_PROBs")
  writeData(wb, "F3_POST_PROBs", cbind(RESPID=XTab_Data[1],SegProc$posterior,SEGM_poLCA=SegProc$predclass))    
  # Add F4_FULL_OUTPUT to Excel
  addWorksheet(wb, "F4_FULL_OUTPUT")
  writeData(wb, "F4_FULL_OUTPUT", SegProc_OUTPUT) 
  
  # Save the Excel file
  print("SAVING ITR")
  setwd(output_working_dir)
  saveWorkbook(wb, paste0(file_run_and_date, "_SUMMARY.xlsx"), overwrite = TRUE)
  output<-list()
  output$directory<- output_working_dir
  output$file_name<-file_run_and_date
  print(output_working_dir)
  print(file_run_and_date)
  return(output)   
  
}  ###-END-IF-SUMMARY-ONLY-CONTROL-###

##################################################
###-END-PROGRAM-##################################
##################################################
##################################################
##################################################
##################################################

for(i in seq(1,1,1)){
  cat("---------------------------------------------------------------------","\n")
  cat("---------------------------------------------------------------------","\n")
  cat("                                                                     ","\n")
  cat("TOTAL-RUN-TIME:",round(difftime(Sys.time(),time_start,units="mins"),digits=2), "MINS" ,"\n",sep=" ")
  cat("                                                                     ","\n")
  cat("---------------------------------------------------------------------","\n")
  cat("---------------------------------------------------------------------","\n")
}
}

# Example usage:
#LCA_SEG_Summary(
#  input_working_dir = "C:/Users/gs36325/Documents/00 My Learnings/09 Iterator",
#  output_working_dir = "C:/Users/gs36325/Documents/00 My Learnings/09 Iterator",
#  input_db_name = "poLCA_CF_input_V2.xlsx",
#  file_run_and_date = "CF_TRIAL_V3",
#  cat_seg_var_list = c(7,8,9,10,33,34,25,39,19),
#  cont_seg_var_list = c(),
#  set_num_segments = 4,
#  set_maxiter = 1000,
#  set_nrep = 10,
#  option_summary_only = 0
#)
