library(readxl)
library(DBI)
library(odbc)
library(dplyr)
library(lubridate)
library(tidyr)
library(openxlsx)
library(writexl)
library(stringr)

#GANTI VALUATION!!!

val_year <- 2024
val_quarter <- 4

#GANTI COHORT!!!

cohort_start <- 2009
cohort_end <- 2024

jumlah_val <- (val_year-1-2009)*4+1+val_quarter

sheet_data <- read_excel("Prior_Sheet.v1.xlsx")

PC_prior_function <- function(cohort, term, sheet_name){
  
  
  ifelse(term == 1, st_or_lt <- "-ST", st_or_lt <- "-LT")
  
  if(sheet_name == "Marine"){
    
    '    #MENGAMBIL DF per Cohort (Partial)
    path_name <- paste0("Workbook/PC ",cohort,st_or_lt,"/PaidClaim-Net.Raw.xlsx")
    data_par_ME <- read_excel(path_name, sheet = "Marine-ECommerce")
    data_par_ME <- as.data.frame(lapply(data_par_ME, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_MO <- read_excel(path_name, sheet = "Marine-Other")
    data_par_MO <- as.data.frame(lapply(data_par_MO, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_MOX <- read_excel(path_name, sheet = "Marine-Other-XL")
    data_par_MOX <- as.data.frame(lapply(data_par_MOX, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_MEX <- read_excel(path_name, sheet = "Marine-ECommerce-XL")
    data_par_MEX <- as.data.frame(lapply(data_par_MEX, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    
    data_par <- data_par_ME
    col_count <- ncol(data_par)
    data_par[,3:col_count] <- NA
    data_par[,3:col_count] <- data_par_ME[,3:col_count]+data_par_MO[,3:col_count]+data_par_MOX[,3:col_count]+data_par_MEX[,3:col_count]
    '
    #MENGAMBIL DF Compiled
    data_comp_M <- read_excel("Workbook/PaidClaim-Net.Prior.wo.ML.xlsx", sheet = "Marine")
    data_comp <- data_comp_M
    
    #MENGAMBIL DF Total (semua cohort)
    path_tot <- "Workbook/PaidClaim-NetTotal.Raw.FIXED-LBLW.v1.xlsx"
    data_tot_ME <- read_excel(path_tot, sheet = "Marine-ECommerce")
    data_tot_ME <- as.data.frame(lapply(data_tot_ME, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_MO <- read_excel(path_tot, sheet = "Marine-Other")
    data_tot_MO <- as.data.frame(lapply(data_tot_MO, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_MOX <- read_excel(path_tot, sheet = "Marine-Other-XL")
    data_tot_MOX <- as.data.frame(lapply(data_tot_MOX, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_MEX <- read_excel(path_tot, sheet = "Marine-ECommerce-XL")
    data_tot_MEX <- as.data.frame(lapply(data_tot_MEX, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    
    data_tot <- data_tot_ME
    col_count <- ncol(data_tot)
    data_tot[,3:col_count] <- NA
    data_tot[,3:col_count] <- data_tot_ME[,3:col_count]+data_tot_MO[,3:col_count]+data_tot_MOX[,3:col_count]+data_tot_MEX[,3:col_count]
    
  }else if(sheet_name == "PA"){
    
    '    #MENGAMBIL DF per Cohort (Partial)
    path_name <- paste0("Workbook/PC ",cohort,st_or_lt,"/PaidClaim-Net.Raw.xlsx")
    data_par_PR <- read_excel(path_name, sheet = "PA-Retail")
    data_par_PR <- as.data.frame(lapply(data_par_PR, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_PC <- read_excel(path_name, sheet = "PA-Commercial")
    data_par_PC <- as.data.frame(lapply(data_par_PC, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_PTR <- read_excel(path_name, sheet = "PA-Travel-Retail")
    data_par_PTR <- as.data.frame(lapply(data_par_PTR, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_PTC <- read_excel(path_name, sheet = "PA-Travel-Commercial")
    data_par_PTC <- as.data.frame(lapply(data_par_PTC, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_PAX <- read_excel(path_name, sheet = "PA-XL")
    data_par_PAX <- as.data.frame(lapply(data_par_PAX, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    
    data_par <- data_par_PR
    col_count <- ncol(data_par)
    data_par[,3:col_count] <- NA
    data_par[,3:col_count] <- data_par_PR[,3:col_count]+data_par_PC[,3:col_count]+data_par_PTR[,3:col_count]+data_par_PTC[,3:col_count]+data_par_PAX[,3:col_count]
    '
    #MENGAMBIL DF Compiled
    data_comp_P <- read_excel("Workbook/PaidClaim-Net.Prior.wo.ML.xlsx", sheet = "PA")
    data_comp <- data_comp_P
    
    #MENGAMBIL DF Total (semua cohort)
    path_tot <- "Workbook/PaidClaim-NetTotal.Raw.FIXED-LBLW.v1.xlsx"
    data_tot_PR <- read_excel(path_tot, sheet = "PA-Retail")
    data_tot_PR <- as.data.frame(lapply(data_tot_PR, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_PC <- read_excel(path_tot, sheet = "PA-Commercial")
    data_tot_PC <- as.data.frame(lapply(data_tot_PC, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_PTR <- read_excel(path_tot, sheet = "PA-Travel-Retail")
    data_tot_PTR <- as.data.frame(lapply(data_tot_PTR, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_PTC <- read_excel(path_tot, sheet = "PA-Travel-Commercial")
    data_tot_PTC <- as.data.frame(lapply(data_tot_PTC, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_PAX <- read_excel(path_tot, sheet = "PA-XL")
    data_tot_PAX <- as.data.frame(lapply(data_tot_PAX, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    
    data_tot <- data_tot_PR
    col_count <- ncol(data_tot)
    data_tot[,3:col_count] <- NA
    data_tot[,3:col_count] <- data_tot_PR[,3:col_count]+data_tot_PC[,3:col_count]+data_tot_PTR[,3:col_count]+data_tot_PTC[,3:col_count]+data_tot_PAX[,3:col_count]
    
  }else if(sheet_name == "Various"){
    
    '    #MENGAMBIL DF per Cohort (Partial)
    path_name <- paste0("Workbook/PC ",cohort,st_or_lt,"/PaidClaim-Net.Raw.xlsx")
    data_par_PR <- read_excel(path_name, sheet = "PA-Retail")
    data_par_PR <- as.data.frame(lapply(data_par_PR, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_PC <- read_excel(path_name, sheet = "PA-Commercial")
    data_par_PC <- as.data.frame(lapply(data_par_PC, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_PTR <- read_excel(path_name, sheet = "PA-Travel-Retail")
    data_par_PTR <- as.data.frame(lapply(data_par_PTR, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_PTC <- read_excel(path_name, sheet = "PA-Travel-Commercial")
    data_par_PTC <- as.data.frame(lapply(data_par_PTC, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_par_PAX <- read_excel(path_name, sheet = "PA-XL")
    data_par_PAX <- as.data.frame(lapply(data_par_PAX, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    
    data_par <- data_par_PR
    col_count <- ncol(data_par)
    data_par[,3:col_count] <- NA
    data_par[,3:col_count] <- data_par_PR[,3:col_count]+data_par_PC[,3:col_count]+data_par_PTR[,3:col_count]+data_par_PTC[,3:col_count]+data_par_PAX[,3:col_count]
    '
    #MENGAMBIL DF Compiled
    data_comp_V <- read_excel("Workbook/PaidClaim-Net.Prior.wo.ML.xlsx", sheet = "Various")
    data_comp <- data_comp_V
    
    #MENGAMBIL DF Total (semua cohort)
    path_tot <- "Workbook/PaidClaim-NetTotal.Raw.FIXED-LBLW.v1.xlsx"
    data_tot_VF <- read_excel(path_tot, sheet = "Various-Fire")
    data_tot_VF <- as.data.frame(lapply(data_tot_VF, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_VL <- read_excel(path_tot, sheet = "Various-Liability")
    data_tot_VL <- as.data.frame(lapply(data_tot_VL, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_VPR <- read_excel(path_tot, sheet = "Various-PA-Retail")
    data_tot_VPR <- as.data.frame(lapply(data_tot_VPR, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_VPC <- read_excel(path_tot, sheet = "Various-PA-Commercial")
    data_tot_VPC <- as.data.frame(lapply(data_tot_VPC, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_VTR <- read_excel(path_tot, sheet = "Various-Travel-Retail")
    data_tot_VTR <- as.data.frame(lapply(data_tot_VTR, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_VTC <- read_excel(path_tot, sheet = "Various-Travel-Commercial")
    data_tot_VTC <- as.data.frame(lapply(data_tot_VTC, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_VTX <- read_excel(path_tot, sheet = "Various-Travel-XL")
    data_tot_VTX <- as.data.frame(lapply(data_tot_VTX, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_VTCr <- read_excel(path_tot, sheet = "Various-Trade Credit")
    data_tot_VTCr <- as.data.frame(lapply(data_tot_VTCr, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_VO <- read_excel(path_tot, sheet = "Various-Other")
    data_tot_VO <- as.data.frame(lapply(data_tot_VO, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    data_tot_VOX <- read_excel(path_tot, sheet = "Various-Other-XL")
    data_tot_VOX <- as.data.frame(lapply(data_tot_VOX, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    
    data_tot <- data_tot_VF
    col_count <- ncol(data_tot)
    data_tot[,3:col_count] <- NA
    data_tot[,3:col_count] <- data_tot_VF[,3:col_count]+data_tot_VL[,3:col_count]+data_tot_VPR[,3:col_count]+data_tot_VPC[,3:col_count]+data_tot_VTR[,3:col_count]+data_tot_VTC[,3:col_count]+data_tot_VTX[,3:col_count]+data_tot_VTCr[,3:col_count]+data_tot_VO[,3:col_count]+data_tot_VOX[,3:col_count]
    
    
  }else{
    
    '    #MENGAMBIL DF per Cohort (Partial)
    path_name <- paste0("Workbook/PC ",cohort,st_or_lt,"/PaidClaim-Net.Raw.xlsx")
    data_par <- read_excel(path_name, sheet = sheet_name)
    data_par <- as.data.frame(lapply(data_par, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))'
    
    #MENGAMBIL DF Compiled
    data_comp <- read_excel("Workbook/PaidClaim-Net.Prior.wo.ML.xlsx", sheet = sheet_name)
    
    #MENGAMBIL DF Total (semua cohort)
    path_tot <- "Workbook/PaidClaim-NetTotal.Raw.FIXED-LBLW.v1.xlsx"
    data_tot <- read_excel(path_tot, sheet = sheet_name)
    data_tot <- as.data.frame(lapply(data_tot, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    
  }
  
  data_ratio <- data_tot
  col_count <- ncol(data_ratio)
  data_ratio[,3:col_count] <- data_comp[,3:col_count]/data_tot[,3:col_count]
  
  data_ratio <- as.data.frame(lapply(data_ratio, function(x) {
    x[is.nan(x)] <- 1
    return(x)
  }))
  
  data_ratio <- as.data.frame(lapply(data_ratio, function(x) {
    x[is.na(x)] <- 1
    return(x)
  }))
  
  for(z in 2:nrow(data_ratio)){
    
    data_ratio[z,(jumlah_val+4-z):col_count] <- NA
    
  }
  
  '  data_prior <- data_par
  col_count <- ncol(data_prior)
  data_prior[,3:col_count] <- NA
  data_prior[,3:col_count] <- data_par[,3:col_count]*data_ratio[,3:col_count]'
  
  return(data_ratio)
  
} #Fungsi END

for(k in cohort_start:cohort_end){
  
  for(l in 0:1){
    
    ifelse(l == 1, st_or_lt <- "-ST", st_or_lt <- "-LT")
    file_path <- paste0("Workbook/PC ", k, st_or_lt, "/PaidClaim-Net.Ratio.xlsx")
    wb <- createWorkbook(file_path) 
    
    for(m in sheet_data$SheetName){
      
      data_ratio_x <- PC_prior_function(k,l,m)
      data_ratio_x <- data_ratio_x %>%
        mutate(across(everything(), ~ ifelse(is.infinite(.), 0, .)))
      addWorksheet(wb, sheet = m)
      writeData(wb, sheet = m, x = data_ratio_x)
      
      'for(i in 1:41){
        
        col_lim <- 44-i
        temp_vec <- data_prior_x[i,3:col_lim]
        wb_add_data(wb, sheet = m, x = t(as.data.frame(temp_vec)), startRow = i+1, startCol = 3, colNames = FALSE, rowNames = FALSE)
        
      }'
      
    }
    
    saveWorkbook(wb, file_path, overwrite = TRUE)
    
  }
  
  
}

