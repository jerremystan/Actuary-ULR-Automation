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

function_raw <- function(sheet_name, condition){
  
  if(sheet_name == "Marine"){
    
    #MENGAMBIL DF Total (semua cohort)
    path_tot <- paste0("Workbook/OSClaim-Net", condition)
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
    
    #MENGAMBIL DF Total (semua cohort)
    path_tot <- paste0("Workbook/OSClaim-Net", condition)
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
    
    #MENGAMBIL DF Total (semua cohort)
    path_tot <- paste0("Workbook/OSClaim-Net", condition)
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
    
    
    #MENGAMBIL DF Total (semua cohort)
    path_tot <- paste0("Workbook/OSClaim-Net", condition)
    data_tot <- read_excel(path_tot, sheet = sheet_name)
    data_tot <- as.data.frame(lapply(data_tot, function(x) {
      x[is.na(x)] <- 0
      return(x)
    }))
    
  }
  
  return(data_tot)
  
} #Fungsi END

wb <- createWorkbook()
wbx <- createWorkbook()

for (p in sheet_data$SheetName){
  print(p)
  data_tot_wo_ML <- function_raw(p, "Total.Raw.FIXED-LBLW.v1.xlsx")
  #data_tot_wo_ML <- data_tot__wo_ML %>%
  #mutate(across(everything(), ~ ifelse(is.infinite(.), 0, .)))
  
  data_tot_w_ML <- function_raw(p, ".Raw.xlsx")
  #data_tot_w_ML <- data_tot_w_ML %>%
  #mutate(across(everything(), ~ ifelse(is.infinite(.), 0, .)))
  
  data_prior_w_ML <- read_excel("Workbook/OSClaim-Net.Prior.xlsx", sheet = p)
  
  ifelse(ncol(data_prior_w_ML) > (nrow(data_prior_w_ML)+2),
         data_prior_w_ML <- data_prior_w_ML[,-(nrow(data_prior_w_ML)+3):-ncol(data_prior_w_ML)],
  )
  
  data_ratio <- data_tot_wo_ML
  col_count <- ncol(data_ratio)
  data_ratio[,3:col_count] <- data_tot_wo_ML[,3:col_count]/data_tot_w_ML[,3:col_count]
  
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
  
  data_prior_wo_ML <- data_tot_wo_ML
  col_count <- ncol(data_prior_wo_ML)
  data_prior_wo_ML[,3:col_count] <- data_prior_w_ML[,3:col_count] * data_ratio[,3:col_count]
  
  addWorksheet(wb, sheet = p)
  addWorksheet(wbx, sheet = p)
  writeData(wb, sheet = p, x = data_ratio)
  writeData(wbx, sheet = p, x = data_prior_wo_ML)
  
}

saveWorkbook(wb, "Workbook/OSClaim-Net.Ratio.wo.ML.xlsx", overwrite = TRUE)
saveWorkbook(wbx, "Workbook/OSClaim-Net.Prior.wo.ML.xlsx", overwrite = TRUE)



