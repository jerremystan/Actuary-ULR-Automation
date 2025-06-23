library(readxl)
library(DBI)
library(odbc)
library(dplyr)
library(lubridate)
library(tidyr)
library(openxlsx)
library(writexl)

#GANTI VALUATION!!!

valuation_date <- "202412"
val_year <- 2024
val_quarter <- 4

#GANTI COHORT!!!

cohort_start <- 2009
cohort_end <- 2024

row_count <- (val_year-1-2009)*4+1+val_quarter

limit_row <- row_count+1

colname_GN <- unlist(read_excel(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/", valuation_date, "/Earned Premium/Workbook/EP 2009-ST/Earned Premium.xlsx"), range = paste0("EP!A1:AD1"), col_names = FALSE))
colname_GN[3:14] <- "Gross"
colname_GN[15:26] <- "Net"
colname_NM <- unlist(read_excel(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/", valuation_date, "/Earned Premium/Workbook/EP 2009-ST/Earned Premium.xlsx"), range = paste0("EP!A2:AD2"), col_names = FALSE))

EP_compiled <- read_excel("Workbook/Earned Premium.Manual.xlsx", range = paste0("EP!A2:AD",limit_row))
EP_compiled[is.na(EP_compiled)] <- 0
EP_compiled[c(3:14,27,29)] <- 0


EP_contoh <- read_excel(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/", valuation_date, "/Earned Premium/Workbook/EP 2009-ST/Earned Premium.xlsx"), range = paste0("EP!A2:AD",limit_row))
EP_contoh[is.na(EP_contoh)] <- 0
EP_contoh[3:ncol(EP_contoh)] <-  lapply(EP_contoh[3:ncol(EP_contoh)], function(x) 0)
EP_manual_template <- EP_contoh
EP_manual_recon <- EP_contoh
EP_manual_recon_result <- EP_contoh

#EP_manual_recon_result[3:ncol(EP_manual_recon_result)] <- 0

EP_temp <- EP_contoh
EP_contoh <- EP_temp


#Menghasilkan xlsx untuk total cohort EP Final

for (c in cohort_start:cohort_end){
  
  for (t in 0:1){
    
    st_lt <- ifelse(t==1, "-ST", "-LT")
    
    EP_final_temp <- read_excel(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/", valuation_date, "/Earned Premium/Workbook/EP ", c, st_lt, "/Earned Premium.xlsx"), range = paste0("EP!A2:AD",limit_row))
    EP_contoh[3:ncol(EP_contoh)] <- EP_contoh[3:ncol(EP_contoh)] + EP_final_temp[3:ncol(EP_final_temp)]
    
    
  }
  
  
}

#write.csv(EP_contoh, "EP_Final_Total_Cohort.csv", row.names = FALSE)

for (c in cohort_start:cohort_end){
  
  for (t in 0:1){
    
    st_lt <- ifelse(t==1, "-ST", "-LT")
    
    wb <- createWorkbook()
    addWorksheet(wb, "EP")
    
    
    
    EP_final_temp <- read_excel(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/", valuation_date, "/Earned Premium/Workbook/EP ", c, st_lt, "/Earned Premium.xlsx"), range = paste0("EP!A2:AD",limit_row))
    EP_manual_template[3:ncol(EP_manual_template)] <- EP_final_temp[3:ncol(EP_final_temp)] * (EP_compiled[3:ncol(EP_compiled)] / EP_contoh[3:ncol(EP_contoh)])
    EP_manual_template <- as.data.frame(lapply(EP_manual_template, function(x) {
      x[is.nan(x)] <- 0
      return(x)
    }))
    
    EP_manual_recon[3:ncol(EP_manual_recon)] <- EP_manual_recon[3:ncol(EP_manual_recon)]+EP_manual_template[3:ncol(EP_manual_template)]
    writeData(wb, sheet = "EP", x = EP_manual_template, startCol = 1, startRow = 3, rowNames = FALSE, colNames = FALSE)
    writeData(wb, sheet = "EP", x = t(colname_GN), startCol = 1, startRow = 1, rowNames = FALSE, colNames = FALSE)
    writeData(wb, sheet = "EP", x = t(colname_NM), startCol = 1, startRow = 2, rowNames = FALSE, colNames = FALSE)
    saveWorkbook(wb, paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/", valuation_date, "/Earned Premium/Workbook/EP ", c, st_lt, "/Earned Premium.Manual.xlsx"), overwrite = TRUE)
    
  }
  
  
}


#RECONCILE

EP_manual_recon_result[3:ncol(EP_manual_recon_result)] <- EP_manual_recon[3:ncol(EP_manual_recon)] - EP_compiled[3:ncol(EP_compiled)]
write.csv(EP_manual_recon_result, "Reconcile_EP_Manual.csv", row.names = FALSE)


