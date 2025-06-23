library(stringr)
library(dplyr)
library(openxlsx)
library(readxl)


val_year <- 2023
val_quarter <- 4

cohort_start <- 2022
cohort_end <- 2023

row_count <- (val_year-1-2009)*4+1+val_quarter

EP_template <- read_excel("Compiled IFRS17/2009-LT/Earned Premium - wo XL.xlsx", sheet = "EP", range = paste0("A2:AD",row_count+1))
EP_template[3:ncol(EP_template)] <- 0
EP_colGN <- unlist(read_excel("Compiled IFRS17/2009-LT/Earned Premium - wo XL.xlsx", sheet = "EP", range = "A1:AD1", col_names = FALSE))
EP_colGN[3:14] <- "Gross"
EP_colGN[15:26] <- "Net"
EP_colLB <- unlist(read_excel("Compiled IFRS17/2009-LT/Earned Premium - wo XL.xlsx", sheet = "EP", range = "A2:AD2", col_names = FALSE))

CL_template <- read_excel("Compiled IFRS17/2009-LT/OSClaim-Gross.xlsx", sheet = "Auto")
CL_template[3:ncol(CL_template)] <- 0



PG_temp <- CL_template
PN_temp <- CL_template

ON_temp <- CL_template

#wb <- loadWorkbook(file = system.file("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/202312/Compiled IFRS17/2009-LT", "OSClaim-Gross.xlsx", package = "openxlsx"))
wb <- loadWorkbook("Compiled IFRS17/2009-LT/OSClaim-Gross.xlsx")
sheet_list <- names(wb)

EP_temp <- EP_template

for (j in c("-ST","-LT")){
  
  #Earned Premium
  
  for (i in cohort_start:cohort_end){
    
    EP_add <- read_excel(paste0("Compiled IFRS17/",i,j,"/Earned Premium - wo XL.xlsx"), sheet = "EP", range = paste0("A2:AD",row_count+1))
    EP_temp[3:ncol(EP_temp)] <- EP_temp[3:ncol(EP_temp)] + EP_add[3:ncol(EP_add)]
    
  }
  
  #Save EP
  
  wb <- createWorkbook()
  addWorksheet(wb, "EP")
  
  writeData(wb, "EP", EP_temp, startCol = 1, startRow = 3, colNames = FALSE, rowNames = FALSE)
  writeData(wb, "EP", t(EP_colGN), startCol = 1, startRow = 1, colNames = FALSE, rowNames = FALSE)
  writeData(wb, "EP", t(EP_colLB), startCol = 1, startRow = 2, colNames = FALSE, rowNames = FALSE)
  
  dir.create(paste0("Compiled IFRS17/",cohort_start,"-",cohort_end,j))
  
  saveWorkbook(wb, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end,j, "/Earned Premium - wo XL.xlsx"), overwrite = TRUE)
  
}



