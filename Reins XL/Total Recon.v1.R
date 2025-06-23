library(stringr)
library(dplyr)
library(openxlsx)
library(readxl)


val_year <- 2024
val_quarter <- 2

cohort_start <- 2009
cohort_end <- 2024

row_count <- (val_year-1-2009)*4+1+val_quarter

EP_template <- read_excel("Compiled IFRS17/2009-LT/Earned Premium.xlsx", sheet = "EP", range = paste0("A2:AD",row_count+1))
EP_template[3:ncol(EP_template)] <- 0
EP_colGN <- unlist(read_excel("Compiled IFRS17/2009-LT/Earned Premium.xlsx", sheet = "EP", range = "A1:AD1", col_names = FALSE))
EP_colGN[3:14] <- "Gross"
EP_colGN[15:26] <- "Net"
EP_colLB <- unlist(read_excel("Compiled IFRS17/2009-LT/Earned Premium.xlsx", sheet = "EP", range = "A2:AD2", col_names = FALSE))

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
    
    EP_add <- read_excel(paste0("Compiled IFRS17/",i,j,"/Earned Premium.xlsx"), sheet = "EP", range = paste0("A2:AD",row_count+1))
    EP_temp[3:ncol(EP_temp)] <- EP_temp[3:ncol(EP_temp)] + EP_add[3:ncol(EP_add)]
    
  }
}

dir.create(paste0("Compiled IFRS17/",cohort_start,"-",cohort_end))

#Save EP

wb <- createWorkbook()
addWorksheet(wb, "EP")

writeData(wb, "EP", EP_temp, startCol = 1, startRow = 3, colNames = FALSE, rowNames = FALSE)
writeData(wb, "EP", t(EP_colGN), startCol = 1, startRow = 1, colNames = FALSE, rowNames = FALSE)
writeData(wb, "EP", t(EP_colLB), startCol = 1, startRow = 2, colNames = FALSE, rowNames = FALSE)

saveWorkbook(wb, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/Earned Premium.xlsx"), overwrite = TRUE)


#OS Gross
wb <- createWorkbook()
for (sheet in sheet_list){
  
  addWorksheet(wb, sheetName = sheet)
  OG_temp <- CL_template
  for (j in c("-ST","-LT")){
    for (i in cohort_start:cohort_end){
    
      CL_add <- read_excel(paste0("Compiled IFRS17/",i,j,"/OSClaim-Gross.xlsx"), sheet = sheet)
      OG_temp[3:ncol(OG_temp)] <- OG_temp[3:ncol(OG_temp)] + CL_add[3:ncol(CL_add)]
    
    }
  }
  writeData(wb, sheet, OG_temp)
  
}
saveWorkbook(wb, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/OSClaim-Gross.xlsx"), overwrite = TRUE)

#OS Net
wb <- createWorkbook()
for (sheet in sheet_list){
  
  addWorksheet(wb, sheetName = sheet)
  ON_temp <- CL_template
  for (j in c("-ST","-LT")){
    for (i in cohort_start:cohort_end){
    
      CL_add <- read_excel(paste0("Compiled IFRS17/",i,j,"/OSClaim-Net.xlsx"), sheet = sheet)
      ON_temp[3:ncol(ON_temp)] <- ON_temp[3:ncol(ON_temp)] + CL_add[3:ncol(CL_add)]
    
    }
  }
  writeData(wb, sheet, ON_temp)
  
}
saveWorkbook(wb, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/OSClaim-Net.xlsx"), overwrite = TRUE)

#Paid Gross
wb <- createWorkbook()
for (sheet in sheet_list){
  
  addWorksheet(wb, sheetName = sheet)
  PG_temp <- CL_template
  for (j in c("-ST","-LT")){
    for (i in cohort_start:cohort_end){
    
      CL_add <- read_excel(paste0("Compiled IFRS17/",i,j,"/PaidClaim-Gross.xlsx"), sheet = sheet)
      PG_temp[3:ncol(PG_temp)] <- PG_temp[3:ncol(PG_temp)] + CL_add[3:ncol(CL_add)]
    
    }
  }
  writeData(wb, sheet, PG_temp)
  
}
saveWorkbook(wb, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/PaidClaim-Gross.xlsx"), overwrite = TRUE)

#Paid Net
wb <- createWorkbook()
for (sheet in sheet_list){
  
  addWorksheet(wb, sheetName = sheet)
  PN_temp <- CL_template
  for (j in c("-ST","-LT")){
    for (i in cohort_start:cohort_end){
    
      CL_add <- read_excel(paste0("Compiled IFRS17/",i,j,"/PaidClaim-Net.xlsx"), sheet = sheet)
      PN_temp[3:ncol(PN_temp)] <- PN_temp[3:ncol(PN_temp)] + CL_add[3:ncol(CL_add)]
    
    }
  }
  writeData(wb, sheet, PN_temp)
  
}
saveWorkbook(wb, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/PaidClaim-Net.xlsx"), overwrite = TRUE)
