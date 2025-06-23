library(stringr)
library(dplyr)
library(openxlsx)
library(readxl)


val_year <- 2023
val_quarter <- 4

cohort_start <- 2009
cohort_end <- 2023

row_count <- (val_year-1-2009)*4+1+val_quarter

CL_template <- read_excel("Compiled IFRS17/2009-LT/OSClaim-Gross.xlsx", sheet = "Auto")
CL_template[3:ncol(CL_template)] <- 0

CL_diff <- CL_template

PG_temp <- CL_template
PN_temp <- CL_template
OG_temp <- CL_template
ON_temp <- CL_template

#wb <- loadWorkbook(file = system.file("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/202312/Compiled IFRS17/2009-LT", "OSClaim-Gross.xlsx", package = "openxlsx"))
wb <- loadWorkbook("Master/OSClaim-Gross.xlsx")
sheet_list <- names(wb)
sheet_list <- sheet_list[sheet_list != "Fire-XL"]

#OS Gross
wb <- createWorkbook()
wbd <- createWorkbook()
for (sheet in sheet_list){
  
  addWorksheet(wb, sheetName = sheet)
  addWorksheet(wbd, sheetName = sheet)
  OG_temp <- CL_template
  CL_diff <- CL_template
  for (j in c("-ST","-LT")){
    for (i in cohort_start:cohort_end){
      print(paste0("OSG-",i,j,sheet))
      CL_add <- read_excel(paste0("Compiled IFRS17/",i,j,"/OSClaim-Gross.xlsx"), sheet = sheet)
      CL_add <- as.data.frame(lapply(CL_add, function(x) {
        x[is.na(x)] <- 0
        return(x)
      }))
      OG_temp[3:ncol(OG_temp)] <- OG_temp[3:ncol(OG_temp)] + CL_add[3:ncol(CL_add)]
    
    }
  }
  CL_comp <- CL_template
  CL_comp <- read_excel(paste0("Reserving Compiled/OSClaim-Gross.xlsx"), sheet = sheet)
  CL_comp <- as.data.frame(lapply(CL_comp, function(x) {
    x[is.na(x)] <- 0
    return(x)
  }))
  CL_diff[3:ncol(CL_diff)] <- OG_temp[3:ncol(OG_temp)] - CL_comp[3:ncol(CL_comp)]
  writeData(wb, sheet, OG_temp)
  writeData(wbd, sheet, CL_diff)
  
}
saveWorkbook(wb, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/OSClaim-Gross-IFRS.xlsx"), overwrite = TRUE)
saveWorkbook(wbd, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/OSClaim-Gross - Checker.xlsx"), overwrite = TRUE)


#OS Net
wb <- createWorkbook()
wbd <- createWorkbook()
for (sheet in sheet_list){
  
  addWorksheet(wb, sheetName = sheet)
  addWorksheet(wbd, sheetName = sheet)
  ON_temp <- CL_template
  CL_diff <- CL_template
  for (j in c("-ST","-LT")){
    for (i in cohort_start:cohort_end){
      print(paste0("OSN-",i,j,sheet))
      CL_add <- read_excel(paste0("Compiled IFRS17/",i,j,"/OSClaim-Net.xlsx"), sheet = sheet)
      CL_add <- as.data.frame(lapply(CL_add, function(x) {
        x[is.na(x)] <- 0
        return(x)
      }))
      ON_temp[3:ncol(ON_temp)] <- ON_temp[3:ncol(ON_temp)] + CL_add[3:ncol(CL_add)]
    
    }
  }
  CL_comp <- CL_template
  CL_comp <- read_excel(paste0("Reserving Compiled/OSClaim-Net.xlsx"), sheet = sheet)
  CL_comp <- as.data.frame(lapply(CL_comp, function(x) {
    x[is.na(x)] <- 0
    return(x)
  }))
  CL_diff[3:ncol(CL_diff)] <- ON_temp[3:ncol(ON_temp)] - CL_comp[3:ncol(CL_comp)]
  writeData(wb, sheet, ON_temp)
  writeData(wbd, sheet, CL_diff)
  
    
}
saveWorkbook(wb, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/OSClaim-Net-IFRS.xlsx"), overwrite = TRUE)
saveWorkbook(wbd, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/OSClaim-Net - Checker.xlsx"), overwrite = TRUE)

#Paid Gross
wb <- createWorkbook()
wbd <- createWorkbook()
for (sheet in sheet_list){
  
  addWorksheet(wb, sheetName = sheet)
  addWorksheet(wbd, sheetName = sheet)
  PG_temp <- CL_template
  CL_diff <- CL_template
  for (j in c("-ST","-LT")){
    for (i in cohort_start:cohort_end){
      print(paste0("PG-",i,j,sheet))
      CL_add <- read_excel(paste0("Compiled IFRS17/",i,j,"/PaidClaim-Gross.xlsx"), sheet = sheet)
      CL_add <- as.data.frame(lapply(CL_add, function(x) {
        x[is.na(x)] <- 0
        return(x)
      }))
      PG_temp[3:ncol(PG_temp)] <- PG_temp[3:ncol(PG_temp)] + CL_add[3:ncol(CL_add)]
    
    }
  }
  CL_comp <- CL_template
  CL_comp <- read_excel(paste0("Reserving Compiled/PaidClaim-Gross.xlsx"), sheet = sheet)
  CL_comp <- as.data.frame(lapply(CL_comp, function(x) {
    x[is.na(x)] <- 0
    return(x)
  }))
  CL_diff[3:ncol(CL_diff)] <- PG_temp[3:ncol(PG_temp)] - CL_comp[3:ncol(CL_comp)]
  writeData(wb, sheet, PG_temp)
  writeData(wbd, sheet, CL_diff)
  
}
saveWorkbook(wb, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/PaidClaim-Gross-IFRS.xlsx"), overwrite = TRUE)
saveWorkbook(wbd, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/PaidClaim-Gross - Checker.xlsx"), overwrite = TRUE)

#Paid Net
wb <- createWorkbook()
wbd <- createWorkbook()
for (sheet in sheet_list){
  
  addWorksheet(wb, sheetName = sheet)
  addWorksheet(wbd, sheetName = sheet)
  PN_temp <- CL_template
  CL_diff <- CL_template
  for (j in c("-ST","-LT")){
    for (i in cohort_start:cohort_end){
      print(paste0("PN-",i,j,sheet))
      CL_add <- read_excel(paste0("Compiled IFRS17/",i,j,"/PaidClaim-Net.xlsx"), sheet = sheet)
      CL_add <- as.data.frame(lapply(CL_add, function(x) {
        x[is.na(x)] <- 0
        return(x)
      }))
      PN_temp[3:ncol(PN_temp)] <- PN_temp[3:ncol(PN_temp)] + CL_add[3:ncol(CL_add)]
    
    }
  }
  CL_comp <- CL_template
  CL_comp <- read_excel(paste0("Reserving Compiled/PaidClaim-Net.xlsx"), sheet = sheet)
  CL_comp <- as.data.frame(lapply(CL_comp, function(x) {
    x[is.na(x)] <- 0
    return(x)
  }))
  CL_diff[3:ncol(CL_diff)] <- PN_temp[3:ncol(PN_temp)] - CL_comp[3:ncol(CL_comp)]
  writeData(wb, sheet, PN_temp)
  writeData(wbd, sheet, CL_diff)
  
}
saveWorkbook(wb, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/PaidClaim-Net-IFRS.xlsx"), overwrite = TRUE)
saveWorkbook(wbd, paste0("Compiled IFRS17/",cohort_start,"-",cohort_end, "/PaidClaim-Net - Checker.xlsx"), overwrite = TRUE)
