
EP_files <- NULL

valuation <- "202312"

cohort_start <- 2009
cohort_end <- 2023

Master_loc <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Master/")
Master_files <- c(
  paste0(Master_loc,"Earned Premium - wo XL.xlsx"),
  paste0(Master_loc,"PaidClaim-Gross.Prior.xlsx"),
  paste0(Master_loc,"PaidClaim-Gross.xlsx"),
  paste0(Master_loc,"PaidClaim-Net.Prior.xlsx"),
  paste0(Master_loc,"PaidClaim-Net.xlsx"),
  paste0(Master_loc,"PaidClaim-Net - wo XL.xlsx"),
  paste0(Master_loc,"OSClaim-Gross.Prior.xlsx"),
  paste0(Master_loc,"OSClaim-Gross.xlsx"),
  paste0(Master_loc,"OSClaim-Net.Prior.xlsx"),
  paste0(Master_loc,"OSClaim-Net.xlsx"),
  paste0(Master_loc,"OSClaim-Net - wo XL.xlsx")
)

for(i in cohort_start:cohort_end){
  
  for(j in 0:1){
    
    st_or_lt <- ifelse(j == 1, "-ST", "-LT")
    EP_loc <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Earned Premium/Workbook/EP ",i,st_or_lt,"/")
    PC_loc <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Paid Claims/Workbook/PC ",i,st_or_lt,"/")
    OS_loc <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Outstanding Claims/Workbook/OS ",i,st_or_lt,"/")
    
    EP_files <- c(
      paste0(EP_loc,"Earned Premium Prior Ratio.xlsx"),
      paste0(EP_loc,"Earned Premium.Manual.xlsx"),
      paste0(EP_loc,"Earned Premium.Prior.xlsx"),
      paste0(EP_loc,"Earned Premium.xlsx"),
      paste0(EP_loc,"EP with FS.xlsx"),
      paste0(EP_loc,"Earned Premium.Raw.xlsx"),
      paste0(EP_loc,"Earned Premium.Total.Raw.xlsx")
    )
    
    # PC_files <- c(
    #   paste0(PC_loc,"PaidClaim-Gross.Ratio.xlsx"),
    #   paste0(PC_loc,"PaidClaim-Gross.Raw.xlsx"),
    #   paste0(PC_loc,"PaidClaim-Net.Ratio.xlsx"),
    #   paste0(PC_loc,"PaidClaim-Net.Raw.xlsx")
    # )

    # OS_files <- c(
    #   paste0(OS_loc,"OSClaim-Gross.Ratio.xlsx"),
    #   paste0(OS_loc,"OSClaim-Gross.Raw.xlsx"),
    #   paste0(OS_loc,"OSClaim-Net.Ratio.xlsx"),
    #   paste0(OS_loc,"OSClaim-Net.Raw.xlsx")
    # )
  
    dir.create(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Compiled IFRS17/"))
    destination <-paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Compiled IFRS17/",i,st_or_lt)
    dir.create(destination)
    file.copy(EP_files, destination, overwrite = TRUE)
    # file.copy(PC_files, destination, overwrite = TRUE)
    # file.copy(OS_files, destination, overwrite = TRUE)
    file.copy(Master_files, destination, overwrite = TRUE)
    
  }
  
}

#COPY the IBNR Compiled Files not separated by cohort and st/lt

'for(i in 2009:2023){
  
  for(j in 0:1){
    
    st_or_lt <- ifelse(j == 1, "-ST", "-LT")
    
    compiled_files <- c(
      "D:/Stanley/Tugas IFRS17/TMI IBNR Preparation IFRS17 R/Cummulative Incurred/Compiled TMI/Cummulative-Incurred-Occurrence.xlsx",
      "D:/Stanley/Tugas IFRS17/TMI IBNR Preparation IFRS17 R/Cummulative Incurred/Compiled TMI/Unrecorded.MajorLoss.OSClaim-Gross.xlsx",
      "D:/Stanley/Tugas IFRS17/TMI IBNR Preparation IFRS17 R/Cummulative Incurred/Compiled TMI/Unrecorded.MajorLoss.OSClaim-Net.xlsx"
    )
    
    dir.create(paste0("D:/Stanley/Tugas IFRS17/TMI IBNR Preparation IFRS17 R/Compiled IFRS17/",valuation))
    destination <-paste0("D:/Stanley/Tugas IFRS17/TMI IBNR Preparation IFRS17 R/Compiled IFRS17/",valuation,"/",i,st_or_lt)
    dir.create(destination)
    file.copy(compiled_files, destination)
    
  }
  
}'
