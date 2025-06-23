
library(fs)

EP_files <- NULL

valuation <- "202412"

cohort_start <- 2009
cohort_end <- 2024

# Master_loc <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Master/")
# Master_files <- c(
#   paste0(Master_loc,"Earned Premium - wo XL.xlsx"),
#   paste0(Master_loc,"PaidClaim-Gross.Prior.xlsx"),
#   paste0(Master_loc,"PaidClaim-Gross.xlsx"),
#   paste0(Master_loc,"PaidClaim-Net.Prior.xlsx"),
#   paste0(Master_loc,"PaidClaim-Net.xlsx"),
#   paste0(Master_loc,"PaidClaim-Net - wo XL.xlsx"),
#   paste0(Master_loc,"OSClaim-Gross.Prior.xlsx"),
#   paste0(Master_loc,"OSClaim-Gross.xlsx"),
#   paste0(Master_loc,"OSClaim-Net.Prior.xlsx"),
#   paste0(Master_loc,"OSClaim-Net.xlsx"),
#   paste0(Master_loc,"OSClaim-Net - wo XL.xlsx")
# )

for(i in cohort_start:cohort_end){
  
  for(j in 0:1){
    
    st_or_lt <- ifelse(j == 1, "-ST", "-LT")
    EP_loc <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Earned Premium/Workbook/EP ",i,st_or_lt,"/")
    # PC_loc <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Paid Claims/Workbook/PC ",i,st_or_lt,"/")
    # OS_loc <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Outstanding Claims/Workbook/OS ",i,st_or_lt,"/")
    # 
    EP_files <- c(
      # paste0(EP_loc,"Earned Premium Prior Ratio.xlsx"),
      # paste0(EP_loc,"Earned Premium.Manual.xlsx"),
      # paste0(EP_loc,"Earned Premium.Prior.xlsx"),
      # paste0(EP_loc,"Earned Premium.xlsx"),
      # paste0(EP_loc,"EP with FS.xlsx"),
      paste0(EP_loc,"Earned Premium.XoL.xlsx")
      # paste0(EP_loc,"Earned Premium.Total.Raw.xlsx")
    )
    
    # PC_files <- c(
    #   paste0(PC_loc,"PaidClaim-Gross.Ratio.xlsx"),
    #   paste0(PC_loc,"PaidClaim-Gross.Raw.xlsx"),
    #   paste0(PC_loc,"PaidClaim-Net.Ratio.xlsx"),
    #   paste0(PC_loc,"PaidClaim-Net.Raw.xlsx")
    # )
    # 
    # OS_files <- c(
    #   paste0(OS_loc,"OSClaim-Gross.Ratio.xlsx"),
    #   paste0(OS_loc,"OSClaim-Gross.Raw.xlsx"),
    #   paste0(OS_loc,"OSClaim-Net.Ratio.xlsx"),
    #   paste0(OS_loc,"OSClaim-Net.Raw.xlsx")
    # )
    # 
    # PCT_files <- c(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Paid Claims/Workbook/PaidClaim-GrossTotal.Raw.FIXED-LBLW.v1.xlsx"),
    #                paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Paid Claims/Workbook/PaidClaim-NetTotal.Raw.FIXED-LBLW.v1.xlsx")
    # )
    # OST_files <- c(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Outstanding Claims/Workbook/OSClaim-GrossTotal.Raw.FIXED-LBLW.v1.xlsx"),
    #                paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Outstanding Claims/Workbook/OSClaim-NetTotal.Raw.FIXED-LBLW.v1.xlsx")    
    # )
    dir.create(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Compiled IFRS17/"))
    destination <-paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Compiled IFRS17/",i,st_or_lt)
    dir.create(destination)
    file.copy(EP_files, destination, overwrite = TRUE)
    # file.copy(PC_files, destination, overwrite = TRUE)
    # file.copy(OS_files, destination, overwrite = TRUE)
    # file.copy(PCT_files, destination, overwrite = TRUE)
    # file.copy(OST_files, destination, overwrite = TRUE)
    # file.copy(Master_files, destination, overwrite = TRUE)
    
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

#COPY THE LAST_COHORT_YEAR TO LAST_COHORT_YEAR + 1

last_coh_loc_st <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Compiled IFRS17/",cohort_end,"-ST")
last_coh_loc_lt <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Compiled IFRS17/",cohort_end,"-LT")

dir.create(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Compiled IFRS17/",cohort_end+1,"-ST"))
dir.create(paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Compiled IFRS17/",cohort_end+1,"-LT"))

last_1_coh_loc_st <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Compiled IFRS17/",cohort_end+1,"-ST")
last_1_coh_loc_lt <- paste0("D:/Stanley/Tugas IFRS17/TMI Reserving Macro Automation/TMI/",valuation,"/Compiled IFRS17/",cohort_end+1,"-LT")

file.copy(from = list.files(last_coh_loc_st, full.names = TRUE),
          to = last_1_coh_loc_st,
          recursive = TRUE)

file.copy(from = list.files(last_coh_loc_lt, full.names = TRUE),
          to = last_1_coh_loc_lt,
          recursive = TRUE)