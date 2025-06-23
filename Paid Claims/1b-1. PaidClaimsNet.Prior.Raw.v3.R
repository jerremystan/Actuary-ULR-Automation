
library(readxl)
library(DBI)
library(odbc)
library(dplyr)
library(lubridate)
library(tidyr)
library(openxlsx)
library(writexl)
library(cellranger)


# Define the connection details
'server <- "ARISTYO-NB"
database <- "TMI.Actuary" #untuk semua db

# Create a connection to the SQL Server database using Windows Authentication
con <- dbConnect(odbc::odbc(),
                 Driver   = "SQL Server",   # This is the ODBC driver name
                 Server   = server,
                 Database = database,
                 Trusted_Connection = "Yes")'

#GANTI VALUATION!!!

val_year <- 2024
val_quarter <- 4

#GANTI COHORT!!!

cohort_start <- 2020
cohort_end <- 2024

row_count <- (val_year-1-2009)*4+1+val_quarter

sql_data <- read.csv("SQLData.csv")
#sql_data2 <- read_excel("SQLData.xlsx")
sql_data$Net <- as.numeric(sql_data$Net)
#sql_data$Net <- as.numeric(sql_data$Net)
sql_data <- na.omit(sql_data)

sheet_df <- read_excel("Mapping_LoB_Sheet.v1.xlsx")

sql_data <- sql_data[sql_data$Cohort != "NULL",]
sql_data <- sql_data[sql_data$IsShortTerm != "NULL",]

result <- sql_data %>%
  group_by(SheetName, ReportingYear, ReportingMonth, LossYear, LossMonth, Cohort, IsShortTerm) %>%
  summarize(
    Gross = sum(Gross)
    ,Net = sum(Net)
  )

sql_data <- result

per_master <- expand.grid(year = seq(2010,(val_year-1)),
                          kuarter = seq(1,4))
per_master_x <- expand.grid(year = val_year,
                            kuarter = seq(1,val_quarter))

per_master_o <- data.frame(year = c(2009),
                           kuarter = c(4))

per_master <- rbind(per_master_o,per_master,per_master_x)
per_master <- per_master[order(per_master$year),]
per_master$masterper <- paste0(per_master$year,per_master$kuarter)

'tahun <- 2012
kuarter <- 3
term <- 1
cohort <- 2012
sheet <- "AUTO"'

sheet_list <- unique(sql_data$SheetName)


PC_function <- function(cohort, term, sheet, tahun, kuarter){   #fungsi term, cohort, sheet, tahun, kuarter
  
  filter_master <- per_master$masterper[per_master$masterper >= paste0(tahun,kuarter)]
  
  df_index <- data.frame(LossPer = rep(paste0(tahun,kuarter),length(filter_master)),
                         RPTPer = filter_master,
                         ColIndex = seq(1,length(filter_master)))
  df_index$RPTYear <- substr(df_index$RPTPer,1,4)
  df_index$RPTQuar <- substr(df_index$RPTPer,5,5)
  
  num_col <- lapply(1:row_count, function(x){
    
    ifelse(length(sql_data$Net[sql_data$LossYear == tahun & sql_data$LossMonth == kuarter & sql_data$ReportingYear == df_index$RPTYear[df_index$ColIndex == x] & sql_data$ReportingMonth == df_index$RPTQuar[df_index$ColIndex == x] & sql_data$Cohort == cohort & sql_data$IsShortTerm == term & sql_data$SheetName == sheet]) > 0, sql_data$Net[sql_data$LossYear == tahun & sql_data$LossMonth == kuarter & sql_data$ReportingYear == df_index$RPTYear[df_index$ColIndex == x] & sql_data$ReportingMonth == df_index$RPTQuar[df_index$ColIndex == x] & sql_data$Cohort == cohort & sql_data$IsShortTerm == term & sql_data$SheetName == sheet],"")
    
  })
  
  new_row <- data.frame(Year = ifelse(tahun == 2009, "Prior", tahun),
                        Quarter = ifelse(tahun == 2009 & kuarter == 4,"Prior", kuarter),
                        setNames(as.data.frame(num_col), as.character(1:row_count)))
  
  return(new_row)
  
}

prior_sheet <- read_excel("Prior_Sheet.v1.xlsx")
excel_range <- cell_limits(c(1,1),c(row_count+1,row_count+2))
template_prior <- read_excel("Workbook/PaidClaim-Net.Prior.xlsx",sheet = "Auto", range = excel_range)
template_prior[,3:ncol(template_prior)] <- 0
for (i in 1:row_count){
  if (i > 1){
  template_prior[i,(row_count+2+2-i):(row_count+2)] <- NA
  }
}

Triangle_raw_function <- function(sheet,cohort,term){
  triangle_temp <- read_excel(paste0("Workbook/PC ",cohort,"-",term,"/PaidClaim-Net.Raw.xlsx"), sheet = sheet, range = excel_range)
  triangle_temp[is.na(triangle_temp)] <- 0
  for (x in 1:row_count){
    if (x > 1){
      triangle_temp[x,(row_count+2+2-x):(row_count+2)] <- NA
    }
  }
  return(triangle_temp)
}

Triangle_prior_function <- function(sheet){
  triangle_temp <- read_excel(paste0("Workbook/","/PaidClaim-Net.Prior.xlsx"), sheet = sheet, range = excel_range)
  for (x in 1:row_count){
    if (x > 1){
      triangle_temp[x,(row_count+2+2-x):(row_count+2)] <- NA
    }
  }
  triangle_temp[is.na(triangle_temp)] <- 0
  for (x in 1:row_count){
    if (x > 1){
      triangle_temp[x,(row_count+2+2-x):(row_count+2)] <- NA
    }
  }
  return(triangle_temp)
}

Triangle_total_raw_function <- function(sheet){
  triangle_temp <- read_excel(paste0("Workbook/","/PaidClaim-Net.Total.Raw.xlsx"), sheet = sheet, range = excel_range)
  for (x in 1:row_count){
    if (x > 1){
      triangle_temp[x,(row_count+2+2-x):(row_count+2)] <- NA
    }
  }
  triangle_temp[is.na(triangle_temp)] <- 0
  for (x in 1:row_count){
    if (x > 1){
      triangle_temp[x,(row_count+2+2-x):(row_count+2)] <- NA
    }
  }
  return(triangle_temp)
}


for ( i in cohort_start:cohort_end){ #COHORT
  
  for (j in 0:1){ #TERM
    
    st_or_lt <- ifelse(j == 1, "-ST","-LT")
    dir_name <- paste0("Workbook/PC ",i,st_or_lt)
    dir.create(dir_name)
    
    wb <- createWorkbook()
    
    for (k in sheet_df$Sheet){ #SHEET 
      
      PC_data <- data.frame() #NEW BLANK temporary every LoB
      sheet_name <- sheet_df$Compiled[sheet_df$Sheet == k]
      addWorksheet(wb, sheet_name)
      
      for (l in 2009:val_year){ #ACCIDENT YEAR
        
        for (m in 1:4){ #ACCIDENT QUARTER
          print(paste0(i,st_or_lt,"-",k,"-",l,"-",m))
          if(l == val_year & m > val_quarter){break}
          else if(l == 2009 & m %in% c(1,2,3)){}
          else{
            result_row <- PC_function(i,j,k,l,m)
            PC_data <- rbind(PC_data, result_row)
          }
          
        } #ACCIDENT QUARTER
        
        
      } #ACCIDENT YEAR
      
      PC_data[,3:ncol(PC_data)] <- lapply(PC_data[,3:ncol(PC_data)], as.numeric)
      colnames(PC_data) <- c("Year","Kuarter",seq(1,nrow(PC_data)))
      writeData(wb, sheet_name, PC_data)
      #addStyle(wb, sheet = sheet_name, style = createStyle(numFmt = "######.######"), rows = 1:nrow(PC_data)+1, cols = 1:ncol(PC_data), gridExpand = TRUE)
      
    } #SHEET 
    
    saveWorkbook(wb, paste0(dir_name,"/PaidClaim-Net.Raw.xlsx"), overwrite = TRUE)
    
    
  } #TERM
  
  #PRIOR
  # if(i == 2020){
  #   
    #Auto
    prior_auto_st <- Triangle_raw_function("Auto",i,"ST")
    prior_auto_lt <- Triangle_raw_function("Auto",i,"LT")

    prior_auto <- Triangle_total_raw_function("Auto")

    comp_prior_auto <- Triangle_prior_function("Auto")

    prior_auto_st_final <- template_prior
    prior_auto_st_final[,3:(row_count + 2)] <- (prior_auto_st[,3:(row_count + 2)]/prior_auto[,3:(row_count + 2)]) * comp_prior_auto[,3:(row_count + 2)]
    prior_auto_st_final <- prior_auto_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_auto_lt_final <- template_prior
    prior_auto_lt_final[,3:(row_count + 2)] <- (prior_auto_lt[,3:(row_count + 2)]/prior_auto[,3:(row_count + 2)]) * comp_prior_auto[,3:(row_count + 2)]
    prior_auto_lt_final <- prior_auto_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))

    #Auto-OD-PartialLoss
    prior_auto_PL_st <- Triangle_raw_function("Auto-OD-PartialLoss",i,"ST")
    prior_auto_PL_lt <- Triangle_raw_function("Auto-OD-PartialLoss",i,"LT")

    prior_auto_PL <- Triangle_total_raw_function("Auto-OD-PartialLoss")

    comp_prior_auto_PL <- Triangle_prior_function("Auto-OD-PartialLoss")

    prior_auto_PL_st_final <- template_prior
    prior_auto_PL_st_final[,3:(row_count + 2)] <- (prior_auto_PL_st[,3:(row_count + 2)]/prior_auto_PL[,3:(row_count + 2)]) * comp_prior_auto_PL[,3:(row_count + 2)]
    prior_auto_PL_st_final <- prior_auto_PL_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_auto_PL_lt_final <- template_prior
    prior_auto_PL_lt_final[,3:(row_count + 2)] <- (prior_auto_PL_lt[,3:(row_count + 2)]/prior_auto_PL[,3:(row_count + 2)]) * comp_prior_auto_PL[,3:(row_count + 2)]
    prior_auto_PL_lt_final <- prior_auto_PL_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))

    #Auto-OD-TotalLoss-MotorCycle
    prior_auto_TLM_st <- Triangle_raw_function("Auto-OD-TotalLoss-MotorCycle",i,"ST")
    prior_auto_TLM_lt <- Triangle_raw_function("Auto-OD-TotalLoss-MotorCycle",i,"LT")

    prior_auto_TLM <- Triangle_total_raw_function("Auto-OD-TotalLoss-MotorCycle")

    comp_prior_auto_TLM <- Triangle_prior_function("Auto-OD-TotalLoss-MotorCycle")

    prior_auto_TLM_st_final <- template_prior
    prior_auto_TLM_st_final[,3:(row_count + 2)] <- (prior_auto_TLM_st[,3:(row_count + 2)]/prior_auto_TLM[,3:(row_count + 2)]) * comp_prior_auto_TLM[,3:(row_count + 2)]
    prior_auto_TLM_st_final <- prior_auto_TLM_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_auto_TLM_lt_final <- template_prior
    prior_auto_TLM_lt_final[,3:(row_count + 2)] <- (prior_auto_TLM_lt[,3:(row_count + 2)]/prior_auto_TLM[,3:(row_count + 2)]) * comp_prior_auto_TLM[,3:(row_count + 2)]
    prior_auto_TLM_lt_final <- prior_auto_TLM_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))

    #Auto-OD-TotalLoss-NonMotorCycle
    prior_auto_TLN_st <- Triangle_raw_function("Auto-OD-TotalLoss-NonMotorCycle",i,"ST")
    prior_auto_TLN_lt <- Triangle_raw_function("Auto-OD-TotalLoss-NonMotorCycle",i,"LT")

    prior_auto_TLN <- Triangle_total_raw_function("Auto-OD-TotalLoss-NonMotorCycle")
      
    comp_prior_auto_TLN <- Triangle_prior_function("Auto-OD-TotalLoss-NonMotorCycle")

    prior_auto_TLN_st_final <- template_prior
    prior_auto_TLN_st_final[,3:(row_count + 2)] <- (prior_auto_TLN_st[,3:(row_count + 2)]/prior_auto_TLN[,3:(row_count + 2)]) * comp_prior_auto_TLN[,3:(row_count + 2)]
    prior_auto_TLN_st_final <- prior_auto_TLN_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_auto_TLN_lt_final <- template_prior
    prior_auto_TLN_lt_final[,3:(row_count + 2)] <- (prior_auto_TLN_lt[,3:(row_count + 2)]/prior_auto_TLN[,3:(row_count + 2)]) * comp_prior_auto_TLN[,3:(row_count + 2)]
    prior_auto_TLN_lt_final <- prior_auto_TLN_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    #Auto-TP
    prior_auto_TP_st <- Triangle_raw_function("Auto-TP",i,"ST")
    prior_auto_TP_lt <- Triangle_raw_function("Auto-TP",i,"LT")
    
    prior_auto_TP <- Triangle_total_raw_function("Auto-TP")
    
    comp_prior_auto_TP <- Triangle_prior_function("Auto-TP")
    
    prior_auto_TP_st_final <- template_prior
    prior_auto_TP_st_final[,3:(row_count + 2)] <- (prior_auto_TP_st[,3:(row_count + 2)]/prior_auto_TP[,3:(row_count + 2)]) * comp_prior_auto_TP[,3:(row_count + 2)]
    prior_auto_TP_st_final <- prior_auto_TP_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_auto_TP_lt_final <- template_prior
    prior_auto_TP_lt_final[,3:(row_count + 2)] <- (prior_auto_TP_lt[,3:(row_count + 2)]/prior_auto_TP[,3:(row_count + 2)]) * comp_prior_auto_TP[,3:(row_count + 2)]
    prior_auto_TP_lt_final <- prior_auto_TP_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    #Engineering
    prior_eng_st <- Triangle_raw_function("Engineering",i,"ST")
    prior_eng_lt <- Triangle_raw_function("Engineering",i,"LT")
    
    prior_eng <- Triangle_total_raw_function("Engineering")
    
    comp_prior_eng <- Triangle_prior_function("Engineering")
    
    prior_eng_st_final <- template_prior
    prior_eng_st_final[,3:(row_count + 2)] <- (prior_eng_st[,3:(row_count + 2)]/prior_eng[,3:(row_count + 2)]) * comp_prior_eng[,3:(row_count + 2)]
    prior_eng_st_final <- prior_eng_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_eng_lt_final <- template_prior
    prior_eng_lt_final[,3:(row_count + 2)] <- (prior_eng_lt[,3:(row_count + 2)]/prior_eng[,3:(row_count + 2)]) * comp_prior_eng[,3:(row_count + 2)]
    prior_eng_lt_final <- prior_eng_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    #Fire
    prior_fire_st <- template_prior
    prior_fire_st[,3:(row_count + 2)] <- Triangle_raw_function("Fire",i,"ST")[,3:(row_count + 2)] + Triangle_raw_function("Various-Fire",i,"ST")[,3:(row_count + 2)]
    prior_fire_lt <- template_prior
    prior_fire_lt[,3:(row_count + 2)] <- Triangle_raw_function("Fire",i,"LT")[,3:(row_count + 2)] + Triangle_raw_function("Various-Fire",i,"LT")[,3:(row_count + 2)]
    
    prior_fire <- template_prior
    prior_fire[,3:(row_count + 2)] <- Triangle_total_raw_function("Fire")[,3:(row_count + 2)] + Triangle_total_raw_function("Various-Fire")[,3:(row_count + 2)]
    
    
    comp_prior_fire <- Triangle_prior_function("Fire")
    
    prior_fire_st_final <- template_prior
    prior_fire_st_final[,3:(row_count + 2)] <- (prior_fire_st[,3:(row_count + 2)]/prior_fire[,3:(row_count + 2)]) * comp_prior_fire[,3:(row_count + 2)]
    prior_fire_st_final <- prior_fire_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_fire_lt_final <- template_prior
    prior_fire_lt_final[,3:(row_count + 2)] <- (prior_fire_lt[,3:(row_count + 2)]/prior_fire[,3:(row_count + 2)]) * comp_prior_fire[,3:(row_count + 2)]
    prior_fire_lt_final <- prior_fire_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    #Marine
    prior_marine_st <- template_prior
    prior_marine_st[,3:(row_count + 2)] <- Triangle_raw_function("Marine-ECommerce",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Marine-Other",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Marine-Other-XL",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Marine-ECommerce-XL",i,"ST")[,3:(row_count + 2)]
    prior_marine_lt <- template_prior
    prior_marine_lt[,3:(row_count + 2)] <- Triangle_raw_function("Marine-ECommerce",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Marine-Other",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Marine-Other-XL",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Marine-ECommerce-XL",i,"LT")[,3:(row_count + 2)]
    
    prior_marine <- template_prior
    prior_marine[,3:(row_count + 2)] <- Triangle_total_raw_function("Marine-ECommerce")[,3:(row_count + 2)] +
                                        Triangle_total_raw_function("Marine-Other")[,3:(row_count + 2)] +
                                        Triangle_total_raw_function("Marine-Other-XL")[,3:(row_count + 2)] +
                                        Triangle_total_raw_function("Marine-ECommerce-XL")[,3:(row_count + 2)]
    
    comp_prior_marine <- Triangle_prior_function("Marine")
    
    prior_marine_st_final <- template_prior
    prior_marine_st_final[,3:(row_count + 2)] <- (prior_marine_st[,3:(row_count + 2)]/prior_marine[,3:(row_count + 2)]) * comp_prior_marine[,3:(row_count + 2)]
    prior_marine_st_final <- prior_marine_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_marine_lt_final <- template_prior
    prior_marine_lt_final[,3:(row_count + 2)] <- (prior_marine_lt[,3:(row_count + 2)]/prior_marine[,3:(row_count + 2)]) * comp_prior_marine[,3:(row_count + 2)]
    prior_marine_lt_final <- prior_marine_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    #Various
    prior_various_st <- template_prior
    prior_various_st[,3:(row_count + 2)] <- Triangle_raw_function("Various-Fire",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Various-Liability",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Various-PA-Retail",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Various-PA-Commercial",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Various-Travel-Retail",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Various-Travel-Commercial",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Various-Travel-XL",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Various-Trade Credit",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Various-Other",i,"ST")[,3:(row_count + 2)] +
                                           Triangle_raw_function("Various-Other-XL",i,"ST")[,3:(row_count + 2)]
    prior_various_lt <- template_prior
    prior_various_lt[,3:(row_count + 2)] <- Triangle_raw_function("Various-Fire",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Various-Liability",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Various-PA-Retail",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Various-PA-Commercial",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Various-Travel-Retail",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Various-Travel-Commercial",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Various-Travel-XL",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Various-Trade Credit",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Various-Other",i,"LT")[,3:(row_count + 2)] +
                                          Triangle_raw_function("Various-Other-XL",i,"LT")[,3:(row_count + 2)]
    
    prior_various <- template_prior
    prior_various[,3:(row_count + 2)] <- Triangle_total_raw_function("Various-Fire")[,3:(row_count + 2)] +
                                          Triangle_total_raw_function("Various-Liability")[,3:(row_count + 2)] +
                                          Triangle_total_raw_function("Various-PA-Retail")[,3:(row_count + 2)] +
                                          Triangle_total_raw_function("Various-PA-Commercial")[,3:(row_count + 2)] +
                                          Triangle_total_raw_function("Various-Travel-Retail")[,3:(row_count + 2)] +
                                          Triangle_total_raw_function("Various-Travel-Commercial")[,3:(row_count + 2)] +
                                          Triangle_total_raw_function("Various-Travel-XL")[,3:(row_count + 2)] +
                                          Triangle_total_raw_function("Various-Trade Credit")[,3:(row_count + 2)] +
                                          Triangle_total_raw_function("Various-Other")[,3:(row_count + 2)] +
                                          Triangle_total_raw_function("Various-Other-XL")[,3:(row_count + 2)]
    
    comp_prior_various <- Triangle_prior_function("Various")
    
    prior_various_st_final <- template_prior
    prior_various_st_final[,3:(row_count + 2)] <- (prior_various_st[,3:(row_count + 2)]/prior_various[,3:(row_count + 2)]) * comp_prior_various[,3:(row_count + 2)]
    prior_various_st_final <- prior_various_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_various_lt_final <- template_prior
    prior_various_lt_final[,3:(row_count + 2)] <- (prior_various_lt[,3:(row_count + 2)]/prior_various[,3:(row_count + 2)]) * comp_prior_various[,3:(row_count + 2)]
    prior_various_lt_final <- prior_various_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    #PA
    prior_pa_st <- template_prior
    prior_pa_st[,3:(row_count + 2)] <- Triangle_raw_function("PA-Retail",i,"ST")[,3:(row_count + 2)] +
                                      Triangle_raw_function("PA-Commercial",i,"ST")[,3:(row_count + 2)] +
                                      Triangle_raw_function("PA-Travel-Retail",i,"ST")[,3:(row_count + 2)] +
                                      Triangle_raw_function("PA-Travel-Commercial",i,"ST")[,3:(row_count + 2)] +
                                      Triangle_raw_function("PA-XL",i,"ST")[,3:(row_count + 2)]
    prior_pa_lt <- template_prior
    prior_pa_lt[,3:(row_count + 2)] <- Triangle_raw_function("PA-Retail",i,"LT")[,3:(row_count + 2)] +
                                      Triangle_raw_function("PA-Commercial",i,"LT")[,3:(row_count + 2)] +
                                      Triangle_raw_function("PA-Travel-Retail",i,"LT")[,3:(row_count + 2)] +
                                      Triangle_raw_function("PA-Travel-Commercial",i,"LT")[,3:(row_count + 2)] +
                                      Triangle_raw_function("PA-XL",i,"LT")[,3:(row_count + 2)]
    
    prior_pa <- template_prior
    prior_pa[,3:(row_count + 2)] <- Triangle_total_raw_function("PA-Retail")[,3:(row_count + 2)] +
                                    Triangle_total_raw_function("PA-Commercial")[,3:(row_count + 2)] +
                                    Triangle_total_raw_function("PA-Travel-Retail")[,3:(row_count + 2)] +
                                    Triangle_total_raw_function("PA-Travel-Commercial")[,3:(row_count + 2)] +
                                    Triangle_total_raw_function("PA-XL")[,3:(row_count + 2)]
    
    comp_prior_pa <- Triangle_prior_function("PA")
    
    prior_pa_st_final <- template_prior
    prior_pa_st_final[,3:(row_count + 2)] <- (prior_pa_st[,3:(row_count + 2)]/prior_pa[,3:(row_count + 2)]) * comp_prior_pa[,3:(row_count + 2)]
    prior_pa_st_final <- prior_pa_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_pa_lt_final <- template_prior
    prior_pa_lt_final[,3:(row_count + 2)] <- (prior_pa_lt[,3:(row_count + 2)]/prior_pa[,3:(row_count + 2)]) * comp_prior_pa[,3:(row_count + 2)]
    prior_pa_lt_final <- prior_pa_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    #Liability
    prior_liab_st <- Triangle_raw_function("Liability",i,"ST")
    prior_liab_lt <- Triangle_raw_function("Liability",i,"LT")
    
    prior_liab <- Triangle_total_raw_function("Liability")
    
    comp_prior_liab <- Triangle_prior_function("Liability")
    
    prior_liab_st_final <- template_prior
    prior_liab_st_final[,3:(row_count + 2)] <- (prior_liab_st[,3:(row_count + 2)]/prior_liab[,3:(row_count + 2)]) * comp_prior_liab[,3:(row_count + 2)]
    prior_liab_st_final <- prior_liab_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_liab_lt_final <- template_prior
    prior_liab_lt_final[,3:(row_count + 2)] <- (prior_liab_lt[,3:(row_count + 2)]/prior_liab[,3:(row_count + 2)]) * comp_prior_liab[,3:(row_count + 2)]
    prior_liab_lt_final <- prior_liab_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    #Health
    prior_health_st <- Triangle_raw_function("Health",i,"ST")
    prior_health_lt <- Triangle_raw_function("Health",i,"LT")
    
    prior_health <- Triangle_total_raw_function("Health")
    
    comp_prior_health <- Triangle_prior_function("Health")
    
    prior_health_st_final <- template_prior
    prior_health_st_final[,3:(row_count + 2)] <- (prior_health_st[,3:(row_count + 2)]/prior_health[,3:(row_count + 2)]) * comp_prior_health[,3:(row_count + 2)]
    prior_health_st_final <- prior_health_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_health_lt_final <- template_prior
    prior_health_lt_final[,3:(row_count + 2)] <- (prior_health_lt[,3:(row_count + 2)]/prior_health[,3:(row_count + 2)]) * comp_prior_health[,3:(row_count + 2)]
    prior_health_lt_final <- prior_health_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    #Various-ECommerce
    prior_vare_st <- template_prior
    prior_vare_st[,3:(row_count + 2)] <- Triangle_raw_function("Various-ECommerce",i,"ST")[,3:(row_count + 2)] +
                                        Triangle_raw_function("Various-ECommerce-XL",i,"ST")[,3:(row_count + 2)]
    prior_vare_lt <- template_prior
    prior_vare_lt[,3:(row_count + 2)] <- Triangle_raw_function("Various-ECommerce",i,"LT")[,3:(row_count + 2)] +
                                        Triangle_raw_function("Various-ECommerce-XL",i,"LT")[,3:(row_count + 2)]
    
    prior_vare <- template_prior
    prior_vare[,3:(row_count + 2)] <- Triangle_total_raw_function("Various-ECommerce")[,3:(row_count + 2)] +
                  Triangle_total_raw_function("Various-ECommerce-XL")[,3:(row_count + 2)]
    
    comp_prior_vare <- Triangle_prior_function("Various-ECommerce")
    
    prior_vare_st_final <- template_prior
    prior_vare_st_final[,3:(row_count + 2)] <- (prior_vare_st[,3:(row_count + 2)]/prior_vare[,3:(row_count + 2)]) * comp_prior_vare[,3:(row_count + 2)]
    prior_vare_st_final <- prior_vare_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_vare_lt_final <- template_prior
    prior_vare_lt_final[,3:(row_count + 2)] <- (prior_vare_lt[,3:(row_count + 2)]/prior_vare[,3:(row_count + 2)]) * comp_prior_vare[,3:(row_count + 2)]
    prior_vare_lt_final <- prior_vare_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    #Fire-XL
    prior_firexl_st <- Triangle_raw_function("Fire-XL",i,"ST")
    prior_firexl_lt <- Triangle_raw_function("Fire-XL",i,"LT")
    
    prior_firexl <- Triangle_total_raw_function("Fire-XL")
    
    # comp_prior_firexl <- Triangle_prior_function("Fire-XL")
    comp_prior_firexl <- template_prior
    
    prior_firexl_st_final <- template_prior
    prior_firexl_st_final[,3:(row_count + 2)] <- (prior_firexl_st[,3:(row_count + 2)]/prior_firexl[,3:(row_count + 2)]) * comp_prior_firexl[,3:(row_count + 2)]
    prior_firexl_st_final <- prior_firexl_st_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    prior_firexl_lt_final <- template_prior
    prior_firexl_lt_final[,3:(row_count + 2)] <- (prior_firexl_lt[,3:(row_count + 2)]/prior_firexl[,3:(row_count + 2)]) * comp_prior_firexl[,3:(row_count + 2)]
    prior_firexl_lt_final <- prior_firexl_lt_final %>%
      mutate(across(everything(), ~ ifelse(is.nan(.), 0, ifelse(is.infinite(.),0,.))))
    
    wbp <- createWorkbook()
    for (s in c("Auto",prior_sheet$SheetName)){
      addWorksheet(wbp,s)
    }
    writeData(wbp,"Auto", prior_auto_st_final)
    writeData(wbp,"Auto-OD-PartialLoss", prior_auto_PL_st_final)
    writeData(wbp,"Auto-OD-TotalLoss-MotorCycle", prior_auto_TLM_st_final)
    writeData(wbp,"Auto-OD-TotalLoss-NonMotorCycle", prior_auto_TLN_st_final)
    writeData(wbp,"Auto-TP", prior_auto_TP_st_final)
    writeData(wbp,"Engineering", prior_eng_st_final)
    writeData(wbp,"Fire", prior_fire_st_final)
    writeData(wbp,"Marine", prior_marine_st_final)
    writeData(wbp,"PA", prior_pa_st_final)
    writeData(wbp,"Liability", prior_liab_st_final)
    writeData(wbp,"Various", prior_various_st_final)
    writeData(wbp,"Various-ECommerce", prior_vare_st_final)
    writeData(wbp,"Fire-XL", prior_firexl_st_final)
    writeData(wbp,"Health", prior_health_st_final)
    saveWorkbook(wbp, paste0("Workbook/PC ",i,"-ST/PaidClaim-Net.Prior.xlsx"),overwrite = TRUE)
    
    wbp <- createWorkbook()
    for (s in c("Auto",prior_sheet$SheetName)){
      addWorksheet(wbp,s)
    }
    writeData(wbp,"Auto", prior_auto_lt_final)
    writeData(wbp,"Auto-OD-PartialLoss", prior_auto_PL_lt_final)
    writeData(wbp,"Auto-OD-TotalLoss-MotorCycle", prior_auto_TLM_lt_final)
    writeData(wbp,"Auto-OD-TotalLoss-NonMotorCycle", prior_auto_TLN_lt_final)
    writeData(wbp,"Auto-TP", prior_auto_TP_lt_final)
    writeData(wbp,"Engineering", prior_eng_lt_final)
    writeData(wbp,"Fire", prior_fire_lt_final)
    writeData(wbp,"Marine", prior_marine_lt_final)
    writeData(wbp,"PA", prior_pa_lt_final)
    writeData(wbp,"Liability", prior_liab_lt_final)
    writeData(wbp,"Various", prior_various_lt_final)
    writeData(wbp,"Various-ECommerce", prior_vare_lt_final)
    writeData(wbp,"Fire-XL", prior_firexl_lt_final)
    writeData(wbp,"Health", prior_health_lt_final)
    saveWorkbook(wbp, paste0("Workbook/PC ",i,"-LT/PaidClaim-Net.Prior.xlsx"),overwrite = TRUE)
    
  # }else{
  #   
  #   wbp <-createWorkbook()
  #   for (s in c("Auto",prior_sheet$SheetName)){
  #     addWorksheet(wbp, s)
  #     writeData(wbp, s, template_prior)
  #   }
  #   saveWorkbook(wbp, paste0("Workbook/PC ",i,"-ST/PaidClaim-Net.Prior.xlsx"),overwrite = TRUE)
  #   saveWorkbook(wbp, paste0("Workbook/PC ",i,"-LT/PaidClaim-Net.Prior.xlsx"),overwrite = TRUE)
  # 
  # }
  
  
} #COHORT



