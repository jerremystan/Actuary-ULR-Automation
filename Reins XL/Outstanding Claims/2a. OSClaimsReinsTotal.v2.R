
library(readxl)
library(DBI)
library(odbc)
library(dplyr)
library(lubridate)
library(tidyr)
library(openxlsx)
library(writexl)
library(stringr)


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

val_year <- 2025
val_quarter <- 1

#GANTI COHORT!!!
# 
# cohort_start <- 2020
# cohort_end <- 2024

row_count <- (val_year-1-2009)*4+1+val_quarter

sql_data <- read.csv("SQLData.csv")
#sql_data2 <- read_excel("SQLData.xlsx")
sql_data$Reins <- as.numeric(sql_data$Reins)
#sql_data$Net <- as.numeric(sql_data$Net)
sql_data <- na.omit(sql_data)

sheet_df <- read_excel("Mapping_LoB_Sheet.v1.xlsx")

sql_data <- sql_data[sql_data$Cohort != "NULL",]
sql_data <- sql_data[sql_data$IsShortTerm != "NULL",]


result <- sql_data %>%
          group_by(SheetName, ReportingYear, ReportingMonth, LossYear, LossMonth) %>%
          summarize(
            Reins = sum(Reins)
            #,Net = sum(Net)
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




OS_function <- function(sheet, tahun, kuarter){   #fungsi term, cohort, sheet, tahun, kuarter
  
  filter_master <- per_master$masterper[per_master$masterper >= paste0(tahun,kuarter)]
  
  df_index <- data.frame(LossPer = rep(paste0(tahun,kuarter),length(filter_master)),
                         RPTPer = filter_master,
                         ColIndex = seq(1,length(filter_master)))
  df_index$RPTYear <- substr(df_index$RPTPer,1,4)
  df_index$RPTQuar <- substr(df_index$RPTPer,5,5)
  
  num_col <- lapply(1:row_count, function(x){
    
    ifelse(length(sql_data$Reins[sql_data$LossYear == tahun & sql_data$LossMonth == kuarter & sql_data$ReportingYear == df_index$RPTYear[df_index$ColIndex == x] & sql_data$ReportingMonth == df_index$RPTQuar[df_index$ColIndex == x] & sql_data$SheetName == sheet]) > 0, sql_data$Reins[sql_data$LossYear == tahun & sql_data$LossMonth == kuarter & sql_data$ReportingYear == df_index$RPTYear[df_index$ColIndex == x] & sql_data$ReportingMonth == df_index$RPTQuar[df_index$ColIndex == x] & sql_data$SheetName == sheet],"")
    
  })
  
  new_row <- data.frame(Year = ifelse(tahun == 2009, "Prior", tahun),
                        Quarter = ifelse(tahun == 2009 & kuarter == 4,"Prior", kuarter),
                        setNames(as.data.frame(num_col), as.character(1:row_count)))
  
  return(new_row)
  
}


wb <- createWorkbook()

for (k in sheet_df$Sheet){ #SHEET
  
  OS_data <- data.frame()
  sheet_name <- sheet_df$Compiled[sheet_df$Sheet == k]
  addWorksheet(wb, sheet_name)
  
      for (l in 2009:val_year){ #ACCIDENT YEAR
        
        for (m in 1:4){ #ACCIDENT QUARTER
          
          if(l == val_year & m > val_quarter){break}
          else if(l == 2009 & m %in% c(1,2,3)){}
          else{
          result_row <- OS_function(k,l,m)
          OS_data <- rbind(OS_data, result_row)
          
          }
          
        } #ACCIDENT QUARTER
        
        
      } #ACCIDENT YEAR
  
  OS_data[,3:ncol(OS_data)] <- lapply(OS_data[,3:ncol(OS_data)], as.numeric)
  colnames(OS_data) <- c("Year","Kuarter",seq(1,nrow(OS_data)))
  writeData(wb, sheet_name, OS_data)    
  
} #SHEET

saveWorkbook(wb, "Workbook/OSClaim-Reins.Total.Raw.xlsx", overwrite = TRUE)

