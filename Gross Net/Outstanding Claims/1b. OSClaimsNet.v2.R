
library(readxl)
library(DBI)
library(odbc)
library(dplyr)
library(lubridate)
library(tidyr)
library(openxlsx)
library(writexl)


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

cohort_start <- 2009
cohort_end <- 2024

row_count <- (val_year-1-2009)*4+1+val_quarter

sql_data <- read.csv("SQLData.csv")
#sql_data2 <- read_excel("SQLData.xlsx")
sql_data$Gross <- as.numeric(sql_data$Gross)
sql_data$Net <- as.numeric(sql_data$Net)
sql_data <- na.omit(sql_data)

sheet_df <- read_excel("Mapping_LoB_Sheet.v1.xlsx")

sql_data <- sql_data[sql_data$Cohort != "NULL",]
sql_data <- sql_data[sql_data$IsShortTerm != "NULL",]

result <- sql_data %>%
  group_by(SheetName, ReportingYear, ReportingMonth, LossYear, LossMonth, Cohort, IsShortTerm) %>%
  summarize(
    Gross = sum(Gross),
    Net = sum(Net)
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


OS_function <- function(cohort, term, sheet, tahun, kuarter){   #fungsi term, cohort, sheet, tahun, kuarter
  
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

for ( i in cohort_start:cohort_end){ #COHORT
  
  for (j in 0:1){ #TERM
    
    st_or_lt <- ifelse(j == 1, "-ST","-LT")
    dir_name <- paste0("Workbook/OS ",i,st_or_lt)
    dir.create(dir_name)
    
    wb <- createWorkbook()
    
    for (k in sheet_df$Sheet){ #SHEET 
      
      OS_data <- data.frame() #NEW BLANK temporary every LoB
      sheet_name <- sheet_df$Compiled[sheet_df$Sheet == k]
      addWorksheet(wb, sheet_name)
      
      for (l in 2009:val_year){ #ACCIDENT YEAR
        
        for (m in 1:4){ #ACCIDENT QUARTER
          print(paste0(i,st_or_lt,"-",k,"-",l,"-",m))
          if(l == val_year & m > val_quarter){break}
          else if(l == 2009 & m %in% c(1,2,3)){}
          else{
          result_row <- OS_function(i,j,k,l,m)
          OS_data <- rbind(OS_data, result_row)
          
          }
          
        } #ACCIDENT QUARTER
        
        
      } #ACCIDENT YEAR
      
      OS_data[,3:ncol(OS_data)] <- lapply(OS_data[,3:ncol(OS_data)], as.numeric)
      colnames(OS_data) <- c("Year","Kuarter",seq(1,nrow(OS_data)))
      writeData(wb, sheet_name, OS_data)
      #addStyle(wb, sheet = sheet_name, style = createStyle(numFmt = "######.######"), rows = 1:nrow(OS_data)+1, cols = 1:ncol(OS_data), gridExpand = TRUE)
      
    } #SHEET 
    
    saveWorkbook(wb, paste0(dir_name,"/OSClaim-Net.Raw.xlsx"), overwrite = TRUE)
    
  } #TERM
  
  
} #COHORT



