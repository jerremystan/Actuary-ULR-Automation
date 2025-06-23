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
                 Trusted_Connection = "Yes")
'
# Check the connection
#if (!dbIsValid(con)) {
#  stop("Failed to connect to database.")
#}

# Define the SQL query to retrieve data
#query <- "SELECT TOP 10 * from dbo.PolicyTransactions"

# Execute the query and retrieve the data into a data frame
#df <- dbGetQuery(con, query)


#GANTI COHORT DISINI

cohort_year_start <- 2023
cohort_year_end <- 2025

#cohort_start <- paste0(cohort_year,"-01-01")
#cohort_end <- paste0(cohort_year,"-12-31")


#GANTI VALUATION DATE DISINI
val_year <- 2025
val_quarter <- 1

#CALL FUNCTION SQL

if(val_quarter == 1){
  ep_end <- paste0(val_year,"-03-31")
} else if(val_quarter == 2){
  ep_end <- paste0(val_year,"-06-30")
} else if(val_quarter == 3){
  ep_end <- paste0(val_year,"-09-30")
} else if(val_quarter == 4){
  ep_end <- paste0(val_year,"-12-31")
}


'query_sql <- paste0("")



writeLines(query_sql, "Query_SQL_check.txt")

sql_data <- dbGetQuery(con, query_sql)'

sql_data <- read_excel("EP_Data.xlsx")

#write.csv(sql_data, "SQL_Data.v1.csv",row.names = FALSE)

sql_data_temp <- sql_data
#sql_data_temp[sql_data_temp$Cohort=="NULL",]

sql_data[is.na(sql_data)] <- 0
sql_data <- sql_data[sql_data$Cohort != "NULL",]
sql_data <- sql_data[sql_data$IsShortTerm != "NULL",]

sql_data$Net <- 0

# sql_data <- sql_data[,-6:-7]

sql_data <- sql_data %>% group_by(LineOfBusiness, Tahun, Kuarter) %>% summarize(Reins = sum(Reins),Net = sum(Net))

#sql_data<-sql_data_temp

#wb <- createWorkbook()



EP_function <- function(tahun,quarter){
  
  #sql_data <- read.csv("EP_SQL_Template.csv")
  
  #sql_data <- sql_data[-3,]
  
  new_row <- data.frame(`Year` = tahun,
                        `Quarter` = quarter,
                        `Reins_Auto` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "AUTO" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0 ,sql_data$Reins[sql_data$LineOfBusiness == "AUTO" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_Marine-Ecommerce` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "MARINE-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "MARINE-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter], 0),
                        `Reins_Marine-Others` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "MARINE-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) >0 ,sql_data$Reins[sql_data$LineOfBusiness == "MARINE-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter], 0),
                        `Reins_Marine-Others-XL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "MARINE-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "MARINE-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter], 0),
                        `Reins_Fire` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0 ,sql_data$Reins[sql_data$LineOfBusiness == "FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter], 0),
                        `Reins_Engineering` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "ENGINEERING" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "ENGINEERING" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_Liability` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_PA-RETAIL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_PA-COMMERCIAL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_PA-TRAVEL-RETAIL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "PA-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "PA-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_PA-TRAVEL-COMMERCIAL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "PA-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "PA-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_PA-XL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "PA-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "PA-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-FIRE` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-LIABILITY` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-PA-RETAIL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-PA-COMMERCIAL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-TRADE CREDIT` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-TRADE CREDIT" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0, sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-TRADE CREDIT" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-TRAVEL-RETAIL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-TRAVEL-COMMERCIAL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-TRAVEL-XL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-OTHERS` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-OTHERS-XL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_Auto` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "AUTO" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "AUTO" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_Marine-Ecommerce` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "MARINE-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "MARINE-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_Marine-Others` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "MARINE-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "MARINE-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_Marine-Others-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "MARINE-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "MARINE-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_Fire` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_Engineering` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "ENGINEERING" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "ENGINEERING" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_Liability` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_PA-RETAIL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_PA-COMMERCIAL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_PA-TRAVEL-RETAIL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "PA-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "PA-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_PA-TRAVEL-COMMERCIAL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "PA-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "PA-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_PA-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "PA-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "PA-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-FIRE` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-LIABILITY` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-PA-RETAIL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-PA-COMMERCIAL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-TRADE CREDIT` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRADE CREDIT" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRADE CREDIT" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-TRAVEL-RETAIL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-TRAVEL-COMMERCIAL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-TRAVEL-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-OTHERS` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-OTHERS-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_HEALTH` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "HEALTH" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "HEALTH" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_HEALTH` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "HEALTH" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "HEALTH" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-ECOMMERCE` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-ECOMMERCE` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_MARINE-ECOMMERCE-XL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "MARINE-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "MARINE-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_MARINE-ECOMMERCE-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "MARINE-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "MARINE-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_VARIOUS-ECOMMERCE-XL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_VARIOUS-ECOMMERCE-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_AUTO-XL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "AUTO-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "AUTO-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_AUTO-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "AUTO-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "AUTO-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_ENGINEERING-XL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "ENGINEERING-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "ENGINEERING-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_ENGINEERING-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "ENGINEERING-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "ENGINEERING-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_FIRE-XL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "FIRE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "FIRE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_FIRE-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "FIRE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "FIRE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_TRADE CREDIT-XL` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "TRADE CREDIT-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "TRADE CREDIT-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_TRADE CREDIT-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "TRADE CREDIT-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "TRADE CREDIT-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Reins_LBLW` = ifelse(length(sql_data$Reins[sql_data$LineOfBusiness == "LBLW" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Reins[sql_data$LineOfBusiness == "LBLW" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0),
                        `Net_LBLW` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "LBLW" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter]) > 0,sql_data$Net[sql_data$LineOfBusiness == "LBLW" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter],0)
  )
  
  return(new_row)
  
}


EP_data_ST <- data.frame()

for (i in 2010:val_year){
  
  for( j in 1:4){
    
    if(i == val_year & j > val_quarter){break}
    
    result_ST <- EP_function(i,j)
    EP_data_ST <- rbind(EP_data_ST,result_ST)
    
  }
  
}

#Short Term
nama_kolom <- colnames(EP_data_ST)
Reins_or_net <- grepl("Reins|Net", nama_kolom)
nama_Reins_net <- gsub("_.*","",nama_kolom)
nama_Reins_net[1:2] <- ""
new_nama <- ifelse(Reins_or_net, nama_kolom, "")
new_nama_1 <- gsub("Reins_","",new_nama)
new_nama_1 <- gsub("Net_","",new_nama_1)
new_nama_1[1:2] <- c("Year","Quarter")
new_nama_1 <- gsub("\\.","-",new_nama_1)
new_nama_1 <- gsub("TRADE-CREDIT","TRADE CREDIT",new_nama_1)

to_excel <- rbind(new_nama_1,EP_data_ST)
colnames(to_excel) <- nama_Reins_net

dir_name <- paste0("Workbook")
dir.create(dir_name)

#write.csv(to_excel, paste0(dir_name,"/Earned Premium.Raw.All.Cohort",sheet_name,".csv"), row.names = FALSE)
#write.xlsx(list("EP" = to_excel),file = paste0(dir_name,"/Earned Premium.Raw. ",".xlsx"))

wb <- createWorkbook()
baris_1 <- data.frame(matrix(ncol = length(nama_Reins_net),nrow=0))
names(baris_1) <- nama_Reins_net
EP_data_ST_xl <- EP_data_ST
colnames(EP_data_ST_xl) <- new_nama_1

addWorksheet(wb, "EP")
writeData(wb, "EP", baris_1)
writeData(wb, "EP", EP_data_ST_xl, startCol = 1, startRow = 2)
saveWorkbook(wb, paste0(dir_name,"/Earned Premium.Reins.Raw.All.Cohort.xlsx"), overwrite = TRUE)
saveWorkbook(wb, paste0(dir_name,"/Earned Premium.Reins.Total.Raw.xlsx"), overwrite = TRUE)

#addWorksheet(wb, sheet_name)
#writeData(wb, sheet_name, to_excel)


#addWorksheet(wb, sheet_name)
#writeData(wb, sheet_name, to_excel)

#addStyle(wb, sheet = as.character(c), style = createStyle(numFmt = "GENERAL"), rows = 1:1, cols = 1:ncol(to_excel), gridExpand = TRUE)
#addStyle(wb, sheet = as.character(c), style = createStyle(numFmt = "NUMERIC"), rows = 2:(nrow(to_excel)+1), cols = 1:ncol(to_excel), gridExpand = TRUE)



#saveWorkbook(wb, "by Cohort/EarnedPremiumRaw_All_Cohort.v8.xlsx", overwrite = TRUE)

'dbDisconnect(con)'

#test_sql_data <- sql_data[sql_data$Cohort == 2014 & sql_data$Tahun == 2014 & sql_data$Kuarter == 2,]
#sql_data[sql_data$LineOfBusiness == "FIRE" & sql_data$Tahun == 2023 & sql_data$Kuarter == 3 & sql_data$Cohort == 2023,]

