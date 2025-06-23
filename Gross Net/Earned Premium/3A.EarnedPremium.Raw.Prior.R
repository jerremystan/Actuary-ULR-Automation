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

cohort_year_start <- 2020
cohort_year_end <- 2024

#cohort_start <- paste0(cohort_year,"-01-01")
#cohort_end <- paste0(cohort_year,"-12-31")

valuation <- "202412"

#GANTI VALUATION DATE DISINI
val_year <- 2024
val_quarter <- 4

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

sql_data <- na.omit(sql_data)
sql_data <- sql_data[sql_data$Cohort != "NULL",]
sql_data <- sql_data[sql_data$IsShortTerm != "NULL",]

result <- sql_data %>%
  group_by(LineOfBusiness, Tahun, Kuarter, Cohort, IsShortTerm) %>%
  summarize(
    Gross = sum(Gross),
    Net = sum(Net)
  )

sql_data <- result

#sql_data<-sql_data_temp

#wb <- createWorkbook()

#Prior EP Data
prior_ep <- read_excel("Workbook/Earned Premium.Prior.xlsx", skip = 1, sheet = "EP")
vector_a <- unlist(read_excel("Workbook/Earned Premium.Prior.xlsx", range = "A2:R2", col_names = FALSE))
vector_a[1:2] <- c("Year", "Quarter")
vector_a[3:9] <- paste0("Gross_",vector_a[3:9])
vector_a[10:16] <- paste0("Net_",vector_a[10:16])
vector_a[17] <- paste0("Gross_",vector_a[17])
vector_a[18] <- paste0("Net_",vector_a[18])
colnames(prior_ep) <- vector_a
prior_ep[is.na(prior_ep)] <- 0

vector_gn_prior <- c("","",rep("Gross",7),rep("Net",7),"Gross","Net")


EP_function <- function(tahun,quarter,term){
  
  #sql_data <- read.csv("EP_SQL_Template.csv")
  
  #sql_data <- sql_data[-3,]
  
  new_row <- data.frame(`Year` = tahun,
                        `Quarter` = quarter,
                        `Gross_Auto` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "AUTO" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0 ,sql_data$Gross[sql_data$LineOfBusiness == "AUTO" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_Marine-Ecommerce` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "MARINE-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "MARINE-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term], 0),
                        `Gross_Marine-Others` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "MARINE-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) >0 ,sql_data$Gross[sql_data$LineOfBusiness == "MARINE-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term], 0),
                        `Gross_Marine-Others-XL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "MARINE-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "MARINE-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term], 0),
                        `Gross_Fire` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0 ,sql_data$Gross[sql_data$LineOfBusiness == "FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term], 0),
                        `Gross_Engineering` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "ENGINEERING" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "ENGINEERING" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_Liability` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_PA-RETAIL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_PA-COMMERCIAL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_PA-TRAVEL-RETAIL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "PA-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "PA-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_PA-TRAVEL-COMMERCIAL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "PA-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "PA-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_PA-XL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "PA-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "PA-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-FIRE` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-LIABILITY` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-PA-RETAIL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-PA-COMMERCIAL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-TRADE CREDIT` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-TRADE CREDIT" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0, sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-TRADE CREDIT" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-TRAVEL-RETAIL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-TRAVEL-COMMERCIAL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-TRAVEL-XL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-OTHERS` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-OTHERS-XL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_Auto` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "AUTO" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "AUTO" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_Marine-Ecommerce` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "MARINE-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "MARINE-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_Marine-Others` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "MARINE-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "MARINE-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_Marine-Others-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "MARINE-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "MARINE-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_Fire` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_Engineering` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "ENGINEERING" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "ENGINEERING" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_Liability` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_PA-RETAIL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_PA-COMMERCIAL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_PA-TRAVEL-RETAIL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "PA-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "PA-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_PA-TRAVEL-COMMERCIAL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "PA-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "PA-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_PA-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "PA-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "PA-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-FIRE` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-FIRE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-LIABILITY` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-LIABILITY" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-PA-RETAIL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-PA-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-PA-COMMERCIAL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-PA-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-TRADE CREDIT` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRADE CREDIT" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRADE CREDIT" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-TRAVEL-RETAIL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-RETAIL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-TRAVEL-COMMERCIAL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-COMMERCIAL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-TRAVEL-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-TRAVEL-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-OTHERS` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-OTHERS" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-OTHERS-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-OTHERS-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_HEALTH` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "HEALTH" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "HEALTH" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_HEALTH` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "HEALTH" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "HEALTH" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-ECOMMERCE` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-ECOMMERCE` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_MARINE-ECOMMERCE-XL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "MARINE-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "MARINE-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_MARINE-ECOMMERCE-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "MARINE-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "MARINE-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_VARIOUS-ECOMMERCE-XL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_VARIOUS-ECOMMERCE-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "VARIOUS-ECOMMERCE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_AUTO-XL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "AUTO-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "AUTO-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_AUTO-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "AUTO-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "AUTO-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_ENGINEERING-XL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "ENGINEERING-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "ENGINEERING-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_ENGINEERING-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "ENGINEERING-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "ENGINEERING-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_FIRE-XL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "FIRE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "FIRE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_FIRE-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "FIRE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "FIRE-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_TRADE CREDIT-XL` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "TRADE CREDIT-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "TRADE CREDIT-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_TRADE CREDIT-XL` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "TRADE CREDIT-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "TRADE CREDIT-XL" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Gross_LBLW` = ifelse(length(sql_data$Gross[sql_data$LineOfBusiness == "LBLW" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Gross[sql_data$LineOfBusiness == "LBLW" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0),
                        `Net_LBLW` = ifelse(length(sql_data$Net[sql_data$LineOfBusiness == "LBLW" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term]) > 0,sql_data$Net[sql_data$LineOfBusiness == "LBLW" & sql_data$Tahun == tahun & sql_data$Kuarter == quarter & sql_data$Cohort == c & sql_data$IsShortTerm == term],0)
                        
  )
  
  return(new_row)
  
}


function_prior <- function(tahun, quarter, term){
  
  EP_function()
  
  
}

EP_prior_STLT <- read_excel("Workbook/Earned Premium.Total.Prior.xlsx", skip = 1)

for (c in cohort_year_start:cohort_year_end){
  
  EP_data_ST <- data.frame()
  EP_data_LT <- data.frame()
  
  for (i in 2010:val_year){
    
    for( j in 1:4){
      
      if(i == val_year & j > val_quarter){break}
      
      result_ST <- EP_function(i,j,1)
      result_LT <- EP_function(i,j,0)
      EP_data_ST <- rbind(EP_data_ST,result_ST)
      EP_data_LT <- rbind(EP_data_LT,result_LT)
      
    }
    
  }
  
  
  #Short Term
  
  #EP RAW
  nama_kolom <- colnames(EP_data_ST)
  gross_or_net <- grepl("Gross|Net", nama_kolom)
  nama_gross_net <- gsub("_.*","",nama_kolom)
  nama_gross_net[1:2] <- ""
  new_nama <- ifelse(gross_or_net, nama_kolom, "")
  new_nama_1 <- gsub("Gross_","",new_nama)
  new_nama_1 <- gsub("Net_","",new_nama_1)
  new_nama_1[1:2] <- c("Year","Quarter")
  new_nama_1 <- gsub("\\.","-",new_nama_1)
  new_nama_1 <- gsub("TRADE-CREDIT","TRADE CREDIT",new_nama_1)
  
  new_nama_2 <- gsub("\\.","-",colnames(EP_data_ST))
  colnames(EP_data_ST) <- new_nama_2
  
  to_excel <- rbind(new_nama_1,EP_data_ST)
  colnames(to_excel) <- nama_gross_net
  
  st_or_lt <- substr(as.character(substitute(EP_data_ST)),9,10)
  sheet_name <- paste0(as.character(c),"-",st_or_lt)
  
  dir_name <- paste0("Workbook/EP ",sheet_name)
  dir.create(dir_name)
  
  write.csv(to_excel, paste0(dir_name,"/Earned Premium.Raw. ",sheet_name,".csv"), row.names = FALSE)
  #write.xlsx(list("EP" = to_excel),file = paste0(dir_name,"/Earned Premium.Raw. ",".xlsx"))
  
  wb <- createWorkbook()
  baris_1 <- data.frame(matrix(ncol = length(nama_gross_net),nrow=0))
  names(baris_1) <- nama_gross_net
  EP_data_ST_xl <- EP_data_ST
  colnames(EP_data_ST_xl) <- new_nama_1
  
  addWorksheet(wb, "EP")
  writeData(wb, "EP", baris_1)
  writeData(wb, "EP", EP_data_ST_xl, startCol = 1, startRow = 2)
  saveWorkbook(wb, paste0(dir_name,"/Earned Premium.Raw",".xlsx"), overwrite = TRUE)
  
  #addWorksheet(wb, sheet_name)
  #writeData(wb, sheet_name, to_excel)
  
  
  
  
  #Long Term
  nama_kolom <- colnames(EP_data_LT)
  gross_or_net <- grepl("Gross|Net", nama_kolom)
  nama_gross_net <- gsub("_.*","",nama_kolom)
  nama_gross_net[1:2] <- ""
  new_nama <- ifelse(gross_or_net, nama_kolom, "")
  new_nama_1 <- gsub("Gross_","",new_nama)
  new_nama_1 <- gsub("Net_","",new_nama_1)
  new_nama_1[1:2] <- c("Year","Quarter")
  new_nama_1 <- gsub("\\.","-",new_nama_1)
  new_nama_1 <- gsub("TRADE-CREDIT","TRADE CREDIT",new_nama_1)
  
  new_nama_2 <- gsub("\\.","-",colnames(EP_data_LT))
  colnames(EP_data_LT) <- new_nama_2
  
  to_excel <- rbind(new_nama_1,EP_data_LT)
  colnames(to_excel) <- nama_gross_net
  
  st_or_lt <- substr(as.character(substitute(EP_data_LT)),9,10)
  sheet_name <- paste0(as.character(c),"-",st_or_lt)
  
  dir_name <- paste0("Workbook/EP ",sheet_name)
  dir.create(dir_name)
  
  write.csv(to_excel, paste0(dir_name,"/Earned Premium.Raw",sheet_name,".csv"), row.names = FALSE)
  #write.xlsx(list("EP" = to_excel),file = paste0(dir_name,"/Earned Premium.Raw. ",".xlsx"))
  
  wb <- createWorkbook()
  baris_1 <- data.frame(matrix(ncol = length(nama_gross_net),nrow=0))
  names(baris_1) <- nama_gross_net
  EP_data_LT_xl <- EP_data_LT
  colnames(EP_data_LT_xl) <- new_nama_1
  
  addWorksheet(wb, "EP")
  writeData(wb, "EP", baris_1)
  writeData(wb, "EP", EP_data_LT_xl, startCol = 1, startRow = 2)
  saveWorkbook(wb, paste0(dir_name,"/Earned Premium.Raw",".xlsx"), overwrite = TRUE)
  
  # if(c == 2020){
    
    EP_data_ST_prior <- EP_data_ST %>%
      transmute(
        
        `Year` = `Year`
        ,`Quarter` = `Quarter`
        ,`Gross_Auto` = `Gross_Auto`
        ,`Gross_Marine` = `Gross_Marine-Ecommerce` + `Gross_MARINE-ECOMMERCE-XL`+ `Gross_Marine-Others` + `Gross_Marine-Others-XL`
        ,`Gross_Fire` = `Gross_Fire`
        ,`Gross_Engineering` = `Gross_Engineering`
        ,`Gross_Liability` = `Gross_Liability` + `Gross_VARIOUS-LIABILITY`
        ,`Gross_Various` = 
          `Gross_VARIOUS-FIRE`
          +`Gross_VARIOUS-LIABILITY` 
          +`Gross_VARIOUS-PA-RETAIL`
          +`Gross_VARIOUS-PA-COMMERCIAL`
          +`Gross_VARIOUS-TRADE-CREDIT`
          +`Gross_VARIOUS-TRAVEL-RETAIL`
          +`Gross_VARIOUS-TRAVEL-COMMERCIAL`
          +`Gross_VARIOUS-TRAVEL-XL`
          +`Gross_VARIOUS-OTHERS`
          +`Gross_VARIOUS-OTHERS-XL`
          +`Gross_VARIOUS-ECOMMERCE`
          +`Gross_VARIOUS-ECOMMERCE-XL`
        ,`Gross_PA` = 
          `Gross_PA-RETAIL`
          +`Gross_PA-COMMERCIAL`
          +`Gross_PA-TRAVEL-RETAIL`
          +`Gross_PA-TRAVEL-COMMERCIAL`
          +`Gross_PA-XL`
        
        ,`Net_Auto` = `Net_Auto`
        ,`Net_Marine` = `Net_Marine-Ecommerce` + `Net_MARINE-ECOMMERCE-XL`+ `Net_Marine-Others` + `Net_Marine-Others-XL`
        ,`Net_Fire` = `Net_Fire`
        ,`Net_Engineering` = `Net_Engineering`
        ,`Net_Liability` = `Net_Liability` + `Net_VARIOUS-LIABILITY`
        ,`Net_Various` = 
          `Net_VARIOUS-FIRE`
          +`Net_VARIOUS-LIABILITY` 
          +`Net_VARIOUS-PA-RETAIL`
          +`Net_VARIOUS-PA-COMMERCIAL`
          +`Net_VARIOUS-TRADE-CREDIT`
          +`Net_VARIOUS-TRAVEL-RETAIL`
          +`Net_VARIOUS-TRAVEL-COMMERCIAL`
          +`Net_VARIOUS-TRAVEL-XL`
          +`Net_VARIOUS-OTHERS`
          +`Net_VARIOUS-OTHERS-XL`
          +`Net_VARIOUS-ECOMMERCE`
          +`Net_VARIOUS-ECOMMERCE-XL`
        ,`Net_PA` = 
          `Net_PA-RETAIL`
          +`Net_PA-COMMERCIAL`
          +`Net_PA-TRAVEL-RETAIL`
          +`Net_PA-TRAVEL-COMMERCIAL`
          +`Net_PA-XL`
        
        ,`Gross_Health` = `Gross_HEALTH`
        ,`Net_Health` = `Net_HEALTH`
      )
      
    EP_data_LT_prior <- EP_data_LT %>%
      transmute(
        
        `Year` = `Year`
        ,`Quarter` = `Quarter`
        ,`Gross_Auto` = `Gross_Auto`
        ,`Gross_Marine` = `Gross_Marine-Ecommerce` + `Gross_MARINE-ECOMMERCE-XL`+ `Gross_Marine-Others` + `Gross_Marine-Others-XL`
        ,`Gross_Fire` = `Gross_Fire`
        ,`Gross_Engineering` = `Gross_Engineering`
        ,`Gross_Liability` = `Gross_Liability` + `Gross_VARIOUS-LIABILITY`
        ,`Gross_Various` = 
          `Gross_VARIOUS-FIRE`
          +`Gross_VARIOUS-LIABILITY` 
          +`Gross_VARIOUS-PA-RETAIL`
          +`Gross_VARIOUS-PA-COMMERCIAL`
          +`Gross_VARIOUS-TRADE-CREDIT`
          +`Gross_VARIOUS-TRAVEL-RETAIL`
          +`Gross_VARIOUS-TRAVEL-COMMERCIAL`
          +`Gross_VARIOUS-TRAVEL-XL`
          +`Gross_VARIOUS-OTHERS`
          +`Gross_VARIOUS-OTHERS-XL`
          +`Gross_VARIOUS-ECOMMERCE`
          +`Gross_VARIOUS-ECOMMERCE-XL`
        ,`Gross_PA` = 
          `Gross_PA-RETAIL`
          +`Gross_PA-COMMERCIAL`
          +`Gross_PA-TRAVEL-RETAIL`
          +`Gross_PA-TRAVEL-COMMERCIAL`
          +`Gross_PA-XL`
        
        ,`Net_Auto` = `Net_Auto`
        ,`Net_Marine` = `Net_Marine-Ecommerce` + `Net_MARINE-ECOMMERCE-XL`+ `Net_Marine-Others` + `Net_Marine-Others-XL`
        ,`Net_Fire` = `Net_Fire`
        ,`Net_Engineering` = `Net_Engineering`
        ,`Net_Liability` = `Net_Liability` + `Net_VARIOUS-LIABILITY`
        ,`Net_Various` = 
          `Net_VARIOUS-FIRE`
          +`Net_VARIOUS-LIABILITY` 
          +`Net_VARIOUS-PA-RETAIL`
          +`Net_VARIOUS-PA-COMMERCIAL`
          +`Net_VARIOUS-TRADE-CREDIT`
          +`Net_VARIOUS-TRAVEL-RETAIL`
          +`Net_VARIOUS-TRAVEL-COMMERCIAL`
          +`Net_VARIOUS-TRAVEL-XL`
          +`Net_VARIOUS-OTHERS`
          +`Net_VARIOUS-OTHERS-XL`
          +`Net_VARIOUS-ECOMMERCE`
          +`Net_VARIOUS-ECOMMERCE-XL`
        ,`Net_PA` = 
          `Net_PA-RETAIL`
          +`Net_PA-COMMERCIAL`
          +`Net_PA-TRAVEL-RETAIL`
          +`Net_PA-TRAVEL-COMMERCIAL`
          +`Net_PA-XL`
          
        ,`Gross_Health` = `Gross_HEALTH`
        ,`Net_Health` = `Net_HEALTH`
      )
    
    #EP_prior_STLT <- EP_data_ST_prior + EP_data_LT_prior
    EP_prior_final_ST <- (EP_data_ST_prior / EP_prior_STLT) * prior_ep
    EP_prior_final_LT <- (EP_data_LT_prior / EP_prior_STLT) * prior_ep
    
    
  # } 
  # else{
  #   
  #   EP_prior_final_ST <- prior_ep
  #   EP_prior_final_ST[,3:ncol(prior_ep)] <- 0
  #   EP_prior_final_LT <- EP_prior_final_ST
  #   
  # }
  
  EP_prior_final_ST[is.nan(as.matrix(EP_prior_final_ST))] <- 0
  EP_prior_final_LT[is.nan(as.matrix(EP_prior_final_LT))] <- 0
  
  #EP PRIOR ST
  colnames(EP_prior_final_ST) <- gsub("Gross_", "", colnames(EP_prior_final_ST))
  colnames(EP_prior_final_ST) <- gsub("Net_", "", colnames(EP_prior_final_ST))
  wb <- createWorkbook()
  addWorksheet(wb, "EP")
  writeData(wb, "EP", t(vector_gn_prior), colNames = FALSE, rowNames = FALSE)
  writeData(wb, "EP", EP_prior_final_ST, startCol = 1, startRow = 2)
  saveWorkbook(wb, paste0(paste0("Workbook/EP ",c,"-","ST"),"/Earned Premium.Prior",".xlsx"), overwrite = TRUE)
  
  #EP PRIOR LT
  colnames(EP_prior_final_LT) <- gsub("Gross_", "", colnames(EP_prior_final_LT))
  colnames(EP_prior_final_LT) <- gsub("Net_", "", colnames(EP_prior_final_LT))
  wb <- createWorkbook()
  addWorksheet(wb, "EP")
  writeData(wb, "EP", t(vector_gn_prior), colNames = FALSE, rowNames = FALSE)
  writeData(wb, "EP", EP_prior_final_LT, startCol = 1, startRow = 2)
  saveWorkbook(wb, paste0(paste0("Workbook/EP ",c,"-","LT"),"/Earned Premium.Prior",".xlsx"), overwrite = TRUE)

}

#Move Data from Master

Master_loc <- paste0("C:/Stanley/Tugas IFRS17/Reserving RCH Enhancement/", valuation, "/Gross Net/Master Gross Net/")
master_files <- c(
  #paste0(Master_loc,"Earned Premium Prior Ratio.xlsx"),
  paste0(Master_loc,"Earned Premium.Manual.xlsx"),
  #paste0(Master_loc,"Earned Premium.Prior.xlsx"),
  paste0(Master_loc,"Earned Premium.xlsx"),
  paste0(Master_loc,"EP with FS.xlsx")
  #paste0(Master_loc,"Earned Premium.Total.Raw.xlsx")
)


for (x in cohort_year_start:cohort_year_end){
  
  for (y in c("-ST","-LT")){
    
    ep_fld_loc <- paste0("Workbook/EP ", x, y)
    file.copy(master_files, ep_fld_loc, overwrite = TRUE)
    file.copy(paste0("C:/Stanley/Tugas IFRS17/Reserving RCH Enhancement/", valuation, "/Gross Net/Earned Premium/Workbook/Earned Premium.Total.Raw.xlsx"), ep_fld_loc, overwrite = TRUE)
    
  }
  
}


