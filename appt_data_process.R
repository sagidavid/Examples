library(tidyverse)
library(stringr)
library(lubridate)
library(janitor)
library(DT)
library(scales)
library(zoo)
library(readxl)

# latest_gpad_zip_url <- "https://files.digital.nhs.uk/20/40049A/Appointments_GP_Daily_Oct21.zip"
# latest_gpad_xlsx_url <- "https://files.digital.nhs.uk/98/BD8126/GP_APPT_Publication_October%20%20_2021.xlsx"
# 
# temp <- tempfile()
# download.file(latest_gpad_zip_url, 
#               temp)
# con <- unzip(temp)
# 
# appts <- bind_rows(lapply(X = con[grepl(x = con, pattern = "CCG")],
#                           FUN = read_csv))
# coverage <- read_csv(con[grepl(x = con, pattern = "COVERAGE")])
# 
# unlink(temp)
# 
# latest_date <- max(as.Date(unique(appts$Appointment_Date), tryFormats = c("%d%b%Y", "%d-%b-%y")),na.rm = T)
# earliest_date <- min(as.Date(unique(appts$Appointment_Date), "%d%b%Y"),na.rm = T)
# latest_month <- floor_date(latest_date, "month")
# 
# temp <- tempfile()
# download.file(latest_gpad_xlsx_url, 
#               temp, mode="wb")
# excel_data <- read_excel(temp,
#                          sheet = 5, skip = 10, col_types = "text", col_names = T)
# 
# working_days <- excel_data %>%
#   slice(9) %>%
#   pivot_longer(cols = 1:32, names_to = "date", values_to =  "workingDays") %>%
#   mutate(date = excel_numeric_to_date(as.numeric(date)),
#          workingDays = as.numeric(workingDays)) %>%
#   filter(!is.na(workingDays)) %>%
#   select(-date)
# 
# working_days <- tibble(working_days, 
#                        month = rev(seq.Date(from = earliest_date, to = latest_date, by = "months")))
# 
# vaccs_count <- excel_data %>%
#   slice(16) %>%
#   pivot_longer(cols = 1:32, names_to = "date", values_to =  "vaccs") %>%
#   mutate(vaccs = as.numeric(vaccs)) %>%
#   filter(!is.na(vaccs)) %>%
#   select(-date)
# 
# vaccs <- tibble(vaccs_count,
#                 month = c(seq(from = floor_date(latest_date, "month"), length.out = nrow(vaccs_count), by = "-1 month")))
# rm(excel_data)
# rm(vaccs_count)
# 
# unlink(temp)
# 
# 
# 
# national_coverage <- coverage %>%
#   mutate(month = floor_date(as.Date(Appointment_Month, "%d%b%Y"), "month")) %>%
#   group_by(month) %>%
#   summarise(coverage = sum(`Patients registered at included practices`)/sum(`Patients registered at open practices`),
#             pract_coverage = sum(`Included Practices`)/sum(`Open Practices`))
# 
# 
# total_appts_month <- appts %>%
#   mutate(month = floor_date(as.Date(Appointment_Date, tryFormats = c("%d%b%Y", "%d-%b-%y")), "month")) %>%
#   group_by(month) %>%
#   summarise(count_appts = sum(COUNT_OF_APPOINTMENTS)) %>%
#   
#   left_join(national_coverage, by = "month") %>%
#   left_join(working_days, by = "month") %>%
#   left_join(vaccs, by = "month") %>%
#   
#   mutate(est_appts = count_appts/coverage,
#          est_apptsWD = (count_appts/coverage)/workingDays)
# 
# 

library(DBI)
library(RSQLite)
library(tidyverse)
library(lubridate)
library(readxl)
library(dbplyr)
library(jsonlite)
library(zoo)
library(RcppRoll)

user <- "TMorris2"
model_path <- paste0("C:/Users/", user, "/Department of Health and Social Care/NW012 - HCAT/General Practice/2. GENERAL PRACTICE ACTIVITY AND DEMAND/Dynamic GP Demand Model/v2/")
db_path <- paste0(model_path, "/gp_db.sqlite")
gp_db <- dbConnect(RSQLite::SQLite(), db_path)

source(paste0(model_path, "data_functions.R"))
source(paste0(model_path, "db_functions.R"))
source(paste0(model_path, "model_functions.R"))

vaccs <- dbReadTable(gp_db, "covid_vaccs") %>% 
    group_by(MONTH, YEAR) %>%
    summarise(VACCS = sum(VACCS)) %>%
    collect() %>%
    ungroup() %>%
    mutate(month = get_date_db(YEAR, MONTH)) %>%
    select(-YEAR, -MONTH, month, vaccs = VACCS)

total_appts_month <- get_appts_ts() %>% 
  collect() %>%
  mutate(date = get_date_db(YEAR, MONTH)) %>%
  select(month = date,
         count_appts = TOTAL_APPTS,
         coverage = COVERAGE,
         workingDays = WORKING_DAYS,
         est_appts = EST_APPTS,
         est_apptsWD = EST_APPTS_WD) %>%
  left_join(vaccs, by = "month")




  




