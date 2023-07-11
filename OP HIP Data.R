#===============================================================================
#OP DATA PULL
#BY: JOLENE WUN
#===============================================================================

#===============================================================================
#SETUP
#===============================================================================
#Libraries
library(gdata)
library(readxl)
library(Dict)
library(logr)
library(data.table)
library(tidyverse)
library(googledrive)
library(filesstrings)
library(googlesheets4)
library(writexl)
library(glue)
library(janitor)
library(lubridate)

#Location of downloaded files
setwd("C:/Users/jwun/Downloads/")

#Today's date (make sure that the files are downloaded the same day the script is run)
timestamp <- as.character(today())

#Reference tables
ref <- as_id("1CrWwsL6X1OPsJR0N5HN4bQW9NDJ6ZKIPc9dCAEbojQg")

#Output file: FY2022 OP/HIP tracker
output <- as_id('1VVKkP4DONiuZcJuf8YNz8tcQiR_K30b4W9bXEtbha7U')

#===============================================================================
#READ IN OP DATA
#===============================================================================
#Upload original OP excel files to google drive
raw_folder <- as_id("1_m2sUGwsoPU6WzpAhl6qPlDoj6Z4dDlS")

drive_upload("LLD data pull - FY2022 OP current.xlsx", name=paste("LLD data pull - FY2022 OP current", timestamp),
             path=raw_folder, overwrite=TRUE)
drive_upload("LLD data pull - FY2022 OP approved.xlsx", name=paste("LLD data pull - FY2022 OP approved", timestamp),
             path=raw_folder, overwrite=TRUE)

#Create consolidated file for FY2022
current_fields <- c("Operating Unit",	"SPSD Area",	"Account",	"USAID Bureau",	
                    "IM Number",	"Implementing Mechanism (IM)",	"IM Narrative",	
                    "Award Number","Prime Partner", "Metrics", "OP Current")

approved_fields <- current_fields[!current_fields=="IM Narrative"] %>%
  replace(approved_fields=="OP Current", "OP Approved")

op <- left_join(
  read_excel("LLD data pull - FY2022 OP current.xlsx", skip=11, col_names=current_fields),
  read_excel("LLD data pull - FY2022 OP approved.xlsx", skip=11, col_names=approved_fields)) %>%
  select(!Metrics) %>%
  filter(`Operating Unit`!="Office of U.S. Foreign Assistance Resources") %>%
  mutate(`IM Number`= as.integer(`IM Number`)) %>%
  mutate(`Award Number` = case_when(
    `Award Number`=="N/A" | `Award Number`=="Not Applicable" ~ NA_character_, 
    TRUE ~ `Award Number`))

#===============================================================================
#READ IN OLD OP DESIGNATIONS
#===============================================================================
#Old OP22 designations from 2022 OP/HIP all Health Programming 
all_hp_status <- read_sheet(as_id('1mOceKSR-qRfSw8PaN2namP0eP-_DkRG4OYDQw2VCQM8'), 
                                   range="'upd w/Bureau & PA'!A3:W5172") %>%
  filter(!is.na(`FY 2022 OP Current`)) %>%
  select(`Operating Unit`,`Account`, `IM Number`,`Staffing or Operations`,
         `SPSD Area`, `Mechansim Status \n(All Health Programming)`,	
         `Staff Signed\n(All Health Programming)`,`Mechansim Status \n(All Health Programming 2020-2022)`) %>%
  rename(`Mechanism Status (All HP)` = `Mechansim Status \n(All Health Programming)`,
         `Staff signed (All HP)` = `Staff Signed\n(All Health Programming)`,
         `Mechanism Status (All HP 2020-2022)` = `Mechansim Status \n(All Health Programming 2020-2022)`)

#check whether there are any duplicate entries for each IM
all_hp_check <- all_hp_status %>%
  mutate(dup = duplicated(select(.,`IM Number`, `Account`, `SPSD Area`), fromLast=FALSE) | 
                             duplicated(select(., `IM Number`, `Account`, `SPSD Area`), fromLast=TRUE))

#TBD sheet (GH Localization: OP/HIP Data _ with # of TBDs per Country)
tbd_status <- read_sheet(as_id('1p-UROilekLk0uX16Qa7nM4Yr4bKIoq_uEsAH1pjMDCI'), sheet="TBDs") %>%
  filter(!is.na(`IM Number`)) %>%
  select(`IM Number`, `Account`, `Mechanism Status (as of April 2023)`, 
         `Signed off by Mission by June 16th (include initials)`,
         `What kind of localization support do you need from GH Bureau?`,
         `Additional Comments`) %>%
  rename(`HQ support needed` = `What kind of localization support do you need from GH Bureau?`) %>%
  mutate(`HQ support needed` = case_when(
    `HQ support needed` %in% c("N/A",
                               "NA",
                               "n/a this award is on hold until further notice due to the non-permissive operating environment",
                               "Not at the moment") ~ NA_character_,
    str_detect(`HQ support needed`, "None") ~ NA_character_, 
    TRUE ~ `HQ support needed`))

#===============================================================================
#READ IN AWARD DATA FROM PHOENIX
#===============================================================================
glaas_vendor <- read_excel("GLAAS Released Award Summary - In Progress 2023.07.03.xlsx", 
                           sheet="Data") %>%
  distinct(`Award Number`, `Vendor Name`) %>%
  rename(`GLAAS Vendor Name` = `Vendor Name`) %>%
  mutate(`TBD flag` = "Award Number not blank, marked as TBD")

#===============================================================================
#COMBINE ALL DATA
#===============================================================================
op_export <- op %>%
  left_join(select(read_sheet(ref, sheet="Cross-bureau funded IMs"), !MECHANISM), multiple="all") %>%
  mutate(`Cross-bureau IM` = !is.na(`Managing GH Entity`)) %>%
  left_join(read_sheet(ref, sheet="SPSD area"), multiple="all") %>%
  mutate(`Managing GH Entity` = ifelse(is.na(`Managing GH Entity`),`SPSD Area Entity`,`Managing GH Entity`)) %>%
  select(!`SPSD Area Entity`) %>%
  left_join(all_hp_status) %>%
  left_join(tbd_status) %>%
  relocate(`USAID Bureau`) %>%
  relocate(`Managing GH Entity`, `Cross-bureau IM`, .after="SPSD Area") %>%
  relocate(`OP Current`, `OP Approved`, .after="Additional Comments") %>%
  mutate(`TBD flag` = case_when(
    `Award Number`=="To be Determined" | str_detect(`Award Number`, "TBD") ~ "TBD",
    str_detect(`Mechanism Status (All HP)`, "TBD") | 
      str_detect(`Mechanism Status (All HP 2020-2022)`, "TBD") | 
      str_detect(`Mechanism Status (as of April 2023)`, "TBD") ~ "Award Number not blank, marked as TBD",
    TRUE ~ NA_character_)) %>%
  left_join(glaas_vendor) %>% #should not have more than one match. If there is, create a separate table
  mutate(`Mechanism Status (final)` = case_when(
    str_detect(`Implementing Mechanism (IM)`, regex('program', ignore_case=TRUE)) & 
      (str_detect(`Mechanism Status (All HP)`, "Other")| `Mechanism Status (as of April 2023)`=="Other") ~ "USAID Staffing/Operations",
    str_detect(`Implementing Mechanism (IM)`,"Foreign Service Limited") ~ "USAID Staffing/Operations",
    str_detect(`Implementing Mechanism (IM)`,"Personal Services Contractor") ~ "USAID Staffing/Operations", 
    str_detect(`Implementing Mechanism (IM)`,regex('STAR|TAMS', ignore_case=FALSE)) ~ "USAID Staffing/Operations",
    str_detect(`Implementing Mechanism (IM)`,"Administration and Oversight") ~ "USAID Staffing/Operations", 
    str_detect(`Mechanism Status (All HP)`, "Other") | str_detect(`Mechanism Status (All HP 2020-2022)`, "Other") | str_detect(`Mechanism Status (as of April 2023)`, "Other") ~ NA_character_,
    `Mechanism Status (All HP)`=="TBD - Unknown" & `Mechanism Status (All HP 2020-2022)`=="TBD - Unknown" & str_detect(`Mechanism Status (as of April 2023)`, "TBD - Expected") ~ `Mechanism Status (as of April 2023)`,
    `Mechanism Status (All HP)`==`Mechanism Status (All HP 2020-2022)` & `Mechanism Status (All HP 2020-2022)`==`Mechanism Status (as of April 2023)` ~ `Mechanism Status (as of April 2023)`,
    `Mechanism Status (All HP)`==`Mechanism Status (All HP 2020-2022)` & is.na(`Mechanism Status (as of April 2023)`) ~ `Mechanism Status (All HP)`,
    `Mechanism Status (All HP)`==`Mechanism Status (as of April 2023)` & is.na(`Mechanism Status (All HP 2020-2022)`) ~ `Mechanism Status (All HP)`,
    TRUE ~ NA_character_), .after="Mechanism Status (as of April 2023)") %>%
  mutate(`Mechanism Status (final)` = ifelse((!str_detect(`Mechanism Status (final)`,"TBD") & 
                                             `TBD flag`=="TBD"),
                                             paste("TBD - Expected",`Mechanism Status (final)`), 
                                             `Mechanism Status (final)`))


sheet_write(op_export, output, sheet="FY 2022")