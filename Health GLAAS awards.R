#===============================================================================
#ALL GLAAS AWARDS WITH HEALTH ELEMENTS
#BY: JOLENE WUN

#PRIOR TO RUNNING, ENSURE THAT ALL OLD PHOENIX EXCEL FILES ARE REMOVED FROM THE
#DOWNLOADS FOLDER, AND DOWNLOAD THE FOLLOWING SAVED QUERIES:
# - Obligations by Original Fiscal Year - HL GLAAS only
# - Obligations by Original Fiscal Year - GH GLAAS only
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

#===============================================================================
#ADD RAW PHOENIX DATA TO DRIVE FOLDER
#===============================================================================
#Folder: Phoenix health data pulls
raw_folder <- as_id('1rZvjHyaA9kkzJludOcG-U1o-Z1Z35l6n')

drive_upload("Obligations by Original Fiscal Year - HL GLAAS only.xlsx", 
             name=paste("Obligations by Original Fiscal Year - HL GLAAS only", timestamp),
             path=raw_folder, overwrite=TRUE)

drive_upload("Obligations by Original Fiscal Year - GH GLAAS only.xlsx", 
             name=paste("Obligations by Original Fiscal Year - GH GLAAS only", timestamp),
             path=raw_folder, overwrite=TRUE)

asofdate <- read_excel("Obligations by Original Fiscal Year - HL GLAAS only.xlsx", 
           sheet="Parameters", range="B7", col_names="asofdate") %>% 
  mutate(asofdate = as.Date(mdy_hms(asofdate)))

#===============================================================================
#PROCESS PHOENIX DATA
#===============================================================================
temp1 <- read_excel("Obligations by Original Fiscal Year - HL GLAAS only.xlsx")

temp2 <- read_excel("Obligations By Original Fiscal Year - GH GLAAS only.xlsx")

glaas <- bind_rows(temp1, temp2) %>%
  distinct() %>%
  filter(!is.na(`Document Number`))

glaas_award_hl <- select(glaas,
    `Document Number`,`Vendor Name`,`Program Area`, `Program Area Name`,
    `Managing Organization`, `COR/AOR Name`, `Contact`, 
    `Obligation Header Description`, `Equivalent SPSD Area Code`) %>%
  distinct() %>%
  mutate(`SPSD Area` = paste(`Program Area`, `Program Area Name`)) %>%
  left_join(read_sheet(ref, sheet="SPSD area"))

area_count <- glaas_award_hl %>%
  group_by(`Document Number`, `Obligation Header Description`) %>%
  summarise(`Multiple SPSD Area Entities` = n_distinct(`SPSD Area Entity`), .groups = "drop") %>%
  mutate(`Multiple SPSD Area Entities` = ifelse(`Multiple SPSD Area Entities`>1,TRUE,FALSE))

area_detail <- glaas_award_hl %>%
  distinct(`Document Number`, `Obligation Header Description`,`SPSD Area Entity`) %>%
  group_by(`Document Number`, `Obligation Header Description`) %>%
  summarise(`SPSD Area Entities` = paste0(`SPSD Area Entity`,collapse="; ")) %>%
  mutate(`SPSD Area Entities` = str_replace_all(`SPSD Area Entities`, "; NA","")) %>%
  mutate(`SPSD Area Entities` = str_replace_all(`SPSD Area Entities`, "NA","")) %>%  
  ungroup()

glaas_awards_full <- select(glaas_award_hl, !contains("Area")) %>%
  unique() %>%
  left_join(area_count) %>%
  left_join(area_detail)

glaas_awards_final <- glaas_awards_full %>%
  distinct(`Document Number`, .keep_all=TRUE) %>%
  bind_cols(asofdate)

#===============================================================================
#ADD PRIOR SUBAWARD ASSIGNMENTS !!MODIFY IN NEXT DATA RUN
#===============================================================================
#Read in Subaward analysis (SubAwards - Active GH OHA Awards - 2023_06Jun_02)
subaward_data <- read_sheet(as_id('1jv-1F5JP7fnuoafzJQ9PePK88jBDpxpliqsPHKth17M'), 
                            sheet="Export Worksheet") %>%
  select(CONTRACT_PIID, CO_AO_NAME, CTO_AOR_PERSON_ID,	CTO_AOR_PERSON_NM, VENDOR_NAME,
         OBLIGATED, SUBAWARD_NUMBER, SUBAWARD_AMOUNT, SUBAWARD_ACTION_DATE,
         SUBAWARDEE_NAME, SUBAWARDEE_COUNTRY_NAME, SUBAWARD_POP_COUNTRY_NAME, 
         SUBAWARD_DESCRIPTION, CTO_AOR_PERSON_ID_PA_OFFICE)

#Determine whether award has multiple SPSD Area Entities
subaward_ofd <- select(subaward_data, CONTRACT_PIID, CTO_AOR_PERSON_ID_PA_OFFICE) %>%
  unique() %>%
  count(CONTRACT_PIID) %>%
  mutate(`Multiple GH offices/divisions` = n>1) %>% select(!n)

single_ofd_a <- select(subaward_data, CONTRACT_PIID, CTO_AOR_PERSON_ID_PA_OFFICE)

single_ofd <- filter(subaward_ofd, `Multiple GH offices/divisions`==FALSE) %>%
  left_join(single_spsd_a) %>%
  rename(`GH office/division` = CTO_AOR_PERSON_ID_PA_OFFICE)

subaward_ofd <- left_join(subaward_ofd, single_ofd)

#===============================================================================
#COMBINE PHOENIX WITH SUBAWARD ASSIGNMENTS AND EXPORT TO REFERENCE TABLE
#===============================================================================
#Merge GLAAS info with subaward info
prime_awards <- left_join(glaas_awards_final, subaward_ofd, by=c("Document Number"="CONTRACT_PIID"))

sheet_write(prime_awards, ref, sheet="Health GLAAS Awards")

#TEMPORARY: Compare previous coding with Phoenix data
check <- prime_awards %>%
  mutate(`Flag 1` = case_when(
    is.na(`SPSD Area Entities (original coding)`) ~ FALSE,
    `SPSD Area Entities (original coding)` %in% c("CII","PPP","OCS","OHS","PDMS") ~ FALSE,
    TRUE ~ !str_detect(`SPSD Area Entities`,`SPSD Area Entities (original coding)`))) %>%
  mutate(`Flag 2` = case_when(
    !str_detect(`Managing Organization`, "GH/") ~ FALSE,
    str_detect(`Managing Organization`, "GH/ID") & str_detect(`SPSD Area Entities (original coding)`,regex("pmi|tb|ntd|ghs", ignore_case=T)) ~ FALSE,
    str_detect(`Managing Organization`, "GH/ID") & !str_detect(`SPSD Area Entities (original coding)`,regex("pmi|tb|ntd|ghs", ignore_case=T)) ~ TRUE,
    TRUE ~ !str_detect(`Managing Organization`,`SPSD Area Entities (original coding)`)
  ))


  