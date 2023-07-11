#===============================================================================
#COP MATRIX
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
output <- as_id('17le6a7ATuSRpNPa1dy59txD_Advil7_83Jlz1kf9c5s')

#===============================================================================
#READ IN COP DATA
#===============================================================================
#Upload original OP excel files to google drive
raw_folder <- as_id("1oES7sPckGwUS5qEnTWVeI4PypuolRzLO")

drive_upload("Standard COP Matrix - LP fields.xlsx", name=paste("Standard COP Matrix - LP fields", timestamp),
             path=raw_folder, overwrite=TRUE)

duns <- count(cop, `Prime Partner DUNS`)
funding_cat <- count(cop, `Funding Category`)

partner_type <- read_sheet(as_id("1MQviknJkJDttGdNEJeNaYPKmHCw6BqPuJ0C5cslV5IE"),sheet="Partner / COP matrix Data") %>%
  select(`Operating Unit`, `Planning Cycle`,`Mechanism ID`, `Funding Agency`, `Partner Name`,
         `Partner Type`) %>%
  unique()

cop <- read_excel("Standard COP Matrix - LP fields.xlsx", skip=8) %>%
  filter(`Planning Cycle`!="Office of U.S. Foreign Assistance Resources") %>%
  left_join(read_sheet(ref, sheet="OUs"), multiple="all") %>%
  relocate(`OU Type`, `OU Designation`,.after="Operating Unit") %>%
  relocate(`Mechanism ID`, `Mechanism Name`, .after="Funding Agency") %>%
  relocate(`Prime Partner DUNS`, `Partner Name`, `Partner Org Type`, `Is Indigenous Partner?`,
           .after="Total Planned Funding") %>%
  mutate(`TBD flag` = `Partner Name`=="TBD" & `Record Type`=="Implementing Mechanism") %>%
  left_join(partner_type) %>%
  rename(`Partner Type (from MechID file)` = `Partner Type`)

cop <- read_excel("Standard COP Matrix.xlsx", skip=8) %>%
  filter(`Planning Cycle`!="Office of U.S. Foreign Assistance Resources") %>%
  rename(`AGYW` = `Prevention among Adolescent Girls & Young Women`,
         `GBV` = `Gender: Preventing and Responding to Gender-based Violence (GBV)`,
         `Gender equality` = `Gender: Gender Equality`,
         `MSM and TG key pops` = `Key Populations: MSM and TG`,
         `SW key pops` = `Key Populations: SW`)

cop_areas <- cop %>% pivot_longer(cols=`AGYW`:`SE`) %>%
  filter(!is.na(value)) %>%
  filter(value>0) %>%
  select(`Mechanism ID`, name) %>% unique() %>%
  group_by(`Mechanism ID`) %>%
  summarise(`Mechanism Areas` = paste0(name,collapse="; "))

cop_final <- left_join(cop, read_sheet(ref, sheet="OUs"), multiple="all") %>%
  left_join(cop_areas) %>%
  relocate(`OU Type`, `OU Designation`,.after="Operating Unit") %>%
  relocate(`Mechanism ID`, `Mechanism Areas`, `Mechanism Name`,  `Award Number`, 
           `Prime Partner DUNS`, `Partner Name`, `Partner Org Type`, `Is Indigenous Partner?`,
           .after="Funding Agency") %>%
  mutate(`TBD flag` = `Partner Name`=="TBD" & `Record Type`=="Implementing Mechanism") %>%
  left_join(partner_type) %>%
  rename(`Partner Type (from MechID file)` = `Partner Type`) %>%
  select(!CIRC:SE)

mech_funding <- select(cop_final, !c(`Initiative`:`Applied Pipeline Amount`)) %>%
  group_by(`Planning Cycle`, `Mechanism ID`, `Partner Name`) %>%
  summarise(`Total Planned Funding` = sum(`Total Planned Funding`), .groups="drop")

mech <- select(cop_final, !c(`Initiative`:`Total Planned Funding`)) %>%
  unique() %>%
  arrange(`Operating Unit`, `Planning Cycle`, `Mechanism ID`) %>%
  left_join(mech_funding) %>%
  relocate(`Total Planned Funding`, .before="Partner Type (from MechID file)") %>%
  mutate(`Partner Type (final)` = case_when(
    `Record Type`=="Management and Operations" ~ "N/A (Management and Operations)",
    is.na(`Partner Name`) & `Record Type`=="Implementing Mechanism" ~ "Check",
    `Partner Type (from MechID file)`=="TBD Local" ~ "TBD - Expected Local NGO",
    `Partner Type (from MechID file)`=="TBD" ~ "TBD - Unknown",
    TRUE ~ NA_character_))

sheet_write(mech, output, sheet="Mechanisms")

sheet_write(select(cop_final, `Planning Cycle`:`TBD flag`), output, 
            sheet="COP matrix")

check <- mech %>%
  filter(`TBD flag`==TRUE & `Planning Cycle`!="2021 COP" & `Total Planned Funding`>0 & str_detect(`Funding Agency`, "USAID")) %>%
  group_by(`Planning Cycle`) %>%
  summarise(n = n_distinct(`Mechanism ID`))