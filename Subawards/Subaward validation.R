#===============================================================================
#Subaward validation
#BY: JOLENE WUN

#PURPOSE: flag GH awards/subawards that need review by AOR
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
#READ IN TABLEAU PREP OUTPUT FILE
#===============================================================================
subaward <- read_csv("~/GitHub/Locally Led Development/Subawards/subaward_init.csv") 

gh_subaward <- filter(subaward, str_detect(`Managing Organization`, "GH/"))

#view(count(gh_subaward, `Managing Organization`))

#===============================================================================
#NEW SUBAWARDS OF AWARDS WITH MULTIPLE GH OFFICES
#===============================================================================
previous_subaward <- read_sheet(as_id('1jv-1F5JP7fnuoafzJQ9PePK88jBDpxpliqsPHKth17M'), 
                                                 sheet="Export Worksheet") %>%
  select(CONTRACT_PIID, SUBAWARD_NUMBER, SUBAWARDEE_NAME, SUBAWARDEE_COUNTRY_NAME, 
         SUBAWARD_POP_COUNTRY_NAME)

check_subaward <- filter(gh_subaward, `Multiple GH offices/divisions`==TRUE) %>%
  anti_join(previous_subaward) %>%
  select(`COR/AOR Name`, `Managing Organization`,	`Document Number`,
         `Obligation Header Description`,	`Vendor Name`, SUBAWARD_NUMBER,
         SUBAWARD_AMOUNT,	SUBAWARD_ACTION_DATE,	SUBAWARDEE_NAME,	SUBAWARD_DESCRIPTION)

#===============================================================================
#EXPORT TO GOOGLE SHEET
#===============================================================================
#New GH subawards: GH designation needed
check_subaward_sheet <- as_id('18lB379Ik33R6pxRVwtRGo4u_59tZCy_lhfXSyqCIKzM')

sheet_copy(check_subaward_sheet,"template", check_subaward_sheet, timestamp)

range_write(check_subaward_sheet, check_subaward, sheet=timestamp)

#GH Awards: new to subaward database
check_gh_award_sheet <- as_id('1VxUuKl7Y0MZDTNcmU6I4rvVPiKuPYxm0xfqWNJCZsQ4')

previous_award <- read_sheet(as_id('1jv-1F5JP7fnuoafzJQ9PePK88jBDpxpliqsPHKth17M'), 
                             sheet="Export Worksheet") %>%
  distinct(CONTRACT_PIID)

check_award <- gh_subaward %>%
  group_by(`COR/AOR Name`, `Managing Organization`,	`Document Number`,	
         `Obligation Header Description`,	`Vendor Name`) %>%
  summarise(`Latest Obligation FY` = max(ORIGINAL_FY), .groups="drop") %>%
  anti_join(previous_award, by=c("Document Number" = "CONTRACT_PIID")) %>%
  arrange(`COR/AOR Name`,`Latest Obligation FY`) %>%
  mutate(`Date added` = timestamp)

sheet_append(check_gh_award_sheet, check_award, sheet="New (raw)")
