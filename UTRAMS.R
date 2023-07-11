
#!!Update with new file name
filename <- "UTRAMS_data_FY19_or_later FILTER for G2G + gov_2023-06-30"

utrams_raw <- read_excel(paste(filename,".xlsx", sep=""), col_types="text") %>%
  mutate(update_date=as.Date(str_sub(filename, -10)))

sheet_write(utrams_raw, as_id("1OAdjLdZ68OwsOWmW-HuRwn_u544A95LFP7HAt58r0bc"),
            sheet="UTRAMS")

#Add file to raw data folder
raw_folder <- as_id("1_m2sUGwsoPU6WzpAhl6qPlDoj6Z4dDlS")

drive_upload(paste(filename,".xlsx", sep=""), path=raw_folder, overwrite=TRUE)

view(count(utrams_raw, `Benefitting Country`))

bc <- filter(utrams_raw, `Benefitting Country`!=`Destination` & 
               `Purpose Category`!="B. Conference/Meeting/Workshop Attendance and Presenting") %>%
  count(`Benefitting Country`, `Destination`)