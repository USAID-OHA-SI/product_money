#### Title
# PURPOSE: Reading in Quarterly Reports
# AUTHOR: jodavis | sch
# LICENSE: MIT
# DATE: 2023-09-22
# NOTES: 

#### LOCALS & SETUP ============================================================================

# Libraries
require(tidyverse)
require(gagglr)
require(here)
require(googledrive)
require(readxl)
#si_setup()

#### FUNCTION ============================================================================  

mfs_muncher = function(file, path = here::here("Data")) {
  
  file <- paste0(path, "/", file)
  
  ta <-  c(
    "Strategy and Planning TA",
    "Forecasting and Supply Planning TA",
    "Procurement TA",
    "Quality Assurance TA",
    "Warehousing and Inventory Management TA",
    "Transportation and Distribution TA",
    "MIS TA",
    "Governance and Financing TA",
    "Monitoring and Evaluation TA",
    "Human Resources Capacity Development TA",
    "Global Standards - GS1 TA"
  )
  
  other <- c(
    "In-country Storage & Distribution",
    "Global Collaboration",
    "Knowledge Management",
    "Staff Development",
    "Country Support",
    "Program Management",
    "Office Operations"
  )
  
  if(length(which(excel_sheets(file) == "TO1-COP"))==0){
    return(NULL)
  }
  
  ## get period
  period <- str_remove(file, paste0(path, "/")) %>% 
    str_remove(".xlsx") %>% 
    str_extract("\\b\\w+ \\d{4}\\b")
  
  ou <- str_remove(file, paste0(path, "/")) %>% 
    str_remove(".xlsx") %>% 
    str_remove("\\b\\w+ \\d{4}\\b") %>%
    trimws()
  
  
  #need col headers
  
  col1 <- read_xlsx(file,
                    sheet = which(str_detect(excel_sheets(file), "TO1-COP"))[1],
                    range = "E9") %>%
    names()
  
  col2 <- read_xlsx(file,
                    sheet = which(str_detect(excel_sheets(file), "TO1-COP"))[1],
                    range = "F9") %>%
    names()
  
  col3 <- read_xlsx(file,
                    sheet = which(str_detect(excel_sheets(file), "TO1-COP"))[1],
                    range = "G9") %>%
    names()
  
  df_commod <- read_xlsx(file,
                         sheet = which(str_detect(excel_sheets(file), "TO1-COP"))[1],
                         skip = 33) %>%
    rename(!!col1 := `...4`,!!col2 := `...5`,!!col3 := `...6`) %>%
    janitor::clean_names() %>%
    select(-x2, -x3) %>%
    filter(
      if_any(everything(), ~ !is.na(.x)),
      if_any(where(is.numeric), ~ !is.na(.x)),!commodities %in% c("Subtotal Commodities", "Grand Total")
    ) %>% 
    rename(subcategory = commodities) %>% 
    pivot_longer(
      cols = where(is.numeric),
      names_to = "financial_activity",
      values_to = "value",
      values_drop_na = TRUE
    ) %>%
    mutate(value = round(value, 0),
           detail = "total") %>%
    filter(value != 0) %>%
    mutate(category = "Commodities")
  
  df_ta <- read_xlsx(file,
                     sheet = which(str_detect(excel_sheets(file), "TO1-COP"))[1],
                     skip = 8) %>%
    janitor::clean_names() %>%
    slice(1:24) %>%
    rename(detail = in_country_sc_activites_ops) %>%
    filter(if_any(everything(), ~ !is.na(.x))) %>%
    mutate(
      subcategory = case_when(
        detail %in% ta ~ "Technical Assistance",
        detail %in% other ~ "Other"
      ),
      category = "Total In-Country SC Activites & Ops"
    ) %>%
    select(c(detail, subcategory, category) | where(is.numeric)) %>%
    filter(!detail %in% c(
      "Technical Assistance",
      "Other",
      "Total In-Country SC Activites & Ops"
    )) %>%
    pivot_longer(
      cols = where(is.double),
      names_to = "financial_activity",
      values_to = "value",
      values_drop_na = TRUE
    ) %>%
    filter(value != 0) %>%
    mutate(value = round(value, 0))
  
  df <- bind_rows(df_ta, df_commod) %>%
    mutate(
      country = ou,
      period = paste0(period, " 01")
    ) %>%
    mutate(period = as.Date(period, format = "%B %Y %d")) %>%
    relocate(country, .before = detail) %>%
    relocate(period, .after = country) %>%
    relocate(category, .before = detail) %>%
    relocate(subcategory, .before = detail)
  
  df <- df %>%
    dplyr::mutate(
      smp = lubridate::quarter(
        x = lubridate::ymd(period),
        with_year = TRUE,
        fiscal_start = 10
      ),
      fiscal_quarter = paste0("FY", substr(smp, 3, 4), "Q", substr(smp, 6, 6))
    ) %>%
    dplyr::select(-smp)
  
  return(df)
}

monthly_muncher <- function(folder, path = here::here("Data")){
  
  # Import files
  files_in_folder <- googledrive::drive_ls(googledrive::as_id(folder))
  for(name in files_in_folder$name){
    glamr::import_drivefile(drive_folder = folder,
                             filename = name,
                             folderpath = path,
                             zip = FALSE)
  }
  
  df <- files_in_folder$name %>%
    map_dfr(~ mfs_muncher(., path = path))
  
  return(df)
}

#### Assemble Historical File ============================================================================  

# Download Files to here("Data")
drive_auth()
year_folder = "1dk5gBttDvX4fl78vfbyfiqjXZabhwkAW"
files_in_folder <- googledrive::drive_ls(googledrive::as_id(year_folder))
for(folder in files_in_folder$id){
  files_in_subfolder <- googledrive::drive_ls(googledrive::as_id(folder))
  for(name in files_in_subfolder$name){
    glamr::import_drivefile(drive_folder = folder,
                            filename = name,
                            folderpath = here("Data"),
                            zip = FALSE)
  }
}

# Map to files to mfs_muncher()
historical_mfrs <- list.files(here("Data")) %>%
  map_dfr(~ mfs_muncher(.x))

historical_mfrs %>%
  write_csv(here("Dataout", "fy23_mfrs.csv"))

# Check for tables that didn't get read in because of the TO1-COP tab issue
tables_exist = historical_mfrs %>%
  distinct(country, period) %>%
  mutate(country = trimws(country)) %>%
  mutate(matches = T)

tables_total = list.files(here("Data")) %>%
  as_tibble() %>%
  mutate(value = str_remove(value, ".xlsx"),
         period = str_extract(value, "\\b\\w+ \\d{4}\\b"),
         country = str_remove(value, "\\b\\w+ \\d{4}\\b"),
         country = trimws(country)) %>%
  select(country, period) %>%
  mutate(period = paste0(period, " 01"),
         period = as.Date(period, "%B %Y %d"))

no_to1_cop <- tables_total %>%
  left_join(tables_exist) %>%
  filter(is.na(matches)) 

no_to1_cop %>%
  write_csv(here("Dataout", "no_to1_cop.csv"))

# Issues with the tables so far:
# - Many OUs don't have a TO1-COP table
# - There are two Liberia November 2022 sheets
# - Rwanda May 2023 TO1-Summary starts a line lower
# - Sheets throughout have tabs that include a space after the name
# - In a few tables Malawi lab commodities are recorded as "."

