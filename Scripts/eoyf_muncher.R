#### Title
# PURPOSE: EOYF Muncher
# AUTHOR: alerichardson | sch
# LICENSE: MIT
# DATE: 2023-11-13
# NOTES: 

#### LOCALS & SETUP ============================================================================

# Libraries
require(tidyverse)
require(gagglr)
require(here)
require(googledrive)
require(readxl)
#si_setup()

#### DOWNLOAD DATA ============================================================================  

# Download Files to here("Data", "EOYF")
drive_auth()
eoyf_folder = "1btJy4izG4h3lZsc50gSw2x8jK36HyoEH"
files_in_folder <- googledrive::drive_ls(googledrive::as_id(eoyf_folder))
for(name in files_in_folder$name){
    glamr::import_drivefile(drive_folder = eoyf_folder,
                            filename = name,
                            folderpath = here("Data", "EOYF"),
                            zip = FALSE)
}

#### EOYF Muncher(s) ============================================================================  

pipeline_muncher <- function(filepath) {
  tabs <- excel_sheets(filepath)
  
  df = read_xlsx(filepath,
                 sheet = tabs[1])
  names(df) <- letters[1:length(df)]
  
  df %>%
    select(a, b) %>%
    pivot_wider(names_from = a, values_from = b) %>%
    return()
}

mechanism_muncher <- function(filepath) {
  if(length(which(str_detect(excel_sheets(filepath), "Mechanism")))==0){
    return(NULL)
  }
  
  df = read_xlsx(filepath,
            sheet = which(str_detect(excel_sheets(filepath), "Mechanism"))[1],
            skip = 2) %>%
    rename_with(str_to_sentence) %>%
    mutate_at(if('% Execution' %in% names(.)) '% Execution' else integer(0), as.numeric)
  
  return(df)
}

#### RUN FUNCTIONS ============================================================================  

pipelines <- files_in_folder$name %>%
  here("Data", "EOYF", .) %>%
  map_dfr(pipeline_muncher)

write_csv(pipelines, here("Dataout", "EOYF_pipelines.csv"))

mechanisms <- files_in_folder$name %>%
  here("Data", "EOYF", .) %>%
  map(mechanism_muncher) %>%
  list_rbind()

write_csv(mechanisms, here("Dataout", "EOYF_mechanisms.csv"))
