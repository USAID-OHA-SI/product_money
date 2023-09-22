extract_mfs <- function(file) {

  ##read in TO1-COP tab
  # create some objects
 ta <-  c("Strategy and Planning TA",
    "Forecasting and Supply Planning TA",
    "Procurement TA",
    "Quality Assurance TA",
    "Warehousing and Inventory Management TA",
    "Transportation and Distribution TA",
    "MIS TA",
    "Governance and Financing TA",
    "Monitoring and Evaluation TA",
    "Human Resources Capacity Development TA",
    "Global Standards - GS1 TA")

 other <- c("In-country Storage & Distribution",
            "Global Collaboration",
            "Knowledge Management",
            "Staff Development",
            "Country Support",
            "Program Management",
            "Office Operations")
 
 ## get period
 period <- read_xlsx(file,
                     sheet = "TO1-COP",
                    range = "D4") %>% 
   names()
 
 ou <- read_xlsx(file,
                 sheet = "TO1-Summary",
                 range = "B1") %>% 
   names()

  

df_ta <- read_xlsx(file,
                    sheet = "TO1-COP",
                    skip = 8) %>%
  janitor::clean_names() %>%
  slice(1:24) %>%
  rename(detail = in_country_sc_activites_ops) %>%
  filter(if_any(everything(), ~ !is.na(.x))) %>%
  mutate(subcategory = case_when(detail %in% ta ~ "Technical Assistance",
                                 detail %in% other ~ "Other"),
         category = "Total In-Country SC Activites & Ops") %>% 
  select(detail,
         total_budget_fy21_excludes_cr,
         total_outlays_fy21,
         monthly_outlay,
         subcategory,
         category) %>%
  filter(!detail %in% c("Technical Assistance", "Other", "Total In-Country SC Activites & Ops")) %>%
  pivot_longer(cols = c(total_budget_fy21_excludes_cr,
                          total_outlays_fy21,
                          monthly_outlay),
                 names_to = "finacial_activity",
                 values_to = "value",
                 values_drop_na = TRUE) %>%
  filter(value != 0) %>% 
  mutate(value =round(value, 0))

#read in commodities freight
df_freight <- read_xlsx(file,
                        sheet = "TO1-COP",
                        skip = 33) %>%
  janitor::clean_names() %>%
  select(commodities, ex_works_value, freight_plus) %>%
  filter(if_any(everything(), ~ !is.na(.x)),
         across(c(ex_works_value, freight_plus), ~ !is.na(.x)),
         !commodities %in% c("Subtotal Commodities", "Grand Total")) %>%
  rename(subcategory = commodities) %>%
  pivot_longer(cols = c(ex_works_value, freight_plus),
               names_to = "detail",
               values_to = "value",
               values_drop_na = TRUE) %>% 
  mutate(value =round(value, 0),
         finacial_activity = "total_outlays_fy21") %>% 
  filter(value != 0) %>% 
  mutate(category = "Commodities")


 
  #read in commodities total
  df_commod <- read_xlsx(file,
                          sheet = "TO1-COP",
                          skip = 33) %>%
    janitor::clean_names() %>%
    rename(total_budget_fy21_excludes_cr = x4,
           total_outlays_fy21 = x5,
           monthly_outlay = x6) %>% 
    select(commodities, total_budget_fy21_excludes_cr, total_outlays_fy21, monthly_outlay) %>%
    filter(if_any(everything(), ~ !is.na(.x)),
           across(c(total_budget_fy21_excludes_cr, total_outlays_fy21, monthly_outlay), ~ !is.na(.x)),
           !commodities %in% c("Subtotal Commodities", "Grand Total")) %>%
    rename(subcategory = commodities) %>%
    pivot_longer(cols = where(is.numeric),
                 names_to = "finacial_activity",
                 values_to = "value",
                 values_drop_na = TRUE) %>% 
    mutate(value =round(value, 0),
           detail = "total") %>% 
    filter(value != 0) %>% 
    mutate(category = "Commodities")
  
#bind commodities and freight
  bind1 <- bind_rows(df_commod, df_freight)
  
#bind commodities incl freight to TA
  
  df <- bind_rows(df_ta, bind1) %>% 
    mutate(country = ou,
           period = period,
           country = stringr::str_extract(country, "\\ - .*"),
           country = stringr::str_remove(country, "\\ -")) %>% 
    relocate(country, .before = detail) %>% 
    relocate(period, .after = country) %>% 
    relocate(category, .before = detail) %>% 
    relocate(subcategory, .before = detail)
  
  return(df)
  
}

  
  
  
  
  
  
  