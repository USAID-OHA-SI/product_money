## for play

file <- here("Data/Zimbabwe July 2023.xlsx")

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


#need col headers

col1 <- read_xlsx(file,
                 sheet = "TO1-COP",
                 range = "E9") %>% 
  names()

col2 <- read_xlsx(file,
                  sheet = "TO1-COP",
                  range = "F9") %>% 
  names()

col3 <- read_xlsx(file,
                  sheet = "TO1-COP",
                  range = "G9") %>%
  names()

df_commod <- read_xlsx(file,
                       sheet = "TO1-COP",
                       skip = 33) %>%
  janitor::clean_names() %>%
  rename(!!col1 := x4,
         !!col2 := x5,
         !!col3 := x6) %>% 
  select(-x2,-x3) %>%
  filter(if_any(everything(), ~!is.na(.x)),
         if_any(where(is.numeric), ~!is.na(.x)),
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

########################
#freight?

# df_freight <- read_xlsx(file,
#                         sheet = "TO1-COP",
#                         skip = 33) %>%
#   janitor::clean_names() %>%
#   select(commodities, ex_works_value, freight_plus) %>%
#   filter(if_any(everything(), ~ !is.na(.x)),
#          if_any(where(is.numeric), ~!is.na(.x)),
#          !commodities %in% c("Subtotal Commodities", "Grand Total")) %>%
#   rename(subcategory = commodities) %>%
#   pivot_longer(cols = c(ex_works_value, freight_plus),
#                names_to = "detail",
#                values_to = "value",
#                values_drop_na = TRUE) %>% 
#   mutate(value =round(value, 0),
#          finacial_activity = "total_outlays_fy21") %>% 
#   filter(value != 0) %>% 
#   mutate(category = "Commodities")


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
  select(c(detail, subcategory, category) | where(is.numeric)) %>%
  filter(!detail %in% c("Technical Assistance", "Other", "Total In-Country SC Activites & Ops")) %>%
  pivot_longer(cols = where(is.double),
               names_to = "finacial_activity",
               values_to = "value",
               values_drop_na = TRUE) %>%
  filter(value != 0) %>% 
  mutate(value =round(value, 0))

df <- bind_rows(df_ta, df_commod) %>% 
  mutate(country = ou,
         period = period,
         country = stringr::str_extract(country, "\\ - .*"),
         country = stringr::str_remove(country, "\\ -")) %>%
  relocate(country, .before = detail) %>% 
  relocate(period, .after = country) %>% 
  relocate(category, .before = detail) %>% 
  relocate(subcategory, .before = detail)

df %>% write_csv("Dataout/zim_mfs_test_july.csv")










