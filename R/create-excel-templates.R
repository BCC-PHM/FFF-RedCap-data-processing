library(readxl)
library(dplyr)
source("R/functions.R")


# Load the data
questions_long <- read_excel("data/fff-question-lookup.xlsx") 

project_table <- read_excel("data/projects-table.xlsx")

create_all_templates(
    questions_long,
    project_table
)

project_name <- "test_project1"

# Get surveys for project
survey_filter <- project_table %>% 
  filter(
    project == project_name
  ) %>%
  pull(surveys) %>% 
  stringr::str_split("; ")

project_id <- project_table %>% 
  filter(
    project == project_name
  ) %>%
  pull(project_id)

# Filter for surveys required by project
questions_long_filtered <- questions_long %>%
  filter(
    survey %in% survey_filter[[1]]
  )

wb <- create_template(
  questions_long,
  project_table,
  project_name
) 

wb <- add_validation(
  wb,
  questions_long_filtered
)

saveWorkbook(wb, "output/check.xlsx", overwrite = TRUE)