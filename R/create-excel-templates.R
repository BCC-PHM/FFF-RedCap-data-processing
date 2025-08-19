library(readxl)
library(dplyr)
source("R/functions.R")

# Load the data
questions_long <- read_excel("data/fff-question-lookup.xlsx") %>%
  select(
    survey, section_header, user_column_name
  )

project_table <- read_excel("data/projects-table.xlsx")

create_all_templates(questions_long, project_table)