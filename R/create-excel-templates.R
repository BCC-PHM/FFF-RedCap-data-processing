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

# 
# questions_long %>%
#   filter(RC_option_code != "Imputed") %>%
#   pull(RC_option_code) %>%
#   stringr::str_split(";") %>%
#   unlist() %>%
#   trimws() %>%
#   unique()