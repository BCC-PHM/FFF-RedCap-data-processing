library(readxl)
library(dplyr)
source("R/functions.R")


# Load the data
questions_long <- read_excel("data/fff-question-lookup.xlsx") 

project_table <- read_excel("data/projects-table.xlsx")

#questions_long$RC_option_code[[10]] <- "1; 0; 2"

create_all_templates(
    questions_long,
    project_table
)