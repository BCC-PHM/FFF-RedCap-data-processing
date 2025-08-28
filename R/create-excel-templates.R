library(readxl)
library(dplyr)
source("R/functions.R")


# Load the data
questions_long <- read_excel("data/fff-question-lookup.xlsx") 



create_all_templates(
    control_file_path = "data/projects-control-table.xlsx",
    project_table
)
