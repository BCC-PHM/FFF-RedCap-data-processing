library(readxl)
library(dplyr)
source("R/functions.R")
source("R/validation-functions.R")

# Load the data
questions_long <- read_excel("data/fff-question-lookup.xlsx") 

project_table <- read_excel("data/projects-table.xlsx")



check_option_numbers(questions_long)
check_RedCap_codes(questions_long)
