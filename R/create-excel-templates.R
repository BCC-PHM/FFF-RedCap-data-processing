library(readxl)
library(dplyr)
source("R/functions.R")


create_all_templates(
    control_file_path = "data/projects-control-table.xlsx",
    project_table
)
