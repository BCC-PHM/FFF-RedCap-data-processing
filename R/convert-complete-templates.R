# Convert completed templates into spreadsheet that is compatible
# with the RedCap system
source("R/convertion/convertion-functions.R")

sheet_path_i <- "data/complete-templates/FFF-template-ALL PROJECT SURVEY TEST-complete.xlsx"

bulk_convet(
  file_path = "data/complete-templates/",
  verbose = T
)