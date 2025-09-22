# Convert completed templates into spreadsheet that is compatible
# with the RedCap system
source("R/convertion/convertion-functions.R")

sheet_path_i <- "data/complete-templates/FFF-template-ALL PROJECT SURVEY TEST-complete.xlsx"

convert_single_sheet(
  sheet_path_i,
  verbose = T
)
# test