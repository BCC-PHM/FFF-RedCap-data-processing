library(readxl)
library(readxl)
library(stringr)
library(writexl)
library(dplyr)

convert_single_sheet <- function(
  sheet_path,
  control_file_path = "data/projects-control-table.xlsx",
  question_lookup_path = "data/fff-question-lookup.xlsx",
  verbose = F
) {
  # load completed sheet
  comp_data <- read_excel(
    sheet_path,
    skip = 2,
    col_types = "text",
    .name_repair = "minimal"
  )
  
  # get project name - Should be the same for all returns
  project_name <- comp_data[[4]][[1]]
  
  # lookup project surveys from projects-control-file
  control_file <- read_excel(
    control_file_path,
    sheet = "Project Control Sheet"
    )
  
  surveys <- control_file %>%
    filter(
      project == project_name
    ) %>%
    pull(surveys) %>% 
    str_split("; ") %>%
    purrr::pluck(1)
  
  questions <- read_excel(question_lookup_path) %>%
    filter(
      survey %in% surveys
    ) %>%
    select(
      FFF_column_name, options, RC_option_code
    )
  
  # Convert table column names to RedCap names
  colnames(comp_data) = questions$FFF_column_name
  
  # Loop over all questions
  for (col_j in colnames(comp_data)) {
    
    options_j <- questions %>% filter(
      FFF_column_name == col_j
    )
    
    # Check if there are multiple options for the question
    if (grepl(";", options_j$options[[1]])) {
      # Create coded option lookup
      lookup_j <- setNames(
        trimws(str_split(options_j$RC_option_code[[1]], ";")[[1]]),
        trimws(str_split(options_j$options[[1]], ";")[[1]])
      )
      # remove any empty options
      lookup_j <- lookup_j[names(lookup_j)!= ""]
      
      # Convert column
      comp_data <- comp_data %>%
        mutate(!!sym(col_j) := recode(.data[[col_j]], !!!lookup_j)) #
    }
  }
  
  output_prefix <- "output/converted-for-RedCap/"
  
  output_path <- paste0(
    output_prefix,
    project_name,
    "-CONVERTED-",
    Sys.Date(),
    ".xlsx"
  )
  
  if (verbose) {
    print(
      paste0(
        "Converted project: \'",
        project_name,
        "\' successfully."
      )
    )
  }
  
  write_xlsx(comp_data, output_path)
  
}

bulk_convet <- function(
  file_path,
  control_file_path = "data/projects-control-table.xlsx",
  question_lookup_path = "data/fff-question-lookup.xlsx",
  verbose = F
) {
  
  # Get all xlsx files in folder
  files <- list.files(file_path, pattern = ".xlsx")
  
  # Convert all xlsx files in the given directory
  for (file_i in files) {
    
    # Create file path
    path_i <- file.path(
      file_path,
      file_i
    )
    
    # Convert sheet
    convert_single_sheet(
      sheet_path = path_i,
      control_file_path = control_file_path,
      question_lookup_path = question_lookup_path,
      verbose = verbose
    )
  }
}