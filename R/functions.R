library(dplyr)
library(openxlsx)
library(tidyr)
#library(openxlsx2)
source("R/validation-functions.R")
source("R/imputation-functions.R")


convert_to_wide <- function(
    questions_long_filtered
) {
  
  # Transpose template table
  questions_wide <- as.data.frame(t(questions_long_filtered), stringsAsFactors = FALSE)
  rownames(questions_wide) <- NULL
  
  return(questions_wide)
}

get_column_positions <- function(
    questions_long_filtered
) {
  # Get start and end column position for each survey (for cell merging)
  survey_col_pos <- questions_long_filtered %>%
    mutate(row_number = row_number()) %>%
    group_by(survey) %>%
    summarise(
      first_col = min(row_number),
      last_col = max(row_number),
      .groups = "drop"
    ) %>%
    arrange(first_col)
  
  # Get start and end column position for each section header (for cell merging)
  header_col_pos <- questions_long_filtered %>%
    mutate(row_number = row_number()) %>%
    group_by(survey, section_header) %>%
    summarise(
      first_col = min(row_number),
      last_col = max(row_number),
      .groups = "drop"
    ) %>%
    arrange(first_col)
  
  return(
    list(survey_col_pos, header_col_pos)
  )
}

create_worksheet <- function(
    questions_wide,
    questions_long_filtered
) {
  

  ## Create a new workbook
  wb <- createWorkbook()
  ## Add a worksheet
  addWorksheet(wb, "Survey Input")
  ## Add wide questions to worksheet
  writeData(wb, sheet = "Survey Input", x = questions_wide, colNames = FALSE)
  
  # get column positions
  column_positions <- get_column_positions(questions_long_filtered)
  survey_col_pos <- column_positions[[1]]
  header_col_pos <- column_positions[[2]]
  
  # Merge survey cells
  for (survey_i in survey_col_pos$survey) {
    first_col <- survey_col_pos %>%
      filter(survey == survey_i) %>%
      pull(first_col)
    last_col <- survey_col_pos %>%
      filter(survey == survey_i) %>%
      pull(last_col)
    mergeCells(wb, "Survey Input", cols = first_col:last_col, rows = 1)
  }
  
  survey_section_pairs <- header_col_pos %>%
    select(survey, section_header) %>%
    distinct() %>%
    split(seq(nrow(.))) %>%
    lapply(as.list)
  
  # Merge section header cells
  for (pair_i in survey_section_pairs) {
    first_col <- header_col_pos %>%
      filter(
        survey == pair_i[[1]],
        section_header == pair_i[[2]]
      ) %>%
      pull(first_col)
    last_col <- header_col_pos %>%
      filter(
        survey == pair_i[[1]],
        section_header == pair_i[[2]]
      ) %>%
      pull(last_col)
    mergeCells(wb, "Survey Input", cols = first_col:last_col, rows = 2)
  }
  
  return(wb)
}

add_styling <- function(
    wb,
    questions_long_filtered
) {
  # Add styling
  string_lengths <- questions_long_filtered %>%
    select(user_column_name) %>%
    mutate(
      str_len = nchar(user_column_name)
    ) %>% 
    pull(str_len)
  
  widths <- case_when(
    string_lengths > 500 ~ 120,
    string_lengths > 300 ~ 90,
    string_lengths > 100 ~ 50,
    string_lengths > 50 ~ 20,
    TRUE ~ 15
  )
  
  num_cols <- nrow(questions_long_filtered)
  
  # Header rows
  toprows <- createStyle(wrapText = TRUE, textDecoration = "bold",border = "bottom",)
  addStyle(wb, sheet = 1, toprows, rows = 1, cols = 1:num_cols)
  addStyle(wb, sheet = 1, toprows, rows = 2, cols = 1:num_cols)
  addStyle(wb, sheet = 1, toprows, rows = 3, cols = 1:num_cols)
  
  ## Right hand borders
  column_positions <- get_column_positions(questions_long_filtered)
  survey_col_pos <- column_positions[[1]]
  header_col_pos <- column_positions[[2]]
  
  # Section borders
  top_right_hand_borders_sec <- createStyle(border = "BottomRight", wrapText = TRUE, 
                                        textDecoration = "bold")
  other_right_hand_borders_sec <- createStyle(border = "right")
  for (col_i in header_col_pos$last_col) {
    addStyle(wb, sheet = 1, top_right_hand_borders_sec, rows = 2:3, 
             cols = col_i)
    addStyle(wb, sheet = 1, other_right_hand_borders_sec, rows = 4:100, 
             cols = col_i)
  }

  # Survey borders
  top_right_hand_borders_surv <- createStyle(
    border = c("bottom", "right"), 
    wrapText = TRUE, 
    textDecoration = "bold", 
    borderStyle = c("thin","thick")
    )
  
  other_right_hand_borders_surv <- createStyle(border = "right", borderStyle = "thick")
  
  for (col_i in survey_col_pos$last_col) {
    addStyle(wb, sheet = 1, top_right_hand_borders_surv, rows = 1:3, 
             cols = col_i)
    addStyle(wb, sheet = 1, other_right_hand_borders_surv, rows = 4:100, 
             cols = col_i)
  }
  
  
  # Column widths
  setColWidths(wb, sheet = 1, widths = widths, cols = 1:nrow(questions_long))
  
  return(wb)

}


# Create new worksheet to store all question options
store_options <- function(
    wb,
    questions_long_filtered_with_options
) {

  addWorksheet(wb, "Survey Options", visible = FALSE)
  
  for (col_i in 1:nrow(questions_long_filtered_with_options)) {
    
    options <- questions_long_filtered_with_options$options[col_i]
    
    if (grepl(";", options)) {
      allowed <- trimws(stringr::str_split(options, ";")[[1]])
      
      writeData(
        wb, 
        sheet = "Survey Options", 
        x = allowed, 
        startCol = col_i,
        startRow = 1, 
        colNames = FALSE
      )
    }
  }
  
  # Lock Survey Options Sheet
  protectWorksheet(
    wb,
    "Survey Options",
    protect = TRUE
  )
  
  return(wb)
}

# Add in cell validation 
# (drop down options for all questions that aren't free text)
add_validation <- function(
    wb,
    questions_long_filtered_with_options
) {

  # Store all options in separate sheet
  wb <- store_options(wb, questions_long_filtered_with_options)

  # Loop over all questions
  for (col_i in 1:nrow(questions_long_filtered_with_options)) {
    
    # Get question options string
    options <- questions_long_filtered_with_options$options[col_i]
    
    # Check if there are multiple options for the question
    if (grepl(";", options)) {
      
      # Get column letter e.g. 34 -> AH
      col_letter <- int2col(col_i)
      
      # Get number of options for this question
      num_options <- length(stringr::str_split(options, ";")[[1]])
      
      # Get start cell code e.g. AH1
      start_cell <- paste0(col_letter, "1")
      
      # Get end cell code e.g. AH8
      end_cell <- paste0(col_letter, num_options)
      
      # Combine start and end cells into validation formula
      validation_formula <- paste0(
        "\'Survey Options\'!",
        start_cell, ":", end_cell
      )
      
      # Implement validation
      dataValidation(
        wb = wb,
        sheet = "Survey Input",
        cols = col_i,
        rows = 4:10000,
        type = "list",
        value = validation_formula,
        showErrorMsg = TRUE
      )
    }
  }
  
  return(wb)
}

impute_data <- function(
  wb,
  questions_long_filtered_with_options,
  project_name,
  project_id
) {

  imputed_options <- questions_long_filtered_with_options %>%
    mutate(
      col_num = row_number()
    ) %>%
    filter(
      options == "Imputed"
    )
  
  project_code <- read_excel("data/projects-table.xlsx", 
                              sheet = "select_project") %>%
    filter(
      select_project == project_name
    ) %>%
    pull(project_RC_code)
  
  for (col_name in imputed_options$FFF_column_name) {
  
    column = imputed_options$col_num[imputed_options$FFF_column_name == col_name]

    if (col_name == "record_id") {
      wb <- add_record_ids(
        wb,
        project_id,
        column
      ) 
      
    } else if (col_name %in% c("select_project", "grant_application_name", 
                               "funding_stream", "project_lead_email_address")) {
      
      impute_item <- read_excel("data/projects-table.xlsx", 
                                 sheet = col_name) %>%
        filter(
          project_RC_code == project_code
        ) %>%
          pull(sym(col_name))
      
      wb <- impute_item(
        wb,
        impute_item,
        column
      ) 
    } else if (col_name == "how_surveys_carried_out") {
      wb <- impute_item(
        wb,
        "Remotely via Electronic Survey",
        column
      ) 
    }
  }

  # writeData(
  #   wb, 
  #   sheet = "Survey Options", 
  #   x = allowed, 
  #   startCol = col_i,
  #   startRow = 1, 
  #   colNames = FALSE
  # )
  return(wb)
}

create_template <- function(
    questions_long,
    project_table,
    project_name
) {
  
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
    ) %>%
    select(
      survey, section_header, user_column_name
    )
  
  questions_long_filtered_with_options <- questions_long %>%
    filter(
      survey %in% survey_filter[[1]]
    )

  # Convert to wide format
  questions_wide <- convert_to_wide(
    questions_long_filtered
  )
  
  wb <- create_worksheet(
    questions_wide,
    questions_long_filtered
  )
  
  wb <- add_styling(
    wb,
    questions_long_filtered
  )
  
  wb <- add_validation(
    wb,
    questions_long_filtered_with_options
  )
  
  wb <- impute_data(
    wb,
    questions_long_filtered_with_options,
    project_name,
    project_id
  )
  
  save_name <- paste0("output/FFF-template-", project_name, ".xlsx")
  
  saveWorkbook(wb, save_name, overwrite = TRUE)
  return(wb)
}


create_all_templates <- function(
    questions_long,
    project_table
) {
  
  # Validate the questions
  question_options_check(questions_long)
  
  # # Select columns needed
  # questions_long <- questions_long %>%
  #   select(
  #     survey, section_header, user_column_name, options
  #   )
  
  project_names <- unique(project_table$project)
  
  for (project_name_i in project_names) {
    wb <- create_template(
      questions_long,
      project_table,
      project_name_i
    )
  }
  
}
