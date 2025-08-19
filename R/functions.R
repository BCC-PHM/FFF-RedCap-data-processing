convert_to_wide <- function(
    questions_long_filtered
) {
  
  # Transpose template table
  questions_wide <- as.data.frame(t(questions_long_filtered), stringsAsFactors = FALSE)
  rownames(questions_wide) <- NULL
  
  return(questions_wide)
}

create_worksheet <- function(
    questions_wide,
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
  
  ## Create a new workbook
  wb <- createWorkbook()
  ## Add a worksheet
  addWorksheet(wb, "Survey Input")
  ## Add wide questions to worksheet
  writeData(wb, sheet = "Survey Input", x = questions_wide, colNames = FALSE)
  
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
    ws,
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
  
  wrapStyle <- createStyle(wrapText = TRUE)
  addStyle(wb, sheet = 1, wrapStyle, rows = 2, cols = 1:1000)
  addStyle(wb, sheet = 1, wrapStyle, rows = 3, cols = 1:1000)
  setColWidths(wb, sheet = 1, widths = widths, cols = 1:nrow(questions_long))
  
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
  
  # Filter for surveys required by project
  questions_long_filtered <- questions_long %>%
    filter(
      survey %in% survey_filter[[1]]
    )
  
  # Convert to wide format
  questions_wide <- convert_to_wide(
    questions_long_filtered
  )
  
  ws <- create_worksheet(
    questions_wide,
    questions_long_filtered
  )
  
  ws <- add_styling(
    ws,
    questions_long_filtered
    )
  
  save_name <- paste0("output/FFF-template-", project_name, ".xlsx")
  
  saveWorkbook(wb, save_name, overwrite = TRUE)
  return(ws)
}

create_all_templates <- function(
    questions_long,
    project_table
) {
  project_names <- unique(project_table$project)
  
  for (project_name_i in project_names) {
    create_template(
      questions_long,
      project_table,
      project_name_i
    )
  }
}