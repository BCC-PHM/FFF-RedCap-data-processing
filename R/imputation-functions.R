# Imputation functions


# Record ids
add_record_ids <- function(
    wb,
    project_id,
    column
) {
  ids <- sprintf("%s - %04d", project_id, 1:997)  # e.g. TEST - 0001 ... TEST - 0097
  
  writeData(
    wb, 
    sheet = "Survey Input", 
    x = ids, 
    startCol = column, 
    startRow = 4, 
    colNames = FALSE
  )
  
  return(wb)
}

add_project_name <- function(
    wb,
    project_name,
    column
) {
  project_names <-  rep(project_name, 997)  # e.g. TEST - 0001 ... TEST - 0097
  
  writeData(
    wb, 
    sheet = "Survey Input", 
    x = project_names, 
    startCol = column, 
    startRow = 4, 
    colNames = FALSE
  )
  
  return(wb)
}