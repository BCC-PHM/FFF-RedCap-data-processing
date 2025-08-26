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

impute_item <- function(
    wb,
    impute_item,
    column
) {

  impute_list <-  rep(impute_item, 997)  # e.g. TEST - 0001 ... TEST - 0097
  
  writeData(
    wb, 
    sheet = "Survey Input", 
    x = impute_list, 
    startCol = column, 
    startRow = 4, 
    colNames = FALSE
  )
  
  return(wb)
}