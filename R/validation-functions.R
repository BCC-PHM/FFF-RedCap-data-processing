# Validation functions

check_option_numbers <- function(
    questions_long  
) {
  # Check if the number of question options matches the number of RedCap codes
  # for that question
  num_check_df <- questions_long %>%
    mutate(
      row = row_number()
    ) %>%
    filter(options != "Imputed") %>%
    mutate(
      num_options = stringr::str_count(options, "; ") + 1,
      num_codes   = stringr::str_count(RCS_option_code, "; ") + 1,
      match       = num_options == num_codes
    ) 
  
  check <- all(num_check_df$match)
  
  if (!check) {
    # Print non-matching rows
    cat("Question found with mismatched options and codes:\n")
    print(
      num_check_df %>%
        filter(!match) %>%
        select(row, options, RC_option_code)
    )
  }
  
  return(check)
}

check_RedCap_codes <- function(
    questions_long
) {
  # Check that all RedCap codes are either a number or in the allowed specical
  # characters
  allowed_special <- c(
    "O",  # Other
    "NA", # Not applicable
    "DK", # Don't know
    "DN"  # Don't want to answer
  )
  
  unique_codes <- questions_long %>%
    filter(RC_option_code != "Imputed") %>%
    pull(RC_option_code) %>%
    stringr::str_split("; ") %>%
    unlist() %>%
    unique()
  
  check <- all( 
    (grepl("^[0-9]+$", unique_codes) | 
       unique_codes %in% allowed_special) 
  )
  
  bad_codes <- unique_codes[!(grepl("^[0-9]+$", unique_codes) | unique_codes %in% allowed_special)]
  
  return(check)
}