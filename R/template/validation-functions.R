# Validation functions

expected_surveys <- c(
  "add_participants", "participant_information_sheet_pis", "consent_form", 
  "demographics", "warwick_edinburgh_mental_wellbeing_scale",
  "short_warwick_edinburgh_mental_wellbeing_scale", "breastfeeding_survey", 
  "training_survey_tool", "blood_pressure_questionnaire", 
  "health_literacy_tool_kit", "healthy_start_tool_kit", 
  "hiv_and_hepatitis_b_and_c_prevention", "immunisation_for_adults",
  "immunisation_for_children", "longacting_reversible_contraception_tool_kit",
  "nhs_health_checks_tool_kit", "nutrition_toolkit", "nutrition_bmi",
  "nutrition_food_literacy", "nutrition_food_insecurity",
  "physical_activity_tool_kit", "smoking_cessation", 
  "social_isolation_checklist", "death_literacy_index", 
  "lubben_social_network_scale_revised_lsnsr", 
  "lubben_social_network_scale_6_lsns_6",
  "asthma_and_home_safety_survey",
  "health_promotion_survey",
  "health_literacy_deaf_health_champions",
  "thank_you_message")

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
      num_options = stringr::str_count(options, ";") + 1,
      num_codes   = stringr::str_count(RC_option_code, ";") + 1,
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

check_surveys_in_lookup <- function(
    questions_long
) {
  # Check that all expected surveys are included in the lookup
  unique_surveys <- unique(questions_long$survey)
  passed1 <- all(unique_surveys %in% expected_surveys)
  passed2 <- all(expected_surveys %in% unique_surveys)
  
  if (!passed1) {
    cat("1 or more unexpected survey in lookup: ")
    missing <- unique_surveys[!(unique_surveys %in% expected_surveys)]
    cat(paste(missing, sep = ", "), "\n")
  }
  
  if (!passed2) {
    cat("1 or more survey missing from lookup: ")
    missing <- expected_surveys[!(expected_surveys %in% unique_surveys)]
    cat(paste(missing, sep = ", "), "\n")
  }

  return(passed1 & passed2)
}

check_RedCap_column_names <- function(
    questions_long
) {
  # Check that all RedCap column names (FFF_column_name) are unique
  passed <- length(questions_long$FFF_column_name) == length(unique(questions_long$FFF_column_name))

  return(passed)
  
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
    "DN", # Don't want to answer
    ""   # Empty
  )
  
  unique_codes <- questions_long %>%
    filter(RC_option_code != "Imputed") %>%
    pull(RC_option_code) %>%
    stringr::str_split(";") %>%
    unlist() %>%
    trimws() %>%
    unique()
  
  passed <- all( 
    (grepl("^[0-9]+$", unique_codes) | 
       unique_codes %in% allowed_special) 
  )
  
  if (!passed) {
    cat("Unexpected codes found: \"")
    bad_codes <- unique_codes[!(grepl("^[0-9]+$", unique_codes) | unique_codes %in% allowed_special)]
    cat(paste(bad_codes), "\"\n")
  }

  return(passed)
}


question_options_check <- function(
    questions_long
) {
  if(!check_RedCap_column_names(questions_long)) {
    stop("2 or more elements in FFF_column_name are not unique")
  }
  
  if (! check_option_numbers(questions_long) ) {
    stop("Mismatch between number of questions and RedCap codes")
  }
  
  if (!  check_RedCap_codes(questions_long) ) {
    stop("Invalid RedCap codes found")
  }
  
  if (! check_surveys_in_lookup(questions_long)) {
    stop("Survey mismatch")
  }
  
}
