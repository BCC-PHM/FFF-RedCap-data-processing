library(readxl)
library(dplyr)
library(openxlsx)
library(tidyr)
library(openxlsx2)

survey_filter <- c("add_participants",
                   "participant_information_sheet_pis",
                   "consent_form",
                   "demographics",
                   "smoking_cessation")

# Load the data
questions_long <- read_excel("data/fff-question-lookup.xlsx") %>%
  select(
    survey, section_header, user_column_name
  )

# Filter for surveys required by project
questions_long_filtered <- questions_long %>%
  filter(
    survey %in% survey_filter
  )

# Get start and end column position for each survey (for cell merging)
survey_col_pos <- questions_long_filtered %>%
  mutate(row_number = row_number()) %>%
  group_by(survey) %>%
  summarise(
    first_col = min(row_number),
    last_col = max(row_number)
  ) %>%
  arrange(first_col)

# Get start and end column position for each section header (for cell merging)
header_col_pos <- questions_long_filtered %>%
  mutate(row_number = row_number()) %>%
  group_by(survey, section_header) %>%
  summarise(
    first_col = min(row_number),
    last_col = max(row_number)
  ) %>%
  arrange(first_col)

# Transpose template table
questions_wide <- as.data.frame(t(questions_long_filtered), stringsAsFactors = FALSE)
#questions_wide$rownames <- rownames(questions_wide)
rownames(questions_wide) <- NULL

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

# Add styling
string_lengths <- questions_long %>%
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
saveWorkbook(wb, "output/check.xlsx", overwrite = TRUE)